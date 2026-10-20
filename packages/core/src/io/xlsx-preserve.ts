import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { formatAddress, formatRange, parseAddress, parseRange, type RangeAddress } from '../address.js';
import { tokenize } from '../formula/tokenizer.js';
import { parseRefToken } from '../formula/rewrite.js';
import { addFutureFunctionPrefixes } from '../formula/rewrite.js';
import { parseFormula } from '../formula/parser.js';
import { hasUnknownFunctions } from '../formula/scope.js';
import type { WorkbookData } from '../ops.js';
import type { CellData } from '../sheet.js';
import type { Workbook } from '../workbook.js';
import { tableNameAt } from '../table.js';
import type { SheetData } from '../sheet.js';
import { validationsXml } from './xlsx-validations.js';
import { arr, attr, encodeOoxmlEscapes, escapeXml, parseXml, resolvePath, relsPathFor, text } from './xml.js';
import { PreservedStyles } from './xlsx-preserve-styles.js';
import { patchLayout } from './xlsx-preserve-layout.js';
import { attribute, canonical, sameValue, child, contents, editXml, nodeAttribute, prefix, protect, raw, setAttribute, setChild, xmlRoot, type XmlEdit, type XmlSpan } from './xlsx-preserve-xml.js';

export function encodeXlsxArchive(bytes: Uint8Array): string {
  const chunks: string[] = [];
  for (let i = 0; i < bytes.length; i += 32768) chunks.push(String.fromCharCode(...bytes.subarray(i, i + 32768)));
  return btoa(chunks.join(''));
}
let decoded: { value: string; bytes: Uint8Array } | undefined;
function decodeArchive(value: string): Uint8Array {
  // Saving repeatedly decodes the same archive; keep the last result.
  if (decoded?.value === value) return decoded.bytes;
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  decoded = { value, bytes };
  return bytes;
}
function payload(cell?: CellData): unknown {
  return cell?.f !== undefined ? { f: cell.f, a: cell.a } : { v: cell?.v ?? null, a: cell?.a };
}
const EMPTY_CELL = { v: null };
function sameCellContent(a: CellData | undefined, b: CellData | undefined): boolean {
  if (a === b) return true;
  const x = a ? { ...payload(a) as object, s: a.s } : EMPTY_CELL, y = b ? { ...payload(b) as object, s: b.s } : EMPTY_CELL;
  return sameValue(x, y);
}

/** Whether two workbooks have the same content apart from the archive and active sheet. */
function sameModel(a: WorkbookData, b: WorkbookData): boolean {
  const strip = ({ xlsxSource: _source, activeSheet: _active, sheets: _sheets, ...rest }: WorkbookData) => rest;
  if (!sameValue(strip(a), strip(b)) || a.sheets.length !== b.sheets.length) return false;
  return a.sheets.every((sa, i) => sameSheetContent(sa, b.sheets[i]!));
}

/** Whether two sheets have the same content; computed formula results are not content. */
function sameSheetContent(sa: SheetData, sb: SheetData): boolean {
  const { cells: ca, ...ra } = sa, { cells: cb, ...rb } = sb;
  if (!sameValue(ra, rb)) return false;
  for (const key in ca) if (!sameCellContent(ca[key], cb[key])) return false;
  for (const key in cb) if (!(key in ca) && !sameCellContent(undefined, cb[key])) return false;
  return true;
}

function sheetMetadata({ cells: _cells, colWidths: _cols, rowHeights: _rows, validations: _validations, ...sheet }: SheetData): unknown {
  return sheet;
}
/** Workbook content apart from cells, sizes and validations, which are patched. */
function metadata(data: WorkbookData, sheetCount = data.sheets.length): unknown {
  const { xlsxSource: _source, activeSheet: _active, ...rest } = data;
  return { ...rest, sheets: rest.sheets.slice(0, sheetCount).map(sheetMetadata) };
}
function relations(files: Record<string, Uint8Array>, part: string): { id: string; target: string; type: string }[] {
  const bytes = files[relsPathFor(part)];
  return arr(bytes ? parseXml(strFromU8(bytes))?.Relationships?.Relationship : undefined)
    .filter((r) => attr(r, 'TargetMode') !== 'External')
    .map((r) => ({ id: attr(r, 'Id') ?? '', target: resolvePath(part, attr(r, 'Target') ?? ''), type: attr(r, 'Type') ?? '' }));
}
/** Shared string items (<si>) that contain more than plain text (<r>, <rPh>, <phoneticPr>). */
export function richStringIndices(xml: string): Set<number> {
  const out = new Set<number>();
  const item = /<(?:[\w.-]+:)?si(?:\s[^>]*)?(\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?si\s*>)/g;
  let index = 0;
  for (let m; (m = item.exec(xml)); index++) {
    if (m[2] && /<(?:[\w.-]+:)?(?:r|rPh|phoneticPr)[\s/>]/.test(m[2])) out.add(index);
  }
  return out;
}

function textValue(value: string): string {
  return escapeXml(encodeOoxmlEscapes(value));
}

/**
 * The cached result (<v> and cell type) of a formula as computed by the engine, or
 * undefined when it must not be cached: dynamic arrays, results without a value and
 * formulas with functions the engine does not know (their result would be invented).
 */
function computedCache(cell: CellData | undefined, p: string, names: Set<string>): { type: string | undefined; body: string } | undefined {
  if (cell?.f === undefined || cell.a) return undefined;
  try {
    if (hasUnknownFunctions(parseFormula(cell.f), (name) => names.has(name.toUpperCase()))) return undefined;
  } catch { return undefined; }
  const v = cell.v;
  if (typeof v === 'number') return Number.isFinite(v) ? { type: undefined, body: `<${p}v>${v}</${p}v>` } : undefined;
  if (typeof v === 'string') return { type: 'str', body: `<${p}v>${textValue(v)}</${p}v>` };
  if (typeof v === 'boolean') return { type: 'b', body: `<${p}v>${v ? 1 : 0}</${p}v>` };
  if (v && typeof v === 'object') return { type: 'e', body: `<${p}v>${escapeXml(v.error)}</${p}v>` };
  return undefined;
}

/**
 * Replaces the cached result of an unchanged formula whose value changed through
 * other edits, so programs that read caches without calculating (data imports,
 * openpyxl, pandas) see current values. Cells with structures beyond a plain or
 * shared formula keep their original XML; Excel recalculates them on opening.
 */
function refreshCache(xml: string, node: XmlSpan, cell: CellData | undefined, p: string, names: Set<string>): string {
  const result = raw(xml, node);
  const root = xmlRoot(result), open = result.slice(0, root.openEnd);
  if (attribute(open, 'cm') !== undefined || attribute(open, 'vm') !== undefined) return result;
  if (!['n', 'b', 'str', 'e'].includes(attribute(open, 't') ?? 'n') || root.children.some((c) => !['f', 'v'].includes(c.name))) return result;
  const f = child(root, 'f');
  const formulaType = f ? attribute(result.slice(f.start, f.openEnd), 't') : undefined;
  if (!f || (formulaType !== undefined && formulaType !== 'normal' && formulaType !== 'shared')) return result;
  const cache = computedCache(cell, p, names);
  if (!cache) return result;
  const v = child(root, 'v');
  return editXml(result, [
    { start: 0, end: root.openEnd, value: setAttribute(open, 't', cache.type) },
    ...(v ? [{ start: v.start, end: v.end, value: '' }] : []),
    { start: f.end, end: f.end, value: cache.body },
  ]);
}

const AFTER_VALIDATIONS = ['hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** Replaces the sheet's <dataValidations>; Excel 2010 rules in extLst cannot be merged safely. */
function patchValidations(xml: string, before: SheetData, after: SheetData): string {
  const rules = after.validations ?? [];
  if (sameValue(before.validations ?? [], rules)) return xml;
  const root = xmlRoot(xml);
  const ext = child(root, 'extLst');
  if (ext && /dataValidations/.test(raw(xml, ext))) protect(`Erweiterte Datenüberprüfung (Excel 2010) in ${before.name}.`, `Extended data validation (Excel 2010) in ${before.name}.`);
  return setChild(xml, 'dataValidations', rules.length ? validationsXml(rules, prefix(root)) : '', AFTER_VALIDATIONS);
}

function updateCell(xml: string, node: XmlSpan | undefined, ref: string, before: CellData | undefined, after: CellData | undefined, p: string, styles: () => PreservedStyles, inheritedStyle: number, richStrings: { has(index: number): boolean }, names: Set<string>, table?: string, styleIndex?: number): string {
  const valueChanged = !sameValue(payload(before), payload(after));
  let result = node ? raw(xml, node) : `<${p}c r="${ref}"/>`;
  let root = xmlRoot(result), open = result.slice(0, root.openEnd);
  if (valueChanged) {
    if (attribute(open, 'cm') !== undefined || attribute(open, 'vm') !== undefined) protect(`Zellmetadaten in ${ref}.`, `Cell metadata in ${ref}.`);
    if (!['n', 'b', 's', 'str', 'inlineStr', 'e'].includes(attribute(open, 't') ?? 'n')) protect(`Nicht unterstützter Zelltyp in ${ref}.`, `Unsupported cell type in ${ref}.`);
    if (root.children.some((c) => !['f', 'v', 'is'].includes(c.name))) protect(`Erweiterte Zellinhalte in ${ref}.`, `Extended cell content in ${ref}.`);
    const inline = child(root, 'is');
    const sharedIndex = attribute(open, 't') === 's' ? Number(text(parseXml(result)?.c?.v)) : -1;
    if (inline?.children.some((c) => c.name !== 't') || richStrings.has(sharedIndex)) protect(`Rich-Text oder phonetische Annotationen in ${ref}.`, `Rich text or phonetic annotations in ${ref}.`);
    const f = child(root, 'f');
    if (f && attribute(result.slice(f.start, f.openEnd), 't') && attribute(result.slice(f.start, f.openEnd), 't') !== 'normal') protect(`Geteilte Formel, Matrixformel oder Datentabelle in ${ref}.`, `Shared formula, array formula or data table in ${ref}.`);
    if (before?.a || after?.a) protect(`Dynamische Matrix in ${ref}.`, `Dynamic array in ${ref}.`);
    let type: string | undefined, body = '';
    if (after?.f !== undefined) {
      body = `<${p}f>${escapeXml(addFutureFunctionPrefixes(after.f, table ? { table } : {}))}</${p}f>`;
      // Cache the computed result only for formulas the engine fully understands, never an invented one.
      const cache = computedCache(after, p, names);
      if (cache) { body += cache.body; type = cache.type; }
    } else if (typeof after?.v === 'string') {
      type = 'inlineStr'; body = `<${p}is><${p}t xml:space="preserve">${textValue(after.v)}</${p}t></${p}is>`;
    } else if (typeof after?.v === 'number') {
      if (!Number.isFinite(after.v)) protect(`Nicht endlicher Zahlenwert in ${ref}.`, `Non-finite number in ${ref}.`);
      body = `<${p}v>${after.v}</${p}v>`;
    } else if (typeof after?.v === 'boolean') { type = 'b'; body = `<${p}v>${after.v ? 1 : 0}</${p}v>`; }
    else if (after?.v && typeof after.v === 'object') { type = 'e'; body = `<${p}v>${escapeXml(after.v.error)}</${p}v>`; }
    const remaining = editXml(result, root.children.filter((n) => ['f', 'v', 'is'].includes(n.name)).map((n) => ({ start: n.start, end: n.end, value: '' })));
    root = xmlRoot(remaining);
    open = setAttribute(remaining.slice(0, root.openEnd), 't', type);
    result = contents(remaining, root, body + remaining.slice(root.openEnd, root.closeStart), open);
  }
  if (styleIndex !== undefined) {
    // Sorted content keeps the exact original format record of the cell it came from.
    root = xmlRoot(result); open = result.slice(0, root.openEnd);
    if (Number(attribute(open, 's') ?? inheritedStyle) !== styleIndex) result = editXml(result, [{ start: 0, end: root.openEnd, value: setAttribute(open, 's', styleIndex) }]);
  } else if (!sameValue(before?.s, after?.s)) {
    root = xmlRoot(result); open = result.slice(0, root.openEnd);
    const index = styles().get(Number(attribute(open, 's') ?? inheritedStyle), before?.s, after?.s);
    result = editXml(result, [{ start: 0, end: root.openEnd, value: setAttribute(open, 's', index) }]);
  }
  return result;
}

/**
 * Ranges of array formulas and data tables in <sheetData>, found by scanning
 * the text so cells of unchanged rows need not be parsed.
 */
function arrayRanges(xml: string, data: XmlSpan, p: string): ReturnType<typeof parseRange>[] {
  const out: ReturnType<typeof parseRange>[] = [];
  const body = xml.slice(data.openEnd, data.closeStart);
  const formula = new RegExp(`^<${p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}f[\\s/>]`);
  const type = /\st\s*=\s*(["'])(?:array|dataTable)\1/g;
  for (let m; (m = type.exec(body));) {
    const lt = body.lastIndexOf('<', m.index);
    const tag = body.slice(lt, body.indexOf('>', m.index) + 1);
    if (!formula.test(tag)) continue;
    const ref = attribute(tag, 'ref');
    if (ref) { out.push(parseRange(ref)); continue; }
    // Without a ref the formula covers its own cell.
    const cellStart = Math.max(body.lastIndexOf(`<${p}c `, lt), body.lastIndexOf(`<${p}c>`, lt));
    const address = cellStart >= 0 ? attribute(body.slice(cellStart, body.indexOf('>', cellStart) + 1), 'r') : undefined;
    if (!address) protect('Matrixformel ohne Zelladresse.', 'Array formula without a cell address.');
    out.push(parseRange(address));
  }
  return out;
}

interface SheetContext {
  /** Editable style table; created on first use when the file has none. */
  styles: () => PreservedStyles;
  /** Existing style table, without creating one. */
  existingStyles: () => PreservedStyles | undefined;
  date1904: boolean;
  hasCalcChain: boolean;
  richStrings: { has(index: number): boolean };
  lockedRanges: string[];
  /** The sheet was sorted since import: moved content needs additional checks. */
  sorted: boolean;
  /** Upper-case defined names (for recognizing formulas). */
  names: Set<string>;
}

function within(a: { row: number; col: number }, ranges: (RangeAddress | undefined)[]): boolean {
  return ranges.some((r) => r && a.row >= r.start.row && a.row <= r.end.row && a.col >= r.start.col && a.col <= r.end.col);
}

/** Style index a cell node uses, falling back to its row or column style. */
function inheritedStyle(xml: string, root: XmlSpan, row: XmlSpan | undefined, col: number): number {
  const columnStyle = child(root, 'cols')?.children.filter((n) => n.name === 'col').find((n) => col + 1 >= Number(nodeAttribute(xml, n, 'min')) && col + 1 <= Number(nodeAttribute(xml, n, 'max')));
  return Number((row && nodeAttribute(xml, row, 's')) ?? (columnStyle && nodeAttribute(xml, columnStyle, 'style')) ?? 0);
}

/** Whether a numeric <v> text is exactly the shortest or a correctly rounded rendering of its double. */
export function exactNumberText(value: string): boolean {
  const n = Number(value);
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(value) || !Number.isFinite(n)) return false;
  const digits = value.replace(/e.*$/i, '').replace(/\D/g, '').replace(/^0+/, '').length;
  return sameDecimal(value, String(n)) || sameDecimal(value, n.toPrecision(Math.min(100, Math.max(1, digits))));
}
function sameDecimal(a: string, b: string): boolean {
  const norm = (text: string) => {
    const m = /^([+-]?)(\d*)(?:\.(\d*))?(?:e([+-]?\d+))?$/i.exec(text);
    if (!m) return text;
    let digits = (m[2] ?? '') + (m[3] ?? ''), point = (m[2] ?? '').length + Number(m[4] ?? 0);
    const lead = /^0*/.exec(digits)![0].length;
    digits = digits.slice(lead).replace(/0+$/, ''); point -= lead;
    return digits ? `${m[1] === '-' ? '-' : ''}${digits}e${point}` : '0';
  };
  return norm(a) === norm(b);
}

/**
 * Safeguards for sorted sheets. Sorting moves values, formulas and formats between
 * cells, but only the model is known, not which original cell XML moved where. The
 * save is therefore allowed only where the model determines the result exactly:
 * no notes, links, native tables, validations or merges in changed cells, row-local
 * formulas with known functions, numbers that keep their exact text, and formats
 * where each displayed format maps to one original format record. The returned map
 * gives that record for each displayed format.
 */
function sortSafeguards(xml: string, root: XmlSpan, data: XmlSpan, p: string, before: SheetData, after: SheetData, changes: string[], rows: Map<number, XmlSpan>, cells: Map<string, XmlSpan>, ctx: SheetContext): Map<string, number> {
  const ranges = (list: string[]) => list.map((r) => parseRange(r)?.range);
  const tables = ranges((before.tables ?? []).map((t) => t.range));
  const validations = ranges((before.validations ?? []).flatMap((v) => v.range.trim().split(/\s+/)));
  const merges = ranges(before.merges ?? []);
  for (const ref of changes) {
    const a = parseAddress(ref)!;
    if (before.notes?.[ref] || before.links?.[ref] || after.notes?.[ref] || after.links?.[ref]) protect(`Sortieren von Zellen mit Notizen oder Hyperlinks bei ${ref}.`, `Sorting cells with notes or hyperlinks at ${ref}.`);
    if (within(a, tables)) protect(`Sortieren in einer nativen Excel-Tabelle bei ${ref}.`, `Sorting inside a native Excel table at ${ref}.`);
    if (within(a, validations)) protect(`Sortieren von Zellen mit Datenüberprüfung bei ${ref}.`, `Sorting cells with data validation at ${ref}.`);
    if (within(a, merges)) protect(`Sortieren verbundener Zellen bei ${ref}.`, `Sorting merged cells at ${ref}.`);
    for (const formula of [before.cells[ref]?.f, after.cells[ref]?.f]) {
      if (formula === undefined) continue;
      let local = true;
      try {
        if (hasUnknownFunctions(parseFormula(formula), (name) => ctx.names.has(name.toUpperCase()))) local = false;
        for (const token of tokenize(formula)) {
          if (token.type !== 'ref') continue;
          const r = parseRefToken(token.text);
          if (r.kind !== 'cols' && [r.start, r.end].some((point) => !point.absRow && point.row !== a.row)) local = false;
        }
      } catch { local = false; }
      if (!local) protect(`Sortierte Formel mit Bezügen auf andere Zeilen oder unbekannten Funktionen bei ${ref}.`, `Sorted formula with references to other rows or unknown functions at ${ref}.`);
    }
  }
  // Moved numbers are written from the model; their original text must not carry more digits.
  const body = xml.slice(data.openEnd, data.closeStart), q = p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const cell = new RegExp(`<${q}c(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</${q}c>)`, 'g'), value = new RegExp(`<${q}v(?:\\s[^>]*)?>([^<]*)</${q}v>`);
  for (let m; (m = cell.exec(body));) {
    if ((attribute(` ${m[1]}`, 't') ?? 'n') !== 'n' || !m[2]) continue;
    const v = value.exec(m[2])?.[1];
    if (v !== undefined && v.trim() !== '' && !exactNumberText(v.trim())) protect('Sortieren von Zahlen, deren gespeicherte Genauigkeit über eine Gleitkommazahl hinausgeht.', 'Sorting numbers stored with more precision than a floating-point number.');
  }
  // Formats: each displayed format must come from exactly one original format record.
  const styles = ctx.existingStyles();
  const fingerprints = new Map<number, string>();
  const formats = new Map<string, { index: number; fingerprint: string }>();
  for (const ref of changes) {
    const a = parseAddress(ref)!;
    const node = cells.get(ref);
    const explicit = node ? nodeAttribute(xml, node, 's') : undefined;
    const index = explicit !== undefined ? Number(explicit) : inheritedStyle(xml, root, rows.get(a.row), a.col);
    let fingerprint = fingerprints.get(index);
    if (fingerprint === undefined) fingerprints.set(index, fingerprint = styles ? styles.fingerprint(index) : `#${index}`);
    const key = canonical(before.cells[ref]?.s ?? null);
    const known = formats.get(key);
    if (known && known.fingerprint !== fingerprint) protect(`Sortieren gleich dargestellter Zellen mit unterschiedlichen Excel-Formaten bei ${ref}.`, `Sorting cells that look alike but have different Excel formats at ${ref}.`);
    if (!known) formats.set(key, { index, fingerprint });
  }
  return new Map([...formats].map(([key, { index }]) => [key, index]));
}

function patchSheet(xml: string, root: XmlSpan, before: WorkbookData['sheets'][number], after: WorkbookData['sheets'][number], ctx: SheetContext): string {
  const { date1904, hasCalcChain, richStrings, lockedRanges } = ctx;
  const changes = [...new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])]
    .filter((key) => !sameValue([payload(before.cells[key]), before.cells[key]?.s], [payload(after.cells[key]), after.cells[key]?.s]));
  // Unchanged formulas whose result changed through edits get a current cache. This
  // never blocks a save: cells that cannot be refreshed safely keep their XML.
  const changed = new Set(changes);
  const refreshes = date1904 || child(root, 'sheetProtection') ? [] : Object.keys(after.cells).filter((key) => {
    const a = after.cells[key], b = before.cells[key];
    return !changed.has(key) && a?.f !== undefined && b?.f === a.f && !a.a && !b.a && !sameValue(a.v ?? null, b.v ?? null);
  });
  if (!changes.length && !refreshes.length) return xml;
  const p = prefix(root);
  if (root.name !== 'worksheet') protect('Nicht unterstützter Blatttyp.', 'Unsupported sheet type.');
  if (child(root, 'sheetProtection')) protect(`Blattschutz in ${before.name}.`, `Sheet protection in ${before.name}.`);
  const data = child(root, 'sheetData'); if (!data) protect('sheetData fehlt.', 'sheetData is missing.');
  const rows = new Map<number, XmlSpan>(), cells = new Map<string, XmlSpan>();
  let rowIndex = -1;
  for (const row of data.children) {
    if (row.name !== 'row') continue;
    if (row.qname !== p + 'row') protect('Nicht unterstützter Zeilen-Namensraum.', 'Unsupported row namespace.');
    rowIndex = Number(nodeAttribute(xml, row, 'r') ?? rowIndex + 2) - 1;
    if (rows.has(rowIndex)) protect('Doppelte Zeilenadresse.', 'Duplicate row address.');
    rows.set(rowIndex, row);
  }
  const arrays = arrayRanges(xml, data, p);
  const changedRows = new Map<number, string[]>();
  for (const ref of changes) {
    const a = parseAddress(ref); if (!a) protect('Ungültige Zelladresse im Modell.', 'Invalid cell address in the model.');
    const valueChanged = !sameValue(payload(before.cells[ref]), payload(after.cells[ref]));
    if (valueChanged) {
      if (lockedRanges.some((ref) => {
        const r = parseRange(ref)?.range;
        return r && a.row >= r.start.row && a.row <= r.end.row && a.col >= r.start.col && a.col <= r.end.col;
      })) protect(`Berechnete Tabellenspalte, Summenzeile oder Pivot-Ausgabe bei ${ref}.`, `Calculated table column, totals row or pivot output at ${ref}.`);
      if (date1904) protect('Die Datei verwendet das noch nicht vollständig unterstützte Datumssystem 1904.', 'The file uses the 1904 date system, which is not fully supported yet.');
      if (arrays.some((r) => r && a.row >= r.range.start.row && a.row <= r.range.end.row && a.col >= r.range.start.col && a.col <= r.range.end.col)) protect(`Matrixformel oder Datentabelle bei ${ref}.`, `Array formula or data table at ${ref}.`);
      if (hasCalcChain && before.cells[ref]?.f !== after.cells[ref]?.f) protect('Formeländerung mit vorhandener Berechnungskette.', 'Formula change in a file with a calculation chain.');
      for (const table of before.tables ?? []) {
        const r = parseRange(table.range)!.range;
        if (table.headerRow && a.row === r.start.row && a.col >= r.start.col && a.col <= r.end.col) protect(`Änderung eines nativen Tabellenkopfs bei ${ref}.`, `Change of a native table header at ${ref}.`);
      }
    }
    const list = changedRows.get(a.row) ?? []; list.push(ref); changedRows.set(a.row, list);
  }
  for (const ref of refreshes) {
    const a = parseAddress(ref);
    if (!a || !rows.has(a.row) || within(a, arrays.map((r) => r?.range))) continue;
    const list = changedRows.get(a.row) ?? []; list.push(ref); changedRows.set(a.row, list);
  }
  // Cells are resolved only in rows that change; other rows are copied unchanged.
  for (const index of changedRows.keys()) {
    const row = rows.get(index);
    if (!row) continue;
    const rowXml = raw(xml, row);
    row.children = xmlRoot(rowXml, 2).children.map((c) => ({ ...c, start: c.start + row.start, openEnd: c.openEnd + row.start, closeStart: c.closeStart + row.start, end: c.end + row.start }));
    let col = -1;
    for (const c of row.children.filter((n) => n.name === 'c')) {
      if (c.qname !== p + 'c') protect('Nicht unterstützter Zell-Namensraum.', 'Unsupported cell namespace.');
      const address = nodeAttribute(xml, c, 'r');
      const a = address ? parseAddress(address) : { row: index, col: col + 1 };
      if (!a || a.row !== index) protect('Ungültige Zelladresse.', 'Invalid cell address.');
      col = a.col;
      const ref = formatAddress(a.row, a.col);
      if (cells.has(ref)) protect('Doppelte Zelladresse.', 'Duplicate cell address.');
      cells.set(ref, c);
    }
  }
  const sortedFormats = ctx.sorted ? sortSafeguards(xml, root, data, p, before, after, changes, rows, cells, ctx) : undefined;
  const rowEdits: XmlEdit[] = [], insertedRows = new Map<number, string>();
  for (const [index, refs] of [...changedRows].sort(([a], [b]) => a - b)) {
    const row = rows.get(index);
    if (!row && [...rows.values()].some((node) => nodeAttribute(xml, node, 'r') === undefined)) protect('Neue Zeilen bei impliziten Original-Zeilenadressen.', 'New rows in a sheet with implicit row addresses.');
    const base = row ? raw(xml, row) : `<${p}row r="${index + 1}"/>`;
    const edits: XmlEdit[] = [], insertions = new Map<number, string>();
    const offset = row?.start ?? 0;
    for (const ref of refs.sort((a, b) => parseAddress(a)!.col - parseAddress(b)!.col)) {
      const node = cells.get(ref);
      if (!changed.has(ref)) {
        // A cache refresh only edits existing formula cells.
        if (node) edits.push({ start: node.start - offset, end: node.end - offset, value: refreshCache(xml, node, after.cells[ref], p, ctx.names) });
        continue;
      }
      if (!node && row?.children.some((c) => c.name === 'c' && nodeAttribute(xml, c, 'r') === undefined)) protect('Neue Zellen bei impliziten Original-Zelladressen.', 'New cells in a row with implicit cell addresses.');
      const col = parseAddress(ref)!.col;
      const at = parseAddress(ref);
      const table = at ? tableNameAt(after.tables ?? [], at.row, at.col) : undefined;
      const styleIndex = sortedFormats?.get(canonical(after.cells[ref]?.s ?? null));
      const value = updateCell(xml, node, ref, before.cells[ref], after.cells[ref], p, ctx.styles, inheritedStyle(xml, root, row, col), richStrings, ctx.names, table, styleIndex);
      if (node) edits.push({ start: node.start - offset, end: node.end - offset, value });
      else {
        const next = row?.children.find((n) => n.name === 'c' && parseAddress(nodeAttribute(xml, n, 'r') ?? '')?.col! > col);
        const last = row?.children.filter((n) => n.name === 'c').at(-1);
        const position = (next?.start ?? last?.end ?? row?.openEnd ?? xmlRoot(base).openEnd) - offset;
        insertions.set(position, (insertions.get(position) ?? '') + value);
      }
    }
    const node = xmlRoot(base);
    let value: string;
    if (node.closeStart === node.openEnd && /\/\s*>$/.test(base.slice(0, node.openEnd))) value = contents(base, node, [...insertions.values()].join(''));
    else value = editXml(base, [...edits, ...[...insertions].map(([start, value]) => ({ start, end: start, value }))]);
    if (row) rowEdits.push({ start: row.start, end: row.end, value });
    else {
      const next = [...rows].find(([r]) => r > index)?.[1];
      const position = next?.start ?? [...rows.values()].at(-1)?.end ?? data.openEnd;
      insertedRows.set(position, (insertedRows.get(position) ?? '') + value);
    }
  }
  let result: string;
  if (data.closeStart === data.openEnd && /\/\s*>$/.test(xml.slice(data.start, data.openEnd))) result = editXml(xml, [{ start: data.start, end: data.end, value: contents(xml, data, [...insertedRows.values()].join('')) }]);
  else result = editXml(xml, [...rowEdits, ...[...insertedRows].map(([start, value]) => ({ start, end: start, value }))]);
  // <dimension> precedes <sheetData>, so its position in the original is unchanged.
  const dimension = child(root, 'dimension');
  if (dimension && dimension.end <= data.start) {
    const range = parseRange(nodeAttribute(result, dimension, 'ref') ?? 'A1')?.range;
    if (range) {
      for (const ref of changes) {
        const a = parseAddress(ref)!;
        range.start.row = Math.min(range.start.row, a.row); range.start.col = Math.min(range.start.col, a.col);
        range.end.row = Math.max(range.end.row, a.row); range.end.col = Math.max(range.end.col, a.col);
      }
      const open = setAttribute(result.slice(dimension.start, dimension.openEnd), 'ref', formatRange(range));
      result = editXml(result, [{ start: dimension.start, end: dimension.openEnd, value: open }]);
    }
  }
  return result;
}

/** Allow a complete rebuild only when the original contains exactly our serializer's output. */
function equivalentPackage(a: Record<string, Uint8Array>, b: Record<string, Uint8Array>): boolean {
  if (!sameValue(Object.keys(a).sort(), Object.keys(b).sort())) return false;
  return Object.keys(a).every((name) => {
    const aa = a[name]!, bb = b[name]!;
    if (name === 'docProps/core.xml') {
      const neutral = (bytes: Uint8Array) => strFromU8(bytes).replace(/(<dcterms:(?:created|modified)\b[^>]*>)[^<]*(<\/dcterms:(?:created|modified)>)/g, '$1$2');
      return neutral(aa) === neutral(bb);
    }
    return aa.length === bb.length && aa.every((value, index) => value === bb[index]);
  });
}

function understandsFormulas(data: WorkbookData): boolean {
  try {
    let supported = true;
    const names = new Set(Object.keys(data.names ?? {}).map((n) => n.toUpperCase()));
    for (const sheet of data.sheets) for (const cell of Object.values(sheet.cells)) {
      if (cell.f !== undefined && hasUnknownFunctions(parseFormula(cell.f), (name) => names.has(name.toUpperCase()))) supported = false;
    }
    return supported;
  } catch { return false; }
}

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
/** Smallest valid style table; used when a file has none and formats are added. */
const MINIMAL_STYLES = XML_HEAD + `<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>` +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

function hasPart(files: Record<string, Uint8Array>, path: string): boolean {
  // OPC part names are case-insensitive.
  const lower = path.toLowerCase();
  return Object.keys(files).some((name) => name.toLowerCase() === lower);
}

/** Appends a child element to the root of an XML part, keeping everything else. */
function appendToRoot(xml: string, value: string): string {
  const root = xmlRoot(xml);
  if (root.closeStart === root.openEnd && /\/\s*>$/.test(xml.slice(root.start, root.openEnd))) return editXml(xml, [{ start: root.start, end: root.end, value: contents(xml, root, value) }]);
  return editXml(xml, [{ start: root.closeStart, end: root.closeStart, value }]);
}

/** Adds an internal relationship to a part and returns its new, unique id. */
function addRelationship(files: Record<string, Uint8Array>, part: string, type: string, target: string): string {
  const path = relsPathFor(part), bytes = files[path];
  if (!bytes) protect('Die Beziehungen der Arbeitsmappe fehlen.', 'The workbook relationships are missing.');
  const xml = strFromU8(bytes), root = xmlRoot(xml);
  if (root.name !== 'Relationships') protect('Ungültige Beziehungsdatei.', 'Invalid relationships part.');
  const ids = new Set(arr(parseXml(xml)?.Relationships?.Relationship).map((r) => attr(r, 'Id')));
  let id = 'rIdSheets1';
  for (let i = 2; ids.has(id); i++) id = `rIdSheets${i}`;
  files[path] = strToU8(appendToRoot(xml, `<${prefix(root)}Relationship Id="${id}" Type="${type}" Target="${escapeXml(target)}"/>`));
  return id;
}

function addContentType(files: Record<string, Uint8Array>, part: string, type: string): void {
  const bytes = files['[Content_Types].xml'];
  if (!bytes) protect('Die Inhaltstypen des Pakets fehlen.', 'The package content types are missing.');
  const xml = strFromU8(bytes), root = xmlRoot(xml);
  const name = `/${part}`.toLowerCase();
  if (arr(parseXml(xml)?.Types?.Override).some((o) => attr(o, 'PartName')?.toLowerCase() === name)) protect('Der neue Paketteil ist bereits registriert.', 'The new package part is already registered.');
  files['[Content_Types].xml'] = strToU8(appendToRoot(xml, `<${prefix(root)}Override PartName="/${escapeXml(part)}" ContentType="${type}"/>`));
}

/** A sheet added in the editor contains nothing but cells, sizes and validations. */
function isPlainNewSheet(sheet: SheetData): boolean {
  const { name: _name, ...rest } = sheetMetadata(sheet) as SheetData;
  return sameValue(rest, {});
}

export function preserveXlsx(wb: Workbook, compression: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9, fresh: (data?: WorkbookData) => Uint8Array): Uint8Array {
  const source = wb.xlsxSource!;
  if (!source.baseline) protect('Die Import-Basis fehlt.', 'The import baseline is missing.');
  const baseline = source.baseline;
  const original = decodeArchive(source.archive), files = unzipSync(original), current = wb.toJSON();
  if (!source.structuralEdits && sameModel(current, baseline)) return original;
  if (Object.keys(files).some((name) => /(?:^_xmlsignatures\/|vbaProjectSignature)/i.test(name))) protect('Die Datei enthält digitale Signaturen.', 'The file contains digital signatures.');
  const rootRels = relations(files, '');
  const workbookPath = rootRels.find((r) => r.type.endsWith('/officeDocument'))?.target ?? 'xl/workbook.xml';
  const wbBytes = files[workbookPath]; if (!wbBytes) protect('Der Workbook-Teil fehlt.', 'The workbook part is missing.');
  let wbXml = strFromU8(wbBytes);
  const wbRoot = xmlRoot(wbXml);
  const wbRels = relations(files, workbookPath);
  const sheetEntries = arr(parseXml(wbXml)?.workbook?.sheets?.sheet);
  const worksheets = sheetEntries.map((sheet) => ({ name: attr(sheet, 'name'), rel: wbRels.find((r) => r.id === attr(sheet, 'id')) })).filter((s) => s.rel?.type.endsWith('/worksheet'));
  // Sheets added at the end are written as new parts; all original sheets keep their order and names.
  const appended = current.sheets.length >= baseline.sheets.length && baseline.sheets.every((sheet, i) => current.sheets[i]!.name === sheet.name)
    ? current.sheets.slice(baseline.sheets.length) : undefined;
  const nonCellChanges = !appended || !sameValue(metadata(current, baseline.sheets.length), metadata(baseline)) || !appended.every(isPlainNewSheet);
  // Keep serializer-owned files canonical even for cell-only edits. Patching
  // their cells in place can introduce inline strings or unused shared strings,
  // causing the next reopen to reject an otherwise safe table/structure edit.
  const generatedBaseline = unzipSync(fresh(baseline));
  if (equivalentPackage(files, generatedBaseline) && understandsFormulas(baseline)) {
    const generated = unzipSync(fresh());
    // Keep original document properties, including creation times, on a safe rebuild.
    for (const name of Object.keys(files)) if (name.startsWith('docProps/')) generated[name] = files[name]!;
    return zipSync(generated, { level: compression });
  }
  if (source.structuralEdits || nonCellChanges) {
    protect('Struktur-, Tabellen-, Diagramm- oder Regeländerung mit nicht vollständig modellierten Excel-Inhalten.', 'Structure, table, chart or rule change in a file with Excel content that is not fully modeled.');
  }
  if (worksheets.length !== baseline.sheets.length) protect('Nicht übereinstimmende Blattzuordnung.', 'Sheet mapping does not match.');
  const stylesPath = wbRels.find((r) => r.type.endsWith('/styles'))?.target;
  const stringsPath = wbRels.find((r) => r.type.endsWith('/sharedStrings'))?.target;
  // Indices of shared strings with rich text or phonetic runs, scanned on first use.
  let richIndices: Set<number> | undefined;
  const richStrings = { has: (index: number) => {
    if (index < 0) return false;
    if (!richIndices) richIndices = stringsPath && files[stringsPath] ? richStringIndices(strFromU8(files[stringsPath]!)) : new Set();
    return richIndices.has(index);
  } };
  const p = prefix(wbRoot);
  const transitional = attribute(wbXml.slice(wbRoot.start, wbRoot.openEnd), p ? `xmlns:${p.slice(0, -1)}` : 'xmlns') === NS_MAIN;
  let styles = stylesPath && files[stylesPath] ? new PreservedStyles(strFromU8(files[stylesPath]!)) : undefined;
  let createdStyles = false;
  const editableStyles = (): PreservedStyles => {
    if (styles) return styles;
    // A file without any style table gets a minimal one; existing parts stay untouched.
    if (stylesPath || !transitional) protect('Die Datei enthält keine bearbeitbare Style-Tabelle.', 'The file has no editable style table.');
    styles = new PreservedStyles(MINIMAL_STYLES);
    createdStyles = true;
    return styles;
  };
  const date1904 = ['1', 'true'].includes(nodeAttribute(wbXml, child(wbRoot, 'workbookPr') ?? wbRoot, 'date1904') ?? '');
  const context = (sheet: string, lockedRanges: string[]): SheetContext => ({
    styles: editableStyles, existingStyles: () => styles, date1904, hasCalcChain: wbRels.some((r) => r.type.endsWith('/calcChain')),
    richStrings, lockedRanges, sorted: !!source.sortedSheets?.[sheet], names: new Set(Object.keys(current.names ?? {}).map((n) => n.toUpperCase())),
  });
  for (let i = 0; i < worksheets.length; i++) {
    const entry = worksheets[i]!, before = baseline.sheets[i]!, after = current.sheets[i]!;
    if (entry.name !== before.name || after.name !== before.name || !entry.rel || !files[entry.rel.target]) protect('Die Blattzuordnung wurde verändert.', 'The sheet mapping was changed.');
    const originalSheet = strFromU8(files[entry.rel.target]!);
    const lockedRanges: string[] = [];
    for (const relation of relations(files, entry.rel.target)) {
      const bytes = files[relation.target]; if (!bytes) continue;
      if (relation.type.endsWith('/pivotTable')) {
        const ref = attr(parseXml(strFromU8(bytes))?.pivotTableDefinition?.location, 'ref');
        if (ref) lockedRanges.push(ref);
      }
      if (relation.type.endsWith('/table')) {
        const table = parseXml(strFromU8(bytes))?.table;
        const range = parseRange(attr(table, 'ref') ?? '')?.range;
        if (!range) continue;
        arr(table.tableColumns?.tableColumn).forEach((column, index) => {
          if (column?.calculatedColumnFormula !== undefined) lockedRanges.push(formatRange({ start: { row: range.start.row + 1, col: range.start.col + index }, end: { row: range.end.row, col: range.start.col + index } }));
        });
        const totals = Number(attr(table, 'totalsRowCount') ?? 0);
        if (totals > 0) lockedRanges.push(formatRange({ start: { row: range.end.row - totals + 1, col: range.start.col }, end: range.end }));
      }
    }
    if (sameValue(before, after)) continue;
    const originalRoot = xmlRoot(originalSheet, 3);
    // Only edits count; new results of formulas that depend on other sheets do not.
    if (child(originalRoot, 'sheetProtection') && !sameSheetContent(before, after)) protect(`Blattschutz in ${before.name}.`, `Sheet protection in ${before.name}.`);
    const patched = patchValidations(patchLayout(patchSheet(originalSheet, originalRoot, before, after, context(before.name, lockedRanges)), before, after), before, after);
    if (patched !== originalSheet) { xmlRoot(patched, 1); files[entry.rel.target] = strToU8(patched); }
  }
  if (appended.length) {
    if (!transitional) protect('Neue Blätter in Strict-Open-XML-Dateien.', 'New sheets in Strict Open XML files.');
    const rPrefix = /\sxmlns:([\w.-]+)\s*=\s*["']http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships["']/.exec(wbXml.slice(wbRoot.start, wbRoot.openEnd))?.[1];
    const sheetsNode = child(wbRoot, 'sheets');
    if (!rPrefix || !sheetsNode) protect('Die Blattliste der Arbeitsmappe lässt sich nicht sicher erweitern.', 'The workbook sheet list cannot be extended safely.');
    // Names must be unique among all sheets, including chart sheets that the editor does not show.
    const taken = new Set(sheetEntries.map((s) => attr(s, 'name')?.toUpperCase()));
    let sheetId = Math.max(0, ...sheetEntries.map((s) => Number(attr(s, 'sheetId') ?? 0)).filter(Number.isFinite));
    const dir = workbookPath.includes('/') ? workbookPath.slice(0, workbookPath.lastIndexOf('/') + 1) : '';
    let entries = '';
    for (const after of appended) {
      if (taken.has(after.name.toUpperCase())) protect(`Ein Blatt namens ${after.name} existiert bereits in der Datei.`, `A sheet named ${after.name} already exists in the file.`);
      taken.add(after.name.toUpperCase());
      let n = 1;
      while (hasPart(files, `${dir}worksheets/sheet${n}.xml`) || hasPart(files, relsPathFor(`${dir}worksheets/sheet${n}.xml`))) n++;
      const part = `${dir}worksheets/sheet${n}.xml`, empty: SheetData = { name: after.name, cells: {} };
      let xml = XML_HEAD + `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><sheetData/></worksheet>`;
      xml = patchValidations(patchLayout(patchSheet(xml, xmlRoot(xml, 3), empty, after, context(after.name, [])), empty, after), empty, after);
      xmlRoot(xml, 1);
      files[part] = strToU8(xml);
      const id = addRelationship(files, workbookPath, `${NS_R}/worksheet`, `worksheets/sheet${n}.xml`);
      addContentType(files, part, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
      entries += `<${p}sheet name="${escapeXml(after.name)}" sheetId="${++sheetId}" ${rPrefix}:id="${id}"/>`;
    }
    wbXml = editXml(wbXml, [{ start: sheetsNode.start, end: sheetsNode.end, value: contents(wbXml, sheetsNode, wbXml.slice(sheetsNode.openEnd, sheetsNode.closeStart) + entries) }]);
  }
  if (styles) {
    xmlRoot(styles.xml);
    if (createdStyles) {
      const dir = workbookPath.includes('/') ? workbookPath.slice(0, workbookPath.lastIndexOf('/') + 1) : '';
      const part = `${dir}styles.xml`;
      if (hasPart(files, part)) protect('Ein nicht verknüpfter Style-Teil ist bereits vorhanden.', 'An unlinked style part already exists.');
      files[part] = strToU8(styles.xml);
      addRelationship(files, workbookPath, `${NS_R}/styles`, 'styles.xml');
      addContentType(files, part, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
    } else files[stylesPath!] = strToU8(styles.xml);
  }
  // Preserve the original formula caches and calculation chain. Excel refreshes
  // dependent caches on opening; unknown formulas and their metadata stay intact.
  const root = xmlRoot(wbXml), cp = child(root, 'calcPr');
  let calc = cp ? raw(wbXml, cp) : `<${p}calcPr/>`;
  const calcRoot = xmlRoot(calc);
  let open = calc.slice(0, calcRoot.openEnd);
  open = setAttribute(setAttribute(open, 'fullCalcOnLoad', 1), 'forceFullCalc', 1);
  calc = editXml(calc, [{ start: 0, end: calcRoot.openEnd, value: open }]);
  wbXml = setChild(wbXml, 'calcPr', calc, ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']);
  files[workbookPath] = strToU8(wbXml);
  return zipSync(files, { level: compression });
}
