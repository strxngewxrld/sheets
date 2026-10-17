import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { formatAddress, formatRange, parseAddress, parseRange } from '../address.js';
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
  return a.sheets.every((sa, i) => {
    const sb = b.sheets[i]!;
    const { cells: ca, ...ra } = sa, { cells: cb, ...rb } = sb;
    if (!sameValue(ra, rb)) return false;
    for (const key in ca) if (!sameCellContent(ca[key], cb[key])) return false;
    for (const key in cb) if (!(key in ca) && !sameCellContent(undefined, cb[key])) return false;
    return true;
  });
}

function metadata(data: WorkbookData): unknown {
  const { xlsxSource: _source, activeSheet: _active, ...rest } = data;
  return { ...rest, sheets: rest.sheets.map(({ cells: _cells, colWidths: _cols, rowHeights: _rows, validations: _validations, ...sheet }) => sheet) };
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

const AFTER_VALIDATIONS = ['hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** Replaces the sheet's <dataValidations>; Excel 2010 rules in extLst cannot be merged safely. */
function patchValidations(xml: string, before: SheetData, after: SheetData): string {
  const rules = after.validations ?? [];
  if (sameValue(before.validations ?? [], rules)) return xml;
  const root = xmlRoot(xml);
  const ext = child(root, 'extLst');
  if (ext && /dataValidations/.test(raw(xml, ext))) protect(`Erweiterte Datenüberprüfung (Excel 2010) in ${before.name}.`);
  return setChild(xml, 'dataValidations', rules.length ? validationsXml(rules, prefix(root)) : '', AFTER_VALIDATIONS);
}

function updateCell(xml: string, node: XmlSpan | undefined, ref: string, before: CellData | undefined, after: CellData | undefined, p: string, styles: PreservedStyles | undefined, inheritedStyle: number, richStrings: { has(index: number): boolean }, table?: string): string {
  const valueChanged = !sameValue(payload(before), payload(after));
  let result = node ? raw(xml, node) : `<${p}c r="${ref}"/>`;
  let root = xmlRoot(result), open = result.slice(0, root.openEnd);
  if (valueChanged) {
    if (attribute(open, 'cm') !== undefined || attribute(open, 'vm') !== undefined) protect(`Zellmetadaten in ${ref}.`);
    if (!['n', 'b', 's', 'str', 'inlineStr', 'e'].includes(attribute(open, 't') ?? 'n')) protect(`Nicht unterstützter Zelltyp in ${ref}.`);
    if (root.children.some((c) => !['f', 'v', 'is'].includes(c.name))) protect(`Erweiterte Zellinhalte in ${ref}.`);
    const inline = child(root, 'is');
    const sharedIndex = attribute(open, 't') === 's' ? Number(text(parseXml(result)?.c?.v)) : -1;
    if (inline?.children.some((c) => c.name !== 't') || richStrings.has(sharedIndex)) protect(`Rich-Text oder phonetische Annotationen in ${ref}.`);
    const f = child(root, 'f');
    if (f && attribute(result.slice(f.start, f.openEnd), 't') && attribute(result.slice(f.start, f.openEnd), 't') !== 'normal') protect(`Geteilte Formel, Matrixformel oder Datentabelle in ${ref}.`);
    if (before?.a || after?.a) protect(`Dynamische Matrix in ${ref}.`);
    let type: string | undefined, body = '';
    if (after?.f !== undefined) {
      body = `<${p}f>${escapeXml(addFutureFunctionPrefixes(after.f, table ? { table } : {}))}</${p}f>`;
      // New/edited formulas are recalculated by Excel; never invent an error cache.
    } else if (typeof after?.v === 'string') {
      type = 'inlineStr'; body = `<${p}is><${p}t xml:space="preserve">${textValue(after.v)}</${p}t></${p}is>`;
    } else if (typeof after?.v === 'number') {
      if (!Number.isFinite(after.v)) protect(`Nicht endlicher Zahlenwert in ${ref}.`);
      body = `<${p}v>${after.v}</${p}v>`;
    } else if (typeof after?.v === 'boolean') { type = 'b'; body = `<${p}v>${after.v ? 1 : 0}</${p}v>`; }
    else if (after?.v && typeof after.v === 'object') { type = 'e'; body = `<${p}v>${escapeXml(after.v.error)}</${p}v>`; }
    const remaining = editXml(result, root.children.filter((n) => ['f', 'v', 'is'].includes(n.name)).map((n) => ({ start: n.start, end: n.end, value: '' })));
    root = xmlRoot(remaining);
    open = setAttribute(remaining.slice(0, root.openEnd), 't', type);
    result = contents(remaining, root, body + remaining.slice(root.openEnd, root.closeStart), open);
  }
  if (!sameValue(before?.s, after?.s)) {
    if (!styles) protect('Die Datei enthält keine bearbeitbare Style-Tabelle.');
    root = xmlRoot(result); open = result.slice(0, root.openEnd);
    const index = styles.get(Number(attribute(open, 's') ?? inheritedStyle), before?.s, after?.s);
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
    if (!address) protect('Matrixformel ohne Zelladresse.');
    out.push(parseRange(address));
  }
  return out;
}

function patchSheet(xml: string, root: XmlSpan, before: WorkbookData['sheets'][number], after: WorkbookData['sheets'][number], styles: PreservedStyles | undefined, date1904: boolean, hasCalcChain: boolean, richStrings: { has(index: number): boolean }, lockedRanges: string[]): string {
  const changes = [...new Set([...Object.keys(before.cells), ...Object.keys(after.cells)])]
    .filter((key) => !sameValue([payload(before.cells[key]), before.cells[key]?.s], [payload(after.cells[key]), after.cells[key]?.s]));
  if (!changes.length) return xml;
  const p = prefix(root);
  if (root.name !== 'worksheet') protect('Nicht unterstützter Blatttyp.');
  if (child(root, 'sheetProtection')) protect(`Blattschutz in ${before.name}.`);
  const data = child(root, 'sheetData'); if (!data) protect('sheetData fehlt.');
  const rows = new Map<number, XmlSpan>(), cells = new Map<string, XmlSpan>();
  let rowIndex = -1;
  for (const row of data.children) {
    if (row.name !== 'row') continue;
    if (row.qname !== p + 'row') protect('Nicht unterstützter Zeilen-Namensraum.');
    rowIndex = Number(nodeAttribute(xml, row, 'r') ?? rowIndex + 2) - 1;
    if (rows.has(rowIndex)) protect('Doppelte Zeilenadresse.');
    rows.set(rowIndex, row);
  }
  const arrays = arrayRanges(xml, data, p);
  const changedRows = new Map<number, string[]>();
  for (const ref of changes) {
    const a = parseAddress(ref); if (!a) protect('Ungültige Zelladresse im Modell.');
    const valueChanged = !sameValue(payload(before.cells[ref]), payload(after.cells[ref]));
    if (valueChanged) {
      if (lockedRanges.some((ref) => {
        const r = parseRange(ref)?.range;
        return r && a.row >= r.start.row && a.row <= r.end.row && a.col >= r.start.col && a.col <= r.end.col;
      })) protect(`Berechnete Tabellenspalte, Summenzeile oder Pivot-Ausgabe bei ${ref}.`);
      if (date1904) protect('Die Datei verwendet das noch nicht vollständig unterstützte Datumssystem 1904.');
      if (arrays.some((r) => r && a.row >= r.range.start.row && a.row <= r.range.end.row && a.col >= r.range.start.col && a.col <= r.range.end.col)) protect(`Matrixformel oder Datentabelle bei ${ref}.`);
      if (hasCalcChain && before.cells[ref]?.f !== after.cells[ref]?.f) protect('Formeländerung mit vorhandener Berechnungskette.');
      for (const table of before.tables ?? []) {
        const r = parseRange(table.range)!.range;
        if (table.headerRow && a.row === r.start.row && a.col >= r.start.col && a.col <= r.end.col) protect(`Änderung eines nativen Tabellenkopfs bei ${ref}.`);
      }
    }
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
      if (c.qname !== p + 'c') protect('Nicht unterstützter Zell-Namensraum.');
      const address = nodeAttribute(xml, c, 'r');
      const a = address ? parseAddress(address) : { row: index, col: col + 1 };
      if (!a || a.row !== index) protect('Ungültige Zelladresse.');
      col = a.col;
      const ref = formatAddress(a.row, a.col);
      if (cells.has(ref)) protect('Doppelte Zelladresse.');
      cells.set(ref, c);
    }
  }
  const rowEdits: XmlEdit[] = [], insertedRows = new Map<number, string>();
  for (const [index, refs] of [...changedRows].sort(([a], [b]) => a - b)) {
    const row = rows.get(index);
    if (!row && [...rows.values()].some((node) => nodeAttribute(xml, node, 'r') === undefined)) protect('Neue Zeilen bei impliziten Original-Zeilenadressen.');
    const base = row ? raw(xml, row) : `<${p}row r="${index + 1}"/>`;
    const edits: XmlEdit[] = [], insertions = new Map<number, string>();
    const offset = row?.start ?? 0;
    for (const ref of refs.sort((a, b) => parseAddress(a)!.col - parseAddress(b)!.col)) {
      const node = cells.get(ref);
      if (!node && row?.children.some((c) => c.name === 'c' && nodeAttribute(xml, c, 'r') === undefined)) protect('Neue Zellen bei impliziten Original-Zelladressen.');
      const col = parseAddress(ref)!.col;
      const columnStyle = child(root, 'cols')?.children.filter((n) => n.name === 'col').find((n) => col + 1 >= Number(nodeAttribute(xml, n, 'min')) && col + 1 <= Number(nodeAttribute(xml, n, 'max')));
      const inheritedStyle = Number((row && nodeAttribute(xml, row, 's')) ?? (columnStyle && nodeAttribute(xml, columnStyle, 'style')) ?? 0);
      const at = parseAddress(ref);
      const table = at ? tableNameAt(after.tables ?? [], at.row, at.col) : undefined;
      const value = updateCell(xml, node, ref, before.cells[ref], after.cells[ref], p, styles, inheritedStyle, richStrings, table);
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

export function preserveXlsx(wb: Workbook, compression: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9, fresh: (data?: WorkbookData) => Uint8Array): Uint8Array {
  const source = wb.xlsxSource!;
  if (!source.baseline) protect('Die Import-Basis fehlt.');
  const original = decodeArchive(source.archive), files = unzipSync(original), current = wb.toJSON();
  if (!source.structuralEdits && sameModel(current, source.baseline)) return original;
  if (Object.keys(files).some((name) => /(?:^_xmlsignatures\/|vbaProjectSignature)/i.test(name))) protect('Die Datei enthält digitale Signaturen.');
  const rootRels = relations(files, '');
  const workbookPath = rootRels.find((r) => r.type.endsWith('/officeDocument'))?.target ?? 'xl/workbook.xml';
  const wbBytes = files[workbookPath]; if (!wbBytes) protect('Der Workbook-Teil fehlt.');
  let wbXml = strFromU8(wbBytes);
  const wbRoot = xmlRoot(wbXml);
  const wbRels = relations(files, workbookPath);
  const worksheets = arr(parseXml(wbXml)?.workbook?.sheets?.sheet).map((sheet) => ({ name: attr(sheet, 'name'), rel: wbRels.find((r) => r.id === attr(sheet, 'id')) })).filter((s) => s.rel?.type.endsWith('/worksheet'));
  const nonCellChanges = !sameValue(metadata(current), metadata(source.baseline));
  if (source.structuralEdits || nonCellChanges) {
    const generatedBaseline = unzipSync(fresh(source.baseline));
    if (!equivalentPackage(files, generatedBaseline) || !understandsFormulas(source.baseline)) protect('Struktur-, Tabellen-, Diagramm- oder Regeländerung mit nicht vollständig modellierten Excel-Inhalten.');
    const generated = unzipSync(fresh());
    // Keep original document properties, including creation times, on a safe rebuild.
    for (const name of Object.keys(files)) if (name.startsWith('docProps/')) generated[name] = files[name]!;
    return zipSync(generated, { level: compression });
  }
  if (worksheets.length !== current.sheets.length) protect('Nicht übereinstimmende Blattzuordnung.');
  const stylesPath = wbRels.find((r) => r.type.endsWith('/styles'))?.target;
  const stringsPath = wbRels.find((r) => r.type.endsWith('/sharedStrings'))?.target;
  // Indices of shared strings with rich text or phonetic runs, scanned on first use.
  let richIndices: Set<number> | undefined;
  const richStrings = { has: (index: number) => {
    if (index < 0) return false;
    if (!richIndices) richIndices = stringsPath && files[stringsPath] ? richStringIndices(strFromU8(files[stringsPath]!)) : new Set();
    return richIndices.has(index);
  } };
  const styles = stylesPath && files[stylesPath] ? new PreservedStyles(strFromU8(files[stylesPath]!)) : undefined;
  const date1904 = ['1', 'true'].includes(nodeAttribute(wbXml, child(wbRoot, 'workbookPr') ?? wbRoot, 'date1904') ?? '');
  for (let i = 0; i < worksheets.length; i++) {
    const entry = worksheets[i]!, before = source.baseline.sheets[i]!, after = current.sheets[i]!;
    if (entry.name !== before.name || after.name !== before.name || !entry.rel || !files[entry.rel.target]) protect('Die Blattzuordnung wurde verändert.');
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
          if (column.calculatedColumnFormula !== undefined) lockedRanges.push(formatRange({ start: { row: range.start.row + 1, col: range.start.col + index }, end: { row: range.end.row, col: range.start.col + index } }));
        });
        const totals = Number(attr(table, 'totalsRowCount') ?? 0);
        if (totals > 0) lockedRanges.push(formatRange({ start: { row: range.end.row - totals + 1, col: range.start.col }, end: range.end }));
      }
    }
    if (sameValue(before, after)) continue;
    const originalRoot = xmlRoot(originalSheet, 3);
    if (child(originalRoot, 'sheetProtection')) protect(`Blattschutz in ${before.name}.`);
    const patched = patchValidations(patchLayout(patchSheet(originalSheet, originalRoot, before, after, styles, date1904, wbRels.some((r) => r.type.endsWith('/calcChain')), richStrings, lockedRanges), before, after), before, after);
    if (patched !== originalSheet) { xmlRoot(patched, 1); files[entry.rel.target] = strToU8(patched); }
  }
  if (stylesPath && styles) { xmlRoot(styles.xml); files[stylesPath] = strToU8(styles.xml); }
  // Preserve the original formula caches and calculation chain. Excel refreshes
  // dependent caches on opening; unknown formulas and their metadata stay intact.
  const cp = child(wbRoot, 'calcPr'), p = prefix(wbRoot);
  let calc = cp ? raw(wbXml, cp) : `<${p}calcPr/>`;
  const calcRoot = xmlRoot(calc);
  let open = calc.slice(0, calcRoot.openEnd);
  open = setAttribute(setAttribute(open, 'fullCalcOnLoad', 1), 'forceFullCalc', 1);
  calc = editXml(calc, [{ start: 0, end: calcRoot.openEnd, value: open }]);
  wbXml = setChild(wbXml, 'calcPr', calc, ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']);
  files[workbookPath] = strToU8(wbXml);
  return zipSync(files, { level: compression });
}
