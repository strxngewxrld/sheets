import { inputToCellData } from '../commands.js';
import { formatAddress } from '../address.js';
import type { Scalar } from '../values.js';
import { formatValue, getDefaultLocale, LOCALE_DE, LOCALE_EN, type FormatLocale } from '../numfmt.js';
import type { CellData, SheetData } from '../sheet.js';
import { Workbook } from '../workbook.js';

export interface CsvReadOptions {
  /** Field delimiter; auto-detected from , ; tab | when omitted. */
  delimiter?: string;
  /** Parse numbers/dates/booleans (default true). When false, everything is text. */
  inferTypes?: boolean;
  /** Treat cells starting with "=" as formulas (default false for safety). */
  formulas?: boolean;
  /** Number locale; defaults to DE when the delimiter is ";" else EN. */
  locale?: FormatLocale;
  sheetName?: string;
}

export interface CsvWriteOptions {
  delimiter?: string;
  /** Write displayed (formatted) values instead of raw values (default false). */
  formatted?: boolean;
  /** Write formulas ("=SUM(...)") instead of their results (default false). */
  formulas?: boolean;
  lineEnding?: '\n' | '\r\n';
  /** End the last line with a line break (default true). */
  finalNewline?: boolean;
  /** Start with a UTF-8 byte order mark (default false). */
  bom?: boolean;
  locale?: FormatLocale;
}

/** Layout details of a CSV text that a save should reproduce. */
export interface CsvLayout {
  lineEnding: '\n' | '\r\n';
  finalNewline: boolean;
  bom: boolean;
}

/** Detects line endings, a final line break and a byte order mark (defaults for empty text). */
export function detectCsvLayout(text: string): CsvLayout {
  const bom = text.charCodeAt(0) === 0xfeff;
  const body = bom ? text.slice(1) : text;
  const first = /\r\n|\n|\r/.exec(body)?.[0];
  return { lineEnding: first === '\r\n' ? '\r\n' : '\n', finalNewline: body === '' || /[\r\n]$/.test(body), bom };
}

export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 10_000).split(/\r?\n/).slice(0, 20);
  let best = ',';
  let bestScore = -1;
  for (const d of [',', ';', '\t', '|']) {
    const counts = sample.filter((l) => l.length).map((l) => parseLine(l, d).length);
    if (!counts.length) continue;
    const first = counts[0]!;
    if (first < 2) continue;
    const consistent = counts.filter((c) => c === first).length;
    const score = consistent * 100 + first;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

function parseLine(line: string, d: string): string[] {
  return parseCsv(line, d)[0] ?? [];
}

/** RFC 4180 parser (quoted fields, escaped quotes, newlines in quotes). */
export function parseCsv(text: string, delimiter = ','): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  if (text.charCodeAt(0) === 0xfeff) i = 1; // BOM
  for (; i < text.length; i++) {
    const ch = text[i]!;
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else inQuotes = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && field === '') inQuotes = true;
    else if (ch === delimiter) {
      row.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export function csvToSheetData(text: string, options: CsvReadOptions = {}): SheetData {
  const delimiter = options.delimiter ?? detectDelimiter(text);
  const locale = options.locale ?? (delimiter === ';' ? LOCALE_DE : LOCALE_EN);
  const rows = parseCsv(text, delimiter);
  const cells: Record<string, CellData> = {};
  rows.forEach((line, r) =>
    line.forEach((raw, c) => {
      if (raw === '') return;
      let data: CellData | null;
      if (options.inferTypes === false) data = { v: raw };
      else if (raw.startsWith('=') && !options.formulas) data = { v: raw };
      else {
        data = inputToCellData(raw, undefined, locale);
        // Keep an inferred type only if it loses nothing, otherwise a plain open + save
        // would rewrite the file (0004711 -> 4711, 2024-01-05 -> 45296, 5% -> 0.05).
        if (data && data.f === undefined && data.v !== undefined && !isLossless(raw, data, locale)) data = { v: raw };
      }
      if (data) cells[formatAddress(r, c)] = data;
    }),
  );
  return { name: options.sheetName ?? 'Sheet1', cells };
}

/** The text the CSV writer emits for a plain (unformatted) value. */
function scalarText(value: Scalar, locale: FormatLocale): string {
  if (typeof value === 'number') return locale.decimal === '.' ? String(value) : String(value).replace('.', locale.decimal);
  return formatValue(value, undefined, LOCALE_EN);
}

const numberPatterns = new Map<string, RegExp>();
function numberPattern(locale: FormatLocale, grouped: boolean): RegExp {
  const key = `${locale.decimal}${locale.group}${grouped}`;
  let re = numberPatterns.get(key);
  if (!re) {
    const esc = (c: string) => c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const integer = grouped ? `[1-9]\\d{0,2}(?:${esc(locale.group)}\\d{3})+` : '(?:0|[1-9]\\d*)?';
    numberPatterns.set(key, re = new RegExp(`^-?${integer}(?:${esc(locale.decimal)}\\d+)?$`));
  }
  return re;
}

function isLossless(raw: string, data: CellData, locale: FormatLocale): boolean {
  const v = data.v as Scalar;
  if (typeof v === 'number') {
    // Dates, percentages and currencies would be written as plain numbers. Correct digit
    // grouping in the file's locale is accepted (1.234,50 -> 1234,5), like typed input.
    const fmt = data.s?.numFmt;
    const grouped = fmt === '#,##0' || fmt === '#,##0.00';
    if (fmt && !grouped) return false;
    // Only the locale's separators; no sign, exponent, spaces, leading zeros or trailing separator.
    if (!numberPattern(locale, grouped).test(raw) || !/\d/.test(raw)) return false;
    if (raw.startsWith('-') && v === 0) return false;
    return raw.replace(/\D/g, '').replace(/^0+/, '').length <= 15; // beyond that doubles lose digits
  }
  return scalarText(v, locale) === raw;
}

/**
 * Amount of workbook content per kind that a CSV/TSV file cannot hold (only the
 * values of the active sheet are written). Callers compare it with the state at
 * load time to warn before content is dropped.
 */
export function csvUnsupportedContent(wb: Workbook): Record<string, number> {
  const out: Record<string, number> = { sheets: wb.sheets.length - 1, charts: 0, formulas: 0, formats: 0, tables: 0, rules: 0, merges: 0, notes: 0, pivots: wb.pivotTables.length, names: wb.names.size };
  for (const sheet of wb.sheets) {
    out.charts! += sheet.charts.length;
    out.tables! += sheet.tables.length;
    out.rules! += sheet.conditionalFormats.length + sheet.validations.length;
    out.merges! += sheet.merges.length;
    out.notes! += sheet.notes.size + sheet.links.size;
    for (const cell of sheet.cells.values()) {
      if (cell.formula !== undefined) out.formulas!++;
      if (cell.style) out.formats!++;
    }
  }
  return out;
}

export function readCsv(text: string, options: CsvReadOptions = {}): Workbook {
  const wb = new Workbook({ empty: true });
  wb.load({ sheets: [csvToSheetData(text, options)] });
  return wb;
}

export function writeCsv(wb: Workbook, sheet: string | number = wb.activeSheetIndex, options: CsvWriteOptions = {}): string {
  const s = wb.requireSheet(sheet);
  const d = options.delimiter ?? ',';
  const eol = options.lineEnding ?? '\r\n';
  const locale = options.locale ?? (d === ';' ? LOCALE_DE : getDefaultLocale());
  const used = s.usedRange();
  if (!used) return options.bom ? '\uFEFF' : '';
  const lines: string[] = [];
  for (let r = 0; r <= used.end.row; r++) {
    const fields: string[] = [];
    for (let c = 0; c <= used.end.col; c++) {
      const cell = s.getCell(r, c);
      let text = '';
      if (cell) {
        if (options.formulas && cell.formula !== undefined) text = `=${cell.formula}`;
        else if (options.formatted) text = formatValue(cell.value, cell.style?.numFmt, locale);
        else text = scalarText(cell.value, locale);
      }
      if (text.includes(d) || /["\r\n]/.test(text) || /^\s|\s$/.test(text)) text = `"${text.replace(/"/g, '""')}"`;
      fields.push(text);
    }
    lines.push(fields.join(d));
  }
  return (options.bom ? '\uFEFF' : '') + lines.join(eol) + (options.finalNewline === false ? '' : eol);
}
