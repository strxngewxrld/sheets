import { inputToCellData } from '../commands.js';
import { formatAddress } from '../address.js';
import { scalarToJson, type Scalar } from '../values.js';
import { formatValue, getDefaultLocale, LOCALE_DE, LOCALE_EN, type FormatLocale } from '../numfmt.js';
import type { CellData, Sheet, SheetData } from '../sheet.js';
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
  /**
   * Records of the file as it was read (see splitCsvRecords). A record whose fields
   * still read as the row's current values is written back verbatim, so quoting,
   * number spelling and short rows of unchanged lines stay as they were.
   */
  original?: (string | null)[];
}

/** Layout details of a CSV text that a save should reproduce. */
export interface CsvLayout {
  lineEnding: '\n' | '\r\n';
  finalNewline: boolean;
  bom: boolean;
}

/** Text encodings of CSV files that are read and written back unchanged. */
export type CsvEncoding = 'utf-8' | 'windows-1252';

/** How an opened CSV/TSV file was written; a save reproduces it. */
export interface CsvSource extends CsvLayout {
  delimiter: string;
  encoding: CsvEncoding;
  /** The sheet holding the file's data; saved even when another sheet is active. */
  sheet: string;
  /** The file's records as written (null for inserted rows), reused for rows that still hold the same values. */
  rows?: (string | null)[];
}

/**
 * Decodes CSV bytes. Valid UTF-8 (with or without BOM) is read as UTF-8; anything
 * else as Windows-1252, the encoding of CSV exports from Excel on German or
 * Western European Windows. Reading such a file as UTF-8 would replace every
 * umlaut with U+FFFD and destroy it on save.
 */
export function decodeCsvBytes(bytes: Uint8Array): { text: string; encoding: CsvEncoding } {
  try {
    return { text: new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(bytes), encoding: 'utf-8' };
  } catch {
    return { text: new TextDecoder('windows-1252').decode(bytes), encoding: 'windows-1252' };
  }
}

// Windows-1252 bytes 0x80-0x9F that differ from Latin-1 (undefined bytes decode to their C1 control).
const CP1252_HIGH = '€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008DŽ\u008F\u0090‘’“”•–—˜™š›œ\u009DžŸ';

/**
 * Encodes text for a CSV file. Returns null when Windows-1252 cannot represent
 * a character, so the caller can fall back to UTF-8 instead of losing it.
 */
export function encodeCsvText(text: string, encoding: CsvEncoding): Uint8Array | null {
  if (encoding === 'utf-8') return new TextEncoder().encode(text);
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    const high = CP1252_HIGH.indexOf(text[i]!);
    if (high >= 0) out[i] = 0x80 + high;
    else if (code < 0x80 || (code >= 0xa0 && code <= 0xff)) out[i] = code;
    else return null;
  }
  return out;
}

/** Characters of the text that Windows-1252 cannot represent (unique, in order of appearance). */
export function unencodableCsvChars(text: string, encoding: CsvEncoding): string[] {
  if (encoding === 'utf-8') return [];
  const found = new Set<string>();
  for (const char of text) if (encodeCsvText(char, encoding) === null) found.add(char);
  return [...found];
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

/**
 * Splits CSV text into its records exactly as written (without line breaks; a BOM
 * is skipped). Quotes are handled like parseCsv, so record i holds row i.
 */
export function splitCsvRecords(text: string, delimiter: string): string[] {
  const records: string[] = [];
  let start = text.charCodeAt(0) === 0xfeff ? 1 : 0;
  let inQuotes = false;
  let fieldEmpty = true;
  for (let i = start; i < text.length; i++) {
    const ch = text[i]!;
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { i++; fieldEmpty = false; } else inQuotes = false;
      } else fieldEmpty = false;
      continue;
    }
    if (ch === '"' && fieldEmpty) inQuotes = true;
    else if (ch === delimiter) fieldEmpty = true;
    else if (ch === '\n' || ch === '\r') {
      records.push(text.slice(start, i));
      if (ch === '\r' && text[i + 1] === '\n') i++;
      start = i + 1;
      fieldEmpty = true;
    } else fieldEmpty = false;
  }
  if (start < text.length) records.push(text.slice(start));
  return records;
}

/** Converts one field like the reader does; also used to verify original records on save. */
function fieldToCellData(raw: string, options: CsvReadOptions, locale: FormatLocale): CellData | null {
  if (raw === '') return null;
  if (options.inferTypes === false) return { v: raw };
  if (raw.startsWith('=') && !options.formulas) return { v: raw };
  const data = inputToCellData(raw, undefined, locale);
  // Keep an inferred type only if it loses nothing, otherwise a plain open + save
  // would rewrite the file (0004711 -> 4711, 2024-01-05 -> 45296, 5% -> 0.05).
  if (data && data.f === undefined && data.v !== undefined && !isLossless(raw, data, locale)) return { v: raw };
  return data;
}

export function csvToSheetData(text: string, options: CsvReadOptions = {}): SheetData {
  const delimiter = options.delimiter ?? detectDelimiter(text);
  const locale = options.locale ?? (delimiter === ';' ? LOCALE_DE : LOCALE_EN);
  const rows = parseCsv(text, delimiter);
  const cells: Record<string, CellData> = {};
  rows.forEach((line, r) =>
    line.forEach((raw, c) => {
      const data = fieldToCellData(raw, options, locale);
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

/** The sheet a CSV/TSV save writes: the opened file's data sheet, else the active sheet. */
export function csvDataSheet(wb: Workbook): Sheet {
  return (wb.csvSource && wb.getSheet(wb.csvSource.sheet)) || wb.activeSheet;
}

/**
 * Amount of workbook content per kind that a CSV/TSV file cannot hold (only the
 * values of the data sheet are written, see csvDataSheet). Callers compare it with the state at
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
  const width = used ? used.end.col + 1 : 0;
  // An original record is reused only if it reads back as exactly the row's current values.
  const original = options.formatted || options.formulas ? undefined : options.original;
  const verbatim = (r: number): string | undefined => {
    const record = original?.[r];
    if (record === undefined || record === null) return undefined;
    const fields = parseCsv(record, d)[0] ?? [];
    for (let c = 0; c < Math.max(fields.length, width); c++) {
      const expected = fields[c] === undefined ? null : fieldToCellData(fields[c]!, {}, locale)?.v ?? null;
      if (JSON.stringify(expected) !== JSON.stringify(scalarToJson(s.getValue(r, c)))) return undefined;
    }
    return record;
  };
  let lastRow = used ? used.end.row : -1;
  // Empty records after the data (blank lines) are kept while they are still empty.
  while (original && lastRow + 1 < original.length && verbatim(lastRow + 1) !== undefined) lastRow++;
  if (lastRow < 0) return options.bom ? '\uFEFF' : '';
  const lines: string[] = [];
  for (let r = 0; r <= lastRow; r++) {
    const kept = verbatim(r);
    if (kept !== undefined) { lines.push(kept); continue; }
    const fields: string[] = [];
    for (let c = 0; c < width; c++) {
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
