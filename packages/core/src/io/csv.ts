import { inputToCellData } from '../commands.js';
import { formatAddress } from '../address.js';
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
  locale?: FormatLocale;
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
      else data = inputToCellData(raw, undefined, locale);
      if (data) cells[formatAddress(r, c)] = data;
    }),
  );
  return { name: options.sheetName ?? 'Sheet1', cells };
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
  if (!used) return '';
  const lines: string[] = [];
  for (let r = 0; r <= used.end.row; r++) {
    const fields: string[] = [];
    for (let c = 0; c <= used.end.col; c++) {
      const cell = s.getCell(r, c);
      let text = '';
      if (cell) {
        if (options.formulas && cell.formula !== undefined) text = `=${cell.formula}`;
        else if (options.formatted) text = formatValue(cell.value, cell.style?.numFmt, locale);
        else if (typeof cell.value === 'number') text = locale.decimal === '.' ? String(cell.value) : String(cell.value).replace('.', locale.decimal);
        else text = formatValue(cell.value, undefined, LOCALE_EN);
      }
      if (text.includes(d) || /["\r\n]/.test(text) || /^\s|\s$/.test(text)) text = `"${text.replace(/"/g, '""')}"`;
      fields.push(text);
    }
    lines.push(fields.join(d));
  }
  return lines.join(eol) + eol;
}
