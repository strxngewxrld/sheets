import { strFromU8, strToU8 } from 'fflate';
import { Workbook } from '../workbook.js';
import { LOCALE_DE, LOCALE_EN } from '../numfmt.js';
import { csvDataSheet, decodeCsvBytes, detectCsvLayout, detectDelimiter, encodeCsvText, readCsv, splitCsvRecords, writeCsv, type CsvEncoding, type CsvReadOptions, type CsvWriteOptions } from './csv.js';
import { readXlsx, type XlsxReadOptions } from './xlsx-read.js';
import { writeXlsx, type XlsxWriteOptions } from './xlsx-write.js';

export type FileFormat = 'xlsx' | 'xlsm' | 'csv' | 'tsv' | 'json';

/** Excel workbook formats (macros in .xlsm are kept but never run). */
export function isXlsxFormat(format: FileFormat): format is 'xlsx' | 'xlsm' {
  return format === 'xlsx' || format === 'xlsm';
}

export function formatFromFileName(name: string): FileFormat {
  const ext = name.toLowerCase().split('.').pop();
  if (ext === 'csv' || ext === 'txt') return 'csv';
  if (ext === 'tsv' || ext === 'tab') return 'tsv';
  if (ext === 'json') return 'json';
  if (ext === 'xlsm') return 'xlsm';
  return 'xlsx';
}

export interface OpenOptions {
  csv?: CsvReadOptions;
  xlsx?: XlsxReadOptions;
}

/** Opens a workbook from file bytes; the format is derived from the file name. */
export function openWorkbook(bytes: Uint8Array | ArrayBuffer, fileName: string, options: OpenOptions = {}): Workbook {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const format = formatFromFileName(fileName);
  const baseName = fileName.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1';
  switch (format) {
    case 'csv':
    case 'tsv': {
      const { text, encoding } = decodeCsvBytes(u8);
      const delimiter = options.csv?.delimiter ?? (format === 'tsv' ? '\t' : detectDelimiter(text));
      const wb = readCsv(text, { sheetName: baseName, ...options.csv, delimiter });
      // Remembered so a save writes the file back the way it was (also through JSON and undo).
      wb.csvSource = { ...detectCsvLayout(text), delimiter, encoding, sheet: wb.sheets[0]!.name, rows: splitCsvRecords(text, delimiter) };
      return wb;
    }
    case 'json':
      return Workbook.fromJSON(JSON.parse(strFromU8(u8)));
    default:
      return readXlsx(u8, options.xlsx);
  }
}

export interface SaveOptions {
  /** `encoding` and `sheet` default to the opened CSV file's encoding and data sheet. */
  csv?: CsvWriteOptions & { encoding?: CsvEncoding; sheet?: string | number };
  xlsx?: XlsxWriteOptions;
}

/**
 * Serializes a workbook into the given format. CSV/TSV write one sheet: the data
 * sheet of an opened CSV file, otherwise the active sheet. A workbook read from
 * CSV/TSV is written with its original delimiter, line endings, final line
 * break, byte order mark and encoding unless options override them.
 * New CSV/TSV conversions protect formula-like text for Excel by default;
 * pass csv.escapeFormulas: false for unchanged data exchange.
 */
export function saveWorkbook(wb: Workbook, format: FileFormat, options: SaveOptions = {}): Uint8Array {
  switch (format) {
    case 'csv':
    case 'tsv':
      return saveCsv(wb, format, options.csv);
    case 'json':
      return strToU8(JSON.stringify(wb.toJSON(), null, 2));
    case 'xlsm':
      return writeXlsx(wb, { ...options.xlsx, macroEnabled: true });
    default:
      return writeXlsx(wb, options.xlsx);
  }
}

function saveCsv(wb: Workbook, format: 'csv' | 'tsv', options: NonNullable<SaveOptions['csv']> = {}): Uint8Array {
  const source = wb.csvSource;
  const { encoding: requested, sheet: requestedSheet, ...write } = options;
  const delimiter = write.delimiter ?? (format === 'tsv' ? '\t' : source && source.delimiter !== '\t' ? source.delimiter : ',');
  const sheet = requestedSheet ?? csvDataSheet(wb).name;
  // Unchanged records of the opened file are written back as they were (same sheet and delimiter only).
  const original = source && delimiter === source.delimiter && wb.getSheet(sheet) === csvDataSheet(wb) ? source.rows : undefined;
  const text = writeCsv(wb, sheet, {
    // New CSV/TSV conversions protect text for spreadsheet viewing. Existing
    // CSV records retain their exact values unless the caller requests protection.
    escapeFormulas: !source || delimiter !== source.delimiter,
    ...(source ? { lineEnding: source.lineEnding, finalNewline: source.finalNewline, bom: source.bom, locale: delimiter === ';' ? LOCALE_DE : LOCALE_EN, original } : {}),
    ...Object.fromEntries(Object.entries(write).filter(([, value]) => value !== undefined)),
    delimiter,
  });
  // Text that Windows-1252 cannot hold is written as UTF-8 with a BOM rather than losing characters.
  return encodeCsvText(text, requested ?? source?.encoding ?? 'utf-8') ?? strToU8(text.startsWith('\uFEFF') ? text : `\uFEFF${text}`);
}

export * from './csv.js';
export { readXlsx, xlsxToData, charWidthToPx, pointsToPx, type XlsxReadOptions } from './xlsx-read.js';
export { XlsxProtectionError, setXlsxMessageLanguage, type XlsxMessageLanguage } from './xlsx-preserve-xml.js';
export { writeXlsx, chartXml, pxToCharWidth, pxToPoints, type XlsxWriteOptions } from './xlsx-write.js';
