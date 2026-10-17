import { strFromU8, strToU8 } from 'fflate';
import { Workbook } from '../workbook.js';
import { readCsv, writeCsv, type CsvReadOptions, type CsvWriteOptions } from './csv.js';
import { readXlsx, type XlsxReadOptions } from './xlsx-read.js';
import { writeXlsx, type XlsxWriteOptions } from './xlsx-write.js';

export type FileFormat = 'xlsx' | 'csv' | 'tsv' | 'json';

export function formatFromFileName(name: string): FileFormat {
  const ext = name.toLowerCase().split('.').pop();
  if (ext === 'csv' || ext === 'txt') return 'csv';
  if (ext === 'tsv' || ext === 'tab') return 'tsv';
  if (ext === 'json') return 'json';
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
      return readCsv(strFromU8(u8), { sheetName: baseName, ...options.csv });
    case 'tsv':
      return readCsv(strFromU8(u8), { delimiter: '\t', sheetName: baseName, ...options.csv });
    case 'json':
      return Workbook.fromJSON(JSON.parse(strFromU8(u8)));
    default:
      return readXlsx(u8, options.xlsx);
  }
}

export interface SaveOptions {
  csv?: CsvWriteOptions;
  xlsx?: XlsxWriteOptions;
}

/** Serializes a workbook into the given format. CSV/TSV write the active sheet. */
export function saveWorkbook(wb: Workbook, format: FileFormat, options: SaveOptions = {}): Uint8Array {
  switch (format) {
    case 'csv':
      return strToU8(writeCsv(wb, wb.activeSheetIndex, options.csv));
    case 'tsv':
      return strToU8(writeCsv(wb, wb.activeSheetIndex, { delimiter: '\t', ...options.csv }));
    case 'json':
      return strToU8(JSON.stringify(wb.toJSON(), null, 2));
    default:
      return writeXlsx(wb, options.xlsx);
  }
}

export * from './csv.js';
export { readXlsx, xlsxToData, charWidthToPx, pointsToPx, type XlsxReadOptions } from './xlsx-read.js';
export { writeXlsx, chartXml, pxToCharWidth, pxToPoints, type XlsxWriteOptions } from './xlsx-write.js';
