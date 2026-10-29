import { describe, expect, it, vi } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { MAX_CSV_BYTES, MAX_CSV_CELLS, openWorkbook, parseCsv, saveWorkbook, Workbook, writeCsv, writeXlsx } from '../src/index.js';

const workbook = (values: (string | number)[]) => {
  const wb = new Workbook();
  wb.load({ sheets: [{ name: 'Sheet1', cells: Object.fromEntries(values.map((v, r) => [`A${r + 1}`, { v }])) }] });
  return wb;
};

describe('CSV resource budgets', () => {
  it('rejects a tiny imported XLSX with a distant cell before visiting the output rectangle', () => {
    const parts = unzipSync(writeXlsx(new Workbook()));
    const part = 'xl/worksheets/sheet1.xml';
    parts[part] = strToU8(strFromU8(parts[part]!).replace(/<sheetData>[\s\S]*?<\/sheetData>/,
      '<sheetData><row r="1048576"><c r="XFD1048576" t="inlineStr"><is><t>x</t></is></c></row></sheetData>'));
    const bytes = zipSync(parts);
    expect(bytes.length).toBeLessThan(4000);
    const wb = openWorkbook(bytes, 'sparse.xlsx');
    expect(wb.activeSheet.cells.size).toBe(1);
    const visit = vi.spyOn(wb.activeSheet, 'getCell').mockImplementation(() => { throw new Error('Output iteration started'); });
    expect(() => writeCsv(wb)).toThrow(/cell limit/);
    expect(() => saveWorkbook(wb, 'tsv')).toThrow(/cell limit/);
    expect(visit).not.toHaveBeenCalled();
  });

  it('allows the cell boundary and rejects one more leading empty cell', () => {
    const wb = new Workbook();
    wb.load({ sheets: [{ name: 'Sheet1', cells: { [`A${MAX_CSV_CELLS}`]: { v: 'x' } } }] });
    expect(writeCsv(wb).length).toBe(MAX_CSV_CELLS * 2 + 1);
    wb.load({ sheets: [{ name: 'Sheet1', cells: { [`A${MAX_CSV_CELLS + 1}`]: { v: 'x' } } }] });
    expect(() => writeCsv(wb)).toThrow(/cell limit/);
  });

  it('counts UTF-8, quote escaping, BOM and record separators in the byte budget', () => {
    // Raw text fits, but doubled quotes do not.
    expect(() => writeCsv(workbook(['"'.repeat(MAX_CSV_BYTES / 2)]))).toThrow(/byte limit/);
    // UTF-16 string length alone undercounts UTF-8 output.
    expect(() => writeCsv(workbook(['é'.repeat(MAX_CSV_BYTES / 2)]))).toThrow(/byte limit/);
    const wb = workbook(['😀'.repeat((MAX_CSV_BYTES - 4) / 4)]);
    expect(new TextEncoder().encode(writeCsv(wb, 0, { finalNewline: false })).length).toBe(MAX_CSV_BYTES - 4);
    expect(() => writeCsv(wb, 0, { bom: true })).toThrow(/byte limit/);
  });

  it('bounds cumulative output and retained original records', () => {
    const value = 'x'.repeat(MAX_CSV_BYTES / 2);
    expect(() => writeCsv(workbook([value, value]))).toThrow(/byte limit/);
    expect(() => writeCsv(workbook([value]), 0, { original: [value.repeat(3)], escapeFormulas: false })).toThrow(/byte limit/);
    const wb = new Workbook();
    expect(() => writeCsv(wb, 0, { original: Array(MAX_CSV_CELLS + 1).fill(''), escapeFormulas: false })).toThrow(/cell limit/);
  });
});

describe('CSV formula text protection', () => {
  it.each([',', ';', '\t'])('protects untrusted text, including leading whitespace and embedded delimiters (%j)', (delimiter) => {
    const dangerous = ['=1+1', '+SUM(A1)', '-1+2', '@SUM(A1)', '\t=1', '\r=1', '\n=1', '  =1', '\u0000=1', '＝1', '＋1', '－1', '＠x', `=x${delimiter}"\nnext`];
    const wb = workbook([...dangerous, -12, 'ordinary text']);
    const original = wb.toJSON();
    const rows = parseCsv(writeCsv(wb, 0, { delimiter, escapeFormulas: true }), delimiter);
    expect(rows.map((row) => row[0])).toEqual([...dangerous.map((text) => `\t${text}`), '-12', 'ordinary text']);
    expect(wb.toJSON()).toEqual(original);
    expect(writeCsv(workbook(['=1+1']), 0, { escapeFormulas: true })).toBe('"\t=1+1"\r\n');
  });

  it('does not reuse unprotected original records in the protected mode', () => {
    const wb = openWorkbook(strToU8('=1+1\n'), 'source.csv');
    expect(strFromU8(saveWorkbook(wb, 'csv'))).toBe('=1+1\n');
    expect(strFromU8(saveWorkbook(wb, 'csv', { csv: { escapeFormulas: true } }))).toBe('"\t=1+1"\n');
  });

  it('protects new CSV conversions and supports explicit raw data exchange', () => {
    const wb = workbook(['=1+1']);
    expect(writeCsv(wb)).toBe('"\t=1+1"\r\n');
    expect(writeCsv(wb, 0, { escapeFormulas: false })).toBe('=1+1\r\n');
    expect(strFromU8(saveWorkbook(wb, 'csv'))).toBe('"\t=1+1"\r\n');
    expect(strFromU8(saveWorkbook(wb, 'csv', { csv: { escapeFormulas: false } }))).toBe('=1+1\r\n');
  });

  it('protects formula text and string formula results without changing numeric results', () => {
    const wb = new Workbook();
    wb.load({ sheets: [{ name: 'Sheet1', cells: { A1: { f: '"=1+1"' }, A2: { f: '-12' } } }] });
    expect(parseCsv(writeCsv(wb, 0, { escapeFormulas: true }))).toEqual([['\t=1+1'], ['-12']]);
    expect(parseCsv(writeCsv(wb, 0, { escapeFormulas: true, formulas: true }))).toEqual([['\t="=1+1"'], ['\t=-12']]);
  });

  it('protects formula-like display text supplied by a numeric cell format', () => {
    const wb = new Workbook();
    wb.load({ sheets: [{ name: 'Sheet1', cells: { A1: { v: 12, s: { numFmt: '"=1+1"' } } } }] });
    expect(parseCsv(writeCsv(wb, 0, { formatted: true }))).toEqual([['\t=1+1']]);
    expect(parseCsv(writeCsv(wb))).toEqual([['12']]);
  });
});
