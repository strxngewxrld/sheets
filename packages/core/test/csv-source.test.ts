import { describe, expect, it } from 'vitest';
import { decodeCsvBytes, encodeCsvText, openWorkbook, saveWorkbook, setCellInput, unencodableCsvChars, Workbook } from '../src/index.js';

const utf8 = (text: string) => new TextEncoder().encode(text);
// "Name;Wert\r\nMüller;1,5\r\nB;2" as Windows-1252, without a final line break.
const CP1252 = Uint8Array.from([0x4e, 0x61, 0x6d, 0x65, 0x3b, 0x57, 0x65, 0x72, 0x74, 0x0d, 0x0a, 0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x3b, 0x31, 0x2c, 0x35, 0x0d, 0x0a, 0x42, 0x3b, 0x32]);

describe('CSV files are written back the way they were read', () => {
  it('keeps delimiter, decimal mark, BOM, CRLF and the missing final line break after an edit', () => {
    const wb = openWorkbook(utf8('﻿Name;Wert\r\nMüller;1,5\r\nB;2'), 'a.csv');
    setCellInput(wb, 0, 2, 1, '3');
    expect(new TextDecoder('utf-8', { ignoreBOM: true }).decode(saveWorkbook(wb, 'csv'))).toBe('﻿Name;Wert\r\nMüller;1,5\r\nB;3');
  });

  it('reads Windows-1252 and writes it back byte for byte', () => {
    const wb = openWorkbook(CP1252, 'a.csv');
    expect(wb.getValue(0, 1, 0)).toBe('Müller');
    expect(wb.csvSource?.encoding).toBe('windows-1252');
    expect([...saveWorkbook(wb, 'csv')]).toEqual([...CP1252]);
    setCellInput(wb, 0, 2, 1, '3');
    const expected = [...CP1252];
    expected[expected.length - 1] = 0x33;
    expect([...saveWorkbook(wb, 'csv')]).toEqual(expected);
  });

  it('round-trips every Windows-1252 character', () => {
    const bytes = Uint8Array.from({ length: 256 - 32 }, (_, i) => i + 32);
    const { text, encoding } = decodeCsvBytes(bytes);
    expect(encoding).toBe('windows-1252');
    expect([...encodeCsvText(text, 'windows-1252')!]).toEqual([...bytes]);
  });

  it('falls back to UTF-8 with BOM instead of losing characters Windows-1252 cannot hold', () => {
    const wb = openWorkbook(CP1252, 'a.csv');
    setCellInput(wb, 0, 2, 0, '北京 €');
    const out = saveWorkbook(wb, 'csv');
    expect([...out.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    const text = new TextDecoder().decode(out);
    expect(text).toContain('北京 €');
    expect(text).toContain('Müller');
    expect(unencodableCsvChars('a€北ü京', 'windows-1252')).toEqual(['北', '京']);
  });

  it('saves the data sheet, not the active one, and follows renames and their undo', () => {
    const wb = openWorkbook(utf8('a,b\n1,2\n'), 'data.csv');
    wb.addSheet('Notes');
    wb.activeSheetIndex = 1;
    setCellInput(wb, 'Notes', 0, 0, 'scratch');
    expect(new TextDecoder().decode(saveWorkbook(wb, 'csv'))).toBe('a,b\n1,2\n');
    wb.apply({ type: 'renameSheet', from: 'data', to: 'Renamed' });
    expect(wb.csvSource?.sheet).toBe('Renamed');
    expect(new TextDecoder().decode(saveWorkbook(wb, 'csv'))).toBe('a,b\n1,2\n');
    wb.undo();
    expect(wb.csvSource?.sheet).toBe('data');
    expect(new TextDecoder().decode(saveWorkbook(wb, 'csv'))).toBe('a,b\n1,2\n');
  });

  it('keeps the CSV format through JSON (backups and snapshots)', () => {
    const wb = Workbook.fromJSON(JSON.parse(JSON.stringify(openWorkbook(CP1252, 'a.csv').toJSON())));
    expect([...saveWorkbook(wb, 'csv')]).toEqual([...CP1252]);
  });

  it('keeps TSV and lets explicit options override the source format', () => {
    const wb = openWorkbook(utf8('a\tb\r\n1\t2\r\n'), 'a.tsv');
    expect(new TextDecoder().decode(saveWorkbook(wb, 'tsv'))).toBe('a\tb\r\n1\t2\r\n');
    expect(new TextDecoder().decode(saveWorkbook(wb, 'csv', { csv: { lineEnding: '\n' } }))).toBe('a,b\n1,2\n');
  });

  it('writes new workbooks as before', () => {
    const wb = new Workbook();
    setCellInput(wb, 0, 0, 0, 'x');
    expect(new TextDecoder().decode(saveWorkbook(wb, 'csv'))).toBe('x\r\n');
  });
});
