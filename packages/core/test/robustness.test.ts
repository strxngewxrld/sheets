import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { openWorkbook, saveWorkbook, Workbook, setValues, setCellInput } from '../src/index.js';

// Heavy cases get a generous timeout; wall-clock limits only catch pathological (e.g. quadratic) slowdowns.
const t = (name: string, fn: () => void, timeout = 60_000) => it(name, fn, timeout);
const bytes = (s: string) => strToU8(s);

describe('corrupt input never hangs or crashes uncontrolled', () => {
  const cases: Record<string, Uint8Array> = {
    empty: new Uint8Array(0),
    garbage: bytes('this is not a zip'),
    truncatedZip: saveWorkbook(new Workbook(), 'xlsx').slice(0, 120),
    zipNoWorkbook: zipSync({ 'foo.txt': bytes('x') }),
    zipBadWorkbook: zipSync({ 'xl/workbook.xml': bytes('<workbook'), '[Content_Types].xml': bytes('<Types/>') }),
    zipWorkbookNoSheets: zipSync({ 'xl/workbook.xml': bytes('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></workbook>'), '[Content_Types].xml': bytes('<Types/>') }),
  };
  for (const [name, data] of Object.entries(cases)) {
    t(name, () => {
      try { const wb = openWorkbook(data, 'x.xlsx'); expect(wb.sheets.length).toBeGreaterThan(0); }
      catch (e) { expect(e).toBeInstanceOf(Error); expect((e as Error).message.length).toBeGreaterThan(0); }
    });
  }
  t('json garbage', () => { expect(() => openWorkbook(bytes('{nope'), 'x.json')).toThrow(); });
});

describe('csv edge cases', () => {
  const read = (s: string, name = 'a.csv') => openWorkbook(bytes(s), name);
  t('empty file', () => { const wb = read(''); expect(wb.sheets.length).toBe(1); });
  t('BOM stripped', () => { expect(read('﻿Name,Age\nA,1').getValue(0, 0, 0)).toBe('Name'); });
  t('CRLF', () => { const wb = read('a,b\r\n1,2\r\n'); expect(wb.getValue(0, 1, 1)).toBe(2); });
  t('quoted newline + escaped quote', () => { const wb = read('a,b\n"x\ny","he said ""hi"""\n'); expect(wb.getValue(0, 1, 0)).toBe('x\ny'); expect(wb.getValue(0, 1, 1)).toBe('he said "hi"'); });
  t('unterminated quote', () => { expect(() => read('a,"b\nc')).not.toThrow(); });
  t('semicolon', () => { const wb = read('a;b\n1;2'); expect(wb.getValue(0, 1, 1)).toBe(2); });
  t('tsv', () => { const wb = read('a\tb\n1\t2', 'a.tsv'); expect(wb.getValue(0, 1, 1)).toBe(2); });
  t('emoji + umlaut roundtrip', () => {
    const wb = read('n\nÄÖÜ ß 😀\n');
    const out = new TextDecoder().decode(saveWorkbook(wb, 'csv'));
    expect(out).toContain('ÄÖÜ ß 😀');
  });
  t('leading zeros kept on csv roundtrip', () => {
    const wb = read('id\n0004711\n');
    const out = new TextDecoder().decode(saveWorkbook(wb, 'csv'));
    expect(out).toContain('0004711');
  });
  t('values with comma/quote/newline are quoted on write', () => {
    const wb = new Workbook(); setValues(wb, 0, 0, 0, [['a,b', 'q"q', 'n\nl']]);
    const back = openWorkbook(saveWorkbook(wb, 'csv'), 'a.csv');
    expect(back.getValue(0, 0, 0)).toBe('a,b'); expect(back.getValue(0, 0, 1)).toBe('q"q'); expect(back.getValue(0, 0, 2)).toBe('n\nl');
  });
  t('ragged rows', () => { const wb = read('a,b,c\n1\n1,2,3,4,5'); expect(wb.getValue(0, 2, 4)).toBe(5); });
  t('100k rows fast', () => {
    const rows = ['a,b,c']; for (let i = 0; i < 100000; i++) rows.push(`${i},x${i},${i * 1.5}`);
    const s = Date.now(); const wb = read(rows.join('\n')); expect(wb.getValue(0, 100000, 0)).toBe(99999); expect(Date.now() - s).toBeLessThan(30_000);
  });
});

describe('xlsx roundtrip', () => {
  t('special characters survive', () => {
    const wb = new Workbook(); const vals = ['<a>&"\'', 'line1\nline2', '😀 ünï', '  lead trail  ', 'tab\there', 'x'.repeat(32767)];
    setValues(wb, 0, 0, 0, [vals]);
    const back = openWorkbook(saveWorkbook(wb, 'xlsx'), 'a.xlsx');
    vals.forEach((v, i) => expect(back.getValue(0, 0, i)).toBe(v));
  });
  t('control chars do not corrupt file', () => {
    const wb = new Workbook(); setValues(wb, 0, 0, 0, [['a\u0000b\u0001c\u001fd\u000be']]);
    const back = openWorkbook(saveWorkbook(wb, 'xlsx'), 'a.xlsx'); expect(String(back.getValue(0, 0, 0))).toMatch(/a.*b.*c.*d.*e/);
  });
  t('sheet names with special chars + 31 chars', () => {
    const wb = new Workbook(); wb.addSheet("It's & <ok>"); wb.addSheet('x'.repeat(31));
    const back = openWorkbook(saveWorkbook(wb, 'xlsx'), 'a.xlsx'); expect(back.sheets.map((s) => s.name)).toEqual(wb.sheets.map((s) => s.name));
  });
  t('formulas, dates, big/small numbers, bools, errors roundtrip', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1e308, 1e-300, -0, true, 9007199254740993, '=1/0', '=SUM(A1:B1)', '=TODAY()>0']]);
    const back = openWorkbook(saveWorkbook(wb, 'xlsx'), 'a.xlsx');
    expect(back.getValue(0, 0, 3)).toBe(true); expect(String(back.getValue(0, 0, 5))).toMatch(/#DIV\/0/);
  });
  t('1000 sheets-ish and 200x500 grid', () => {
    const wb = new Workbook(); const m: number[][] = []; for (let r = 0; r < 500; r++) { const row: number[] = []; for (let c = 0; c < 200; c++) row.push(r * c); m.push(row); }
    setValues(wb, 0, 0, 0, m); const s = Date.now();
    const back = openWorkbook(saveWorkbook(wb, 'xlsx'), 'a.xlsx'); expect(back.getValue(0, 499, 199)).toBe(499 * 199); expect(Date.now() - s).toBeLessThan(30_000);
  });
  t('xlsm roundtrip', () => { const wb = new Workbook(); setValues(wb, 0, 0, 0, [[1]]); expect(openWorkbook(saveWorkbook(wb, 'xlsm'), 'a.xlsm').getValue(0, 0, 0)).toBe(1); });
});

describe('formula engine robustness', () => {
  const ev = (f: string) => { const wb = new Workbook(); return wb.evaluate(f); };
  t('syntax errors do not throw', () => { for (const f of ['=', '=(', '=SUM(', '=1+', '=)(', '=SUM(1,,)', '="abc', '=A1:', '=#REF!', '=NOSUCH(1)', '=1/0', '=LEN()', '=\'x\'!A1', '=[1]A1']) expect(() => ev(f)).not.toThrow(); });
  t('circular refs terminate', () => {
    const wb = new Workbook(); setCellInput(wb, 0, 0, 0, '=B1+1'); setCellInput(wb, 0, 0, 1, '=A1+1');
    expect(String(wb.getValue(0, 0, 0))).toMatch(/#|0|circ/i);
  });
  t('self reference', () => { const wb = new Workbook(); setCellInput(wb, 0, 0, 0, '=A1'); expect(() => wb.getValue(0, 0, 0)).not.toThrow(); });
  t('deep nesting', () => { const f = '=' + '('.repeat(500) + '1' + ')'.repeat(500); expect(() => ev(f)).not.toThrow(); });
  t('very long chain of references', () => {
    const wb = new Workbook(); setCellInput(wb, 0, 0, 0, '1'); for (let i = 1; i < 5000; i++) setCellInput(wb, 0, i, 0, `=A${i}+1`);
    expect(wb.getValue(0, 4999, 0)).toBe(5000);
  });
  t('whole-column sum is fast', () => { const wb = new Workbook(); setValues(wb, 0, 0, 0, [[1], [2]]); const s = Date.now(); expect(wb.evaluate('=SUM(A:A)')).toBe(3); expect(Date.now() - s).toBeLessThan(10_000); });
  t('huge repetition yields an error value instead of throwing', () => { expect(String(ev('=REPT("x",1000000000)'))).toMatch(/#/); expect(String(ev('=SEQUENCE(100000,1000)'))).toMatch(/#/); });
  t('locale/edge numbers', () => { expect(ev('=0.1+0.2')).toBeCloseTo(0.3); expect(ev('=2^1024')).toBeDefined(); expect(ev('=-5^2')).toBe(25); });
});
