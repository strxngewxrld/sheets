import { describe, expect, it } from 'vitest';
import {
  applyStyle,
  formatValue,
  insertChartFromRange,
  LOCALE_DE,
  parseInput,
  readCsv,
  readXlsx,
  setValues,
  toRange,
  Workbook,
  writeCsv,
  writeXlsx,
} from '../src/index.js';

describe('number formats', () => {
  it('formats numbers, percents, dates', () => {
    expect(formatValue(1234.567, '#,##0.00')).toBe('1,234.57');
    expect(formatValue(-1234.5, '#,##0.00')).toBe('-1,234.50');
    expect(formatValue(0.256, '0.0%')).toBe('25.6%');
    expect(formatValue(45322, 'dd.mm.yyyy')).toBe('31.01.2024');
    expect(formatValue(0.5, 'hh:mm')).toBe('12:00');
    expect(formatValue(1234.5, '#,##0.00 "€"', LOCALE_DE)).toBe('1.234,50 €');
    expect(formatValue(-5, '0;(0)')).toBe('(5)');
    expect(formatValue(12345678, '0.00E+00')).toBe('1.23E+07');
    expect(formatValue(0.1 + 0.2)).toBe('0.3');
  });

  it('parses input', () => {
    expect(parseInput('12%')).toEqual({ value: 0.12, numFmt: '0%' });
    expect(parseInput('1,234.5').value).toBe(1234.5);
    expect(parseInput('1.234,5', LOCALE_DE).value).toBe(1234.5);
    expect(parseInput('2024-01-31')).toEqual({ value: 45322, numFmt: 'yyyy-mm-dd' });
    expect(parseInput('=A1+1')).toEqual({ value: null, formula: 'A1+1' });
    expect(parseInput('hello').value).toBe('hello');
  });
});

describe('csv', () => {
  it('detects delimiter and types', () => {
    const wb = readCsv('name;amount\n"Müller, A";1.234,50\nB;2\n');
    const s = wb.requireSheet(0);
    expect(s.getValue(1, 0)).toBe('Müller, A');
    expect(s.getValue(1, 1)).toBe(1234.5);
    expect(writeCsv(wb, 0, { delimiter: ',' })).toBe('name,amount\r\n"Müller, A",1234.5\r\nB,2\r\n');
  });
});

describe('xlsx roundtrip', () => {
  it('preserves values, formulas, styles, merges and charts', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [
      ['Month', 'Revenue', 'Cost'],
      ['Jan', 100, 60],
      ['Feb', 150, 80],
      ['Mar', 120, 70],
      ['Total', '=SUM(B2:B4)', '=SUM(C2:C4)'],
      ['Flag', true, 'x & <y>'],
    ]);
    applyStyle(wb, 0, toRange('A1:C1'), { bold: true, fill: '#DDEEFF', hAlign: 'center' });
    applyStyle(wb, 0, toRange('B2:C5'), { numFmt: '#,##0.00 "€"' });
    wb.apply({ type: 'setColWidths', sheet: 'Sheet1', sizes: [{ index: 0, size: 120 }] });
    wb.apply({ type: 'setMerges', sheet: 'Sheet1', merges: ['E1:F2'] });
    wb.apply({ type: 'setFrozen', sheet: 'Sheet1', rows: 1, cols: 0 });
    wb.addSheet('Second');
    setValues(wb, 'Second', 0, 0, [['=Sheet1!B5*2', '=XLOOKUP("Feb",Sheet1!A2:A4,Sheet1!B2:B4)']]);
    insertChartFromRange(wb, 0, toRange('A1:C4'), 'column', { title: 'Revenue vs Cost' });

    const bytes = writeXlsx(wb);
    const back = readXlsx(bytes);
    const s = back.requireSheet('Sheet1');
    expect(s.getValue(4, 1)).toBe(370);
    expect(s.getCell(4, 1)!.formula).toBe('SUM(B2:B4)');
    expect(s.getValue(5, 1)).toBe(true);
    expect(s.getValue(5, 2)).toBe('x & <y>');
    expect(s.getCell(0, 0)!.style).toMatchObject({ bold: true, fill: '#DDEEFF', hAlign: 'center' });
    expect(s.getCell(1, 1)!.style?.numFmt).toBe('#,##0.00 "€"');
    expect(s.colWidth(0)).toBe(120);
    expect(s.merges).toHaveLength(1);
    expect(s.frozen).toEqual({ rows: 1, cols: 0 });
    expect(back.getValue('Second', 0, 0)).toBe(740);
    expect(back.getValue('Second', 0, 1)).toBe(150);
    expect(s.charts).toHaveLength(1);
    expect(s.charts[0]).toMatchObject({ type: 'column', title: 'Revenue vs Cost' });
    expect(s.charts[0]!.series).toHaveLength(2);
    expect(s.charts[0]!.series[0]!.values).toBe('Sheet1!$B$2:$B$4');
  });
});

describe('macro-enabled workbooks', () => {
  it('keeps the VBA project and content types when editing an .xlsm file', async () => {
    const { strFromU8, strToU8, unzipSync, zipSync } = await import('fflate');
    const { formatFromFileName, isXlsxFormat, openWorkbook, saveWorkbook, setValues, Workbook } = await import('../src/index.js');
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Macro', 1]]);
    const parts = unzipSync(saveWorkbook(wb, 'xlsm'));
    expect(strFromU8(parts['[Content_Types].xml']!)).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml');
    // Simulate an Excel macro project.
    parts['xl/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]);
    parts['[Content_Types].xml'] = strToU8(strFromU8(parts['[Content_Types].xml']!).replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'));
    const original = zipSync(parts);
    expect(formatFromFileName('Budget.XLSM')).toBe('xlsm');
    expect(isXlsxFormat('xlsm')).toBe(true);
    const opened = openWorkbook(original, 'Budget.xlsm');
    expect(saveWorkbook(opened, 'xlsm')).toEqual(original);
    setValues(opened, 0, 0, 1, [[2]]);
    const saved = unzipSync(saveWorkbook(opened, 'xlsm'));
    expect(saved['xl/vbaProject.bin']).toEqual(parts['xl/vbaProject.bin']);
    expect(strFromU8(saved['[Content_Types].xml']!)).toContain('vbaProject');
    expect(openWorkbook(zipSync(saved), 'Budget.xlsm').getValue(0, 0, 1)).toBe(2);
  });
});
