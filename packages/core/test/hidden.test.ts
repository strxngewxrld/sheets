import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { formatTable, hasHidden, hideCols, hideRows, readXlsx, setTableFilter, setValues, toRange, unhideCols, unhideRows, Workbook, writeXlsx } from '../src/index.js';

function book() {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [['Item', 'Value'], ['a', 1], ['b', 2], ['c', 4], ['=SUBTOTAL(9,B2:B4)', '=SUBTOTAL(109,B2:B4)'], ['=AGGREGATE(9,5,B2:B4)', '=AGGREGATE(9,4,B2:B4)']]);
  return wb;
}

describe('hidden rows and columns', () => {
  it('hides and unhides rows and columns as one undo step each', () => {
    const wb = book();
    const sheet = wb.activeSheet;
    sheet.rowHeights.set(5, 30);
    hideRows(wb, 0, 1, 2);
    expect(sheet.rowHeight(1)).toBe(0);
    expect(sheet.rowHeight(2)).toBe(0);
    expect(hasHidden(sheet, 'row', 0, 3)).toBe(true);
    unhideRows(wb, 0, 0, 10);
    expect(sheet.rowHeight(1)).toBe(sheet.defaultRowHeight);
    expect(sheet.rowHeight(5)).toBe(30);
    wb.undo();
    expect(sheet.rowHeight(2)).toBe(0);
    hideCols(wb, 0, 1);
    expect(sheet.colWidth(1)).toBe(0);
    unhideCols(wb, 0, 0, 3);
    expect(sheet.colWidth(1)).toBe(sheet.defaultColWidth);
    expect(hasHidden(sheet, 'col', 0, 3)).toBe(false);
  });

  it('lets SUBTOTAL 101-111 and AGGREGATE option 5 skip manually hidden rows', () => {
    const wb = book();
    expect([wb.getValue(0, 4, 0), wb.getValue(0, 4, 1), wb.getValue(0, 5, 0), wb.getValue(0, 5, 1)]).toEqual([7, 7, 7, 7]);
    hideRows(wb, 0, 2);
    expect([wb.getValue(0, 4, 0), wb.getValue(0, 4, 1), wb.getValue(0, 5, 0), wb.getValue(0, 5, 1)]).toEqual([7, 5, 5, 7]);
    wb.undo();
    expect(wb.getValue(0, 4, 1)).toBe(7);
  });

  it('always skips rows hidden by a table filter in SUBTOTAL', () => {
    const wb = book();
    formatTable(wb, 0, toRange('A1:B4'));
    setTableFilter(wb, 0, 'Table1', 0, ['a', 'c']);
    expect(wb.getValue(0, 4, 0)).toBe(5);
    expect(wb.getValue(0, 4, 1)).toBe(5);
    expect(wb.getValue(0, 5, 1)).toBe(7);
  });

  it('writes hidden rows and columns to XLSX and reads them back', () => {
    const wb = book();
    hideRows(wb, 0, 2);
    hideCols(wb, 0, 1);
    const bytes = writeXlsx(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toMatch(/<row r="3"[^>]*hidden="1"/);
    expect(xml).toMatch(/<col min="2" max="2"[^>]*hidden="1"/);
    const back = readXlsx(bytes);
    expect(back.activeSheet.rowHeight(2)).toBe(0);
    expect(back.activeSheet.colWidth(1)).toBe(0);
    expect(back.getValue(0, 4, 1)).toBe(5);
  });

  it('patches only visibility when unhiding in a preserved XLSX', () => {
    const wb = book();
    hideRows(wb, 0, 2);
    const original = readXlsx(writeXlsx(wb));
    unhideRows(original, 0, 0, 10);
    const xml = strFromU8(unzipSync(writeXlsx(original))['xl/worksheets/sheet1.xml']!);
    expect(xml).not.toMatch(/hidden="1"/);
  });
});
