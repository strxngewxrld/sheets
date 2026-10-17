import { describe, expect, it } from 'vitest';
import { formatTable, setValues, toRange, Workbook } from '@sheets/core';
import { currentRegion, resolveSortContext } from '../src/sheet-actions.js';

describe('sheet sort selection', () => {
  it('sorts complete records when the user selects one text column', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'City'], ['Zoe', 'Berlin'], ['Anna', 'Hamburg']]);
    const context = resolveSortContext(wb.activeSheet, toRange('A2:A3'), { row: 1, col: 0 });
    expect(context).toMatchObject({ range: toRange('A1:B3'), hasHeader: true });
  });

  it('honors the explicit header flag on formatted tables even for all-text data', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Zoe', 'Berlin'], ['Anna', 'Hamburg']]);
    formatTable(wb, 0, toRange('A1:B2'), { headerRow: false });
    expect(resolveSortContext(wb.activeSheet, toRange('B1'), { row: 0, col: 1 })).toMatchObject({ range: toRange('A1:B2'), hasHeader: false });
  });

  it('finds a table below blank cells when a whole column is selected', () => {
    const wb = new Workbook();
    setValues(wb, 0, 3, 0, [['Name', 'Value'], ['B', 2], ['A', 1]]);
    formatTable(wb, 0, toRange('A4:B6'));
    expect(resolveSortContext(wb.activeSheet, toRange('B:B'), { row: 0, col: 1 })).toMatchObject({ range: toRange('A4:B6'), hasHeader: true });
  });

  it('keeps explicit rectangular selections and respects worksheet boundaries', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, 2], [3, 4], [5, 6]]);
    expect(resolveSortContext(wb.activeSheet, toRange('A2:B3'), { row: 1, col: 0 })).toMatchObject({ range: toRange('A2:B3'), hasHeader: false });
    setValues(wb, 0, 1048575, 16383, [[1]]);
    expect(currentRegion(wb.activeSheet, { row: 1048575, col: 16383 })).toEqual(toRange('XFD1048576'));
  });
});
