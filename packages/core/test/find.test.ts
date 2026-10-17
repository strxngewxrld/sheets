import { describe, expect, it } from 'vitest';
import { findCells, replaceCells, setValues, toRange, Workbook } from '../src/index.js';

function book() {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [['Apple pie', 12, '=SUM(B1:B2)'], ['apple', 3, '=A1&"!"'], ['Banana', 'x$y', null]]);
  wb.addSheet('Other');
  setValues(wb, 'Other', 0, 0, [['apple tart']]);
  return wb;
}

describe('find', () => {
  it('searches values or formulas, optionally within a range', () => {
    const wb = book();
    expect(findCells(wb, 'apple').map((r) => `${r.sheet}!${r.address}`)).toEqual(['Sheet1!A1', 'Sheet1!A2', 'Sheet1!C2', 'Other!A1']);
    expect(findCells(wb, 'apple', { caseSensitive: true, sheet: 'Sheet1' }).map((r) => r.address)).toEqual(['A2']);
    expect(findCells(wb, 'sum', { inFormulas: true }).map((r) => r.address)).toEqual(['C1']);
    expect(findCells(wb, 'apple', { range: toRange('A2:C3') }).map((r) => r.address)).toEqual(['A2', 'C2']);
  });
});

describe('replace', () => {
  it('replaces in constants as one undoable step', () => {
    const wb = book();
    expect(replaceCells(wb, 'apple', 'Pear')).toBe(3);
    expect(wb.getValue(0, 0, 0)).toBe('Pear pie');
    expect(wb.getValue(0, 1, 0)).toBe('Pear');
    expect(wb.getValue('Other', 0, 0)).toBe('Pear tart');
    expect(wb.getValue(0, 1, 2)).toBe('Pear pie!');
    expect(wb.activeSheet.getCell(1, 2)?.formula).toBe('A1&"!"');
    wb.undo();
    expect(wb.getValue(0, 0, 0)).toBe('Apple pie');
    expect(wb.getValue('Other', 0, 0)).toBe('apple tart');
  });

  it('replaces in formulas only when searching formulas', () => {
    const wb = book();
    expect(replaceCells(wb, 'SUM', 'MAX', { inFormulas: true, sheet: 'Sheet1' })).toBe(1);
    expect(wb.activeSheet.getCell(0, 2)?.formula).toBe('MAX(B1:B2)');
    expect(wb.getValue(0, 0, 2)).toBe(12);
  });

  it('honours whole-cell, case, selected cells and literal dollar signs', () => {
    const wb = book();
    expect(replaceCells(wb, 'apple', 'Kiwi', { wholeCell: true, caseSensitive: true })).toBe(1);
    expect(wb.getValue(0, 1, 0)).toBe('Kiwi');
    expect(replaceCells(wb, '$', '$1', { sheet: 'Sheet1' })).toBe(1);
    expect(wb.getValue(0, 2, 1)).toBe('x$1y');
    expect(replaceCells(wb, 'a', '_', { cells: [{ sheet: 'Sheet1', row: 2, col: 0 }] })).toBe(1);
    expect(wb.getValue(0, 2, 0)).toBe('B_n_n_');
    expect(replaceCells(wb, '12', '15')).toBe(1);
    expect(wb.getValue(0, 0, 1)).toBe(15);
    expect(wb.getValue(0, 0, 2)).toBe(18);
  });

  it('supports regular expressions with groups', () => {
    const wb = book();
    expect(replaceCells(wb, '(\\w+) pie', '$1 cake', { regex: true })).toBe(1);
    expect(wb.getValue(0, 0, 0)).toBe('Apple cake');
  });
});
