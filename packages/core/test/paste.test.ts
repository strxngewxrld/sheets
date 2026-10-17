import { describe, expect, it } from 'vitest';
import { applyStyle, copyRange, pasteRange, setValues, toRange, Workbook } from '../src/index.js';

function values(wb: Workbook, range: string) {
  const { start, end } = toRange(range);
  return Array.from({ length: end.row - start.row + 1 }, (_, r) =>
    Array.from({ length: end.col - start.col + 1 }, (_, c) => wb.getValue(0, start.row + r, start.col + c)));
}

describe('paste special', () => {
  it('transposes values and shifts formulas by each cell\'s new position', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, 2, 3], [10, 20, '=C1*2']]);
    const clip = copyRange(wb, 0, toRange('A1:C2'));
    const pasted = pasteRange(wb, clip, 0, toRange('E1'), { transpose: true });
    expect(pasted).toEqual(toRange('E1:F3'));
    expect(values(wb, 'E1:F3')).toEqual([[1, 10], [2, 20], [3, 6]]);
    expect(wb.activeSheet.getCell(2, 5)?.formula).toBe('E3*2');
  });

  it('skips blank source cells and keeps the target content', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, null, 3]]);
    setValues(wb, 0, 2, 0, [['a', 'b', 'c']]);
    pasteRange(wb, copyRange(wb, 0, toRange('A1:C1')), 0, toRange('A3'), { skipBlanks: true });
    expect(values(wb, 'A3:C3')).toEqual([[1, 'b', 3]]);
    wb.undo();
    pasteRange(wb, copyRange(wb, 0, toRange('A1:C1')), 0, toRange('A3'));
    expect(values(wb, 'A3:C3')).toEqual([[1, null, 3]]);
  });

  it('combines modes with transpose', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[2, '=A1*3']]);
    applyStyle(wb, 0, toRange('A1:B1'), { bold: true });
    pasteRange(wb, copyRange(wb, 0, toRange('A1:B1')), 0, toRange('D1'), { mode: 'values', transpose: true });
    expect(values(wb, 'D1:D2')).toEqual([[2], [6]]);
    expect(wb.activeSheet.getCell(1, 3)?.formula).toBeUndefined();
    expect(wb.activeSheet.getCell(1, 3)?.style).toBeUndefined();
    pasteRange(wb, copyRange(wb, 0, toRange('A1:B1')), 0, toRange('F1'), 'formats');
    expect(wb.activeSheet.getCell(0, 6)?.style?.bold).toBe(true);
    expect(wb.getValue(0, 0, 6)).toBe(null);
  });
});
