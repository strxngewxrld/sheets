import { describe, expect, it } from 'vitest';
import { applyStyle, clearRange, copyRange, pasteRange, readXlsx, setValues, sortRange, toRange, Workbook, writeXlsx } from '../src/index.js';
import { strFromU8, unzipSync } from 'fflate';

function values(wb: Workbook, range: string) {
  const { start, end } = toRange(range);
  return Array.from({ length: end.row - start.row + 1 }, (_, r) =>
    Array.from({ length: end.col - start.col + 1 }, (_, c) => wb.getValue(0, start.row + r, start.col + c)));
}

describe('dynamic arrays', () => {
  it('spills sequences, literals, references and existing array functions', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=SEQUENCE(2,3,10,-2)']]);
    expect(values(wb, 'A1:C2')).toEqual([[10, 8, 6], [4, 2, 0]]);
    setValues(wb, 0, 3, 0, [['=TRANSPOSE(A1:C2)']]);
    expect(values(wb, 'A4:B6')).toEqual([[10, 4], [8, 2], [6, 0]]);
    setValues(wb, 0, 0, 4, [['={1,2;3,4}']]);
    expect(values(wb, 'E1:F2')).toEqual([[1, 2], [3, 4]]);
    setValues(wb, 0, 3, 4, [['=E1:F2*2']]);
    expect(values(wb, 'E4:F5')).toEqual([[2, 4], [6, 8]]);
  });

  it('filters rows and columns, sorts and removes duplicates', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['b', 2], ['A', 1], ['b', 2], ['c', 3]]);
    setValues(wb, 0, 0, 3, [['=SORT(UNIQUE(FILTER(A1:B4,B1:B4>1)),2,-1)']]);
    expect(values(wb, 'D1:E2')).toEqual([['c', 3], ['b', 2]]);
    setValues(wb, 0, 0, 6, [['=FILTER(A1:B4,{FALSE,TRUE})']]);
    expect(values(wb, 'G1:G4')).toEqual([[2], [1], [2], [3]]);
    setValues(wb, 0, 0, 8, [['=UNIQUE(A1:B4,FALSE,TRUE)']]);
    expect(values(wb, 'I1:J2')).toEqual([['A', 1], ['c', 3]]);
    setValues(wb, 0, 6, 0, [['=SORT({3,1,2;30,10,20},1,1,TRUE)']]);
    expect(values(wb, 'A7:C8')).toEqual([[1, 2, 3], [10, 20, 30]]);
    setValues(wb, 0, 10, 0, [['=UNIQUE({1,1,2;3,3,4},TRUE)']]);
    expect(values(wb, 'A11:B12')).toEqual([[1, 2], [3, 4]]);
  });

  it('updates ordinary and spill references across sheets, independent of formula order', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=D3*2', '=SUM(D1#)', '=SUM(D2:D5)', '=SEQUENCE(F1) ', null, 3]]);
    expect(values(wb, 'A1:C1')).toEqual([[6, 6, 5]]);
    wb.addSheet('Results');
    setValues(wb, 'Results', 0, 0, [['=SUM(Sheet1!D1#)']]);
    setValues(wb, 0, 0, 5, [[5]]);
    expect(values(wb, 'A1:C1')).toEqual([[6, 15, 14]]);
    expect(wb.getValue('Results', 0, 0)).toBe(15);
    setValues(wb, 0, 0, 5, [[2]]);
    expect(values(wb, 'D1:D5')).toEqual([[1], [2], [null], [null], [null]]);
    expect(values(wb, 'A1:C1')).toEqual([[0, 3, 2]]);
    wb.undo();
    expect(wb.getValue('Results', 0, 0)).toBe(15);
  });

  it('rejects partial content writes atomically but permits formatting', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=SEQUENCE(3)']]);
    expect(() => setValues(wb, 0, 1, 0, [[99]])).toThrow(/part of a spilled array/);
    expect(() => setValues(wb, 0, 1, 0, [[99, 'unchanged']])).toThrow();
    expect(wb.getValue(0, 1, 1)).toBe(null);
    applyStyle(wb, 0, toRange('A2:A3'), { fill: '#FF0000' });
    expect(values(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
    expect(wb.activeSheet.getCell(1, 0)?.style?.fill).toBe('#FF0000');
    wb.undo();
    expect(wb.activeSheet.getCell(1, 0)?.style?.fill).toBeUndefined();
    clearRange(wb, 0, toRange('A1:A3'));
    expect(values(wb, 'A1:A3')).toEqual([[null], [null], [null]]);
    wb.undo();
    expect(values(wb, 'A1:A3')).toEqual([[1], [2], [3]]);
  });

  it('recovers from blocked, merged and out-of-bounds output', () => {
    const wb = new Workbook();
    setValues(wb, 0, 1, 0, [['blocked']]);
    setValues(wb, 0, 0, 0, [['=SEQUENCE(3)']]);
    expect(wb.getValue(0, 0, 0)).toMatchObject({ code: '#SPILL!' });
    setValues(wb, 0, 1, 0, [[null]]);
    expect(wb.getValue(0, 2, 0)).toBe(3);
    wb.apply({ type: 'setMerges', sheet: 'Sheet1', merges: ['A2:B2'] });
    expect(wb.getValue(0, 0, 0)).toMatchObject({ code: '#SPILL!' });
    expect(wb.getValue(0, 2, 0)).toBe(null);
    wb.undo();
    expect(wb.getValue(0, 2, 0)).toBe(3);
    setValues(wb, 0, 1048575, 1, [['=SEQUENCE(2)']]);
    expect(wb.getValue(0, 1048575, 1)).toMatchObject({ code: '#SPILL!' });
  });

  it('handles empty filters, invalid masks, circular spill references and array IF', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=FILTER({1;2},{FALSE;FALSE})', '=FILTER({1;2},{FALSE;FALSE},"none")']]);
    expect(wb.getValue(0, 0, 0)).toMatchObject({ code: '#CALC!' });
    expect(wb.getValue(0, 0, 1)).toBe('none');
    expect(wb.evaluate('=FILTER({1;2},{TRUE,FALSE})')).toMatchObject({ code: '#VALUE!' });
    expect(wb.evaluate('=FILTER({1;2},{TRUE;#N/A})')).toMatchObject({ code: '#N/A' });
    setValues(wb, 0, 3, 0, [['=IF(SEQUENCE(3)>1,SEQUENCE(3)*10,0)']]);
    expect(values(wb, 'A4:A6')).toEqual([[0], [20], [30]]);
    setValues(wb, 0, 0, 3, [['=D1#']]);
    expect(wb.getValue(0, 0, 3)).toMatchObject({ code: '#CIRC!' });
    expect(wb.evaluate('=SUM(Z1#)')).toMatchObject({ code: '#REF!' });
  });

  it('regenerates derived cells after JSON restoration and structural changes', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=SEQUENCE(3)']]);
    expect(wb.toJSON().sheets[0]!.cells.A2).toBeUndefined();
    const restored = Workbook.fromJSON(wb.toJSON());
    expect(values(restored, 'A1:A3')).toEqual([[1], [2], [3]]);
    wb.apply({ type: 'insertRows', sheet: 'Sheet1', index: 0, count: 1 });
    expect(values(wb, 'A1:A4')).toEqual([[null], [1], [2], [3]]);
    wb.undo();
    expect(values(wb, 'A1:A4')).toEqual([[1], [2], [3], [null]]);
  });

  it('roundtrips array metadata and spill references through XLSX without blocking cached output', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=SEQUENCE(3,2)', null, '=SUM(A1#)']]);
    const bytes = writeXlsx(wb);
    const files = unzipSync(bytes);
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('t="array" ref="A1:B3"');
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('_xlfn.ANCHORARRAY(A1)');
    expect(files['xl/metadata.xml']).toBeDefined();
    const cached = readXlsx(bytes, { recalculate: false });
    expect(values(cached, 'A1:B3')).toEqual([[1, 2], [3, 4], [5, 6]]);
    const back = readXlsx(bytes);
    expect(values(back, 'A1:B3')).toEqual([[1, 2], [3, 4], [5, 6]]);
    expect(back.getValue(0, 0, 2)).toBe(21);
    setValues(back, 0, 0, 0, [['=SEQUENCE(2,2)']]);
    expect(values(back, 'A1:B3')).toEqual([[1, 2], [3, 4], [null, null]]);
    expect(back.getValue(0, 0, 2)).toBe(10);
  });

  it('copies complete arrays as formulas and permits pasting their values separately', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=SEQUENCE(3)']]);
    const clip = copyRange(wb, 0, toRange('A1:A3'));
    pasteRange(wb, clip, 0, toRange('C1'));
    expect(values(wb, 'C1:C3')).toEqual([[1], [2], [3]]);
    pasteRange(wb, clip, 0, toRange('E1'), 'values');
    expect(values(wb, 'E1:E3')).toEqual([[1], [2], [3]]);
    expect(wb.activeSheet.getCell(0, 4)?.formula).toBeUndefined();
    expect(() => sortRange(wb, 0, toRange('A1:A3'), [{ col: 0 }])).toThrow(/SORT formula/);
  });
});

describe('array shaping functions', () => {
  function spill(formula: string, setup?: (wb: Workbook) => void) {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, 2, 3], [4, 5, 6], [7, 8, 9]]);
    setup?.(wb);
    setValues(wb, 0, 10, 0, [[formula]]);
    const range = wb.activeSheet.getCell(10, 0)?.spillRange;
    if (!range) return wb.getValue(0, 10, 0);
    return Array.from({ length: range.end.row - range.start.row + 1 }, (_, r) =>
      Array.from({ length: range.end.col - range.start.col + 1 }, (_, c) => wb.getValue(0, range.start.row + r, range.start.col + c)));
  }
  const NA = expect.objectContaining({ code: '#N/A' });

  it('stacks arrays and pads with #N/A', () => {
    expect(spill('=VSTACK({1,2}, {3})')).toEqual([[1, 2], [3, NA]]);
    expect(spill('=HSTACK({1;2}, {3})')).toEqual([[1, 3], [2, NA]]);
    expect(spill('=VSTACK(A1:C1, A3:C3)')).toEqual([[1, 2, 3], [7, 8, 9]]);
  });

  it('takes, drops and chooses rows and columns', () => {
    expect(spill('=TAKE(A1:C3, 2)')).toEqual([[1, 2, 3], [4, 5, 6]]);
    expect(spill('=TAKE(A1:C3, -1, -2)')).toEqual([[8, 9]]);
    expect(spill('=TAKE(A1:C3, , 1)')).toEqual([[1], [4], [7]]);
    expect(spill('=DROP(A1:C3, 1, -1)')).toEqual([[4, 5], [7, 8]]);
    expect(spill('=DROP(A1:C3, 3)')).toMatchObject({ code: '#CALC!' });
    expect(spill('=CHOOSEROWS(A1:C3, 3, -3)')).toEqual([[7, 8, 9], [1, 2, 3]]);
    expect(spill('=CHOOSECOLS(A1:C3, {2,1})')).toEqual([[2, 1], [5, 4], [8, 7]]);
    expect(spill('=CHOOSECOLS(A1:C3, 4)')).toMatchObject({ code: '#VALUE!' });
  });

  it('flattens, wraps and expands', () => {
    expect(spill('=TOROW(A1:B2)')).toEqual([[1, 2, 4, 5]]);
    expect(spill('=TOCOL(A1:B2, 0, TRUE)')).toEqual([[1], [4], [2], [5]]);
    expect(spill('=TOCOL(D1:E2, 1)', (wb) => setValues(wb, 0, 0, 3, [['a', null], [null, '=1/0']]))).toEqual([['a'], [expect.objectContaining({ code: '#DIV/0!' })]]);
    expect(spill('=TOCOL(D1:E2, 3)', (wb) => setValues(wb, 0, 0, 3, [['a', null], [null, '=1/0']]))).toEqual([['a']]);
    expect(spill('=WRAPROWS(SEQUENCE(5), 2, 0)')).toEqual([[1, 2], [3, 4], [5, 0]]);
    expect(spill('=WRAPCOLS({1,2,3}, 2)')).toEqual([[1, 3], [2, NA]]);
    expect(spill('=EXPAND({1,2}, 2, 3, "-")')).toEqual([[1, 2, '-'], ['-', '-', '-']]);
    expect(spill('=EXPAND(A1:C3, 2)')).toMatchObject({ code: '#VALUE!' });
  });

  it('sorts by other arrays with stable multi-key order', () => {
    const setup = (wb: Workbook) => setValues(wb, 0, 0, 4, [['b', 2], ['a', 2], ['c', 1]]);
    expect(spill('=SORTBY(E1:E3, F1:F3, 1, E1:E3, -1)', setup)).toEqual([['c'], ['b'], ['a']]);
    expect(spill('=SORTBY({"x","y","z"}, {3,1,2})')).toEqual([['y', 'z', 'x']]);
    expect(spill('=SORTBY(E1:E3, F1:F2)', setup)).toMatchObject({ code: '#VALUE!' });
  });
});
