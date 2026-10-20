import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { readXlsx, setValues, toRange, Workbook, writeXlsx } from '../src/index.js';

function values(wb: Workbook, range: string) {
  const { start, end } = toRange(range);
  return Array.from({ length: end.row - start.row + 1 }, (_, r) =>
    Array.from({ length: end.col - start.col + 1 }, (_, c) => wb.getValue(0, start.row + r, start.col + c)));
}

function calc(formula: string, setup?: (wb: Workbook) => void) {
  const wb = new Workbook();
  setup?.(wb);
  setValues(wb, 0, 20, 20, [[formula]]);
  return wb.getValue(0, 20, 20);
}

describe('LET and LAMBDA', () => {
  it.each([
    ['IFS', 'IFS(n=0, "done", TRUE, DOWN(n-1))'],
    ['SWITCH', 'SWITCH(n, 0, "done", DOWN(n-1))'],
    ['CHOOSE', 'CHOOSE(IF(n=0, 1, 2), "done", DOWN(n-1))'],
    ['LET', 'LET(next, n-1, IF(n=0, "done", DOWN(next)))'],
    ['IFERROR fallback', 'IF(n=0, "done", IFERROR(1/0, DOWN(n-1)))'],
    ['IFNA fallback', 'IF(n=0, "done", IFNA(NA(), DOWN(n-1)))'],
  ])('keeps tail recursion through %s independent of the native stack', (_name, body) => {
    const wb = new Workbook();
    wb.apply({ type: 'setName', name: 'DOWN', ref: `LAMBDA(n, ${body})` });
    expect(wb.evaluate('=DOWN(900)')).toBe('done');
    expect(String(wb.evaluate('=DOWN(1000)'))).toBe('#CALC!');
    expect(wb.evaluate('=DOWN(900)')).toBe('done');
  });

  it('preserves argument values and LET bindings across tail calls', () => {
    const wb = new Workbook();
    wb.apply({ type: 'setName', name: 'TOTAL', ref: 'LAMBDA(n, acc, LET(next, n-1, IF(n=0, acc, TOTAL(next, acc+n))))' });
    expect(wb.evaluate('=TOTAL(900, 7)')).toBe(405457);
    expect(wb.evaluate('=TOTAL(3, 7)')).toBe(13);
  });

  it('trampolines mutual recursion and calls of computed LAMBDAs', () => {
    const wb = new Workbook();
    wb.apply({ type: 'setName', name: 'EVEN_N', ref: 'LAMBDA(n, IF(n=0, TRUE, ODD_N(n-1)))' });
    wb.apply({ type: 'setName', name: 'ODD_N', ref: 'LAMBDA(n, IF(n=0, FALSE, EVEN_N(n-1)))' });
    expect(wb.evaluate('=EVEN_N(900)')).toBe(true);
    expect(wb.evaluate('=EVEN_N(899)')).toBe(false);
    wb.apply({ type: 'setName', name: 'DOWN', ref: 'LAMBDA(n, IF(n=0, "done", LAMBDA(x, DOWN(x))(n-1)))' });
    expect(wb.evaluate('=DOWN(450)')).toBe('done');
    expect(String(wb.evaluate('=DOWN(500)'))).toBe('#CALC!');
  });

  it('counts nested non-tail calls and resets depth between higher-order callbacks', () => {
    const wb = new Workbook();
    wb.apply({ type: 'setName', name: 'ZERO', ref: 'LAMBDA(n, IF(n=0, 0, ZERO(n-1)))' });
    wb.apply({ type: 'setName', name: 'OUTER', ref: 'LAMBDA(n, 1+ZERO(n))' });
    expect(wb.evaluate('=OUTER(998)')).toBe(1);
    expect(String(wb.evaluate('=OUTER(999)'))).toBe('#CALC!');
    expect(wb.evaluate('=OUTER(998)')).toBe(1);
    expect(wb.evaluate('=SUM(MAP({900,900}, LAMBDA(n, ZERO(n)+1)))')).toBe(2);
    expect(wb.evaluate('=ZERO(999)')).toBe(0);
  });

  it('does not defer array branches or eager arguments', () => {
    const wb = new Workbook();
    expect(wb.evaluate('=SUM(LAMBDA(n, IF({TRUE,FALSE}, n+1, n+2))(10))')).toBe(23);
    expect(wb.evaluate('=LAMBDA(n, IFERROR(n, 5))(1/0)')).toBe(5);
    expect(wb.evaluate('=LAMBDA(n, SWITCH(n, 0, 1, ))(2)')).toBe(0);
  });

  it('binds names sequentially and shadows outer names', () => {
    expect(calc('=LET(x, 2, y, x*3, x+y)')).toBe(8);
    expect(calc('=LET(x, 2, LET(x, 5, x)+x)')).toBe(7);
    expect(calc('=LET(rng, A1:A3, SUM(rng))', (wb) => setValues(wb, 0, 0, 0, [[1], [2], [3]]))).toBe(6);
    expect(calc('=LET(x, 1, y)')).toMatchObject({ code: '#NAME?' });
    expect(calc('=LET(x, 1)')).toMatchObject({ code: '#VALUE!' });
    expect(calc('=LET(1, 1, 2)')).toMatchObject({ code: '#VALUE!' });
  });

  it('calls lambdas bound by LET and defined names, and errors for uncalled lambdas', () => {
    expect(calc('=LET(sq, LAMBDA(n, n*n), sq(4) + sq(2))')).toBe(20);
    expect(calc('=LAMBDA(n, n*n)')).toMatchObject({ code: '#CALC!' });
    expect(calc('=LET(f, LAMBDA(a, b, a-b), f(10, 3, 1))')).toMatchObject({ code: '#VALUE!' });
    expect(calc('=Double(21)', (wb) => wb.apply({ type: 'setName', name: 'Double', ref: 'LAMBDA(x, x*2)' }))).toBe(42);
    expect(calc('=LET(add, LAMBDA(a, LAMBDA(b, a+b)), inc, add(1), inc(41))')).toBe(42);
  });

  it('maps, folds and builds arrays with lambdas', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, 2], [3, 4]]);
    setValues(wb, 0, 0, 3, [['=MAP(A1:B2, LAMBDA(v, v*10))']]);
    expect(values(wb, 'D1:E2')).toEqual([[10, 20], [30, 40]]);
    setValues(wb, 0, 0, 6, [['=BYROW(A1:B2, LAMBDA(r, SUM(r)))']]);
    expect(values(wb, 'G1:G2')).toEqual([[3], [7]]);
    setValues(wb, 0, 3, 0, [['=BYCOL(A1:B2, LAMBDA(c, MAX(c)))']]);
    expect(values(wb, 'A4:B4')).toEqual([[3, 4]]);
    setValues(wb, 0, 5, 0, [['=REDUCE(0, A1:B2, LAMBDA(acc, v, acc + v*v))']]);
    expect(wb.getValue(0, 5, 0)).toBe(30);
    setValues(wb, 0, 6, 0, [['=SCAN(0, {1,2,3}, LAMBDA(acc, v, acc + v))']]);
    expect(values(wb, 'A7:C7')).toEqual([[1, 3, 6]]);
    setValues(wb, 0, 7, 0, [['=MAKEARRAY(2, 3, LAMBDA(r, c, r*c))']]);
    expect(values(wb, 'A8:C9')).toEqual([[1, 2, 3], [2, 4, 6]]);
    setValues(wb, 0, 9, 0, [['=MAP(A1:B2, {1,2;3,4}, LAMBDA(a, b, a*b))']]);
    expect(values(wb, 'A10:B11')).toEqual([[1, 4], [9, 16]]);
    setValues(wb, 0, 11, 0, [['=MAP(A1:B2, 5)']]);
    expect(wb.getValue(0, 11, 0)).toMatchObject({ code: '#VALUE!' });
  });

  it('recalculates when referenced cells change', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[2, '=LET(x, A1, x*x)']]);
    expect(wb.getValue(0, 0, 1)).toBe(4);
    setValues(wb, 0, 0, 0, [[5]]);
    expect(wb.getValue(0, 0, 1)).toBe(25);
  });

  it('writes Excel prefixes for functions and parameters and reads them back', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['=LET(x, 3, sq, LAMBDA(n, n*n), sq(x))', '=BYROW({1,2;3,4}, LAMBDA(r, SUM(r)))']]);
    const bytes = writeXlsx(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('_xlfn.LET(_xlpm.x, 3, _xlpm.sq, _xlfn.LAMBDA(_xlpm.n, _xlpm.n*_xlpm.n), _xlpm.sq(_xlpm.x))');
    expect(xml).toContain('_xlfn.BYROW(');
    const back = readXlsx(bytes);
    expect(back.activeSheet.getCell(0, 0)?.formula).toBe('LET(x, 3, sq, LAMBDA(n, n*n), sq(x))');
    expect(back.getValue(0, 0, 0)).toBe(9);
    expect(values(back, 'B1:B2')).toEqual([[3], [7]]);
  });
});
