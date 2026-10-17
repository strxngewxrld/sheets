import { describe, expect, it } from 'vitest';
import { CellError, listFunctions, readXlsx, setValues, Workbook, writeXlsx } from '../src/index.js';

const evaluate = (formula: string) => {
  const value = new Workbook().evaluate(formula);
  return value instanceof CellError ? value.code : value;
};

describe('financial functions', () => {
  it.each([
    ['=PMT(8%/12,10,10000)', -1037.03208935916],
    ['=PMT(8%/12,10,10000,,1)', -1030.16432717797],
    ['=FV(6%/12,10,-200,-500,1)', 2581.40337406012],
    ['=NPV(10%,-10000,3000,4200,6800)', 1188.44341233522],
    ['=IPMT(10%/12,1,3*12,8000)', -66.66666666667],
    ['=PPMT(10%/12,1,3*12,8000)', -191.47083088403],
    ['=NPER(12%/12,-100,-1000,10000,1)', 59.67386567429],
  ])('matches the reference result of %s', (formula, expected) => {
    expect(evaluate(formula)).toBeCloseTo(expected, 8);
  });

  it.each([0, 1])('solves the same annuity for payment type %i', (type) => {
    const wb = new Workbook();
    // Test positive, negative and very small interest without subtractive cancellation.
    for (const rate of [0, 0.005, -0.002, 1e-12]) {
      const pmt = wb.evaluate(`=PMT(${rate},120,10000,-500,${type})`) as number;
      expect(wb.evaluate(`=PV(${rate},120,${pmt},-500,${type})`)).toBeCloseTo(10000, 7);
      expect(wb.evaluate(`=FV(${rate},120,${pmt},10000,${type})`)).toBeCloseTo(-500, 7);
      expect(wb.evaluate(`=NPER(${rate},${pmt},10000,-500,${type})`)).toBeCloseTo(120, 7);
    }
  });

  it('handles zero interest, omitted values and savings', () => {
    expect(evaluate('=PMT(0,10,1000)')).toBe(-100);
    expect(evaluate('=PV(0,10,-100)')).toBe(1000);
    expect(evaluate('=FV(0,10,-100)')).toBe(1000);
    expect(evaluate('=NPER(0,-100,1000)')).toBe(10);
    expect(evaluate('=NPER(10%,0,-100,121)')).toBeCloseTo(2, 10);
    expect(evaluate('=IPMT(0,5,10,1000)')).toBe(0);
    expect(evaluate('=PPMT(0,5,10,1000)')).toBe(-100);
    expect(evaluate('=FV(0.01,0,-100,-1000)')).toBe(1000);
  });

  it.each([0, 1])('balances a full amortization schedule for payment type %i', (type) => {
    const wb = new Workbook();
    const rate = 0.01;
    const pmt = wb.evaluate(`=PMT(${rate},12,1000,,${type})`) as number;
    let balance = 1000;
    let principal = 0;
    for (let per = 1; per <= 12; per++) {
      const interest = wb.evaluate(`=IPMT(${rate},${per},12,1000,,${type})`) as number;
      const repayment = wb.evaluate(`=PPMT(${rate},${per},12,1000,,${type})`) as number;
      const expectedInterest = type === 1 && per === 1 ? 0 : -balance * rate;
      expect(interest).toBeCloseTo(expectedInterest, 10);
      expect(interest + repayment).toBeCloseTo(pmt, 10);
      principal += repayment;
      balance += repayment;
    }
    expect(principal).toBeCloseTo(-1000, 8);
    expect(balance).toBeCloseTo(0, 8);
  });

  it('discounts ranges in order, skipping text, booleans and empty cells', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[100, 'ignore'], [null, true], [200, 0]]);
    expect(wb.evaluate('=NPV(10%,A1:B3)')).toBeCloseTo(100 / 1.1 + 200 / 1.1 ** 2, 10);
    expect(wb.evaluate('=NPV(0,A1:B3)')).toBe(300);
    setValues(wb, 0, 1, 0, [['=1/0']]);
    expect(wb.evaluate('=NPV(10%,A1:B3)')).toBeCloseTo(100 / 1.1 + 200 / 1.1 ** 2, 10);
    expect(wb.evaluate('=NPV(10%,TRUE,"100",100)')).toBeCloseTo(100 / 1.1, 10);
  });

  it.each([
    ['=PMT(0,0,100)', '#DIV/0!'],
    ['=NPER(0,0,100)', '#DIV/0!'],
    ['=NPER(10%,-1,100)', '#NUM!'],
    ['=NPER(-100%,-10,100)', '#NUM!'],
    ['=NPV(-100%,100)', '#DIV/0!'],
    ['=PMT(10%,12,100,,2)', '#NUM!'],
    ['=IPMT(10%,0,12,100)', '#NUM!'],
    ['=PPMT(10%,13,12,100)', '#NUM!'],
    ['=PMT("bad",12,100)', '#VALUE!'],
    ['=FV(1/0,12,100)', '#DIV/0!'],
    ['=PV(10%,12,100,#N/A)', '#N/A'],
    ['=PMT(10%,12)', '#VALUE!'],
    ['=FV(10%,1000000,100)', '#NUM!'],
  ])('returns a spreadsheet error for %s', (formula, expected) => {
    expect(evaluate(formula)).toBe(expected);
  });

  it('recalculates, supports undo and survives xlsx roundtrips', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[0, '=PMT(A1,12,1200)', '=IPMT(A1,1,12,1200)', '=PPMT(A1,1,12,1200)']]);
    expect(wb.getValue(0, 0, 1)).toBe(-100);
    setValues(wb, 0, 0, 0, [[0.01]]);
    expect(wb.getValue(0, 0, 1)).toBeCloseTo(-106.61854641401, 8);
    const imported = readXlsx(writeXlsx(wb));
    expect(imported.getValue(0, 0, 1)).toBeCloseTo(wb.getValue(0, 0, 1) as number, 10);
    expect(imported.getValue(0, 0, 2)).toBe(-12);
    wb.undo();
    expect(wb.getValue(0, 0, 1)).toBe(-100);
    expect(wb.getValue(0, 0, 2)).toBe(0);
  });

  it('exposes metadata for autocomplete and function discovery', () => {
    expect(listFunctions().filter((f) => f.category === 'financial').map((f) => f.name))
      .toEqual(['FV', 'IPMT', 'NPER', 'NPV', 'PMT', 'PPMT', 'PV']);
  });
});
