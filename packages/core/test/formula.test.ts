import { describe, expect, it } from 'vitest';
import { setValues, Workbook, CellError, shiftFormula, adjustFormulaForStructure, parseFormula } from '../src/index.js';

function wbWith(values: (string | number | boolean | null)[][]): Workbook {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, values);
  return wb;
}

const ev = (wb: Workbook, f: string) => {
  const v = wb.evaluate(f);
  return v instanceof CellError ? v.code : v;
};

describe('parser', () => {
  it('follows Excel precedence', () => {
    const wb = new Workbook();
    expect(ev(wb, '=1+2*3')).toBe(7);
    expect(ev(wb, '=-2^2')).toBe(4);
    expect(ev(wb, '=2^3^2')).toBe(64);
    expect(ev(wb, '=10%')).toBe(0.1);
    expect(ev(wb, '="a"&1+1')).toBe('a2');
    expect(ev(wb, '=1<2')).toBe(true);
  });

  it('parses sheet-qualified ranges and functions', () => {
    const ast = parseFormula("SUM('My Sheet'!A1:B2, Sheet2!C3)");
    expect(ast.type).toBe('func');
  });
});

describe('functions', () => {
  const wb = wbWith([
    ['Name', 'Region', 'Sales'],
    ['Anna', 'North', 100],
    ['Ben', 'South', 200],
    ['Cara', 'North', 300],
    ['Dan', 'West', null],
  ]);

  it('aggregates', () => {
    expect(ev(wb, '=SUM(C2:C5)')).toBe(600);
    expect(ev(wb, '=AVERAGE(C2:C5)')).toBe(200);
    expect(ev(wb, '=COUNT(C:C)')).toBe(3);
    expect(ev(wb, '=COUNTA(A:A)')).toBe(5);
    expect(ev(wb, '=MAX(C2:C5)-MIN(C2:C5)')).toBe(200);
    expect(ev(wb, '=ROUND(2.345, 2)')).toBe(2.35);
    expect(ev(wb, '=ROUND(-1.005, 2)')).toBe(-1.01);
  });

  it('conditional aggregates', () => {
    expect(ev(wb, '=SUMIF(B2:B5,"North",C2:C5)')).toBe(400);
    expect(ev(wb, '=COUNTIF(C2:C5,">150")')).toBe(2);
    expect(ev(wb, '=SUMIFS(C2:C5,B2:B5,"North",C2:C5,">100")')).toBe(300);
    expect(ev(wb, '=COUNTIF(A2:A5,"?a*")')).toBe(2);
    expect(ev(wb, '=SUMPRODUCT((B2:B5="North")*(C2:C5))')).toBe(400);
  });

  it('lookups', () => {
    expect(ev(wb, '=VLOOKUP("Ben",A2:C5,3,FALSE)')).toBe(200);
    expect(ev(wb, '=INDEX(C2:C5,MATCH("Cara",A2:A5,0))')).toBe(300);
    expect(ev(wb, '=XLOOKUP("Dan",A2:A5,B2:B5)')).toBe('West');
    expect(ev(wb, '=XLOOKUP("Zed",A2:A5,B2:B5,"none")')).toBe('none');
    expect(ev(wb, '=VLOOKUP("Zed",A2:C5,3,FALSE)')).toBe('#N/A');
  });

  it('logic, text and errors', () => {
    expect(ev(wb, '=IF(C2>50,"big","small")')).toBe('big');
    expect(ev(wb, '=IFERROR(1/0,"div")')).toBe('div');
    expect(ev(wb, '=1/0')).toBe('#DIV/0!');
    expect(ev(wb, '=UPPER(LEFT(A2,2))&LEN(A3)')).toBe('AN3');
    expect(ev(wb, '=TEXTJOIN(", ",TRUE,A2:A5)')).toBe('Anna, Ben, Cara, Dan');
    expect(ev(wb, '=TEXT(1234.5,"#,##0.00")')).toBe('1,234.50');
    expect(ev(wb, '=NOSUCHFN(1)')).toBe('#NAME?');
  });

  it('dates', () => {
    expect(ev(wb, '=DATE(2024,1,31)')).toBe(45322);
    expect(ev(wb, '=EOMONTH(DATE(2024,1,31),1)')).toBe(45351);
    expect(ev(wb, '=TEXT(DATE(2024,3,5),"dd.mm.yyyy")')).toBe('05.03.2024');
    expect(ev(wb, '=DATEDIF(DATE(2020,1,15),DATE(2024,3,1),"Y")')).toBe(4);
  });
});

describe('recalculation', () => {
  it('updates dependents and detects cycles', () => {
    const wb = wbWith([[1, '=A1*2', '=SUM(A1:B1)']]);
    expect(wb.getValue(0, 0, 2)).toBe(3);
    setValues(wb, 0, 0, 0, [[10]]);
    expect(wb.getValue(0, 0, 1)).toBe(20);
    expect(wb.getValue(0, 0, 2)).toBe(30);
    setValues(wb, 0, 0, 0, [['=C1']]);
    expect((wb.getValue(0, 0, 2) as CellError).code).toBe('#CIRC!');
  });

  it('handles long dependency chains without stack overflow', () => {
    const wb = new Workbook();
    const rows: (string | number)[][] = [[1]];
    for (let i = 1; i < 5000; i++) rows.push([`=A${i}+1`]);
    setValues(wb, 0, 0, 0, rows);
    expect(wb.getValue(0, 4999, 0)).toBe(5000);
    setValues(wb, 0, 0, 0, [[0]]);
    expect(wb.getValue(0, 4999, 0)).toBe(4999);
  });

  it('cross-sheet references', () => {
    const wb = wbWith([[5]]);
    wb.addSheet('Data');
    setValues(wb, 'Data', 0, 0, [['=Sheet1!A1*3']]);
    expect(wb.getValue('Data', 0, 0)).toBe(15);
    wb.apply({ type: 'renameSheet', from: 'Sheet1', to: 'Input' });
    expect(wb.requireSheet('Data').getCell(0, 0)!.formula).toBe('Input!A1*3');
    setValues(wb, 'Input', 0, 0, [[2]]);
    expect(wb.getValue('Data', 0, 0)).toBe(6);
  });
});

describe('structural changes and history', () => {
  it('adjusts references on insert/delete and undoes', () => {
    const wb = wbWith([[1], [2], [3], ['=SUM(A1:A3)']]);
    wb.apply({ type: 'insertRows', sheet: 'Sheet1', index: 1, count: 2 });
    expect(wb.requireSheet(0).getCell(5, 0)!.formula).toBe('SUM(A1:A5)');
    wb.apply({ type: 'deleteRows', sheet: 'Sheet1', index: 0, count: 1 });
    expect(wb.requireSheet(0).getCell(4, 0)!.formula).toBe('SUM(A1:A4)');
    expect(wb.getValue(0, 4, 0)).toBe(5);
    wb.undo();
    wb.undo();
    expect(wb.requireSheet(0).getCell(3, 0)!.formula).toBe('SUM(A1:A3)');
    expect(wb.getValue(0, 3, 0)).toBe(6);
    wb.redo();
    expect(wb.requireSheet(0).getCell(5, 0)!.formula).toBe('SUM(A1:A5)');
  });

  it('rewrites formulas', () => {
    expect(shiftFormula('A1+$B$2+C$3', 1, 1)).toBe('B2+$B$2+D$3');
    expect(shiftFormula('A1', -1, 0)).toBe('#REF!');
    expect(adjustFormulaForStructure('B5*2', 'S', { sheet: 'S', axis: 'row', index: 4, count: -1 })).toBe('#REF!*2');
  });
});
