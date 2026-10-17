import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { formatTable, insertRows, parseFormula, readXlsx, setValues, tokenize, toRange, Workbook, writeXlsx } from '../src/index.js';

function book() {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [['Product', 'Unit price', 'Qty', 'Total'], ['A', 2, 3], ['B', 5, 1], ['C', 1, 10]]);
  formatTable(wb, 0, toRange('A1:D4'), { name: 'Sales' });
  return wb;
}

describe('structured references', () => {
  it('tokenizes table references without confusing them with cells', () => {
    const kinds = (f: string) => tokenize(f).filter((t) => t.type !== 'ws').map((t) => `${t.type}:${t.text}`);
    expect(kinds('SUM(T1[Qty])')).toEqual(['func:SUM', 'lparen:(', 'structured:T1[Qty]', 'rparen:)']);
    expect(kinds('[@Qty]*2')).toEqual(['structured:[@Qty]', 'op:*', 'number:2']);
    expect(kinds("Sales[[#Headers],[Unit price]:[Qty]]")).toEqual(["structured:Sales[[#Headers],[Unit price]:[Qty]]"]);
    expect(kinds("Sales['#Count]")).toEqual(["structured:Sales['#Count]"]);
    expect(parseFormula("Sales[[#This Row],['[x']]]")).toEqual({ type: 'structured', table: 'Sales', areas: ['#This Row'], columns: ['[x]', '[x]'] });
    expect(parseFormula('Sales[@[Unit price]]')).toEqual({ type: 'structured', table: 'Sales', areas: ['#This Row'], columns: ['Unit price', 'Unit price'] });
  });

  it('evaluates columns, areas and this-row references', () => {
    const wb = book();
    expect(wb.evaluate('=SUM(Sales[Qty])')).toBe(14);
    expect(wb.evaluate('=SUMPRODUCT(Sales[Unit price],Sales[Qty])')).toBe(21);
    expect(wb.evaluate('=ROWS(Sales[#All])')).toBe(4);
    expect(wb.evaluate('=COLUMNS(Sales[#All])')).toBe(4);
    expect(wb.evaluate('=ROWS(Sales[])')).toBe(3);
    expect(wb.evaluate('=COLUMNS(Sales[[Unit price]:[Qty]])')).toBe(2);
    expect(wb.evaluate('=INDEX(Sales[#Headers],2)')).toBe('Unit price');
    expect(wb.evaluate('=sum(sales[QTY])')).toBe(14);
    expect((wb.evaluate('=SUM(Sales[Missing])') as { code: string }).code).toBe('#REF!');
    expect((wb.evaluate('=SUM(Other[Qty])') as { code: string }).code).toBe('#REF!');
    expect((wb.evaluate('=Sales[#Totals]') as { code: string }).code).toBe('#REF!');
    setValues(wb, 0, 1, 3, [['=[@[Unit price]]*[@Qty]']]);
    expect(wb.getValue(0, 1, 3)).toBe(6);
    setValues(wb, 0, 2, 3, [['=Sales[@Qty]']]);
    expect(wb.getValue(0, 2, 3)).toBe(1);
    setValues(wb, 0, 2, 5, [['=Sales[@Qty]']]);
    expect(wb.getValue(0, 2, 5)).toBe(1);
    setValues(wb, 0, 9, 5, [['=Sales[@Qty]']]);
    expect(wb.getValue(0, 9, 5)).toMatchObject({ code: '#VALUE!' });
    setValues(wb, 0, 8, 0, [['=[@Qty]']]);
    expect(wb.getValue(0, 8, 0)).toMatchObject({ code: '#REF!' });
  });

  it('recalculates when table cells or the table range change', () => {
    const wb = book();
    setValues(wb, 0, 6, 0, [['=SUM(Sales[Qty])']]);
    setValues(wb, 0, 1, 2, [[30]]);
    expect(wb.getValue(0, 6, 0)).toBe(41);
    insertRows(wb, 0, 2, 1);
    setValues(wb, 0, 2, 2, [[100]]);
    expect(wb.activeSheet.tables[0]!.range).toBe('A1:D5');
    expect(wb.getValue(0, 7, 0)).toBe(141);
  });

  it('stores qualified file forms in XLSX and shows the short form after reading', () => {
    const wb = book();
    setValues(wb, 0, 1, 3, [['=[@[Unit price]]*[@Qty]']]);
    setValues(wb, 0, 6, 0, [['=SUM(Sales[Qty])', '=Sales[@Qty]']]);
    const bytes = writeXlsx(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<f>Sales[[#This Row],[Unit price]]*Sales[[#This Row],[Qty]]</f>');
    expect(xml).toContain('<f>SUM(Sales[Qty])</f>');
    const back = readXlsx(bytes);
    expect(back.activeSheet.getCell(1, 3)?.formula).toBe('Sales[@[Unit price]]*Sales[@Qty]');
    expect(back.getValue(0, 1, 3)).toBe(6);
    expect(back.getValue(0, 6, 0)).toBe(14);
    // Editing the preserved package patches the formula in its qualified form.
    setValues(back, 0, 2, 3, [['=[@Qty]*10']]);
    expect(back.getValue(0, 2, 3)).toBe(10);
    const patched = strFromU8(unzipSync(writeXlsx(back))['xl/worksheets/sheet1.xml']!);
    expect(patched).toContain('Sales[[#This Row],[Qty]]*10</');
  });
});
