import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { deleteRows, insertRows, invalidCells, isValidValue, readXlsx, setDataValidations, setValues, validationAt, validationListValues, Workbook, writeXlsx, type DataValidation } from '../src/index.js';

function book(rules: DataValidation[]) {
  const wb = new Workbook();
  setValues(wb, 0, 0, 3, [['Red'], ['Green'], ['Blue']]);
  setDataValidations(wb, 0, rules);
  return wb;
}
const check = (wb: Workbook, row: number, col: number, value: string | number | null) => isValidValue(wb, 0, validationAt(wb.activeSheet, row, col)!, row, col, value);

describe('data validation', () => {
  it('validates list rules from literals and ranges', () => {
    const wb = book([
      { range: 'A1:A10', type: 'list', formula1: '"Yes,No, Maybe"' },
      { range: 'B1:B10', type: 'list', formula1: '$D$1:$D$3', allowBlank: false },
    ]);
    expect(validationListValues(wb, 0, validationAt(wb.activeSheet, 0, 0)!, 0, 0)).toEqual(['Yes', 'No', 'Maybe']);
    expect(validationListValues(wb, 0, validationAt(wb.activeSheet, 4, 1)!, 4, 1)).toEqual(['Red', 'Green', 'Blue']);
    expect(check(wb, 0, 0, 'yes')).toBe(true);
    expect(check(wb, 0, 0, 'Perhaps')).toBe(false);
    expect(check(wb, 0, 0, null)).toBe(true);
    expect(check(wb, 0, 1, 'Green')).toBe(true);
    expect(check(wb, 0, 1, null)).toBe(false);
    expect(validationAt(wb.activeSheet, 10, 0)).toBeUndefined();
  });

  it('compares numbers, dates, times and text lengths', () => {
    const wb = book([
      { range: 'A1', type: 'whole', operator: 'between', formula1: '1', formula2: '10' },
      { range: 'A2', type: 'decimal', operator: 'greaterThan', formula1: 'D5' },
      { range: 'A3', type: 'textLength', operator: 'lessThanOrEqual', formula1: '3' },
      { range: 'A4', type: 'date', operator: 'greaterThanOrEqual', formula1: 'DATE(2024,1,1)' },
      { range: 'A5', type: 'time', operator: 'lessThan', formula1: '0.5' },
      { range: 'A6', type: 'whole', operator: 'notBetween', formula1: '5', formula2: '1' },
    ]);
    setValues(wb, 0, 4, 3, [[2.5]]);
    expect([check(wb, 0, 0, 5), check(wb, 0, 0, 5.5), check(wb, 0, 0, 11), check(wb, 0, 0, 'x')]).toEqual([true, false, false, false]);
    expect([check(wb, 1, 0, 2.6), check(wb, 1, 0, 2.5)]).toEqual([true, false]);
    expect([check(wb, 2, 0, 'abc'), check(wb, 2, 0, 'abcd'), check(wb, 2, 0, 1234)]).toEqual([true, false, false]);
    expect([check(wb, 3, 0, 45292), check(wb, 3, 0, 45291)]).toEqual([true, false]);
    expect([check(wb, 4, 0, 0.25), check(wb, 4, 0, 0.75), check(wb, 4, 0, 1.25)]).toEqual([true, false, false]);
    expect([check(wb, 5, 0, 6), check(wb, 5, 0, 3)]).toEqual([true, false]);
  });

  it('evaluates custom formulas relative to each cell and lists invalid cells', () => {
    const wb = book([{ range: 'B1:B4', type: 'custom', formula1: 'B1<=C1' }]);
    setValues(wb, 0, 0, 1, [[1, 5], [9, 5], [5, 5], [null, 1]]);
    expect(invalidCells(wb, 0)).toEqual([{ row: 1, col: 1 }]);
  });

  it('moves ranges and formulas with inserted and deleted rows and undoes rule changes', () => {
    const wb = book([{ range: 'B2:B5 D2', type: 'custom', formula1: 'B2>A2' }]);
    insertRows(wb, 0, 0, 2);
    expect(wb.activeSheet.validations[0]).toMatchObject({ range: 'B4:B7 D4', formula1: 'B4>A4' });
    deleteRows(wb, 0, 3, 1);
    expect(wb.activeSheet.validations[0]).toMatchObject({ range: 'B4:B6', formula1: 'B4>A4' });
    setDataValidations(wb, 0, []);
    expect(wb.activeSheet.validations).toEqual([]);
    wb.undo();
    expect(wb.activeSheet.validations).toHaveLength(1);
    expect(() => setDataValidations(wb, 0, [{ range: 'A1', type: 'whole', formula1: '1' }])).toThrow(/minimum and a maximum/);
  });

  it('roundtrips rules through XLSX', () => {
    const wb = book([
      { range: 'A1:A10', type: 'list', formula1: '"Yes,No"', dropdown: false, prompt: 'Pick one', promptTitle: 'Choice' },
      { range: 'B1', type: 'decimal', operator: 'lessThan', formula1: '100', errorStyle: 'warning', error: 'Too large', allowBlank: false },
    ]);
    const bytes = writeXlsx(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<dataValidation type="list" allowBlank="1" showDropDown="1" showInputMessage="1" showErrorMessage="1" promptTitle="Choice" prompt="Pick one" sqref="A1:A10"><formula1>&quot;Yes,No&quot;</formula1></dataValidation>');
    expect(xml.indexOf('<dataValidations')).toBeLessThan(xml.indexOf('<pageMargins'));
    expect(readXlsx(bytes).activeSheet.validations).toEqual(wb.activeSheet.validations);
  });

  it('patches only the validation element of a preserved workbook', () => {
    const original = readFileSync(new URL('./fixtures/excel-features.xlsx', import.meta.url));
    const wb = readXlsx(original);
    expect(wb.activeSheet.validations).toEqual([{ range: 'E2:E20', type: 'list', formula1: '"A,B,C"', allowBlank: false, showError: false, showPrompt: false }]);
    setDataValidations(wb, 0, [{ range: 'E2:E30', type: 'list', formula1: '"A,B,C,D"' }]);
    const before = unzipSync(original);
    const after = unzipSync(writeXlsx(wb));
    const sheet = strFromU8(after['xl/worksheets/sheet1.xml']!);
    expect(sheet).toContain('sqref="E2:E30"><formula1>&quot;A,B,C,D&quot;</formula1>');
    expect(sheet.replace(/<dataValidations[\s\S]*?<\/dataValidations>/, '')).toBe(strFromU8(before['xl/worksheets/sheet1.xml']!).replace(/<dataValidations[\s\S]*?<\/dataValidations>/, ''));
    for (const name of Object.keys(before)) if (!['xl/worksheets/sheet1.xml', 'xl/workbook.xml'].includes(name)) expect(strFromU8(after[name]!)).toBe(strFromU8(before[name]!));
  });
});
