import { describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { applyStyle, formatTable, insertChartFromRange, readXlsx, removeTable, setConditionalFormats, setValues, sortRange, TABLE_THEMES, toRange, Workbook, writeXlsx } from '../src/index.js';

describe('table highlighting and sheet sorting', () => {
  it('keeps stripes on row positions after sorting complete records, preserving the header', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'City', 'Amount', 'Double'], ['Zoe', 'Berlin', 30, '=C2*2'], ['Anna', 'Hamburg', 10, '=C3*2'], ['Ben', 'Berlin', 20, '=C4*2']]);
    formatTable(wb, 0, toRange('A1:D4'));
    const stripe = wb.getCellStyle(0, 1, 0)?.fill;
    expect(stripe).toBe(TABLE_THEMES.blue!.stripe.fill);
    expect(wb.getCellStyle(0, 0, 0)).toMatchObject(TABLE_THEMES.blue!.header);
    expect(wb.activeSheet.getCell(1, 0)?.style).toBeUndefined();
    sortRange(wb, 0, toRange('A1:D4'), [{ col: 0 }], { hasHeader: true });
    expect(wb.getValue(0, 0, 0)).toBe('Name');
    expect(wb.getValue(0, 1, 0)).toBe('Anna');
    expect(wb.getValue(0, 1, 1)).toBe('Hamburg');
    expect(wb.getValue(0, 1, 3)).toBe(20);
    expect(wb.activeSheet.getCell(1, 3)?.formula).toBe('C2*2');
    expect(wb.getCellStyle(0, 1, 0)?.fill).toBe(stripe);
    expect(wb.getCellStyle(0, 2, 0)?.fill).toBe('#FFFFFF');
    wb.undo();
    expect(wb.getValue(0, 1, 0)).toBe('Zoe');
    expect(wb.activeSheet.tables).toHaveLength(1);
  });

  it('combines table styles, explicit cell formats and conditional rules by priority', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 10]]);
    formatTable(wb, 0, toRange('A1:B2'));
    applyStyle(wb, 0, toRange('A1'), { bold: false });
    expect(wb.getCellStyle(0, 0, 0)?.bold).toBe(false);
    applyStyle(wb, 0, toRange('B2'), { fill: '#FFFF00', numFmt: '0.000 "kg"' });
    expect(wb.getCellStyle(0, 1, 1)).toMatchObject({ fill: '#FFFF00', numFmt: '0.000 "kg"' });
    setConditionalFormats(wb, 0, [{ id: 'high', range: 'B2', type: 'cellIs', operator: 'greaterThan', formulas: ['5'], style: { fill: '#FF0000' } }]);
    expect(wb.getCellStyle(0, 1, 1)?.fill).toBe('#FF0000');
    const back = readXlsx(writeXlsx(wb));
    expect(back.getCellStyle(0, 0, 0)?.bold).toBe(false);
    expect(back.getCellStyle(0, 1, 1)?.numFmt).toBe('0.000 "kg"');
  });

  it('normalizes table headers, removes highlighting without deleting data, and undoes both', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'Name', null], ['A', 'B', 2]]);
    formatTable(wb, 0, toRange('A1:C2'));
    expect(wb.getValue(0, 0, 1)).toBe('Name2');
    expect(wb.getValue(0, 0, 2)).toBe('Column3');
    wb.undo();
    expect(wb.getValue(0, 0, 1)).toBe('Name');
    expect(wb.getValue(0, 0, 2)).toBe(null);
    expect(wb.activeSheet.tables).toHaveLength(0);
    wb.redo();
    removeTable(wb, 0, 'Table1');
    expect(wb.activeSheet.tables).toHaveLength(0);
    expect(wb.getValue(0, 1, 2)).toBe(2);
    expect(wb.getCellStyle(0, 1, 0)).toBeUndefined();
    wb.undo();
    expect(wb.activeSheet.tables).toHaveLength(1);
  });

  it('roundtrips native table parts, custom styles and drawing relationships together', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 1], ['B', 2]]);
    formatTable(wb, 0, toRange('A1:B3'), { name: 'Sales', theme: TABLE_THEMES.green });
    insertChartFromRange(wb, 0, toRange('A1:B3'), 'column');
    const bytes = writeXlsx(wb);
    const parts = unzipSync(bytes);
    expect(strFromU8(parts['xl/tables/table1.xml']!)).toContain('displayName="Sales"');
    expect(strFromU8(parts['xl/styles.xml']!)).toContain('<tableStyleElement type="firstRowStripe"');
    expect(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!)).toContain('/table"');
    expect(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!)).toContain('/drawing"');
    const back = readXlsx(bytes);
    expect(back.activeSheet.tables).toEqual(wb.activeSheet.tables);
    expect(back.activeSheet.charts).toHaveLength(1);
    expect(back.getCellStyle(0, 1, 0)?.fill).toBe(TABLE_THEMES.green!.stripe.fill);
    expect(Workbook.fromJSON(wb.toJSON()).activeSheet.tables).toEqual(wb.activeSheet.tables);
  });

  it('adjusts ranges during structural edits and respects headerless tables', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[10], [20]]);
    formatTable(wb, 0, toRange('A1:A2'), { headerRow: false });
    expect(wb.getValue(0, 0, 0)).toBe(10);
    wb.apply({ type: 'insertRows', sheet: 'Sheet1', index: 0, count: 1 });
    expect(wb.activeSheet.tables[0]?.range).toBe('A2:A3');
    wb.undo();
    expect(wb.activeSheet.tables[0]?.range).toBe('A1:A2');
    const back = readXlsx(writeXlsx(wb));
    expect(back.activeSheet.tables[0]?.headerRow).toBe(false);
    expect(back.getValue(0, 0, 0)).toBe(10);
  });

  it('rejects overlapping tables and unsafe sorts without partial writes', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Name', 'Value'], ['A', 1], ['B', 2]]);
    formatTable(wb, 0, toRange('A1:B3'));
    const snapshot = wb.toJSON();
    expect(() => formatTable(wb, 0, toRange('B1:C3'))).toThrow(/overlap/);
    expect(() => sortRange(wb, 0, toRange('A1:B3'), [{ col: 3 }])).toThrow(/inside/);
    expect(wb.toJSON()).toEqual(snapshot);
    wb.apply({ type: 'setMerges', sheet: 'Sheet1', merges: ['A2:B2'] });
    expect(() => sortRange(wb, 0, toRange('A1:B3'), [{ col: 0 }])).toThrow(/merged/);
  });

  it('uses multiple sort keys, keeps ties stable, and puts empty values last', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Team', 'Score', 'ID'], ['A', 2, 'first'], ['B', 8, 'second'], ['A', 9, 'third'], ['A', 9, 'fourth'], ['A', null, 'empty']]);
    sortRange(wb, 0, toRange('A1:C6'), [{ col: 0 }, { col: 1, descending: true }], { hasHeader: true });
    expect(Array.from({ length: 5 }, (_, r) => wb.getValue(0, r + 1, 2))).toEqual(['third', 'fourth', 'first', 'empty', 'second']);
  });
});
