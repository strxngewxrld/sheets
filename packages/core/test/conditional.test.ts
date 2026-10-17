import { describe, expect, it } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import { applyStyle, ConditionalFormatEvaluator, readXlsx, setConditionalFormats, setValues, toRange, Workbook, writeXlsx, type ConditionalFormatRule } from '../src/index.js';

describe('conditional formatting', () => {
  it('combines rules by priority, honors stopIfTrue and preserves base formatting', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[5], [15], [25]]);
    applyStyle(wb, 0, toRange('A1:A3'), { bold: true, fill: '#FFFFFF' });
    const rules: ConditionalFormatRule[] = [
      { id: 'high', range: 'A1:A3', type: 'cellIs', operator: 'greaterThan', formulas: ['20'], style: { fill: '#FF0000' }, stopIfTrue: true },
      { id: 'medium', range: 'A1:A3', type: 'cellIs', operator: 'between', formulas: ['10', '30'], style: { fill: '#FFFF00', color: '#000000', bold: false } },
      { id: 'all', range: 'A1:A3', type: 'expression', formulas: ['TRUE'], style: { italic: true, fill: '#00FF00' } },
    ];
    setConditionalFormats(wb, 0, rules);
    expect(wb.getCellStyle(0, 0, 0)).toMatchObject({ bold: true, italic: true, fill: '#00FF00' });
    expect(wb.getCellStyle(0, 1, 0)).toMatchObject({ bold: false, italic: true, fill: '#FFFF00' });
    expect(wb.getCellStyle(0, 2, 0)).toEqual({ bold: true, fill: '#FF0000' });
    expect(wb.activeSheet.getCell(2, 0)?.style?.fill).toBe('#FFFFFF');
    wb.undo();
    expect(wb.activeSheet.conditionalFormats).toHaveLength(0);
    wb.redo();
    expect(wb.activeSheet.conditionalFormats).toHaveLength(3);
  });

  it('shifts relative formulas from the range anchor and updates after input changes', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1, 5, 10], [20, 5]]);
    setConditionalFormats(wb, 0, [{ id: 'formula', range: 'A1:B2', type: 'expression', formulas: ['$A1>$C$1'], style: { bold: true } }]);
    expect(wb.getCellStyle(0, 0, 1)).toBeUndefined();
    expect(wb.getCellStyle(0, 1, 1)?.bold).toBe(true);
    setValues(wb, 0, 0, 2, [[30]]);
    expect(wb.getCellStyle(0, 1, 1)).toBeUndefined();
    wb.apply({ type: 'insertRows', sheet: 'Sheet1', index: 0, count: 1 });
    expect(wb.activeSheet.conditionalFormats[0]).toMatchObject({ range: 'A2:B3', formulas: ['$A2>$C$2'] });
    wb.undo();
    expect(wb.activeSheet.conditionalFormats[0]?.range).toBe('A1:B2');
  });

  it('uses a cached min/max for numeric color scales and ignores nonnumeric cells', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[0], [50], [100], ['text']]);
    setConditionalFormats(wb, 0, [{ id: 'scale', range: 'A1:A4', type: 'colorScale', colors: ['#FF0000', '#FFFF00', '#00FF00'] }]);
    const resolve = new ConditionalFormatEvaluator(wb);
    expect(resolve.style(wb.activeSheet, 0, 0)?.fill).toBe('#FF0000');
    expect(resolve.style(wb.activeSheet, 1, 0)?.fill).toBe('#FFFF00');
    expect(resolve.style(wb.activeSheet, 2, 0)?.fill).toBe('#00FF00');
    expect(resolve.style(wb.activeSheet, 3, 0)).toBeUndefined();
  });

  it('roundtrips native XLSX rules and differential styles without baking them into cells', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1], [20], [30]]);
    setConditionalFormats(wb, 0, [
      { id: 'threshold', range: 'A1:A3', type: 'cellIs', operator: 'greaterThanOrEqual', formulas: ['20'], style: { fill: '#FFFF00', color: '#000000', bold: false, wrap: false, numFmt: '0.00', border: { bottom: { style: 'thin', color: '#FF0000' } } }, stopIfTrue: true },
      { id: 'formula', range: 'B1:B3', type: 'expression', formulas: ['A1>10'], style: { italic: true, fontSize: 14, hAlign: 'center', vAlign: 'middle' } },
      { id: 'scale', range: 'A1:A3', type: 'colorScale', colors: ['#FF0000', '#00FF00'] },
    ]);
    const bytes = writeXlsx(wb);
    const files = unzipSync(bytes);
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('<conditionalFormatting sqref="A1:A3">');
    expect(strFromU8(files['xl/styles.xml']!)).toContain('<dxfs count="2">');
    const back = readXlsx(bytes);
    expect(back.activeSheet.conditionalFormats).toHaveLength(3);
    expect(back.getCellStyle(0, 1, 0)).toMatchObject({ fill: '#FFFF00', color: '#000000', bold: false, wrap: false, numFmt: '0.00' });
    expect(back.getCellStyle(0, 1, 1)).toMatchObject({ italic: true, fontSize: 14, hAlign: 'center', vAlign: 'middle' });
    expect(back.activeSheet.getCell(1, 0)?.style).toBeUndefined();
    const jsonBack = Workbook.fromJSON(wb.toJSON());
    expect(jsonBack.activeSheet.conditionalFormats).toEqual(wb.activeSheet.conditionalFormats);
  });

  it('rejects invalid definitions without changing rules and ignores formula errors', () => {
    const wb = new Workbook();
    expect(() => setConditionalFormats(wb, 0, [{ id: 'bad', range: 'no!', type: 'colorScale', colors: ['#000000', '#FFFFFF'] }])).toThrow(/range/);
    expect(() => setConditionalFormats(wb, 0, [{ id: 'bad', range: 'A1', type: 'cellIs', operator: 'between', formulas: ['1'], style: {} }])).toThrow(/2 formula/);
    expect(wb.activeSheet.conditionalFormats).toHaveLength(0);
    setConditionalFormats(wb, 0, [{ id: 'error', range: 'A1', type: 'expression', formulas: ['1/0'], style: { bold: true } }]);
    expect(wb.getCellStyle(0, 0, 0)).toBeUndefined();
  });

  it('rebases relative formulas when the first row of a rule range is deleted', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[1], [20], [30]]);
    setConditionalFormats(wb, 0, [{ id: 'relative', range: 'A1:A3', type: 'expression', formulas: ['A1>10'], style: { bold: true } }]);
    wb.apply({ type: 'deleteRows', sheet: 'Sheet1', index: 0, count: 1 });
    expect(wb.activeSheet.conditionalFormats[0]).toMatchObject({ range: 'A1:A2', formulas: ['A1>10'] });
    expect(wb.getCellStyle(0, 0, 0)?.bold).toBe(true);
    wb.undo();
    expect(wb.getCellStyle(0, 0, 0)).toBeUndefined();
  });
});
