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

describe('data bars and icon sets', () => {
  it('computes bar lengths and icons from range statistics', async () => {
    const { ConditionalFormatEvaluator, setConditionalFormats, setValues, Workbook } = await import('../src/index.js');
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [[0], [25], [50], [100], ['x']]);
    setConditionalFormats(wb, 0, [
      { id: 'bar', range: 'A1:A5', type: 'dataBar', color: '#638EC6' },
      { id: 'icons', range: 'A1:A5', type: 'iconSet', iconSet: '3Arrows' },
      { id: 'fixed', range: 'B1:B2', type: 'dataBar', color: '#FF0000', min: { type: 'num', value: 0 }, max: { type: 'num', value: 10 }, showValue: false },
    ]);
    setValues(wb, 0, 0, 1, [[5], [20]]);
    const ev = new ConditionalFormatEvaluator(wb);
    const sheet = wb.activeSheet;
    expect([0, 1, 2, 3].map((r) => ev.visual(sheet, r, 0)?.bar?.fraction)).toEqual([0, 0.25, 0.5, 1]);
    expect([0, 1, 2, 3].map((r) => ev.visual(sheet, r, 0)?.icon?.index)).toEqual([0, 0, 1, 2]);
    expect(ev.visual(sheet, 4, 0)).toBeUndefined();
    expect(ev.visual(sheet, 0, 1)).toEqual({ bar: { fraction: 0.5, color: '#FF0000' }, hideValue: true });
    expect(ev.visual(sheet, 1, 1)?.bar?.fraction).toBe(1);
    expect(ev.style(sheet, 1, 0)).toBeUndefined();
    setConditionalFormats(wb, 0, [{ id: 'r', range: 'A1:A4', type: 'iconSet', iconSet: '4Rating', reverse: true, thresholds: [{ type: 'num', value: 0 }, { type: 'percentile', value: 25 }, { type: 'percentile', value: 50 }, { type: 'max' }] }]);
    const ev2 = new ConditionalFormatEvaluator(wb);
    expect([0, 1, 2, 3].map((r) => ev2.visual(sheet, r, 0)?.icon?.index)).toEqual([3, 2, 1, 0]);
    expect(() => setConditionalFormats(wb, 0, [{ id: 'bad', range: 'A1', type: 'iconSet', iconSet: '3Arrows', thresholds: [{ type: 'min' }] }])).toThrow(/needs 3 thresholds/);
  });

  it('roundtrips through XLSX and reads Excel data bars', async () => {
    const { readFileSync } = await import('node:fs');
    const { strFromU8, unzipSync } = await import('fflate');
    const { readXlsx, setConditionalFormats, Workbook, writeXlsx } = await import('../src/index.js');
    const wb = new Workbook();
    const rules = [
      { id: 'a', range: 'A1:A9', type: 'dataBar' as const, color: '#63BE7B', min: { type: 'num' as const, value: 0 }, showValue: false },
      { id: 'b', range: 'B1:B9', type: 'iconSet' as const, iconSet: '5Quarters' as const, reverse: true },
    ];
    setConditionalFormats(wb, 0, rules);
    const bytes = writeXlsx(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<dataBar showValue="0"><cfvo type="num" val="0"/><cfvo type="max"/><color rgb="FF63BE7B"/></dataBar>');
    expect(xml).toContain('<iconSet iconSet="5Quarters" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="20"/>');
    expect(readXlsx(bytes).activeSheet.conditionalFormats.map(({ id: _id, ...r }) => r)).toEqual(rules.map(({ id: _id, ...r }) => ({ ...r, stopIfTrue: false })));
    const fixture = readXlsx(readFileSync(new URL('./fixtures/excel-features.xlsx', import.meta.url)));
    expect(fixture.activeSheet.conditionalFormats).toMatchObject([{ range: 'B2:B20', type: 'dataBar', color: '#638EC6', min: { type: 'num', value: 0 }, max: { type: 'num', value: 100 } }]);
  });
});
