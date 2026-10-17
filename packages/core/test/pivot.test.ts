import { describe, expect, it } from 'vitest';
import { createPivotTable, evaluatePivot, readXlsx, refreshPivotTable, removePivotTable, setValues, Workbook, writeXlsx, type PivotTableSpec } from '../src/index.js';

function example(): { wb: Workbook; spec: PivotTableSpec } {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [
    ['Region', 'Product', 'Sales'],
    ['North', 'Apple', 10], ['South', 'Apple', 30], ['North', 'Pear', 20], ['North', 'Apple', 40],
  ]);
  wb.addSheet('Report');
  return { wb, spec: { id: 'sales', source: { sheet: 'Sheet1', range: 'A1:C5' }, target: { sheet: 'Report', cell: 'B2' }, rows: ['Region'], columns: ['Product'], values: [{ field: 'Sales', aggregate: 'sum' }] } };
}

describe('pivot reports', () => {
  it('groups rows and columns with overall totals', () => {
    const { wb, spec } = example();
    expect(evaluatePivot(wb, spec)).toEqual([
      ['Region', 'Apple / SUM Sales', 'Pear / SUM Sales', 'Grand Total / SUM Sales'],
      ['North', 50, 20, 70], ['South', 30, 0, 30], ['Grand Total', 80, 20, 100],
    ]);
    const pivot = createPivotTable(wb, spec);
    expect(pivot.outputRange).toBe('B2:E5');
    expect(wb.getValue('Report', 4, 4)).toBe(100);
    wb.undo();
    expect(wb.getValue('Report', 4, 4)).toBe(null);
    expect(wb.pivotTables).toHaveLength(0);
    wb.redo();
    expect(wb.getValue('Report', 4, 4)).toBe(100);
    expect(wb.pivotTables).toHaveLength(1);
  });

  it('computes weighted totals, multiple aggregates and filters', () => {
    const { wb, spec } = example();
    spec.columns = [];
    spec.values = [{ field: 2, aggregate: 'average', label: 'Mean' }, { field: 'Sales', aggregate: 'count' }, { field: 'Sales', aggregate: 'min' }, { field: 'Sales', aggregate: 'max' }];
    expect(evaluatePivot(wb, spec)).toEqual([
      ['Region', 'Mean', 'COUNT Sales', 'MIN Sales', 'MAX Sales'],
      ['North', 70 / 3, 3, 10, 40], ['South', 30, 1, 30, 30], ['Grand Total', 25, 4, 10, 40],
    ]);
    spec.filters = [{ field: 'Product', criteria: 'A*' }, { field: 'Sales', criteria: '>10' }];
    expect(evaluatePivot(wb, spec).at(-1)).toEqual(['Grand Total', 35, 2, 30, 40]);
  });

  it('refreshes explicitly and clears obsolete output with a single undo step', () => {
    const { wb, spec } = example();
    createPivotTable(wb, spec);
    setValues(wb, 0, 2, 0, [['North']]);
    expect(wb.getValue('Report', 3, 1)).toBe('South');
    refreshPivotTable(wb, 'sales');
    expect(wb.getValue('Report', 3, 1)).toBe('Grand Total');
    expect(wb.getValue('Report', 4, 4)).toBe(null);
    wb.undo();
    expect(wb.getValue('Report', 3, 1)).toBe('South');
    expect(wb.getValue('Report', 4, 4)).toBe(100);
    removePivotTable(wb, 'sales');
    expect(wb.pivotTables).toHaveLength(0);
    expect(wb.getValue('Report', 1, 1)).toBe(null);
    wb.undo();
    expect(wb.pivotTables).toHaveLength(1);
    expect(wb.getValue('Report', 1, 1)).toBe('Region');
  });

  it('protects sources and occupied destinations without partial mutation', () => {
    const { wb, spec } = example();
    setValues(wb, 'Report', 2, 2, [['keep']]);
    const snapshot = wb.toJSON();
    expect(() => createPivotTable(wb, spec)).toThrow(/blocked/);
    expect(wb.toJSON()).toEqual(snapshot);
    expect(() => createPivotTable(wb, { ...spec, target: { sheet: 'Sheet1', cell: 'A1' } })).toThrow(/overlap its source/);
    expect(() => evaluatePivot(wb, { ...spec, rows: ['Missing'] })).toThrow(/field/);
  });

  it('roundtrips definitions in JSON and XLSX and can refresh after reopening', () => {
    const { wb, spec } = example();
    createPivotTable(wb, spec);
    for (const back of [Workbook.fromJSON(wb.toJSON()), readXlsx(writeXlsx(wb))]) {
      expect(back.pivotTables).toEqual(wb.pivotTables);
      expect(back.getValue('Report', 4, 4)).toBe(100);
      setValues(back, 0, 1, 2, [[110]]);
      refreshPivotTable(back, 'sales');
      expect(back.getValue('Report', 4, 4)).toBe(200);
    }
  });

  it('adjusts sheet names, ranges and targets for structural edits', () => {
    const { wb, spec } = example();
    createPivotTable(wb, spec);
    wb.apply({ type: 'renameSheet', from: 'Sheet1', to: 'Data' });
    wb.apply({ type: 'insertRows', sheet: 'Data', index: 0, count: 1 });
    wb.apply({ type: 'insertRows', sheet: 'Report', index: 0, count: 1 });
    expect(wb.pivotTables[0]).toMatchObject({ source: { sheet: 'Data', range: 'A2:C6' }, target: { sheet: 'Report', cell: 'B3' }, outputRange: 'B3:E6' });
    refreshPivotTable(wb, 'sales');
    expect(wb.getValue('Report', 5, 4)).toBe(100);
    wb.apply({ type: 'removeSheet', name: 'Data' });
    expect(wb.pivotTables).toHaveLength(0);
    wb.undo();
    expect(wb.pivotTables).toHaveLength(1);
  });

  it('groups blank values, skips blank source rows and handles ungrouped summaries', () => {
    const { wb, spec } = example();
    setValues(wb, 0, 1, 0, [[null]]);
    spec.columns = [];
    expect(evaluatePivot(wb, spec)[1]).toEqual(['(blank)', 10]);
    spec.rows = [];
    expect(evaluatePivot(wb, spec)).toEqual([['SUM Sales'], [100]]);
  });

  it('keeps numeric field offsets attached to their source columns', () => {
    const { wb, spec } = example();
    spec.rows = [0]; spec.columns = [1]; spec.values = [{ field: 2, aggregate: 'sum' }];
    createPivotTable(wb, spec);
    wb.apply({ type: 'insertCols', sheet: 'Sheet1', index: 1, count: 1 });
    expect(wb.pivotTables[0]).toMatchObject({ rows: [0], columns: [2], values: [{ field: 3 }], source: { range: 'A1:D5' } });
    refreshPivotTable(wb, 'sales');
    expect(wb.getValue('Report', 4, 4)).toBe(100);
    wb.apply({ type: 'deleteCols', sheet: 'Sheet1', index: 1, count: 1 });
    expect(wb.pivotTables[0]).toMatchObject({ rows: [0], columns: [1], values: [{ field: 2 }] });
    wb.apply({ type: 'deleteCols', sheet: 'Sheet1', index: 2, count: 1 });
    expect(wb.pivotTables).toHaveLength(0);
    wb.undo();
    expect(wb.pivotTables).toHaveLength(1);
  });

  it('rejects excessive output before allocating the cross product', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Row', 'Col', 'Value'], ...Array.from({ length: 1100 }, (_, i) => [i, i, 1])]);
    expect(() => evaluatePivot(wb, { id: 'large', source: { sheet: 'Sheet1', range: 'A1:C1101' }, target: { sheet: 'Sheet1', cell: 'E1' }, rows: [0], columns: [1], values: [{ field: 2, aggregate: 'sum' }] })).toThrow(/cell limit/);
  });
});
