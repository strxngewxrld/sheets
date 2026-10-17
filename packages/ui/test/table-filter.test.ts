import { describe, expect, it } from 'vitest';
import { formatTable, setTableFilter, setValues, toRange, Workbook } from '@sheets/core';
import { Geometry } from '../src/geometry.js';

describe('table filter geometry', () => {
  it('collapses filtered rows and restores their custom heights after clearing', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Country'], ['Canada'], ['Germany'], ['Canada']]);
    formatTable(wb, 0, toRange('A1:A4'));
    const s = wb.activeSheet; s.rowHeights.set(2, 40);
    setTableFilter(wb, 0, 'Table1', 0, ['Canada']);
    const filtered = new Geometry(s);
    expect(filtered.rows.size(2)).toBe(0);
    expect(filtered.rows.indexAt(filtered.rows.offset(2))).toBe(3);
    expect(s.rowHeights.get(2)).toBe(40);
    setTableFilter(wb, 0, 'Table1', 0, null);
    expect(new Geometry(s).rows.size(2)).toBe(40);
  });
});

describe('zoomed geometry', () => {
  it('scales custom and default sizes, keeping hidden rows hidden', () => {
    const wb = new Workbook();
    const s = wb.activeSheet;
    s.colWidths.set(1, 100);
    s.rowHeights.set(2, 0);
    const geo = new Geometry(s, 1.5);
    expect(geo.zoom).toBe(1.5);
    expect(geo.cols.size(0)).toBe(s.defaultColWidth * 1.5);
    expect(geo.cols.size(1)).toBe(150);
    expect(geo.cols.offset(2)).toBe(s.defaultColWidth * 1.5 + 150);
    expect(geo.rows.size(2)).toBe(0);
    expect(geo.rows.offset(3)).toBe(s.defaultRowHeight * 1.5 * 2);
    expect(s.colWidths.get(1)).toBe(100);
  });
});
