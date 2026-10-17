import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { formatTable, readXlsx, setTableFilter, setValues, sortRange, tableFilteredRows, toRange, Workbook, writeXlsx } from '../src/index.js';

function sample() {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [['Segment', 'Country', 'Sales'], ['Government', 'Canada', 10], ['Government', 'Germany', 20], ['Midmarket', 'Canada', 30], ['Midmarket', null, 0]]);
  formatTable(wb, 0, toRange('A1:C5'), { name: 'financials' });
  return wb;
}

describe('table value filters', () => {
  it('combines columns and supports empty values, clearing, undo and redo without changing data', () => {
    const wb = sample();
    const cells = structuredClone(wb.activeSheet.toJSON().cells);
    setTableFilter(wb, 0, 'financials', 1, ['Canada']);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([2, 4]);
    setTableFilter(wb, 0, 'financials', 0, ['Government']);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([2, 3, 4]);
    wb.undo();
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([2, 4]);
    wb.redo();
    setTableFilter(wb, 0, 'financials', 0, null);
    setTableFilter(wb, 0, 'financials', 1, ['']);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([1, 2, 3]);
    setTableFilter(wb, 0, 'financials', 1, []);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([1, 2, 3, 4]);
    setTableFilter(wb, 0, 'financials', 1, null);
    expect(tableFilteredRows(wb.activeSheet).size).toBe(0);
    expect(wb.activeSheet.toJSON().cells).toEqual(cells);
  });

  it('exports native filter metadata and hidden rows and restores row heights on reimport', () => {
    const wb = sample();
    wb.activeSheet.rowHeights.set(2, 35);
    wb.activeSheet.rowHeights.set(4, 0);
    setTableFilter(wb, 0, 'financials', 1, ['Canada', '']);
    const bytes = writeXlsx(wb), parts = unzipSync(bytes);
    expect(strFromU8(parts['xl/tables/table1.xml']!)).toContain('<filterColumn colId="1"><filters blank="1"><filter val="Canada"/>');
    expect(strFromU8(parts['xl/worksheets/sheet1.xml']!)).toContain('<row r="3" hidden="1" ht="26.25" customHeight="1">');
    const back = readXlsx(bytes);
    expect(back.activeSheet.tables).toEqual(wb.activeSheet.tables);
    expect(back.activeSheet.rowHeights.get(2)).toBe(35);
    expect(back.activeSheet.rowHeights.get(4)).toBe(0);
    expect([...tableFilteredRows(back.activeSheet)]).toEqual([2]);
    setTableFilter(back, 0, 'financials', 1, null);
    expect(tableFilteredRows(back.activeSheet).size).toBe(0);
    expect(back.activeSheet.rowHeights.get(2)).toBe(35);
    expect(Workbook.fromJSON(wb.toJSON()).activeSheet.tables).toEqual(wb.activeSheet.tables);
  });

  it('re-evaluates visibility after editing and sorting records', () => {
    const wb = sample();
    setTableFilter(wb, 0, 'financials', 1, ['Canada']);
    setValues(wb, 0, 2, 1, [['Canada']]);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([4]);
    sortRange(wb, 0, toRange('A1:C5'), [{ col: 2, descending: true }], { hasHeader: true });
    expect(wb.getValue(0, 1, 2)).toBe(30);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([4]);
  });

  it('moves filter columns with structural changes and removes deleted criteria', () => {
    const wb = sample();
    setTableFilter(wb, 0, 'financials', 1, ['Canada']);
    wb.apply({ type: 'insertCols', sheet: 'Sheet1', index: 1, count: 1 });
    expect(wb.activeSheet.tables[0]?.filters?.[0]?.column).toBe(2);
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([2, 4]);
    wb.apply({ type: 'deleteCols', sheet: 'Sheet1', index: 2, count: 1 });
    expect(tableFilteredRows(wb.activeSheet).size).toBe(0);
    wb.undo();
    expect([...tableFilteredRows(wb.activeSheet)]).toEqual([2, 4]);
  });

  it('rejects invalid filter columns without mutation', () => {
    const wb = sample(), before = wb.toJSON();
    expect(() => setTableFilter(wb, 0, 'financials', 3, ['Canada'])).toThrow(/filter/);
    expect(wb.toJSON()).toEqual(before);
  });

  it('imports Excel built-in table styles and relationships as used by Financial Sample', () => {
    const parts = unzipSync(writeXlsx(sample()));
    parts['xl/tables/table1.xml'] = strToU8(strFromU8(parts['xl/tables/table1.xml']!).replace('SheetsTableStyle1', 'TableStyleLight9').replace(' headerRowCount="1"', ''));
    parts['xl/styles.xml'] = strToU8(strFromU8(parts['xl/styles.xml']!).replace(/<tableStyles\b[\s\S]*?<\/tableStyles>/, ''));
    const wb = readXlsx(zipSync(parts));
    expect(wb.activeSheet.tables[0]).toMatchObject({ name: 'financials', range: 'A1:C5', headerRow: true, bandedRows: true });
    expect(wb.getCellStyle(0, 0, 0)?.fill).toBeTruthy();
    expect(wb.getCellStyle(0, 1, 0)?.fill).not.toBe(wb.getCellStyle(0, 2, 0)?.fill);
  });
});
