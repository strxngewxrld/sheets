import { applyBorders, applyStyle, insertChartFromRange, setValues, toRange, Workbook } from '@sheets/core';

/** Demo workbook showing formulas, formatting and a chart. */
export function createSampleWorkbook(): Workbook {
  const wb = new Workbook({ sheetName: 'Umsatz' });
  const s = 'Umsatz';
  wb.transaction(() => {
    setValues(wb, s, 0, 0, [
      ['Monat', 'Umsatz', 'Kosten', 'Gewinn', 'Marge'],
      ['Jan', 12400, 8100, '=B2-C2', '=D2/B2'],
      ['Feb', 13800, 8600],
      ['Mär', 15100, 9200],
      ['Apr', 14300, 9000],
      ['Mai', 16900, 9800],
      ['Jun', 18200, 10400],
      ['Summe', '=SUM(B2:B7)', '=SUM(C2:C7)', '=SUM(D2:D7)', '=D8/B8'],
    ]);
    // Fill the formulas down like a user would with the fill handle.
    setValues(wb, s, 2, 3, [2, 3, 4, 5, 6].map((r) => [`=B${r + 1}-C${r + 1}`, `=D${r + 1}/B${r + 1}`]));
    applyStyle(wb, s, toRange('A1:E1'), { bold: true, fill: '#DDEBF7', hAlign: 'center' });
    applyStyle(wb, s, toRange('A8:E8'), { bold: true });
    applyStyle(wb, s, toRange('B2:D8'), { numFmt: '#,##0 "€"' });
    applyStyle(wb, s, toRange('E2:E8'), { numFmt: '0.0%' });
    applyBorders(wb, s, toRange('A8:E8'), 'outline');
    wb.apply({ type: 'setColWidths', sheet: s, sizes: [1, 2, 3].map((index) => ({ index, size: 90 })) });
    wb.apply({ type: 'setFrozen', sheet: s, rows: 1, cols: 0 });
    setValues(wb, s, 9, 0, [
      ['Bester Monat', '=INDEX(A2:A7,MATCH(MAX(B2:B7),B2:B7,0))'],
      ['Ø Umsatz', '=AVERAGE(B2:B7)'],
      ['Monate > 15k', '=COUNTIF(B2:B7,">15000")'],
    ]);
    applyStyle(wb, s, toRange('B11'), { numFmt: '#,##0 "€"' });
    insertChartFromRange(wb, s, toRange('A1:C7'), 'column', {
      title: 'Umsatz vs. Kosten',
      anchor: { from: { row: 1, col: 6 }, to: { row: 17, col: 13 } },
    });
  });
  wb.clearHistory();
  return wb;
}
