import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { cellKey, insertRows, readXlsx, setValues, Workbook, writeXlsx } from '../src/index.js';

describe('notes and hyperlinks', () => {
  it('reads notes and hyperlinks and keeps them when saving unchanged parts', () => {
    const original = readFileSync(new URL('./fixtures/excel-features.xlsx', import.meta.url));
    const wb = readXlsx(original);
    const sheet = wb.activeSheet;
    expect(sheet.notes.get(cellKey(1, 0))).toEqual({ author: 'Author', text: 'Keep this comment' });
    expect(sheet.links.get(cellKey(2, 0))).toEqual({ target: 'https://example.com/original' });
    setValues(wb, 0, 1, 1, [[11]]);
    const files = unzipSync(writeXlsx(wb));
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('<hyperlink');
    expect(files['xl/comments/comment1.xml']).toBeDefined();
  });

  it('writes hyperlinks in new workbooks and moves notes and links with rows', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Docs'], ['Jump']]);
    wb.activeSheet.links.set(cellKey(0, 0), { target: 'https://example.com/a?b=1&c=2', tooltip: 'Open docs' });
    wb.activeSheet.links.set(cellKey(1, 0), { location: 'Sheet1!C5' });
    const bytes = writeXlsx(wb);
    const files = unzipSync(bytes);
    const xml = strFromU8(files['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<hyperlinks><hyperlink ref="A1" r:id="rIdLink1" tooltip="Open docs"/><hyperlink ref="A2" location="Sheet1!C5"/></hyperlinks>');
    expect(strFromU8(files['xl/worksheets/_rels/sheet1.xml.rels']!)).toContain('Target="https://example.com/a?b=1&amp;c=2" TargetMode="External"');
    const back = readXlsx(bytes).activeSheet;
    expect(back.links.get(cellKey(0, 0))).toEqual({ target: 'https://example.com/a?b=1&c=2', tooltip: 'Open docs' });
    back.notes.set(cellKey(1, 0), { text: 'note' });
    const copy = Workbook.fromJSON({ sheets: [back.toJSON()] });
    insertRows(copy, 0, 0, 2);
    expect(copy.activeSheet.links.get(cellKey(2, 0))?.target).toBe('https://example.com/a?b=1&c=2');
    expect(copy.activeSheet.notes.get(cellKey(3, 0))).toEqual({ text: 'note' });
    copy.undo();
    expect(copy.activeSheet.notes.get(cellKey(1, 0))).toEqual({ text: 'note' });
  });
});
