import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { readXlsx, setCellInput, setValues, writeXlsx } from '../src/index.js';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const sheet = (rows: string, extra = '') => `<worksheet xmlns="${ns}"><sheetData>${rows}</sheetData>${extra}</worksheet>`;

/** A two-sheet package as Excel would write it; caches are what Excel last computed. */
function book(sheet1: string, sheet2: string): Record<string, Uint8Array> {
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Input" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': sheet1,
    'xl/worksheets/sheet2.xml': sheet2,
  };
  return Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)]));
}

const INPUT = sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c></row>');
const save = (parts: Record<string, Uint8Array>, edit: (wb: ReturnType<typeof readXlsx>) => void) => {
  const wb = readXlsx(zipSync(parts));
  edit(wb);
  return unzipSync(writeXlsx(wb));
};
const text = (parts: Record<string, Uint8Array>, name: string) => strFromU8(parts[name]!);

describe('formula caches follow edits', () => {
  it('refreshes results on other sheets and keeps everything else of the cell', () => {
    const report = sheet('<row r="1"><c r="A1" s="0"><f>Input!A1*10</f><v>20</v></c><c r="B1"><f>IF(Input!B1&gt;5,"high","low")</f><v>low</v></c><c r="C1"><f>Input!A1/0</f><v>#DIV/0!</v></c><c r="D1"><f>Input!B1=3</f><v>1</v></c></row>');
    const parts = book(INPUT, report.replace('<v>low</v>', '<v>low</v>').replace('<c r="B1">', '<c r="B1" t="str">').replace('<c r="C1">', '<c r="C1" t="e">').replace('<c r="D1">', '<c r="D1" t="b">'));
    const out = save(parts, (wb) => setValues(wb, 'Input', 0, 0, [[7, 9]]));
    const xml = text(out, 'xl/worksheets/sheet2.xml');
    expect(xml).toContain('<c r="A1" s="0"><f>Input!A1*10</f><v>70</v></c>');
    expect(xml).toContain('<c r="B1" t="str"><f>IF(Input!B1&gt;5,"high","low")</f><v>high</v></c>');
    expect(xml).toContain('<c r="C1" t="e"><f>Input!A1/0</f><v>#DIV/0!</v></c>');
    expect(xml).toContain('<c r="D1" t="b"><f>Input!B1=3</f><v>0</v></c>');
    expect(Object.keys(out).filter((name) => Buffer.compare(out[name]!, parts[name]!) !== 0).sort()).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
  });

  it('does not touch formulas whose result did not change', () => {
    const report = sheet('<row r="1"><c r="A1"><f>Input!B1*1</f><v>3</v></c></row>');
    const parts = book(INPUT, report);
    const out = save(parts, (wb) => setValues(wb, 'Input', 0, 0, [[7]]));
    expect(Buffer.compare(out['xl/worksheets/sheet2.xml']!, parts['xl/worksheets/sheet2.xml']!)).toBe(0);
  });

  it('keeps the cache of formulas with functions the engine does not know', () => {
    const input = sheet('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1+FUTUREFN(A1)</f><v>99</v></c><c r="C1"><f>A1*2</f><v>4</v></c></row>');
    const parts = book(input, sheet(''));
    const out = save(parts, (wb) => setValues(wb, 'Input', 0, 0, [[5]]));
    const xml = text(out, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="B1"><f>A1+FUTUREFN(A1)</f><v>99</v></c>');
    expect(xml).toContain('<c r="C1"><f>A1*2</f><v>10</v></c>');
  });

  it('refreshes shared formulas without changing their definition', () => {
    const input = sheet('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>');
    const out = save(book(input, sheet('')), (wb) => setValues(wb, 'Input', 1, 0, [[10]]));
    const xml = text(out, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c>');
    expect(xml).toContain('<c r="B2"><f t="shared" si="0"/><v>20</v></c>');
  });

  it('caches the computed result of new and edited formulas', () => {
    const out = save(book(INPUT, sheet('')), (wb) => {
      setCellInput(wb, 'Input', 0, 2, '=A1+B1');
      setCellInput(wb, 'Input', 0, 3, '="x"&A1');
      setCellInput(wb, 'Input', 0, 4, '=A1+FUTUREFN(1)');
    });
    const xml = text(out, 'xl/worksheets/sheet1.xml');
    expect(xml).toContain('<c r="C1"><f>A1+B1</f><v>5</v></c>');
    expect(xml).toContain('<c r="D1" t="str"><f>&quot;x&quot;&amp;A1</f><v>x2</v></c>');
    // An unknown function gets no invented cache.
    expect(xml).toContain('<c r="E1"><f>A1+FUTUREFN(1)</f></c>');
  });

  it('saves edits that change results on a protected sheet without touching that sheet', () => {
    const report = sheet('<row r="1"><c r="A1"><f>Input!A1*10</f><v>20</v></c></row>', '<sheetProtection sheet="1" objects="1" scenarios="1"/>');
    const parts = book(INPUT, report);
    const out = save(parts, (wb) => setValues(wb, 'Input', 0, 0, [[7]]));
    expect(Buffer.compare(out['xl/worksheets/sheet2.xml']!, parts['xl/worksheets/sheet2.xml']!)).toBe(0);
    expect(text(out, 'xl/worksheets/sheet1.xml')).toContain('<c r="A1"><v>7</v></c>');
  });

  it('still refuses edits on the protected sheet itself', () => {
    const report = sheet('<row r="1"><c r="A1"><v>1</v></c></row>', '<sheetProtection sheet="1"/>');
    const wb = readXlsx(zipSync(book(INPUT, report)));
    setValues(wb, 'Report', 0, 0, [[2]]);
    expect(() => writeXlsx(wb)).toThrow(/Blattschutz|Sheet protection/);
  });
});
