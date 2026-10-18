import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { applyStyle, readXlsx, setValues, toRange, writeXlsx } from '../src/index.js';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
function minimalXlsx(overrides: Record<string, string> = {}): Uint8Array {
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>A1+B1</f><v>5</v></c></row></sheetData></worksheet>`,
  };
  return zipSync(Object.fromEntries(Object.entries({ ...parts, ...overrides }).map(([k, v]) => [k, strToU8(v)])));
}

describe('minimal worksheets', () => {
  it('reads a worksheet that contains nothing but sheetData', () => {
    const wb = readXlsx(minimalXlsx());
    const sheet = wb.sheets[0]!;
    expect(sheet.name).toBe('Tabelle1');
    expect(sheet.getValue(0, 0)).toBe(2);
    expect(sheet.getCell(0, 2)?.formula).toBe('A1+B1');
  });
});

const sheetData = '<sheetData><row r="1"><c r="A1"><v>2</v></c></row></sheetData>';
const stylesRel = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`;

describe('empty and whitespace-only elements', () => {
  // The XML parser turns elements without attributes or children into strings ('' or whitespace).
  const variants: Record<string, Record<string, string>> = {
    'whitespace around sheetData': { 'xl/worksheets/sheet1.xml': `<worksheet xmlns="${ns}">\n  ${sheetData}\n</worksheet>` },
    'worksheet without namespace': { 'xl/worksheets/sheet1.xml': `<worksheet>${sheetData}</worksheet>` },
    'empty collections': { 'xl/worksheets/sheet1.xml': `<worksheet xmlns="${ns}"><sheetViews/><sheetFormatPr/><cols/>${sheetData}<mergeCells/><conditionalFormatting/><dataValidations/><hyperlinks/><tableParts/></worksheet>` },
    'whitespace collections': { 'xl/worksheets/sheet1.xml': `<worksheet xmlns="${ns}"><sheetViews> </sheetViews><cols> </cols>${sheetData}<mergeCells> </mergeCells><extLst> </extLst></worksheet>` },
    'empty style collections': { 'xl/_rels/workbook.xml.rels': stylesRel, 'xl/styles.xml': `<styleSheet xmlns="${ns}"><numFmts/><fonts/><fills/><borders/><cellXfs/><dxfs/><tableStyles/></styleSheet>` },
    'whitespace fonts': { 'xl/_rels/workbook.xml.rels': stylesRel, 'xl/styles.xml': `<styleSheet xmlns="${ns}"><fonts count="1"><font> </font></fonts><fills><fill/></fills><borders><border> </border></borders><cellXfs count="1"><xf/></cellXfs><dxfs><dxf><font> </font><fill> </fill></dxf></dxfs></styleSheet>` },
    'empty style sheet': { 'xl/_rels/workbook.xml.rels': stylesRel, 'xl/styles.xml': '<styleSheet/>' },
    'empty definedNames and bookViews': { 'xl/workbook.xml': `<workbook xmlns="${ns}" xmlns:r="${rel}"><bookViews/><sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets><definedNames/><calcPr/></workbook>` },
  };
  for (const [name, overrides] of Object.entries(variants)) {
    it(`opens and saves edits with ${name}`, () => {
      const bytes = minimalXlsx(overrides);
      const wb = readXlsx(bytes);
      expect(wb.getValue(0, 0, 0)).toBe(2);
      expect(writeXlsx(wb)).toEqual(bytes);
      setValues(wb, 0, 1, 0, [[7]]);
      applyStyle(wb, 0, toRange('B2'), { italic: true });
      // Either the edit is saved or it is refused by the protection, never an uncontrolled crash.
      let saved: Uint8Array | undefined;
      try { saved = writeXlsx(wb); } catch (e) { expect((e as Error).message).toMatch(/^XLSX-Schutz: /); }
      if (saved) expect(readXlsx(saved).getValue(0, 1, 0)).toBe(7);
    });
  }

  it('reports a workbook part without sheets as a clear error or an empty, protected workbook', () => {
    expect(() => readXlsx(minimalXlsx({ 'xl/workbook.xml': '<workbook/>' }))).toThrow(/workbook part missing/);
    const wb = readXlsx(minimalXlsx({ 'xl/workbook.xml': `<workbook xmlns="${ns}"><sheets/></workbook>` }));
    expect(wb.sheets).toHaveLength(1);
    setValues(wb, 0, 0, 0, [[1]]);
    expect(() => writeXlsx(wb)).toThrow(/XLSX-Schutz/);
  });
});
