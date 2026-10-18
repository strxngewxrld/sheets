import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, describe, expect, it } from 'vitest';
import { applyStyle, readXlsx, setValues, setXlsxMessageLanguage, sortRange, toRange, Workbook, writeXlsx, XlsxProtectionError } from '../src/index.js';
import { exactNumberText } from '../src/io/xlsx-preserve.js';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/excel-features.xlsx', import.meta.url)));

/** A valid package without styles, shared strings or document properties. */
function minimalParts(sheetXml = `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><v>3</v></c><c r="C1"><f>A1+B1</f><v>5</v></c></row></sheetData></worksheet>`, extraSheets = ''): Record<string, Uint8Array> {
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/>${extraSheets}</sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    'xl/worksheets/sheet1.xml': sheetXml,
  };
  return Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)]));
}

/** Names of entries whose bytes differ, plus added and removed entries. */
function changedParts(before: Record<string, Uint8Array>, after: Record<string, Uint8Array>): string[] {
  const names = new Set([...Object.keys(before), ...Object.keys(after)]);
  return [...names].filter((name) => !before[name] || !after[name] || Buffer.compare(before[name]!, after[name]!) !== 0).sort();
}

afterEach(() => setXlsxMessageLanguage('de'));

describe('formatting files without a style table', () => {
  it('adds a minimal style part and leaves every other entry untouched', () => {
    const parts = minimalParts();
    const wb = readXlsx(zipSync(parts));
    applyStyle(wb, 0, toRange('A1'), { bold: true });
    const saved = writeXlsx(wb), out = unzipSync(saved);
    expect(changedParts(parts, out)).toEqual(['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    expect(strFromU8(out['xl/_rels/workbook.xml.rels']!)).toContain(`Type="${rel}/styles" Target="styles.xml"`);
    expect(strFromU8(out['[Content_Types].xml']!)).toContain('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>');
    // The original relationship and content type entries are kept verbatim.
    expect(strFromU8(out['xl/_rels/workbook.xml.rels']!)).toContain(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', ''));
    const back = readXlsx(saved);
    expect(back.getCellStyle(0, 0, 0)?.bold).toBe(true);
    expect(back.getCellStyle(0, 0, 1)?.bold).toBeUndefined();
    expect(back.getValue(0, 0, 2)).toBe(5);
    expect(back.sheets[0]!.getCell(0, 2)?.formula).toBe('A1+B1');
  });

  it('still refuses when a style relationship exists but its part is missing', () => {
    const parts = minimalParts();
    parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', `<Relationship Id="rId9" Type="${rel}/styles" Target="styles.xml"/></Relationships>`));
    const wb = readXlsx(zipSync(parts));
    applyStyle(wb, 0, toRange('A1'), { bold: true });
    expect(() => writeXlsx(wb)).toThrow(/Style-Tabelle/);
  });
});

describe('adding sheets to preserved files', () => {
  it('appends a new worksheet part without touching original sheets', () => {
    const parts = minimalParts();
    const wb = readXlsx(zipSync(parts));
    const sheet = wb.addSheet('Neu & <frei>');
    setValues(wb, sheet.name, 0, 0, [['Text', 4], ['=Tabelle1!A1*2', true]]);
    applyStyle(wb, sheet.name, toRange('A1'), { bold: true });
    wb.apply({ type: 'setColWidths', sheet: sheet.name, sizes: [{ index: 0, size: 120 }] });
    const saved = writeXlsx(wb), out = unzipSync(saved);
    expect(changedParts(parts, out)).toEqual(['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet2.xml']);
    const workbook = strFromU8(out['xl/workbook.xml']!);
    expect(workbook).toContain('<sheet name="Tabelle1" sheetId="1" r:id="rId1"/><sheet name="Neu &amp; &lt;frei&gt;" sheetId="2" r:id="rIdSheets1"/>');
    const back = readXlsx(saved);
    expect(back.sheets.map((s) => s.name)).toEqual(['Tabelle1', 'Neu & <frei>']);
    expect(back.getValue(1, 0, 0)).toBe('Text');
    expect(back.getValue(1, 1, 0)).toBe(4);
    expect(back.sheets[1]!.getCell(1, 0)?.formula).toBe('Tabelle1!A1*2');
    expect(back.getValue(1, 1, 1)).toBe(true);
    expect(back.getCellStyle(1, 0, 0)?.bold).toBe(true);
    expect(back.sheets[1]!.colWidth(0)).toBe(120);
  });

  it('keeps every original entry of an independently authored workbook', () => {
    const bytes = fixture(), original = unzipSync(bytes);
    const wb = readXlsx(bytes);
    setValues(wb, wb.addSheet('Extra').name, 0, 0, [['x', 1]]);
    const out = unzipSync(writeXlsx(wb));
    expect(changedParts(original, out)).toEqual(['[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet3.xml']);
    const workbook = strFromU8(out['xl/workbook.xml']!);
    // Defined names with localSheetId and the very hidden sheet are unchanged.
    expect(workbook).toContain(`<definedName name="LocalName" localSheetId="0">'Data'!$B$2</definedName>`);
    expect(workbook).toContain('<sheet name="Hidden" sheetId="2" state="veryHidden" r:id="rId2" /><sheet name="Extra" sheetId="3"');
    expect(readXlsx(writeXlsx(wb)).getValue('Extra', 0, 1)).toBe(1);
  });

  it('returns the original bytes after undoing the new sheet', () => {
    const bytes = zipSync(minimalParts());
    const wb = readXlsx(bytes);
    wb.addSheet('Weg');
    wb.undo();
    expect(writeXlsx(wb)).toEqual(bytes);
  });

  it('blocks names of hidden chart sheets, inserted positions, renames and rich new sheets', () => {
    const chart = minimalParts(undefined, '<sheet name="Diagramm" sheetId="2" r:id="rId7"/>');
    chart['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(chart['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', `<Relationship Id="rId7" Type="${rel}/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>`));
    chart['xl/chartsheets/sheet1.xml'] = strToU8(`<chartsheet xmlns="${ns}"/>`);
    const withChart = readXlsx(zipSync(chart));
    withChart.addSheet('diagramm');
    expect(() => writeXlsx(withChart)).toThrow(/existiert bereits/);

    const inserted = readXlsx(zipSync(minimalParts()));
    inserted.apply({ type: 'addSheet', name: 'Vorne', index: 0 });
    expect(() => writeXlsx(inserted)).toThrow(/XLSX-Schutz/);

    const renamed = readXlsx(zipSync(minimalParts()));
    renamed.addSheet('Neu');
    renamed.apply({ type: 'renameSheet', from: 'Tabelle1', to: 'Umbenannt' });
    expect(() => writeXlsx(renamed)).toThrow(/XLSX-Schutz/);

    const merged = readXlsx(zipSync(minimalParts()));
    merged.addSheet('Neu');
    merged.apply({ type: 'setMerges', sheet: 'Neu', merges: ['A1:B1'] });
    expect(() => writeXlsx(merged)).toThrow(/XLSX-Schutz/);

    const removed = readXlsx(zipSync(minimalParts()));
    removed.addSheet('Neu');
    removed.apply({ type: 'removeSheet', name: 'Tabelle1' });
    expect(() => writeXlsx(removed)).toThrow(/XLSX-Schutz/);
  });
});

describe('sorting preserved files', () => {
  const styles = `<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/><vertAlign val="superscript"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment textRotation="45"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>`;
  const rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Wert</t></is></c><c r="C1" t="inlineStr"><is><t>Doppelt</t></is></c></row>' +
    '<row r="2"><c r="A2" s="1" t="inlineStr"><is><t>b</t></is></c><c r="B2"><v>1.5</v></c><c r="C2"><f>B2*2</f><v>3</v></c></row>' +
    '<row r="3"><c r="A3" t="inlineStr"><is><t>c</t></is></c><c r="B3"><v>2</v></c><c r="C3"><f>B3*2</f><v>4</v></c></row>' +
    '<row r="4"><c r="A4" t="inlineStr"><is><t>a</t></is></c><c r="B4"><v>3</v></c><c r="C4"><f>B4*2</f><v>6</v></c></row>';
  function sortable(sheetData = rows, extra = ''): Record<string, Uint8Array> {
    const parts = minimalParts(`<worksheet xmlns="${ns}"><dimension ref="A1:C4"/><sheetData>${sheetData}</sheetData>${extra}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
    parts['xl/styles.xml'] = strToU8(styles);
    parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', `<Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`));
    return parts;
  }

  it('moves values, row-local formulas and exact format records', () => {
    const parts = sortable();
    const wb = readXlsx(zipSync(parts));
    sortRange(wb, 0, toRange('A1:C4'), [{ col: 0 }], { hasHeader: true });
    const saved = writeXlsx(wb), out = unzipSync(saved);
    expect(changedParts(parts, out)).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    const xml = strFromU8(out['xl/worksheets/sheet1.xml']!);
    // "b" kept its own format record (rotation and superscript are not modeled) and moved to row 3.
    expect(xml).toContain('<c r="A3" t="inlineStr" s="1"><is><t xml:space="preserve">b</t></is></c>');
    expect(xml).toContain('<c r="A2" s="0" t="inlineStr"><is><t xml:space="preserve">a</t></is></c>');
    // Row-local formulas read the same after the move, so their XML stays untouched.
    expect(xml).toContain('<c r="C2"><f>B2*2</f><v>3</v></c>');
    expect(xml).toContain('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
    const back = readXlsx(saved);
    expect([1, 2, 3].map((r) => back.getValue(0, r, 0))).toEqual(['a', 'b', 'c']);
    expect([1, 2, 3].map((r) => back.getValue(0, r, 2))).toEqual([6, 3, 4]);
    expect(back.getCellStyle(0, 2, 0)?.bold).toBe(true);
    expect(back.getCellStyle(0, 1, 0)?.bold).toBeUndefined();
  });

  it('sorts values in the independently authored workbook', () => {
    const bytes = fixture(), original = unzipSync(bytes);
    const wb = readXlsx(bytes);
    sortRange(wb, 'Data', toRange('B1:C3'), [{ col: 1, descending: true }], { hasHeader: true });
    const saved = writeXlsx(wb);
    expect(changedParts(original, unzipSync(saved))).toEqual(['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    const back = readXlsx(saved);
    expect([back.getValue('Data', 1, 1), back.getValue('Data', 2, 1)]).toEqual([20, 10]);
    expect(back.sheets[0]!.getCell(1, 2)?.formula).toBe('B2*2');
    expect(back.getValue('Data', 1, 0)).toBe('Original');
  });

  it('returns the original bytes after undoing a sort', () => {
    const bytes = zipSync(sortable());
    const wb = readXlsx(bytes);
    sortRange(wb, 0, toRange('A1:C4'), [{ col: 0 }], { hasHeader: true });
    wb.undo();
    expect(writeXlsx(wb)).toEqual(bytes);
  });

  const blocked: [string, () => Record<string, Uint8Array>, RegExp][] = [
    ['formulas referring to other rows', () => sortable(rows.replace('<f>B3*2</f>', '<f>B2*2</f>')), /andere Zeilen/],
    ['unknown functions', () => sortable(rows.replace('<f>B3*2</f>', '<f>_xlfn.UNIMPLEMENTED(B3)</f>')), /unbekannten Funktionen/],
    ['numbers stored more precisely than a double', () => sortable(rows.replace('<v>2</v>', '<v>9007199254740993</v>')), /Genauigkeit/],
    ['cells that look alike but use different records', () => sortable(rows.replace('<c r="A3" t="inlineStr">', '<c r="A3" s="2" t="inlineStr">').replace('<c r="A2" s="1" t="inlineStr">', '<c r="A2" s="2" t="inlineStr">').replace('<c r="A4" t="inlineStr">', '<c r="A4" s="1" t="inlineStr">')), /unterschiedlichen Excel-Formaten/],
    ['data validation', () => sortable(rows, '<dataValidations count="1"><dataValidation type="list" sqref="A2:A9"><formula1>"a,b,c"</formula1></dataValidation></dataValidations>'), /Datenüberprüfung/],
    ['moved shared formulas', () => sortable(rows.replace('<f>B2*2</f>', '<f t="shared" ref="C2:C3" si="0">B2*2</f>').replace('<f>B3*2</f>', '<f t="shared" si="0"/>').replace('<f>B4*2</f>', '<f>B4*3</f>')), /Geteilte Formel/],
    ['rich text', () => sortable(rows.replace('<is><t>c</t></is>', '<is><r><rPr><b/></rPr><t>c</t></r></is>')), /Rich-Text/],
  ];
  for (const [name, parts, reason] of blocked) {
    it(`blocks sorting with ${name}`, () => {
      const wb = readXlsx(zipSync(parts()));
      sortRange(wb, 0, toRange('A1:C4'), [{ col: 0 }], { hasHeader: true });
      expect(() => writeXlsx(wb)).toThrow(reason);
    });
  }

  it('blocks sorting cells with notes or hyperlinks', () => {
    const wb = readXlsx(fixture());
    sortRange(wb, 'Data', toRange('A1:C3'), [{ col: 1, descending: true }], { hasHeader: true });
    expect(() => writeXlsx(wb)).toThrow(/Notizen oder Hyperlinks/);
  });

  it('keeps cut and paste blocked as a structural move', () => {
    const wb = readXlsx(zipSync(sortable()));
    wb.apply({ type: 'setCells', sheet: 'Tabelle1', cells: [{ row: 5, col: 0, data: { v: 'x' } }], reordersCells: true });
    expect(() => writeXlsx(wb)).toThrow(/XLSX-Schutz/);
  });
});

describe('protection messages', () => {
  it('are available in English and German with a stable error type', () => {
    const wb = readXlsx(fixture());
    wb.apply({ type: 'insertRows', sheet: 'Data', index: 1, count: 1 });
    setXlsxMessageLanguage('en');
    let error: unknown;
    try { writeXlsx(wb); } catch (e) { error = e; }
    expect(error).toBeInstanceOf(XlsxProtectionError);
    expect((error as XlsxProtectionError).message).toMatch(/^XLSX protection: Structure, table, chart or rule change .* The original file was not overwritten\.$/);
    expect((error as XlsxProtectionError).reason).toMatch(/^Structure, table/);
    setXlsxMessageLanguage('de-DE');
    expect(() => writeXlsx(wb)).toThrow(/^XLSX-Schutz: Struktur-/);
  });
});

describe('exact number text', () => {
  it('accepts renderings of a double and rejects extra digits', () => {
    for (const text of ['0', '-0', '1.5', '1.10', '1E-3', '0.30000000000000004', '1.1000000000000001', '123456789012345', '-2.5e10', '.5']) expect(exactNumberText(text), text).toBe(true);
    for (const text of ['9007199254740993', '0.1000000000000000000001', '1e400', 'abc', '', '1,5']) expect(exactNumberText(text), text).toBe(false);
  });
});

describe('workbooks created in the editor', () => {
  it('are unaffected by sort tracking', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['b'], ['a']]);
    sortRange(wb, 0, toRange('A1:A2'), [{ col: 0 }]);
    expect(readXlsx(writeXlsx(wb)).getValue(0, 0, 0)).toBe('a');
  });
});
