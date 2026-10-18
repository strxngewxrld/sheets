import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { applyStyle, formatTable, readXlsx, setCellInput, setValues, sortRange, toRange, Workbook, writeXlsx, xlsxToData } from '../src/index.js';

const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
function fixture() {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [['Name', 'Amount'], ['Original', 10], ['Untouched', '=B2*2']]);
  const parts = unzipSync(writeXlsx(wb));
  const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!)
    .replace('<c r="A2"', '<c r="A2" s="1"')
    .replace('</worksheet>', '<sheetProtection sheet="0" objects="0"/>' + '</worksheet>')
    .replace('<sheetProtection sheet="0" objects="0"/>', '')
    .replace('</worksheet>', '<dataValidations count="1"><dataValidation type="list" sqref="D2:D9"><formula1>"A,B,C"</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="A1" r:id="rHyper"/></hyperlinks><drawing r:id="rPicture"/><legacyDrawing r:id="rComments"/><extLst><ext uri="custom"><u:payload xmlns:u="urn:unknown">keep me</u:payload></ext></extLst></worksheet>');
  parts['xl/worksheets/sheet1.xml'] = strToU8(sheet);
  parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rHyper" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/><Relationship Id="rComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/comments.vml"/></Relationships>');
  parts['xl/drawings/drawing1.xml'] = strToU8('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:extLst/></xdr:wsDr>');
  parts['xl/drawings/comments.vml'] = strToU8('<xml><custom-comment>Original comment</custom-comment></xml>');
  parts['xl/media/image1.png'] = Uint8Array.from([137, 80, 78, 71, 1, 2, 3, 255]);
  parts['xl/pivotCache/pivotCacheDefinition1.xml'] = strToU8('<pivotCacheDefinition xmlns="'+ns+'" recordCount="1" refreshOnLoad="1"/>');
  parts['xl/externalLinks/externalLink1.xml'] = strToU8('<externalLink xmlns="'+ns+'"><externalBook/></externalLink>');
  parts['customXml/item1.xml'] = strToU8('<unknown xmlns="urn:vendor" a="&amp;">secret preserved</unknown>');
  parts['xl/connections.xml'] = strToU8('<connections xmlns="'+ns+'"><connection id="1" name="Original Connection"/></connections>');
  parts['xl/styles.xml'] = strToU8(`<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/><charset val="204"/><vertAlign val="superscript"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"><color rgb="FF0000FF"/></diagonal></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" quotePrefix="1"><alignment textRotation="45" shrinkToFit="1"/><protection locked="0"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  return { bytes: zipSync(parts), parts };
}

describe('lossless XLSX preservation', () => {
  it('preserves a real independently authored workbook with images, comments and advanced features', () => {
    const bytes = new Uint8Array(readFileSync(new URL('./fixtures/excel-features.xlsx', import.meta.url)));
    const wb = readXlsx(bytes), original = unzipSync(bytes);
    expect(writeXlsx(wb)).toEqual(bytes);
    setValues(wb, 'Data', 1, 1, [[42]]);
    applyStyle(wb, 'Data', toRange('A2'), { fill: '#FFCC00', bold: false, wrap: true });
    const output = unzipSync(writeXlsx(wb));
    expect(Object.keys(output).sort()).toEqual(Object.keys(original).sort());
    const changed = Object.keys(original).filter((name) => canonicalBytes(output[name]!) !== canonicalBytes(original[name]!));
    expect(changed.sort()).toEqual(['xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
    expect(readXlsx(writeXlsx(wb)).getValue('Data', 1, 1)).toBe(42);
  });
  it('returns the exact original archive for an unedited file, including JSON handoffs', () => {
    const { bytes } = fixture();
    const wb = readXlsx(bytes);
    expect(writeXlsx(wb)).toEqual(bytes);
    expect(writeXlsx(Workbook.fromJSON(wb.toJSON()))).toEqual(bytes);
    expect(writeXlsx(Workbook.fromJSON(xlsxToData(bytes)))).toEqual(bytes);
  });

  it('encodes literal OOXML escapes, control characters and UTF-16 text in new files', () => {
    const wb = new Workbook();
    const values = ['_x0041_', '_x005F_', 'A\u0000B\u000BC\rD', 'Müller 😀', '\uD800'];
    setValues(wb, 0, 0, 0, [values], { parse: false });
    const back = readXlsx(writeXlsx(wb));
    for (let col = 0; col < values.length; col++) expect(back.getValue(0, 0, col)).toBe(values[col]);
  });

  it('patches only edited cells and preserves every other package part', () => {
    const { bytes, parts } = fixture();
    const wb = readXlsx(bytes);
    setValues(wb, 0, 1, 0, [['Müller & <test> _x0041_\r\u000B', 42]], { parse: false });
    const saved = writeXlsx(wb), back = unzipSync(saved);
    expect(Object.keys(back).sort()).toEqual(Object.keys(parts).sort());
    for (const name of Object.keys(parts)) {
      if (name !== 'xl/worksheets/sheet1.xml' && name !== 'xl/workbook.xml') expect(back[name]).toEqual(parts[name]);
    }
    const xml = strFromU8(back['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<dataValidations count="1">');
    expect(xml).toContain('<u:payload xmlns:u="urn:unknown">keep me</u:payload>');
    // The unchanged formula keeps its XML; its cached result follows the edited input (B2 = 42).
    expect(xml).toContain('<c r="B3"><f>B2*2</f><v>84</v></c>');
    const reopened = readXlsx(saved);
    expect(reopened.getValue(0, 1, 0)).toBe('Müller & <test> _x0041_\r\u000B');
    expect(reopened.getValue(0, 2, 1)).toBe(84);
  });

  it('appends style resources while preserving unknown font, border and alignment properties', () => {
    const { bytes, parts } = fixture();
    const wb = readXlsx(bytes);
    applyStyle(wb, 0, toRange('A2'), { fill: '#FFCC00', bold: false, wrap: true, numFmt: '0.000 "kg"' });
    const saved = writeXlsx(wb), back = unzipSync(saved);
    const styles = strFromU8(back['xl/styles.xml']!);
    expect(styles).toContain('<alignment textRotation="45" shrinkToFit="1" wrapText="1"/>');
    expect(styles).toContain('<protection locked="0"/>');
    expect(styles).toContain('quotePrefix="1"');
    expect(styles.match(/<charset val="204"\/>/g)).toHaveLength(2);
    expect(styles).toContain('<vertAlign val="superscript"/>');
    expect(styles).toContain('<border diagonalUp="1">');
    expect(back['xl/media/image1.png']).toEqual(parts['xl/media/image1.png']);
    expect(readXlsx(saved).getCellStyle(0, 1, 0)).toMatchObject({ fill: '#FFCC00', bold: false, wrap: true, numFmt: '0.000 "kg"' });
  });

  it('inserts and clears cells without rebuilding neighboring XML', () => {
    const { bytes } = fixture();
    const wb = readXlsx(bytes);
    setValues(wb, 0, 0, 2, [['C1', 'D1']]);
    setValues(wb, 0, 1, 0, [[null]]);
    setValues(wb, 0, 8, 0, [['A9', true], ['A10', 99]]);
    const saved = writeXlsx(wb), back = readXlsx(saved);
    expect(back.getValue(0, 0, 2)).toBe('C1');
    expect(back.getValue(0, 0, 3)).toBe('D1');
    expect(back.getValue(0, 1, 0)).toBe(null);
    expect(back.getValue(0, 8, 0)).toBe('A9');
    expect(back.getValue(0, 8, 1)).toBe(true);
    expect(back.getValue(0, 9, 1)).toBe(99);
    expect(strFromU8(unzipSync(saved)['xl/worksheets/sheet1.xml']!)).toContain('dimension ref="A1:D10"');
  });

  it('blocks unsafe structural/metadata edits and allows saving again after undo', () => {
    const { bytes } = fixture();
    const wb = readXlsx(bytes);
    wb.apply({ type: 'insertRows', sheet: 'Sheet1', index: 100, count: 1 });
    expect(() => writeXlsx(wb)).toThrow(/XLSX-Schutz/);
    wb.undo();
    expect(writeXlsx(wb)).toEqual(bytes);
    formatTable(wb, 0, toRange('A1:B3'));
    expect(() => writeXlsx(wb)).toThrow(/Tabellen/);
    wb.undo();
    expect(writeXlsx(wb)).toEqual(bytes);
    sortRange(wb, 0, toRange('A1:B3'), [{ col: 1, descending: true }], { hasHeader: true });
    expect(() => writeXlsx(wb)).toThrow(/XLSX-Schutz/);
    wb.undo();
    expect(writeXlsx(wb)).toEqual(bytes);
  });

  it('resizes columns/rows without losing outline, inherited styles or other attributes', () => {
    const { parts } = fixture();
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!)
      .replace('<sheetData>', '<cols><col min="3" max="5" width="12.34" customWidth="1" style="1" outlineLevel="2" collapsed="1"/></cols><sheetData>')
      .replace('<row r="3">', '<row r="3" s="1" customFormat="1" outlineLevel="1" thickBot="1">'));
    const wb = readXlsx(zipSync(parts));
    wb.apply({ type: 'setColWidths', sheet: 'Sheet1', sizes: [{ index: 3, size: 150 }] });
    wb.apply({ type: 'setRowHeights', sheet: 'Sheet1', sizes: [{ index: 2, size: 36 }] });
    applyStyle(wb, 0, toRange('D4'), { wrap: true });
    const saved = writeXlsx(wb), output = unzipSync(saved);
    const xml = strFromU8(output['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<col min="3" max="3" width="12.34" customWidth="1" style="1" outlineLevel="2" collapsed="1"/>');
    expect(xml).toContain('<col min="4" max="4" width="20.71" customWidth="1" style="1" outlineLevel="2" collapsed="1"/>');
    expect(xml).toContain('<col min="5" max="5" width="12.34" customWidth="1" style="1" outlineLevel="2" collapsed="1"/>');
    expect(xml).toContain('<row r="3" s="1" customFormat="1" outlineLevel="1" thickBot="1" ht="27" customHeight="1">');
    expect(strFromU8(output['xl/styles.xml']!)).toContain('<alignment textRotation="45" shrinkToFit="1" wrapText="1"/>');
    expect(readXlsx(saved).activeSheet.rowHeight(2)).toBe(36);
  });

  it('still allows a complete rebuild of files proven to contain only native serializer output', () => {
    const original = new Workbook(); setValues(original, 0, 0, 0, [['Name', 'Value'], ['A', 2]]);
    const wb = readXlsx(writeXlsx(original));
    formatTable(wb, 0, toRange('A1:B2'));
    wb.apply({ type: 'insertRows', sheet: 'Sheet1', index: 1, count: 1 });
    expect(readXlsx(writeXlsx(wb)).activeSheet.tables[0]?.range).toBe('A1:B3');
  });

  it('protects array output in other rows without parsing unchanged rows', () => {
    const base = new Workbook();
    setValues(base, 0, 0, 0, [['x', 1], ['y', 2], ['z', 3], ['w', 4]]);
    const parts = unzipSync(writeXlsx(base));
    const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!);
    parts['xl/worksheets/sheet1.xml'] = strToU8(sheet
      .replace(/<c r="B1"[^]*?<\/c>/, '<c r="B1"><f t="array" ref="B1:B3">_xlfn.UNIMPLEMENTED(A1)</f><v>1</v></c>')
      .replace(/<c r="B4"[^]*?<\/c>/, '<c r="B4"><f t="array">_xlfn.UNIMPLEMENTED(A4)</f><v>4</v></c>'));
    const bytes = zipSync(parts);
    for (const [row, blocked] of [[2, true], [3, true], [3, false]] as const) {
      const wb = readXlsx(bytes);
      if (blocked) setCellInput(wb, 0, row, 1, '9');
      else setCellInput(wb, 0, row, 0, 'changed');
      if (blocked) expect(() => writeXlsx(wb)).toThrow(/Matrixformel/);
      else expect(readXlsx(writeXlsx(wb)).getValue(0, 3, 0)).toBe('changed');
    }
  });

  it('retains unsupported formula XML and rejects edits to array members and signed files', () => {
    const { parts } = fixture();
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('<c r="B3"><f>B2*2</f><v>20</v></c>', '<c r="B3"><f t="array" ref="B3:C3">_xlfn.UNIMPLEMENTED(B2)</f><v>123</v></c><c r="C3"><v>456</v></c>'));
    const bytes = zipSync(parts), wb = readXlsx(bytes);
    setCellInput(wb, 0, 1, 1, '43');
    const saved = unzipSync(writeXlsx(wb));
    expect(strFromU8(saved['xl/worksheets/sheet1.xml']!)).toContain('<f t="array" ref="B3:C3">_xlfn.UNIMPLEMENTED(B2)</f><v>123</v>');
    setCellInput(wb, 0, 2, 2, '99');
    expect(() => writeXlsx(wb)).toThrow(/Matrixformel/);
    parts['_xmlsignatures/sig1.xml'] = strToU8('<Signature/>');
    const signed = zipSync(parts), signedWb = readXlsx(signed);
    expect(writeXlsx(signedWb)).toEqual(signed);
    setCellInput(signedWb, 0, 1, 1, '43');
    expect(() => writeXlsx(signedWb)).toThrow(/Signaturen/);
  });

  it('keeps rich text intact and refuses flattening it through value edits', () => {
    const { parts } = fixture();
    const rich = '<si><r><rPr><b/></rPr><t>Original</t></r><r><t> text</t></r></si>';
    parts['xl/sharedStrings.xml'] = strToU8(strFromU8(parts['xl/sharedStrings.xml']!).replace('<si><t>Original</t></si>', rich));
    const wb = readXlsx(zipSync(parts));
    setValues(wb, 0, 1, 1, [[42]]);
    expect(unzipSync(writeXlsx(wb))['xl/sharedStrings.xml']).toEqual(parts['xl/sharedStrings.xml']);
    setValues(wb, 0, 1, 0, [['New text']]);
    expect(() => writeXlsx(wb)).toThrow(/Rich-Text/);
  });

  it('retains numeric lexical precision during unrelated edits and formatting', () => {
    const { parts } = fixture();
    parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('<v>10</v>', '<v>9007199254740993</v>'));
    const wb = readXlsx(zipSync(parts));
    setValues(wb, 0, 1, 0, [['New text']]);
    applyStyle(wb, 0, toRange('B2'), { fill: '#FFFF00' });
    expect(strFromU8(unzipSync(writeXlsx(wb))['xl/worksheets/sheet1.xml']!)).toContain('<v>9007199254740993</v>');
  });

  it('blocks protected sheets and 1904 date edits without changing the source', () => {
    const { parts } = fixture();
    parts['xl/workbook.xml'] = strToU8(strFromU8(parts['xl/workbook.xml']!).replace('<workbookPr/>', '<workbookPr date1904="1"/>'));
    const original = zipSync(parts), wb = readXlsx(original);
    expect(writeXlsx(wb)).toEqual(original);
    setValues(wb, 0, 1, 1, [[42]]);
    expect(() => writeXlsx(wb)).toThrow(/1904/);
    const protectedParts = fixture().parts;
    protectedParts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(protectedParts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<sheetProtection sheet="1"/></worksheet>'));
    const protectedWb = readXlsx(zipSync(protectedParts));
    applyStyle(protectedWb, 0, toRange('A2'), { bold: false });
    expect(() => writeXlsx(protectedWb)).toThrow(/Blattschutz/);
  });

  it('keeps an existing calculation chain but blocks changes to its formula graph', () => {
    const { parts } = fixture();
    parts['xl/calcChain.xml'] = strToU8(`<calcChain xmlns="${ns}"><c r="B3" i="1"/></calcChain>`);
    parts['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(parts['xl/_rels/workbook.xml.rels']!).replace('</Relationships>', '<Relationship Id="rCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'));
    const wb = readXlsx(zipSync(parts));
    setValues(wb, 0, 1, 1, [[42]]);
    expect(unzipSync(writeXlsx(wb))['xl/calcChain.xml']).toEqual(parts['xl/calcChain.xml']);
    setValues(wb, 0, 2, 1, [['=B2*3']]);
    expect(() => writeXlsx(wb)).toThrow(/Berechnungskette/);
  });

  it('rejects writes inside unimplemented pivot output ranges', () => {
    const { parts } = fixture();
    parts['xl/worksheets/_rels/sheet1.xml.rels'] = strToU8(strFromU8(parts['xl/worksheets/_rels/sheet1.xml.rels']!).replace('</Relationships>', '<Relationship Id="rPivot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>'));
    parts['xl/pivotTables/pivotTable1.xml'] = strToU8(`<pivotTableDefinition xmlns="${ns}" name="OriginalPivot" cacheId="1"><location ref="B2:C3"/></pivotTableDefinition>`);
    const wb = readXlsx(zipSync(parts));
    setValues(wb, 0, 1, 1, [[99]]);
    expect(() => writeXlsx(wb)).toThrow(/Pivot-Ausgabe/);
  });
});

function canonicalBytes(value: Uint8Array): string { return Buffer.from(value).toString('base64'); }

describe('shared string scanning', () => {
  it('finds rich and phonetic items like a full parse', async () => {
    const { richStringIndices } = await import('../src/io/xlsx-preserve.js');
    const { arr, parseXml } = await import('../src/io/xml.js');
    const xml = '<sst xmlns="ns"><si><t>plain</t></si><si><r><rPr><b/></rPr><t>bold</t></r></si><si/>' +
      '<si><t xml:space="preserve"> a&lt;r&gt; </t></si><si><t>ふ</t><rPh sb="0" eb="1"><t>フ</t></rPh><phoneticPr fontId="1"/></si>' +
      '<si><t>&lt;r&gt;</t></si></sst>';
    const expected = new Set<number>();
    arr(parseXml(xml)?.sst?.si).forEach((si: unknown, index: number) => {
      if (si && typeof si === 'object' && Object.keys(si).some((key) => !key.startsWith('@_') && key !== 't')) expected.add(index);
    });
    expect([...richStringIndices(xml)]).toEqual([1, 4]);
    expect([...expected]).toEqual([1, 4]);
    expect([...richStringIndices('<x:sst xmlns:x="ns"><x:si><x:t>a</x:t></x:si><x:si><x:r><x:t>b</x:t></x:r></x:si></x:sst>')]).toEqual([1]);
  });
});
