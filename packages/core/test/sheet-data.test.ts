import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { strFromU8, unzipSync } from 'fflate';
import { applyStyle, setValues, toRange, Workbook, writeXlsx, xlsxToData } from '../src/index.js';
import { scanSheetData } from '../src/io/sheet-data.js';
import { arr, parseXml } from '../src/io/xml.js';

/** Rows as the generic parser returns them. */
const generic = (xml: string) => arr(parseXml(xml)?.worksheet?.sheetData?.row);

const SAMPLES = [
  '<worksheet xmlns="ns"><dimension ref="A1"/><sheetData/><pageMargins left="1"/></worksheet>',
  '<worksheet xmlns="ns"><sheetData><row r="1" spans="1:3" ht="20" customHeight="1" x14ac:dyDescent="0.25"><c r="A1" s="2" t="s"><v>0</v></c><c r="B1"><f>SUM(C1:C9)&amp;"x"</f><v>3</v></c><c r="C1" t="str"><f t="shared" ref="C1:C3" si="0">A1&amp;B1</f><v>ab</v></c></row>' +
    '<row r="2"><c r="C2" t="str"><f t="shared" si="0"/><v></v></c><c r="D2" t="inlineStr"><is><r><rPr><b/></rPr><t>rich</t></r><t xml:space="preserve"> text</t></is></c><c r="E2" s="1"/></row>' +
    '<row r="4" hidden="1"/>\n  <row r="5"><c r="A5" t="e"><v>#N/A</v></c><c r="B5"><f t="array" ref="B5:B6">SEQUENCE(2)</f><v>1</v></c><c r="C5"><f></f><v>2</v></c></row></sheetData></worksheet>',
  '<x:worksheet xmlns:x="ns"><x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>pre</x:t></x:is></x:c><x:c r="B1"><x:v>5</x:v></x:c></x:row></x:sheetData></x:worksheet>',
];

describe('sheetData scanner', () => {
  it.each(SAMPLES.map((xml, i) => [i, xml]))('matches the generic parser for sample %i', (_i, xml) => {
    const scanned = scanSheetData(xml as string);
    expect(scanned).not.toBeNull();
    expect(JSON.parse(JSON.stringify(scanned!.rows))).toEqual(JSON.parse(JSON.stringify(generic(xml as string))));
    const rest = (xml as string).slice(0, scanned!.start) + (xml as string).slice(scanned!.end);
    expect(rest).not.toContain('sheetData');
  });

  it('falls back for comments and unexpected content', () => {
    expect(scanSheetData('<worksheet><sheetData><!-- x --><row r="1"/></sheetData></worksheet>')).toBeNull();
    expect(scanSheetData('<worksheet><sheetData><row r="1">text<c r="A1"/></row></sheetData></worksheet>')).toBeNull();
    expect(scanSheetData('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v>')).toBeNull();
  });

  it('reads real workbooks exactly as before', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['Text', 1.5, true, '=A1&B1', '=SEQUENCE(2)'], [null, '=1/0', 'x"y<z>&', null, null]]);
    applyStyle(wb, 0, toRange('A1:C2'), { bold: true, fill: '#FFEE00' });
    const files = [writeXlsx(wb), new Uint8Array(readFileSync(new URL('./fixtures/excel-features.xlsx', import.meta.url)))];
    for (const bytes of files) {
      const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
      const scanned = scanSheetData(xml);
      expect(scanned).not.toBeNull();
      expect(JSON.parse(JSON.stringify(scanned!.rows))).toEqual(JSON.parse(JSON.stringify(generic(xml))));
      expect(xlsxToData(bytes).sheets[0]!.cells).toBeDefined();
    }
  });
});
