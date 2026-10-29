import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { applyBorders, copyRange, fillRange, insertChartFromRange, openWorkbook, parseRange, pasteRange, rangeToTSV, setValues, toRange, Workbook, writeXlsx } from '../src/index.js';
import { MAX_ZIP_ENTRIES, MAX_ZIP_ENTRY_BYTES, unzipLimited } from '../src/io/zip.js';

/** Overwrites the uncompressed size in every local and central ZIP header. */
function forgeSize(zip: Uint8Array, size: number): Uint8Array {
  const out = zip.slice();
  const view = new DataView(out.buffer);
  for (let i = 0; i + 4 <= out.length; i++) {
    const signature = view.getUint32(i, true);
    if (signature === 0x04034b50) view.setUint32(i + 22, size, true);
    if (signature === 0x02014b50) view.setUint32(i + 24, size, true);
  }
  return out;
}

describe('ZIP bombs', () => {
  it('refuses stored parts with understated original sizes', () => {
    const forged = forgeSize(zipSync({ 'xl/workbook.xml': strToU8('<workbook/>') }, { level: 0 }), 1);
    // fflate copies stored data by compressed size, ignoring the declared original size.
    expect(unzipSync(forged)['xl/workbook.xml']!.length).toBeGreaterThan(1);
    expect(() => unzipLimited(forged)).toThrow(/inconsistent stored part size/);
  });

  it('validates later headers before trying to inflate earlier corrupt parts', () => {
    const zip = zipSync({ first: new Uint8Array(1000), second: strToU8('x') });
    const view = new DataView(zip.buffer);
    // Corrupt the first DEFLATE block and declare an oversized later entry.
    zip[30 + 'first'.length] = 7;
    let entry = 0;
    for (let i = 0; i + 28 <= zip.length; i++) {
      if (view.getUint32(i, true) === 0x02014b50 && ++entry === 2) view.setUint32(i + 24, MAX_ZIP_ENTRY_BYTES + 1, true);
    }
    expect(() => unzipLimited(zip)).toThrow(/too large/);
  });
  it('refuses a part whose uncompressed size exceeds the limit before inflating it', () => {
    const bomb = forgeSize(zipSync({ 'xl/workbook.xml': strToU8('<workbook/>') }), MAX_ZIP_ENTRY_BYTES + 1);
    expect(() => openWorkbook(bomb, 'bomb.xlsx')).toThrow(/too large/);
  });

  it('refuses archives with too many parts', () => {
    const files: Record<string, Uint8Array> = {};
    for (let i = 0; i <= MAX_ZIP_ENTRIES; i++) files[`p${i}`] = new Uint8Array(0);
    expect(() => openWorkbook(zipSync(files, { level: 0 }), 'many.xlsx')).toThrow(/parts/);
  });
});

describe('chart colors from files', () => {
  it('ignores a series color that is not a hex RGB value', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['k', 'v'], ['a', 1], ['b', 2]]);
    insertChartFromRange(wb, 0, parseRange('A1:B3')!.range, 'column');
    const files = unzipSync(writeXlsx(wb));
    const part = Object.keys(files).find((name) => /charts\/chart\d+\.xml$/.test(name))!;
    const fill = (val: string) => {
      const xml = strFromU8(files[part]!).replace(/(<c:ser>[\s\S]*?<\/c:tx>)/, `$1<c:spPr><a:solidFill><a:srgbClr val="${val}"/></a:solidFill></c:spPr>`);
      return openWorkbook(zipSync({ ...files, [part]: strToU8(xml) }), 'chart.xlsx').sheets[0]!.charts[0]!.series[0]!.color;
    };
    expect(fill('1A2B3C')).toBe('#1A2B3C');
    expect(fill('000000&quot;/&gt;&lt;foreignObject&gt;')).toBeUndefined();
  });
});

describe('malicious XLSX dimensions', () => {
  const parts = unzipSync(writeXlsx(new Workbook()));
  const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!);
  const readWith = (xml: string) => openWorkbook(zipSync({ ...parts, 'xl/worksheets/sheet1.xml': strToU8(xml) }), 'dimensions.xlsx');

  it.each(['Infinity', '-Infinity', 'NaN', '1e309', '9007199254740992', '16385', '-1', '0', '1.5'])(
    'rejects invalid column bounds before iteration: %s', (value) => {
      expect(() => readWith(sheet.replace('<sheetData>', `<cols><col min="${value}" max="${value}" width="10"/></cols><sheetData>`)))
        .toThrow(/invalid (numeric attribute|column bounds)/);
    });

  it.each(['Infinity', 'NaN', '9007199254740992', '1048577', '-1', '0', '1.5'])(
    'rejects invalid row indices: %s', (value) => {
      expect(() => readWith(sheet.replace('<sheetData>', `<sheetData><row r="${value}" ht="20" customHeight="1"/>`)))
        .toThrow(/invalid (numeric attribute|row index)/);
    });

  it('retains legal edge indices and widths', () => {
    const wb = readWith(sheet.replace('<sheetData>', '<cols><col min="16384" max="16384" width="10"/></cols><sheetData><row r="1048576"><c r="XFD1048576"><v>7</v></c></row>'));
    expect(wb.getValue(0, 1048575, 16383)).toBe(7);
    expect(wb.activeSheet.colWidths.get(16383)).toBe(75);
  });

  it('rejects nonfinite numeric layout attributes', () => {
    expect(() => readWith(sheet.replace(/<sheetFormatPr\b[^>]*\/>/, '<sheetFormatPr defaultColWidth="Infinity"/>'))).toThrow(/invalid numeric attribute/);
  });
});

describe('oversized user selections', () => {
  it('refuses dense operations before allocating or changing cells', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['keep']]);
    const original = wb.toJSON();
    const range = toRange('A1:XFC1048575');
    const clip = copyRange(wb, 0, toRange('A1'));
    for (const operation of [
      () => applyBorders(wb, 0, range, 'all'),
      () => copyRange(wb, 0, range),
      () => rangeToTSV(wb, 0, range),
      () => fillRange(wb, 0, toRange('A1'), range),
      () => pasteRange(wb, clip, 0, range),
    ]) {
      expect(operation).toThrow(/cell limit/);
      expect(wb.toJSON()).toEqual(original);
    }
  });

  it('clips whole-column borders and copies to the used cells', () => {
    const wb = new Workbook();
    setValues(wb, 0, 0, 0, [['value']]);
    applyBorders(wb, 0, toRange('A:A'), 'all');
    expect(wb.activeSheet.cells.size).toBe(1);
    expect(copyRange(wb, 0, toRange('A:A')).cells).toHaveLength(1);
    expect(copyRange(new Workbook(), 0, toRange('A:XFD')).cells).toEqual([[null]]);
  });
});
