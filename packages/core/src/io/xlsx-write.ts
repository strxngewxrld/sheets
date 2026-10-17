import { strToU8, zipSync } from 'fflate';
import { formatAddress, formatRange, keyToAddress, quoteSheetName } from '../address.js';
import type { ChartSeries, ChartSpec } from '../chart.js';
import { addFutureFunctionPrefixes } from '../formula/rewrite.js';
import type { Sheet } from '../sheet.js';
import { styleKey, type BorderSide, type CellStyle } from '../style.js';
import { isError } from '../values.js';
import type { Workbook } from '../workbook.js';
import { toArgb } from './colors.js';
import { escapeXml } from './xml.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const BUILTIN_FMT_IDS: Record<string, number> = {
  General: 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11, '@': 49,
};

export const DEFAULT_FONT = { name: 'Calibri', size: 11 };

/** Pixels → Excel character width. */
export function pxToCharWidth(px: number): number {
  return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
}

/** Pixels → points. */
export function pxToPoints(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100;
}

class StyleWriter {
  private fonts: string[] = [];
  private fontIdx = new Map<string, number>();
  private fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  private fillIdx = new Map<string, number>();
  private borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  private borderIdx = new Map<string, number>();
  private numFmts = new Map<string, number>();
  private xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  private xfIdx = new Map<string, number>();

  constructor() {
    this.font({});
  }

  private intern(list: string[], map: Map<string, number>, xml: string): number {
    let idx = map.get(xml);
    if (idx === undefined) {
      idx = list.length;
      list.push(xml);
      map.set(xml, idx);
    }
    return idx;
  }

  private font(s: CellStyle): number {
    const xml =
      '<font>' +
      (s.bold ? '<b/>' : '') +
      (s.italic ? '<i/>' : '') +
      (s.strike ? '<strike/>' : '') +
      (s.underline ? '<u/>' : '') +
      `<sz val="${s.fontSize ?? DEFAULT_FONT.size}"/>` +
      (s.color ? `<color rgb="${toArgb(s.color)}"/>` : '<color theme="1"/>') +
      `<name val="${escapeXml(s.fontName ?? DEFAULT_FONT.name)}"/>` +
      '<family val="2"/>' +
      (s.fontName ? '' : '<scheme val="minor"/>') +
      '</font>';
    return this.intern(this.fonts, this.fontIdx, xml);
  }

  private fill(color: string): number {
    const xml = `<fill><patternFill patternType="solid"><fgColor rgb="${toArgb(color)}"/><bgColor indexed="64"/></patternFill></fill>`;
    return this.intern(this.fills, this.fillIdx, xml);
  }

  private border(b: NonNullable<CellStyle['border']>): number {
    const side = (tag: string, s?: BorderSide) =>
      s ? `<${tag} style="${s.style}"><color rgb="${toArgb(s.color ?? '#000000')}"/></${tag}>` : `<${tag}/>`;
    const xml = `<border>${side('left', b.left)}${side('right', b.right)}${side('top', b.top)}${side('bottom', b.bottom)}<diagonal/></border>`;
    return this.intern(this.borders, this.borderIdx, xml);
  }

  private numFmt(code: string): number {
    if (code in BUILTIN_FMT_IDS) return BUILTIN_FMT_IDS[code]!;
    let id = this.numFmts.get(code);
    if (id === undefined) {
      id = 164 + this.numFmts.size;
      this.numFmts.set(code, id);
    }
    return id;
  }

  /** Returns the cellXfs index for a style. */
  get(style: CellStyle | undefined): number {
    if (!style) return 0;
    const key = styleKey(style);
    const cached = this.xfIdx.get(key);
    if (cached !== undefined) return cached;
    const fontId = this.font(style);
    const fillId = style.fill ? this.fill(style.fill) : 0;
    const borderId = style.border && Object.keys(style.border).length ? this.border(style.border) : 0;
    const numFmtId = style.numFmt ? this.numFmt(style.numFmt) : 0;
    const align: string[] = [];
    if (style.hAlign) align.push(`horizontal="${style.hAlign}"`);
    if (style.vAlign) align.push(`vertical="${style.vAlign === 'middle' ? 'center' : style.vAlign}"`);
    if (style.wrap) align.push('wrapText="1"');
    let xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`;
    if (numFmtId) xml += ' applyNumberFormat="1"';
    if (fontId) xml += ' applyFont="1"';
    if (fillId) xml += ' applyFill="1"';
    if (borderId) xml += ' applyBorder="1"';
    xml += align.length ? ` applyAlignment="1"><alignment ${align.join(' ')}/></xf>` : '/>';
    const idx = this.xfs.length;
    this.xfs.push(xml);
    this.xfIdx.set(key, idx);
    return idx;
  }

  xml(): string {
    const numFmts = this.numFmts.size
      ? `<numFmts count="${this.numFmts.size}">${[...this.numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`).join('')}</numFmts>`
      : '';
    return (
      XML_HEAD +
      `<styleSheet xmlns="${NS_MAIN}">` +
      numFmts +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0"/>' +
      '</styleSheet>'
    );
  }
}

class SharedStrings {
  private list: string[] = [];
  private index = new Map<string, number>();
  count = 0;
  get(s: string): number {
    this.count++;
    let i = this.index.get(s);
    if (i === undefined) {
      i = this.list.length;
      this.list.push(s);
      this.index.set(s, i);
    }
    return i;
  }
  xml(): string {
    const items = this.list.map((s) => `<si><t${/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : ''}>${escapeXml(s)}</t></si>`).join('');
    return XML_HEAD + `<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.list.length}">${items}</sst>`;
  }
}

function sheetXml(sheet: Sheet, styles: StyleWriter, strings: SharedStrings, isActive: boolean, drawingRelId: string | null): string {
  const used = sheet.usedRange();
  const parts: string[] = [XML_HEAD, `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`];
  parts.push(`<dimension ref="${used ? formatRange(used) : 'A1'}"/>`);

  const { rows: fr, cols: fc } = sheet.frozen;
  let pane = '';
  if (fr || fc) {
    const activePane = fr && fc ? 'bottomRight' : fr ? 'bottomLeft' : 'topRight';
    pane =
      `<pane${fc ? ` xSplit="${fc}"` : ''}${fr ? ` ySplit="${fr}"` : ''} topLeftCell="${formatAddress(fr, fc)}" activePane="${activePane}" state="frozen"/>` +
      `<selection pane="${activePane}"/>`;
  }
  parts.push(`<sheetViews><sheetView workbookViewId="0"${isActive ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>`);
  const defaultCol = sheet.defaultColWidth !== 64 ? ` defaultColWidth="${pxToCharWidth(sheet.defaultColWidth)}"` : '';
  parts.push(`<sheetFormatPr defaultRowHeight="${pxToPoints(sheet.defaultRowHeight)}"${defaultCol}${sheet.defaultRowHeight !== 20 ? ' customHeight="1"' : ''}/>`);

  if (sheet.colWidths.size) {
    const cols = [...sheet.colWidths.entries()].sort((a, b) => a[0] - b[0]);
    const groups: { min: number; max: number; px: number }[] = [];
    for (const [c, px] of cols) {
      const last = groups.at(-1);
      if (last && last.max === c - 1 && last.px === px) last.max = c;
      else groups.push({ min: c, max: c, px });
    }
    parts.push(
      '<cols>' +
        groups
          .map((g) =>
            g.px === 0
              ? `<col min="${g.min + 1}" max="${g.max + 1}" width="${pxToCharWidth(sheet.defaultColWidth)}" hidden="1" customWidth="1"/>`
              : `<col min="${g.min + 1}" max="${g.max + 1}" width="${pxToCharWidth(g.px)}" customWidth="1"/>`,
          )
          .join('') +
        '</cols>',
    );
  }

  // Group cells by row.
  const rows = new Map<number, number[]>();
  for (const key of sheet.cells.keys()) {
    const { row } = keyToAddress(key);
    let list = rows.get(row);
    if (!list) rows.set(row, (list = []));
    list.push(key);
  }
  for (const r of sheet.rowHeights.keys()) if (!rows.has(r)) rows.set(r, []);
  parts.push('<sheetData>');
  for (const r of [...rows.keys()].sort((a, b) => a - b)) {
    const h = sheet.rowHeights.get(r);
    const rowAttrs = h === undefined ? '' : h === 0 ? ' hidden="1"' : ` ht="${pxToPoints(h)}" customHeight="1"`;
    parts.push(`<row r="${r + 1}"${rowAttrs}>`);
    for (const key of rows.get(r)!.sort((a, b) => a - b)) {
      const cell = sheet.cells.get(key)!;
      const { col } = keyToAddress(key);
      const ref = formatAddress(r, col);
      const s = styles.get(cell.style);
      const sAttr = s ? ` s="${s}"` : '';
      const v = cell.value;
      if (cell.formula !== undefined) {
        const f = `<f>${escapeXml(addFutureFunctionPrefixes(cell.formula))}</f>`;
        if (typeof v === 'number') parts.push(`<c r="${ref}"${sAttr}>${f}<v>${v}</v></c>`);
        else if (typeof v === 'string') parts.push(`<c r="${ref}"${sAttr} t="str">${f}<v>${escapeXml(v)}</v></c>`);
        else if (typeof v === 'boolean') parts.push(`<c r="${ref}"${sAttr} t="b">${f}<v>${v ? 1 : 0}</v></c>`);
        else if (isError(v) && !['#CIRC!', '#SPILL!', '#CALC!'].includes(v.code)) parts.push(`<c r="${ref}"${sAttr} t="e">${f}<v>${v.code}</v></c>`);
        else parts.push(`<c r="${ref}"${sAttr}>${f}</c>`);
      } else if (typeof v === 'number') parts.push(`<c r="${ref}"${sAttr}><v>${v}</v></c>`);
      else if (typeof v === 'string') parts.push(`<c r="${ref}"${sAttr} t="s"><v>${strings.get(v)}</v></c>`);
      else if (typeof v === 'boolean') parts.push(`<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`);
      else if (isError(v)) parts.push(`<c r="${ref}"${sAttr} t="e"><v>${v.code}</v></c>`);
      else if (s) parts.push(`<c r="${ref}"${sAttr}/>`);
    }
    parts.push('</row>');
  }
  parts.push('</sheetData>');

  if (sheet.merges.length) {
    parts.push(`<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${formatRange(m)}"/>`).join('')}</mergeCells>`);
  }
  parts.push('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
  if (drawingRelId) parts.push(`<drawing r:id="${drawingRelId}"/>`);
  parts.push('</worksheet>');
  return parts.join('');
}

// ---------------------------------------------------------------------------
// Charts (DrawingML)
// ---------------------------------------------------------------------------

function richText(text: string): string {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/>`;
}

function qualify(ref: string, sheetName: string): string {
  return ref.includes('!') ? ref : `${quoteSheetName(sheetName)}!${ref}`;
}

function seriesXml(s: ChartSeries, i: number, spec: ChartSpec, sheetName: string): string {
  const lineLike = spec.type === 'line' || spec.type === 'scatter';
  let xml = `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>`;
  if (s.nameRef) xml += `<c:tx><c:strRef><c:f>${escapeXml(qualify(s.nameRef, sheetName))}</c:f></c:strRef></c:tx>`;
  else if (s.name) xml += `<c:tx><c:v>${escapeXml(s.name)}</c:v></c:tx>`;
  if (spec.type === 'scatter') {
    const color = s.color ? `<a:solidFill><a:srgbClr val="${s.color.replace('#', '')}"/></a:solidFill>` : '';
    xml += `<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/>${color ? `<c:spPr>${color}</c:spPr>` : ''}</c:marker>`;
  } else if (s.color) {
    const fill = `<a:solidFill><a:srgbClr val="${s.color.replace('#', '')}"/></a:solidFill>`;
    xml += lineLike ? `<c:spPr><a:ln w="28575" cap="rnd">${fill}</a:ln></c:spPr>` : `<c:spPr>${fill}</c:spPr>`;
  }
  if (spec.type === 'line') xml += '<c:marker><c:symbol val="none"/></c:marker>';
  if (spec.type === 'column' || spec.type === 'bar') xml += '<c:invertIfNegative val="0"/>';
  if (spec.type === 'scatter') {
    if (s.categories) xml += `<c:xVal><c:numRef><c:f>${escapeXml(qualify(s.categories, sheetName))}</c:f></c:numRef></c:xVal>`;
    xml += `<c:yVal><c:numRef><c:f>${escapeXml(qualify(s.values, sheetName))}</c:f></c:numRef></c:yVal><c:smooth val="0"/>`;
  } else {
    if (s.categories) xml += `<c:cat><c:strRef><c:f>${escapeXml(qualify(s.categories, sheetName))}</c:f></c:strRef></c:cat>`;
    xml += `<c:val><c:numRef><c:f>${escapeXml(qualify(s.values, sheetName))}</c:f></c:numRef></c:val>`;
    if (spec.type === 'line') xml += '<c:smooth val="0"/>';
  }
  return xml + '</c:ser>';
}

function axisXml(kind: 'catAx' | 'valAx', id: number, cross: number, pos: string, title: string | undefined, extra: { gridlines?: boolean; midCat?: boolean }): string {
  let xml = `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>`;
  if (extra.gridlines) xml += '<c:majorGridlines/>';
  if (title) xml += `<c:title>${richText(title)}</c:title>`;
  xml += '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
  xml += `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/>`;
  if (kind === 'catAx') xml += '<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/>';
  else xml += `<c:crossBetween val="${extra.midCat ? 'midCat' : 'between'}"/>`;
  return xml + `</c:${kind}>`;
}

export function chartXml(spec: ChartSpec, sheetName: string): string {
  const series = spec.series.map((s, i) => seriesXml(s, i, spec, sheetName)).join('');
  const stacked = !!spec.stacked;
  let plot: string;
  const axIds = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  const horizontal = spec.type === 'bar';
  switch (spec.type) {
    case 'column':
    case 'bar':
      plot =
        `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${series}` +
        `<c:gapWidth val="150"/>${stacked ? '<c:overlap val="100"/>' : ''}${axIds}</c:barChart>`;
      break;
    case 'line':
      plot = `<c:lineChart><c:grouping val="${stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axIds}</c:lineChart>`;
      break;
    case 'area':
      plot = `<c:areaChart><c:grouping val="${stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${series}${axIds}</c:areaChart>`;
      break;
    case 'pie':
      plot = `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'doughnut':
      plot = `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
      break;
    case 'scatter':
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${axIds}</c:scatterChart>`;
      break;
  }
  let axes = '';
  if (spec.type === 'scatter') {
    axes = axisXml('valAx', 500000001, 500000002, 'b', spec.xAxisTitle, { midCat: true }) + axisXml('valAx', 500000002, 500000001, 'l', spec.yAxisTitle, { gridlines: true, midCat: true });
  } else if (spec.type !== 'pie' && spec.type !== 'doughnut') {
    axes = axisXml('catAx', 500000001, 500000002, horizontal ? 'l' : 'b', spec.xAxisTitle, {}) + axisXml('valAx', 500000002, 500000001, horizontal ? 'b' : 'l', spec.yAxisTitle, { gridlines: true });
  }
  const legendPos = { right: 'r', left: 'l', top: 't', bottom: 'b' } as const;
  const legend = spec.legend && spec.legend !== 'none' ? `<c:legend><c:legendPos val="${legendPos[spec.legend]}"/><c:overlay val="0"/></c:legend>` : '';
  return (
    XML_HEAD +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<c:roundedCorners val="0"/><c:chart>' +
    (spec.title ? `<c:title>${richText(spec.title)}</c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>') +
    `<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
  );
}

function drawingXml(charts: { spec: ChartSpec; relId: string }[]): string {
  const anchors = charts
    .map(({ spec, relId }, i) => {
      const { from, to } = spec.anchor;
      return (
        '<xdr:twoCellAnchor editAs="oneCell">' +
        `<xdr:from><xdr:col>${from.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>${to.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${to.row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="${escapeXml(spec.title ?? `Chart ${i + 1}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${NS_R}" r:id="${relId}"/></a:graphicData></a:graphic>` +
        '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
      );
    })
    .join('');
  return (
    XML_HEAD +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    anchors +
    '</xdr:wsDr>'
  );
}

function relsXml(rels: { id: string; type: string; target: string }[]): string {
  return XML_HEAD + `<Relationships xmlns="${NS_PKG_R}">${rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escapeXml(r.target)}"/>`).join('')}</Relationships>`;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export interface XlsxWriteOptions {
  /** Zip compression level 0-9 (default 6). */
  compression?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  application?: string;
}

export function writeXlsx(wb: Workbook, options: XlsxWriteOptions = {}): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const put = (path: string, xml: string) => (files[path] = strToU8(xml));
  const styles = new StyleWriter();
  const strings = new SharedStrings();
  const overrides: string[] = [];
  const wbRels: { id: string; type: string; target: string }[] = [];
  let chartNo = 0;
  let drawingNo = 0;

  wb.sheets.forEach((sheet, i) => {
    const sheetPath = `xl/worksheets/sheet${i + 1}.xml`;
    let drawingRelId: string | null = null;
    if (sheet.charts.length) {
      drawingNo++;
      drawingRelId = 'rId1';
      const drawingPath = `xl/drawings/drawing${drawingNo}.xml`;
      const chartRels: { id: string; type: string; target: string }[] = [];
      const entries = sheet.charts.map((spec, j) => {
        chartNo++;
        const relId = `rId${j + 1}`;
        put(`xl/charts/chart${chartNo}.xml`, chartXml(spec, sheet.name));
        overrides.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
        chartRels.push({ id: relId, type: `${REL}/chart`, target: `../charts/chart${chartNo}.xml` });
        return { spec, relId };
      });
      put(drawingPath, drawingXml(entries));
      put(`xl/drawings/_rels/drawing${drawingNo}.xml.rels`, relsXml(chartRels));
      put(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, relsXml([{ id: 'rId1', type: `${REL}/drawing`, target: `../drawings/drawing${drawingNo}.xml` }]));
      overrides.push(`<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    }
    put(sheetPath, sheetXml(sheet, styles, strings, i === wb.activeSheetIndex, drawingRelId));
    overrides.push(`<Override PartName="/${sheetPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
    wbRels.push({ id: `rId${i + 1}`, type: `${REL}/worksheet`, target: `worksheets/sheet${i + 1}.xml` });
  });

  const n = wb.sheets.length;
  wbRels.push({ id: `rId${n + 1}`, type: `${REL}/styles`, target: 'styles.xml' });
  wbRels.push({ id: `rId${n + 2}`, type: `${REL}/sharedStrings`, target: 'sharedStrings.xml' });

  const sheetsXml = wb.sheets.map((s, i) => `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  const names = [...wb.names.values()];
  const namesXml = names.length ? `<definedNames>${names.map((d) => `<definedName name="${escapeXml(d.name)}">${escapeXml(d.ref)}</definedName>`).join('')}</definedNames>` : '';
  put(
    'xl/workbook.xml',
    XML_HEAD +
      `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><workbookPr/>` +
      `<bookViews><workbookView activeTab="${wb.activeSheetIndex}"/></bookViews>` +
      `<sheets>${sheetsXml}</sheets>${namesXml}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
  );
  put('xl/_rels/workbook.xml.rels', relsXml(wbRels));
  put('xl/styles.xml', styles.xml());
  put('xl/sharedStrings.xml', strings.xml());

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  put(
    'docProps/core.xml',
    XML_HEAD +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
  );
  put(
    'docProps/app.xml',
    XML_HEAD +
      `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>${escapeXml(options.application ?? 'Sheets')}</Application></Properties>`,
  );
  put(
    '_rels/.rels',
    relsXml([
      { id: 'rId1', type: `${REL}/officeDocument`, target: 'xl/workbook.xml' },
      { id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
      { id: 'rId3', type: `${REL}/extended-properties`, target: 'docProps/app.xml' },
    ]),
  );
  put(
    '[Content_Types].xml',
    XML_HEAD +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      overrides.join('') +
      '</Types>',
  );
  return zipSync(files, { level: options.compression ?? 6 });
}
