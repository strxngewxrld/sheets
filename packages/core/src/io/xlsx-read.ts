import { strFromU8, unzipSync } from 'fflate';
import { formatAddress, parseAddress, parseRange } from '../address.js';
import type { ChartAnchor, ChartSeries, ChartSpec, ChartType, LegendPosition } from '../chart.js';
import { newChartId } from '../chart.js';
import { dateToSerial, timeToSerial } from '../dates.js';
import { shiftFormula, stripFutureFunctionPrefixes } from '../formula/rewrite.js';
import type { WorkbookData } from '../ops.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, type CellData, type SheetData } from '../sheet.js';
import type { BorderLineStyle, BorderSide, CellBorders, CellStyle, HAlign, VAlign } from '../style.js';
import { errorFromCode, scalarToJson, type JsonScalar } from '../values.js';
import { Workbook } from '../workbook.js';
import { applyTint, DEFAULT_THEME, INDEXED_COLORS } from './colors.js';
import { arr, attr, collectText, decodeOoxmlEscapes, flag, numAttr, parseXml, relsPathFor, resolvePath, text, type XNode } from './xml.js';

export interface XlsxReadOptions {
  /** Recalculate all formulas after loading (default true). Values Excel cached are kept for unsupported functions. */
  recalculate?: boolean;
}

const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy-mm-dd', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'hh:mm', 21: 'hh:mm:ss',
  22: 'yyyy-mm-dd hh:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
};

/** Converts Excel column width (characters) to pixels. */
export function charWidthToPx(width: number): number {
  return Math.round(width * 7 + 5);
}

/** Converts row height in points to pixels. */
export function pointsToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

class Package {
  private files: Record<string, Uint8Array>;
  constructor(bytes: Uint8Array) {
    this.files = unzipSync(bytes);
  }
  has(path: string): boolean {
    return path in this.files;
  }
  xml(path: string): XNode | undefined {
    const f = this.files[path];
    return f ? parseXml(strFromU8(f)) : undefined;
  }
  rels(part: string): Map<string, { target: string; type: string }> {
    const out = new Map<string, { target: string; type: string }>();
    const doc = this.xml(relsPathFor(part));
    for (const r of arr(doc?.Relationships?.Relationship)) {
      const id = attr(r, 'Id');
      const target = attr(r, 'Target');
      if (id && target && attr(r, 'TargetMode') !== 'External') {
        out.set(id, { target: resolvePath(part, target), type: attr(r, 'Type') ?? '' });
      }
    }
    return out;
  }
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

class StyleReader {
  styles: (CellStyle | undefined)[] = [];
  private theme: string[];

  constructor(doc: XNode, theme: string[]) {
    this.theme = theme;
    const ss = doc?.styleSheet;
    if (!ss) return;
    const numFmts = new Map<number, string>();
    for (const nf of arr(ss.numFmts?.numFmt)) numFmts.set(numAttr(nf, 'numFmtId')!, attr(nf, 'formatCode') ?? 'General');
    const fonts = arr(ss.fonts?.font);
    const fills = arr(ss.fills?.fill);
    const borders = arr(ss.borders?.border);
    const baseFont = this.font(fonts[0]);
    this.styles = arr(ss.cellXfs?.xf).map((xf) => {
      const style: CellStyle = {};
      const numFmtId = numAttr(xf, 'numFmtId') ?? 0;
      const fmt = numFmts.get(numFmtId) ?? BUILTIN_NUMFMTS[numFmtId];
      if (fmt && fmt !== 'General') style.numFmt = fmt;
      const font = this.font(fonts[numAttr(xf, 'fontId') ?? 0]);
      for (const [k, v] of Object.entries(font) as [keyof CellStyle, never][]) {
        if (baseFont[k] !== v) (style as Record<string, unknown>)[k] = v;
      }
      const fill = fills[numAttr(xf, 'fillId') ?? 0];
      const pf = fill?.patternFill;
      if (pf && attr(pf, 'patternType') && attr(pf, 'patternType') !== 'none') {
        const color = this.color(pf.fgColor) ?? this.color(pf.bgColor);
        if (color) style.fill = color;
      }
      const border = this.border(borders[numAttr(xf, 'borderId') ?? 0]);
      if (border) style.border = border;
      const al = xf.alignment;
      if (al) {
        const h = attr(al, 'horizontal');
        if (h === 'left' || h === 'center' || h === 'right' || h === 'justify') style.hAlign = h as HAlign;
        else if (h === 'centerContinuous') style.hAlign = 'center';
        const v = attr(al, 'vertical');
        if (v === 'top') style.vAlign = 'top';
        else if (v === 'center') style.vAlign = 'middle';
        if (attr(al, 'wrapText') === '1' || attr(al, 'wrapText') === 'true') style.wrap = true;
      }
      return Object.keys(style).length ? style : undefined;
    });
  }

  color(c: XNode): string | undefined {
    if (!c || typeof c !== 'object') return undefined;
    let hex: string | undefined;
    const rgb = attr(c, 'rgb');
    if (rgb) hex = `#${rgb.slice(-6).toUpperCase()}`;
    else if (attr(c, 'theme') !== undefined) hex = this.theme[numAttr(c, 'theme')!];
    else if (attr(c, 'indexed') !== undefined) {
      const idx = numAttr(c, 'indexed')!;
      if (idx === 64 || idx === 65) return undefined; // system foreground/background
      hex = INDEXED_COLORS[idx];
    }
    if (!hex) return undefined;
    const tint = numAttr(c, 'tint');
    return tint ? applyTint(hex, tint) : hex;
  }

  private font(f: XNode): Partial<CellStyle> {
    const out: Partial<CellStyle> = {};
    if (!f) return out;
    if (flag(f, 'b')) out.bold = true;
    if (flag(f, 'i')) out.italic = true;
    if ('u' in f && attr(f.u, 'val') !== 'none') out.underline = true;
    if (flag(f, 'strike')) out.strike = true;
    const sz = numAttr(f.sz, 'val');
    if (sz) out.fontSize = sz;
    const name = attr(f.name, 'val');
    if (name) out.fontName = name;
    const color = this.color(f.color);
    if (color && color !== '#000000') out.color = color;
    return out;
  }

  private border(b: XNode): CellBorders | undefined {
    if (!b) return undefined;
    const out: CellBorders = {};
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const n = b[side];
      const style = attr(n, 'style');
      if (!style || style === 'none') continue;
      const mapped: BorderLineStyle =
        style === 'medium' || style === 'mediumDashed' ? 'medium'
          : style === 'thick' ? 'thick'
            : style === 'double' ? 'double'
              : style.includes('ash') ? 'dashed'
                : style === 'dotted' ? 'dotted'
                  : style === 'hair' ? 'hair'
                    : 'thin';
      const bs: BorderSide = { style: mapped };
      const color = this.color(n.color);
      if (color) bs.color = color;
      out[side] = bs;
    }
    return Object.keys(out).length ? out : undefined;
  }
}

function readTheme(doc: XNode): string[] {
  const scheme = doc?.theme?.themeElements?.clrScheme;
  if (!scheme) return DEFAULT_THEME;
  const get = (key: string, fallback: string) => {
    const n = scheme[key];
    const v = attr(n?.srgbClr, 'val') ?? attr(n?.sysClr, 'lastClr');
    return v ? `#${v.toUpperCase()}` : fallback;
  };
  return [
    get('lt1', DEFAULT_THEME[0]!), get('dk1', DEFAULT_THEME[1]!), get('lt2', DEFAULT_THEME[2]!), get('dk2', DEFAULT_THEME[3]!),
    get('accent1', DEFAULT_THEME[4]!), get('accent2', DEFAULT_THEME[5]!), get('accent3', DEFAULT_THEME[6]!), get('accent4', DEFAULT_THEME[7]!),
    get('accent5', DEFAULT_THEME[8]!), get('accent6', DEFAULT_THEME[9]!), get('hlink', DEFAULT_THEME[10]!), get('folHlink', DEFAULT_THEME[11]!),
  ];
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------

const CHART_ELEMENTS: Record<string, ChartType> = {
  barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter',
};
const LEGEND_POS: Record<string, LegendPosition> = { r: 'right', l: 'left', t: 'top', b: 'bottom', tr: 'right' };

function readChart(doc: XNode, anchor: ChartAnchor): ChartSpec | null {
  const chart = doc?.chartSpace?.chart;
  const plot = chart?.plotArea;
  if (!plot) return null;
  const key = Object.keys(CHART_ELEMENTS).find((k) => plot[k]);
  if (!key) return null;
  const el = arr(plot[key])[0];
  let type = CHART_ELEMENTS[key]!;
  if ((key === 'barChart' || key === 'bar3DChart') && attr(el.barDir, 'val') === 'bar') type = 'bar';
  const series: ChartSeries[] = [];
  for (const ser of arr(el.ser)) {
    const values = text(ser.val?.numRef?.f) || text(ser.yVal?.numRef?.f);
    if (!values) continue;
    const s: ChartSeries = { values };
    const cat = text(ser.cat?.strRef?.f) || text(ser.cat?.numRef?.f) || text(ser.xVal?.numRef?.f) || text(ser.xVal?.strRef?.f);
    if (cat) s.categories = cat;
    const nameRef = text(ser.tx?.strRef?.f);
    if (nameRef) s.nameRef = nameRef;
    else if (ser.tx?.v !== undefined) s.name = text(ser.tx.v);
    const color = attr(ser.spPr?.solidFill?.srgbClr, 'val') ?? attr(ser.spPr?.ln?.solidFill?.srgbClr, 'val');
    if (color) s.color = `#${color}`;
    series.push(s);
  }
  if (!series.length) return null;
  const spec: ChartSpec = { id: newChartId(), type, series, anchor };
  const grouping = attr(el.grouping, 'val');
  if (grouping === 'stacked' || grouping === 'percentStacked') spec.stacked = true;
  const title = chart.title ? collectText(chart.title.tx?.rich) : '';
  if (title) spec.title = title;
  else if (chart.title && series.length === 1) spec.title = series[0]!.name ?? '';
  if (!spec.title) delete spec.title;
  spec.legend = chart.legend ? (LEGEND_POS[attr(chart.legend.legendPos, 'val') ?? 'r'] ?? 'right') : 'none';
  const axisTitle = (ax: XNode) => (ax?.title ? collectText(ax.title.tx?.rich) : '');
  const catAx = arr(plot.catAx)[0] ?? arr(plot.valAx)[0];
  const valAx = arr(plot.valAx).at(-1);
  const xt = axisTitle(catAx);
  const yt = axisTitle(valAx);
  if (xt) spec.xAxisTitle = xt;
  if (yt && valAx !== catAx) spec.yAxisTitle = yt;
  return spec;
}

function readAnchor(a: XNode, ext?: XNode): ChartAnchor {
  const from = { row: Number(text(a.from?.row) || 0), col: Number(text(a.from?.col) || 0) };
  if (a.to) return { from, to: { row: Number(text(a.to.row) || 0), col: Number(text(a.to.col) || 0) } };
  // oneCellAnchor: approximate the size with default cell dimensions (EMU → px).
  const cx = (numAttr(ext, 'cx') ?? 4_572_000) / 9525;
  const cy = (numAttr(ext, 'cy') ?? 2_743_200) / 9525;
  return { from, to: { row: from.row + Math.round(cy / DEFAULT_ROW_HEIGHT), col: from.col + Math.round(cx / DEFAULT_COL_WIDTH) } };
}

function readDrawingCharts(pkg: Package, drawingPath: string): ChartSpec[] {
  const doc = pkg.xml(drawingPath);
  const root = doc?.wsDr;
  if (!root) return [];
  const rels = pkg.rels(drawingPath);
  const out: ChartSpec[] = [];
  const anchors = [...arr(root.twoCellAnchor), ...arr(root.oneCellAnchor)];
  for (const a of anchors) {
    for (const frame of arr(a.graphicFrame)) {
      const rid = attr(frame?.graphic?.graphicData?.chart, 'id');
      const rel = rid ? rels.get(rid) : undefined;
      if (!rel) continue;
      const spec = readChart(pkg.xml(rel.target), readAnchor(a, a.ext));
      if (spec) out.push(spec);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Worksheets
// ---------------------------------------------------------------------------

function readSheet(pkg: Package, path: string, name: string, sharedStrings: string[], styles: StyleReader): SheetData {
  const doc = pkg.xml(path);
  const ws = doc?.worksheet ?? {};
  const sheet: SheetData = { name, cells: {} };

  const fmt = ws.sheetFormatPr;
  const defaultColWidth = numAttr(fmt, 'defaultColWidth');
  const baseColWidth = numAttr(fmt, 'baseColWidth');
  if (defaultColWidth) sheet.defaultColWidth = charWidthToPx(defaultColWidth);
  else if (baseColWidth && baseColWidth !== 8) sheet.defaultColWidth = charWidthToPx(baseColWidth + 0.71);
  const defaultRowHeight = numAttr(fmt, 'defaultRowHeight');
  if (defaultRowHeight && Math.abs(defaultRowHeight - 15) > 0.01) sheet.defaultRowHeight = pointsToPx(defaultRowHeight);

  const colWidths: Record<number, number> = {};
  for (const col of arr(ws.cols?.col)) {
    const min = numAttr(col, 'min')!;
    const max = Math.min(numAttr(col, 'max')!, min + 16_383);
    const hidden = attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true';
    const width = numAttr(col, 'width');
    if (!hidden && width === undefined) continue;
    const px = hidden ? 0 : charWidthToPx(width!);
    if (!hidden && px === (sheet.defaultColWidth ?? DEFAULT_COL_WIDTH)) continue;
    if (max - min > 1000) continue; // whole-sheet column styles
    for (let c = min; c <= max; c++) colWidths[c - 1] = px;
  }
  if (Object.keys(colWidths).length) sheet.colWidths = colWidths;

  const rowHeights: Record<number, number> = {};
  const shared = new Map<string, { formula: string; row: number; col: number }>();
  let rowIndex = -1;
  for (const row of arr(ws.sheetData?.row)) {
    rowIndex = numAttr(row, 'r') !== undefined ? numAttr(row, 'r')! - 1 : rowIndex + 1;
    const ht = numAttr(row, 'ht');
    const hidden = attr(row, 'hidden') === '1' || attr(row, 'hidden') === 'true';
    if (hidden) rowHeights[rowIndex] = 0;
    else if (ht !== undefined && (attr(row, 'customHeight') === '1' || attr(row, 'customHeight') === 'true')) rowHeights[rowIndex] = pointsToPx(ht);
    let colIndex = -1;
    for (const c of arr(row.c)) {
      const ref = attr(c, 'r');
      const addr = ref ? parseAddress(ref) : null;
      colIndex = addr ? addr.col : colIndex + 1;
      const r = addr ? addr.row : rowIndex;
      const data: CellData = {};
      const styleIdx = numAttr(c, 's');
      const style = styleIdx !== undefined ? styles.styles[styleIdx] : undefined;
      if (style) data.s = style;

      const t = attr(c, 't') ?? 'n';
      const raw = c.v !== undefined ? text(c.v) : undefined;
      let v: JsonScalar | undefined;
      switch (t) {
        case 's':
          v = raw !== undefined ? (sharedStrings[Number(raw)] ?? '') : undefined;
          break;
        case 'str':
          v = raw !== undefined ? decodeOoxmlEscapes(raw) : undefined;
          break;
        case 'inlineStr':
          v = decodeOoxmlEscapes(c.is?.t !== undefined ? arr(c.is.t).map(text).join('') : collectText(c.is));
          break;
        case 'b':
          v = raw === undefined ? undefined : raw === '1' || raw === 'true';
          break;
        case 'e':
          v = raw !== undefined ? scalarToJson(errorFromCode(raw)) : undefined;
          break;
        case 'd': {
          const m = raw ? /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}))?/.exec(raw) : null;
          v = m ? dateToSerial(+m[1]!, +m[2]!, +m[3]!) + (m[4] ? timeToSerial(+m[4], +m[5]!, +m[6]!) : 0) : undefined;
          break;
        }
        default:
          v = raw !== undefined && raw !== '' ? Number(raw) : undefined;
      }
      if (v !== undefined) data.v = v;

      const f = c.f;
      if (f !== undefined) {
        const ftype = attr(f, 't');
        let formula = text(f);
        const si = attr(f, 'si');
        if (ftype === 'shared' && si !== undefined) {
          if (formula) shared.set(si, { formula, row: r, col: colIndex });
          else {
            const master = shared.get(si);
            if (master) formula = shiftFormula(master.formula, r - master.row, colIndex - master.col);
          }
        }
        if (formula) data.f = stripFutureFunctionPrefixes(formula);
      }
      if (data.v !== undefined || data.f !== undefined || data.s) sheet.cells[formatAddress(r, colIndex)] = data;
    }
  }
  if (Object.keys(rowHeights).length) sheet.rowHeights = rowHeights;

  const merges = arr(ws.mergeCells?.mergeCell).map((m) => attr(m, 'ref')).filter((m): m is string => !!m && !!parseRange(m));
  if (merges.length) sheet.merges = merges;

  const pane = arr(ws.sheetViews?.sheetView)[0]?.pane;
  if (pane && (attr(pane, 'state') === 'frozen' || attr(pane, 'state') === 'frozenSplit')) {
    sheet.frozen = { rows: numAttr(pane, 'ySplit') ?? 0, cols: numAttr(pane, 'xSplit') ?? 0 };
  }

  const drawingId = attr(ws.drawing, 'id');
  if (drawingId) {
    const rel = pkg.rels(path).get(drawingId);
    if (rel) {
      const charts = readDrawingCharts(pkg, rel.target);
      // Unqualified chart references point to this sheet.
      for (const ch of charts) {
        for (const s of ch.series) {
          for (const k of ['values', 'categories', 'nameRef'] as const) {
            const ref = s[k];
            if (ref && !ref.includes('!')) s[k] = `${name}!${ref}`;
          }
        }
      }
      if (charts.length) sheet.charts = charts;
    }
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

export function xlsxToData(bytes: Uint8Array): WorkbookData {
  const pkg = new Package(bytes);
  const rootRels = pkg.rels('');
  const officeDoc = [...rootRels.values()].find((r) => r.type.endsWith('/officeDocument'))?.target ?? 'xl/workbook.xml';
  const wbDoc = pkg.xml(officeDoc)?.workbook;
  if (!wbDoc) throw new Error('Not a valid xlsx file: workbook part missing');
  const rels = pkg.rels(officeDoc);

  const relOfType = (suffix: string) => [...rels.values()].find((r) => r.type.endsWith(suffix))?.target;
  const ssPath = relOfType('/sharedStrings');
  const sharedStrings = arr(ssPath ? pkg.xml(ssPath)?.sst?.si : undefined).map((si: XNode) =>
    decodeOoxmlEscapes(si.t !== undefined ? arr(si.t).map(text).join('') : arr(si.r).map((r: XNode) => arr(r.t).map(text).join('')).join('')),
  );
  const themePath = relOfType('/theme');
  const theme = readTheme(themePath ? pkg.xml(themePath) : undefined);
  const stylesPath = relOfType('/styles');
  const styles = new StyleReader(stylesPath ? pkg.xml(stylesPath) : undefined, theme);

  const sheets: SheetData[] = [];
  for (const s of arr(wbDoc.sheets?.sheet)) {
    const rel = rels.get(attr(s, 'id') ?? '');
    const name = attr(s, 'name') ?? `Sheet${sheets.length + 1}`;
    if (!rel || !pkg.has(rel.target)) continue; // chart sheets, dialog sheets etc.
    if (!rel.type.endsWith('/worksheet')) continue;
    sheets.push(readSheet(pkg, rel.target, name, sharedStrings, styles));
  }

  const names: Record<string, string> = {};
  for (const dn of arr(wbDoc.definedNames?.definedName)) {
    const name = attr(dn, 'name');
    if (!name || name.startsWith('_xlnm.') || attr(dn, 'localSheetId') !== undefined) continue;
    names[name] = text(dn);
  }
  const data: WorkbookData = { sheets };
  if (Object.keys(names).length) data.names = names;
  const activeTab = numAttr(arr(wbDoc.bookViews?.workbookView)[0], 'activeTab');
  if (activeTab) data.activeSheet = Math.min(activeTab, sheets.length - 1);
  return data;
}

/** Reads an .xlsx file into a Workbook. */
export function readXlsx(bytes: Uint8Array | ArrayBuffer, options: XlsxReadOptions = {}): Workbook {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const data = xlsxToData(u8);
  const wb = new Workbook({ empty: true });
  wb.load(data, options);
  return wb;
}
