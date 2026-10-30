import { formatRange, parseRange, type CellAddress, type RangeAddress } from './address.js';
import { formatValue } from './numfmt.js';
import type { Sheet } from './sheet.js';
import { isError, type Scalar } from './values.js';
import type { Workbook } from './workbook.js';

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
export const CHART_TYPES: readonly ChartType[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

export interface ChartSeries {
  /** Literal series name. */
  name?: string;
  /** Reference to a cell containing the series name, e.g. "Sheet1!$B$1". */
  nameRef?: string;
  /** Sheet-qualified reference to the values, e.g. "Sheet1!$B$2:$B$10". */
  values: string;
  /** Category labels (x values for scatter charts). */
  categories?: string;
  /** CSS color override. */
  color?: string;
}

/** Cell-based anchor: the chart spans from the top-left of `from` to the top-left of `to`. */
export interface ChartAnchor {
  from: CellAddress;
  to: CellAddress;
}

export type LegendPosition = 'top' | 'bottom' | 'left' | 'right' | 'none';

export interface ChartSpec {
  id: string;
  type: ChartType;
  title?: string;
  series: ChartSeries[];
  anchor: ChartAnchor;
  stacked?: boolean;
  legend?: LegendPosition;
  xAxisTitle?: string;
  yAxisTitle?: string;
}

export interface ResolvedSeries {
  name: string;
  values: (number | null)[];
  /** X values for scatter charts. */
  x?: (number | null)[];
  color?: string;
}

export interface ResolvedChart {
  categories: string[];
  series: ResolvedSeries[];
}

let chartCounter = 0;
export function newChartId(): string {
  chartCounter++;
  return `chart-${Date.now().toString(36)}-${chartCounter}`;
}

/** Points per reference a chart shows, like Excel's limit for 2-D charts; files may reference whole sheets. */
export const MAX_CHART_POINTS = 32_000;

/** Visits the cells of a reference within the sheet's used area, at most MAX_CHART_POINTS. */
function eachRefCell(wb: Workbook, ref: string | undefined, visit: (sheet: Sheet, row: number, col: number) => void): void {
  if (!ref) return;
  const parsed = parseRange(ref);
  const sheet = parsed && wb.getSheet(parsed.sheet ?? wb.sheets[0]!.name);
  if (!parsed || !sheet) return;
  const { start, end } = parsed.range;
  const used = sheet.usedRange();
  const lastRow = Math.min(end.row, used?.end.row ?? start.row);
  const lastCol = Math.min(end.col, used?.end.col ?? start.col);
  let points = 0;
  for (let r = start.row; r <= lastRow; r++) {
    for (let c = start.col; c <= lastCol; c++) {
      if (points++ >= MAX_CHART_POINTS) return;
      visit(sheet, r, c);
    }
  }
}

function refValues(wb: Workbook, ref: string | undefined): Scalar[] {
  const out: Scalar[] = [];
  eachRefCell(wb, ref, (sheet, r, c) => out.push(sheet.getValue(r, c)));
  return out;
}

function refDisplay(wb: Workbook, ref: string | undefined): string[] {
  const out: string[] = [];
  eachRefCell(wb, ref, (sheet, r, c) => {
    const cell = sheet.getCell(r, c);
    out.push(cell ? formatValue(cell.value, cell.style?.numFmt) : '');
  });
  return out;
}

const toNum = (v: Scalar): number | null => (typeof v === 'number' && !isError(v) ? v : null);

/** Resolves the live data behind a chart. */
export function resolveChart(wb: Workbook, spec: ChartSpec): ResolvedChart {
  const categories = refDisplay(wb, spec.series[0]?.categories);
  const series = spec.series.map((s, i): ResolvedSeries => {
    const nameVal = s.nameRef ? refValues(wb, s.nameRef)[0] : undefined;
    const name = s.name ?? (nameVal !== undefined && nameVal !== null ? String(nameVal) : `Series ${i + 1}`);
    const out: ResolvedSeries = { name, values: refValues(wb, s.values).map(toNum) };
    if (spec.type === 'scatter') out.x = refValues(wb, s.categories).map(toNum);
    if (s.color) out.color = s.color;
    return out;
  });
  const maxLen = Math.max(0, ...series.map((s) => s.values.length));
  while (categories.length < maxLen) categories.push(String(categories.length + 1));
  return { categories: categories.slice(0, Math.max(maxLen, 0)), series };
}

/**
 * Builds a chart spec from a data range the way Excel does: a text header row
 * becomes series names, a leading text column becomes the categories.
 */
export function chartFromRange(
  wb: Workbook,
  sheetName: string,
  range: RangeAddress,
  type: ChartType,
  options: { title?: string; anchor?: ChartAnchor; seriesInRows?: boolean } = {},
): ChartSpec {
  const sheet = wb.getSheet(sheetName);
  if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
  const { start, end } = range;
  const val = (r: number, c: number) => sheet.getValue(r, c);
  const isText = (v: Scalar) => typeof v === 'string' && v !== '';
  const byRows = options.seriesInRows ?? false;

  // Normalize orientation into a "major" (series) and "minor" (points) axis.
  const majorStart = byRows ? start.row : start.col;
  const majorEnd = byRows ? end.row : end.col;
  const minorStart = byRows ? start.col : start.row;
  const minorEnd = byRows ? end.col : end.row;
  const at = (major: number, minor: number) => (byRows ? val(major, minor) : val(minor, major));
  const ref = (major: number, m0: number, m1: number) =>
    formatAbs(sheetName, byRows ? { start: { row: major, col: m0 }, end: { row: major, col: m1 } } : { start: { row: m0, col: major }, end: { row: m1, col: major } });

  let hasHeader = false;
  for (let m = majorStart; m <= majorEnd; m++) {
    if (isText(at(m, minorStart)) && minorEnd > minorStart) hasHeader = true;
  }
  const firstMajorText = (() => {
    for (let n = minorStart + (hasHeader ? 1 : 0); n <= minorEnd; n++) if (isText(at(majorStart, n))) return true;
    return false;
  })();
  const hasCategories = majorEnd > majorStart && (firstMajorText || type === 'scatter');
  const dataStart = minorStart + (hasHeader ? 1 : 0);
  const categories = hasCategories ? ref(majorStart, dataStart, minorEnd) : undefined;

  const series: ChartSeries[] = [];
  for (let m = majorStart + (hasCategories ? 1 : 0); m <= majorEnd; m++) {
    const s: ChartSeries = { values: ref(m, dataStart, minorEnd) };
    if (hasHeader) s.nameRef = formatAbs(sheetName, byRows ? { start: { row: m, col: minorStart }, end: { row: m, col: minorStart } } : { start: { row: minorStart, col: m }, end: { row: minorStart, col: m } });
    if (categories) s.categories = categories;
    series.push(s);
  }
  const anchor = options.anchor ?? {
    from: { row: start.row, col: end.col + 2 },
    to: { row: start.row + 15, col: end.col + 9 },
  };
  const spec: ChartSpec = { id: newChartId(), type, series, anchor, legend: series.length > 1 || type === 'pie' || type === 'doughnut' ? 'right' : 'none' };
  if (options.title) spec.title = options.title;
  else if (series.length === 1 && hasHeader) {
    const t = at(majorStart + (hasCategories ? 1 : 0), minorStart);
    if (typeof t === 'string') spec.title = t;
  }
  return spec;
}

function formatAbs(sheet: string, range: RangeAddress): string {
  const plain = formatRange(range, sheet);
  return plain.replace(/(!|:)([A-Z]+)(\d+)/g, '$1$$$2$$$3');
}
