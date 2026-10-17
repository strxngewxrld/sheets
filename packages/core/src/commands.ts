import {
  formatAddress,
  formatRange,
  intersectRanges,
  keyToAddress,
  MAX_COLS,
  MAX_ROWS,
  parseRange,
  rangesIntersect,
  type RangeAddress,
} from './address.js';
import { chartFromRange, type ChartSpec, type ChartType, type ChartAnchor } from './chart.js';
import { shiftFormula, transposeFormula } from './formula/rewrite.js';
import { formatValue, getDefaultLocale, parseInput, type FormatLocale } from './numfmt.js';
import type { CellPatch, Operation } from './ops.js';
import { Sheet, type CellData } from './sheet.js';
import { applyStylePatch, type CellStyle, type StylePatch } from './style.js';
import { compareScalars } from './formula/evaluator.js';
import { isError, scalarToJson, type Scalar } from './values.js';
import type { ApplyOptions, Workbook } from './workbook.js';

/**
 * High-level editing commands shared by the UI and the agent API. Each
 * command translates into primitive operations applied in one transaction.
 */

/** Clips whole-row/column ranges to the used area of a sheet. */
export function clipToUsed(sheet: Sheet, range: RangeAddress): RangeAddress | null {
  const used = sheet.usedRange();
  const huge = range.end.row === MAX_ROWS - 1 || range.end.col === MAX_COLS - 1;
  if (!huge) return range;
  if (!used) return null;
  return intersectRanges(range, { start: { row: 0, col: 0 }, end: used.end });
}

function sheetName(wb: Workbook, sheet: string | number): string {
  return wb.requireSheet(sheet).name;
}

// ---------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------

/** Converts typed user input ("12%", "=SUM(A1:A3)", "2024-01-31") into cell data. */
export function inputToCellData(text: string, existingStyle?: CellStyle, locale: FormatLocale = getDefaultLocale()): CellData | null {
  const parsed = parseInput(text, locale);
  const data: CellData = {};
  if (parsed.formula !== undefined) data.f = parsed.formula;
  else if (parsed.value !== null) data.v = scalarToJson(parsed.value);
  let style = existingStyle;
  if (parsed.numFmt && !existingStyle?.numFmt) style = { ...(existingStyle ?? {}), numFmt: parsed.numFmt };
  if (style) data.s = style;
  return data.f === undefined && data.v === undefined && !data.s ? null : data;
}

/** Sets a cell from user input text (like typing into Excel). */
export function setCellInput(wb: Workbook, sheet: string | number, row: number, col: number, text: string, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const data = inputToCellData(text, s.getCell(row, col)?.style);
  wb.apply({ type: 'setCells', sheet: s.name, cells: [{ row, col, data }] }, options);
}

export type InputValue = string | number | boolean | null | undefined;

/**
 * Writes a 2D block of values starting at (row, col). Strings starting with
 * "=" become formulas; other strings are parsed like typed input when
 * `parse` is true (default), otherwise stored as text. Existing styles are kept.
 */
export function setValues(
  wb: Workbook,
  sheet: string | number,
  row: number,
  col: number,
  values: InputValue[][],
  options: ApplyOptions & { parse?: boolean } = {},
): void {
  const s = wb.requireSheet(sheet);
  const cells: CellPatch[] = [];
  values.forEach((rowVals, i) =>
    rowVals.forEach((v, j) => {
      if (v === undefined) return;
      const r = row + i;
      const c = col + j;
      const style = s.getCell(r, c)?.style;
      let data: CellData | null;
      if (typeof v === 'string') {
        if (v.startsWith('=') || options.parse !== false) data = inputToCellData(v, style);
        else data = style ? { v, s: style } : { v };
      } else if (v === null) data = style ? { s: style } : null;
      else data = style ? { v, s: style } : { v };
      cells.push({ row: r, col: c, data });
    }),
  );
  wb.apply({ type: 'setCells', sheet: s.name, cells }, options);
}

export type ClearMode = 'contents' | 'formats' | 'all';

export function clearRange(wb: Workbook, sheet: string | number, range: RangeAddress, mode: ClearMode = 'contents', options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const cells: CellPatch[] = [];
  for (const [key, cell] of s.cells) {
    const { row, col } = keyToAddress(key);
    if (row < range.start.row || row > range.end.row || col < range.start.col || col > range.end.col) continue;
    let data: CellData | null = null;
    if (mode === 'contents') data = cell.style ? { s: cell.style } : null;
    else if (mode === 'formats') {
      data = Sheet.cellToData(cell);
      if (data) delete data.s;
    }
    cells.push({ row, col, data: data && Object.keys(data).length ? data : null, styleOnly: mode === 'formats' });
  }
  wb.apply({ type: 'setCells', sheet: s.name, cells }, options);
}

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

export function applyStyle(wb: Workbook, sheet: string | number, range: RangeAddress, patch: StylePatch, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const target = clipToUsed(s, range);
  if (!target) return;
  if ((target.end.row - target.start.row + 1) * (target.end.col - target.start.col + 1) > 1_000_000) throw new Error('Format range exceeds the cell limit');
  const cells: CellPatch[] = [];
  for (let r = target.start.row; r <= target.end.row; r++) {
    for (let c = target.start.col; c <= target.end.col; c++) {
      const cell = s.getCell(r, c);
      const data = Sheet.cellToData(cell) ?? {};
      const style = applyStylePatch(cell?.style, patch);
      if (style) data.s = style;
      else delete data.s;
      cells.push({ row: r, col: c, data: Object.keys(data).length ? data : null, styleOnly: true });
    }
  }
  wb.apply({ type: 'setCells', sheet: s.name, cells }, options);
}

/** Applies a border around / inside a range. */
export function applyBorders(
  wb: Workbook,
  sheet: string | number,
  range: RangeAddress,
  kind: 'all' | 'outline' | 'none',
  side: { style: 'thin' | 'medium' | 'thick'; color?: string } = { style: 'thin', color: '#000000' },
  options: ApplyOptions = {},
): void {
  const s = wb.requireSheet(sheet);
  const cells: CellPatch[] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    for (let c = range.start.col; c <= range.end.col; c++) {
      const cell = s.getCell(r, c);
      const data = Sheet.cellToData(cell) ?? {};
      let border = kind === 'none' ? undefined : { ...(cell?.style?.border ?? {}) };
      if (border) {
        if (kind === 'all' || r === range.start.row) border.top = side;
        if (kind === 'all' || r === range.end.row) border.bottom = side;
        if (kind === 'all' || c === range.start.col) border.left = side;
        if (kind === 'all' || c === range.end.col) border.right = side;
      }
      const style = applyStylePatch(cell?.style, { border: border ?? null });
      if (style) data.s = style;
      else delete data.s;
      cells.push({ row: r, col: c, data: Object.keys(data).length ? data : null, styleOnly: true });
    }
  }
  wb.apply({ type: 'setCells', sheet: s.name, cells }, options);
}

export function mergeCells(wb: Workbook, sheet: string | number, range: RangeAddress, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  if (s.tables.some((table) => rangesIntersect(parseRange(table.range)!.range, range))) throw new Error('Cannot merge cells inside a table');
  const merges = s.merges.filter((m) => !rangesIntersect(m, range)).map((m) => formatRange(m));
  merges.push(formatRange(range));
  // Excel keeps only the top-left value.
  const cells: CellPatch[] = [];
  for (const [key, cell] of s.cells) {
    const { row, col } = keyToAddress(key);
    if (row === range.start.row && col === range.start.col) continue;
    if (row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col && (cell.value !== null || cell.formula)) {
      cells.push({ row, col, data: cell.style ? { s: cell.style } : null });
    }
  }
  wb.transaction(() => {
    if (cells.length) wb.apply({ type: 'setCells', sheet: s.name, cells });
    wb.apply({ type: 'setMerges', sheet: s.name, merges });
  }, options);
}

export function unmergeCells(wb: Workbook, sheet: string | number, range: RangeAddress, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const merges = s.merges.filter((m) => !rangesIntersect(m, range)).map((m) => formatRange(m));
  wb.apply({ type: 'setMerges', sheet: s.name, merges }, options);
}

// ---------------------------------------------------------------------------
// Clipboard
// ---------------------------------------------------------------------------

export interface ClipboardData {
  sheet: string;
  range: RangeAddress;
  cells: (CellData | null)[][];
  cut?: boolean;
  /** Relative output cells whose formula anchor is included in this copy. */
  derived?: { row: number; col: number }[];
}

export function copyRange(wb: Workbook, sheet: string | number, range: RangeAddress): ClipboardData {
  const s = wb.requireSheet(sheet);
  const r = clipToUsed(s, range) ?? range;
  const cells: (CellData | null)[][] = [];
  const derived: { row: number; col: number }[] = [];
  for (let row = r.start.row; row <= r.end.row; row++) {
    const line: (CellData | null)[] = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = s.getCell(row, col);
      line.push(Sheet.cellToData(cell));
      if (cell?.spillParent !== undefined) {
        const parent = keyToAddress(cell.spillParent);
        if (parent.row >= r.start.row && parent.row <= r.end.row && parent.col >= r.start.col && parent.col <= r.end.col) derived.push({ row: row - r.start.row, col: col - r.start.col });
      }
    }
    cells.push(line);
  }
  return { sheet: s.name, range: r, cells, ...(derived.length ? { derived } : {}) };
}

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas';

export interface PasteSpecial {
  mode?: PasteMode;
  /** Swap rows and columns. */
  transpose?: boolean;
  /** Keep target cells where the source cell is empty. */
  skipBlanks?: boolean;
}

/**
 * Pastes clipboard data at the target cell. Relative references in formulas
 * are shifted. If the target selection is a multiple of the source size the
 * data is tiled (like Excel).
 */
export function pasteRange(
  wb: Workbook,
  clip: ClipboardData,
  sheet: string | number,
  target: RangeAddress,
  special: PasteMode | PasteSpecial = 'all',
  options: ApplyOptions = {},
): RangeAddress {
  const { mode = 'all', transpose = false, skipBlanks = false } = typeof special === 'string' ? { mode: special } : special;
  const s = wb.requireSheet(sheet);
  const srcRows = transpose ? clip.cells[0]?.length ?? 0 : clip.cells.length;
  const srcCols = transpose ? clip.cells.length : clip.cells[0]?.length ?? 0;
  const tRows = target.end.row - target.start.row + 1;
  const tCols = target.end.col - target.start.col + 1;
  const repR = tRows % srcRows === 0 ? tRows / srcRows : 1;
  const repC = tCols % srcCols === 0 ? tCols / srcCols : 1;
  const cells: CellPatch[] = [];
  const written = new Set<string>();
  const derived = new Set(clip.derived?.map((p) => `${p.row},${p.col}`));
  for (let rr = 0; rr < repR; rr++) {
    for (let rc = 0; rc < repC; rc++) {
      clip.cells.forEach((line, i) =>
        line.forEach((src, j) => {
          const [oi, oj] = transpose ? [j, i] : [i, j];
          const row = target.start.row + rr * srcRows + oi;
          const col = target.start.col + rc * srcCols + oj;
          if (row >= MAX_ROWS || col >= MAX_COLS) return;
          if (skipBlanks && (!src || (src.v === undefined && src.f === undefined && (mode !== 'formats' || !src.s)))) return;
          const existing = Sheet.cellToData(s.getCell(row, col));
          let data: CellData | null;
          const dRow = row - (clip.range.start.row + i);
          const dCol = col - (clip.range.start.col + j);
          const shifted: CellData | null = src ? { ...src } : null;
          if (shifted && derived.has(`${i},${j}`) && (mode === 'all' || mode === 'formulas')) delete shifted.v;
          if (shifted?.f !== undefined && !clip.cut) {
            const fromRow = clip.range.start.row + i, fromCol = clip.range.start.col + j;
            shifted.f = transpose ? transposeFormula(shifted.f, fromRow, fromCol, row, col) : shiftFormula(shifted.f, dRow, dCol);
          }
          switch (mode) {
            case 'values':
              data = { ...(existing?.s ? { s: existing.s } : {}), ...(src?.v !== undefined ? { v: src.v } : {}) };
              break;
            case 'formats':
              data = { ...(existing ?? {}), ...(src?.s ? { s: src.s } : {}) };
              if (!src?.s) delete data.s;
              break;
            case 'formulas':
              data = { ...(existing?.s ? { s: existing.s } : {}), ...(shifted?.f !== undefined ? { f: shifted.f } : shifted?.v !== undefined ? { v: shifted.v } : {}) };
              break;
            default:
              data = shifted;
          }
          written.add(`${row},${col}`);
          cells.push({ row, col, data: data && Object.keys(data).length ? data : null, styleOnly: mode === 'formats' });
        }),
      );
    }
  }
  const pasted: RangeAddress = {
    start: target.start,
    end: { row: target.start.row + srcRows * repR - 1, col: target.start.col + srcCols * repC - 1 },
  };
  wb.transaction(() => {
    if (clip.cut) {
      const src = wb.requireSheet(clip.sheet);
      const clear: CellPatch[] = [];
      for (let row = clip.range.start.row; row <= clip.range.end.row; row++) {
        for (let col = clip.range.start.col; col <= clip.range.end.col; col++) {
          if (src === s && written.has(`${row},${col}`)) continue;
          if (src.getCell(row, col)) clear.push({ row, col, data: null });
        }
      }
      if (clear.length) wb.apply({ type: 'setCells', sheet: src.name, cells: clear, reordersCells: true });
    }
    wb.apply({ type: 'setCells', sheet: s.name, cells, ...(clip.cut ? { reordersCells: true } : {}) });
  }, options);
  return pasted;
}

/** Plain text (TSV) representation of a range, using displayed values. */
export function rangeToTSV(wb: Workbook, sheet: string | number, range: RangeAddress): string {
  const s = wb.requireSheet(sheet);
  const lines: string[] = [];
  for (let r = range.start.row; r <= range.end.row; r++) {
    const vals: string[] = [];
    for (let c = range.start.col; c <= range.end.col; c++) {
      const cell = s.getCell(r, c);
      let text = cell ? formatValue(cell.value, cell.style?.numFmt) : '';
      if (/[\t\n"]/.test(text)) text = `"${text.replace(/"/g, '""')}"`;
      vals.push(text);
    }
    lines.push(vals.join('\t'));
  }
  return lines.join('\n');
}

// ---------------------------------------------------------------------------
// Fill (drag handle)
// ---------------------------------------------------------------------------

function seriesStep(values: Scalar[]): number | null {
  if (values.length < 2 || !values.every((v) => typeof v === 'number')) return null;
  const nums = values as number[];
  const step = nums[1]! - nums[0]!;
  for (let i = 2; i < nums.length; i++) if (Math.abs(nums[i]! - nums[i - 1]! - step) > 1e-9) return null;
  return step;
}

/**
 * Extends `source` into `target` (which must contain source and extend it in
 * one direction). Numeric sequences continue linearly, "Item 1" style texts
 * increment, formulas are shifted, everything else is repeated.
 */
export function fillRange(wb: Workbook, sheet: string | number, source: RangeAddress, target: RangeAddress, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const vertical = target.start.col === source.start.col && target.end.col === source.end.col;
  const forward = vertical ? target.end.row > source.end.row : target.end.col > source.end.col;
  const cells: CellPatch[] = [];
  const lanes = vertical ? [source.start.col, source.end.col] : [source.start.row, source.end.row];
  const srcStart = vertical ? source.start.row : source.start.col;
  const srcEnd = vertical ? source.end.row : source.end.col;
  const len = srcEnd - srcStart + 1;
  for (let lane = lanes[0]!; lane <= lanes[1]!; lane++) {
    const at = (i: number) => (vertical ? s.getCell(i, lane) : s.getCell(lane, i));
    const srcCells = Array.from({ length: len }, (_, k) => at(srcStart + k));
    const srcVals = srcCells.map((c) => (c?.formula === undefined ? (c?.value ?? null) : null));
    const step = seriesStep(srcVals);
    const textMatch = srcCells.map((c) => (typeof c?.value === 'string' && c.formula === undefined ? /^(.*?)(\d+)$/.exec(c.value) : null));
    const from = forward ? srcEnd + 1 : vertical ? target.start.row : target.start.col;
    const to = forward ? (vertical ? target.end.row : target.end.col) : srcStart - 1;
    for (let pos = from; pos <= to; pos++) {
      const offset = pos - srcStart;
      const k = ((offset % len) + len) % len;
      const src = srcCells[k];
      const delta = pos - (srcStart + k);
      let data = Sheet.cellToData(src);
      if (data) {
        data = { ...data };
        if (data.f !== undefined) {
          data.f = vertical ? shiftFormula(data.f, delta, 0) : shiftFormula(data.f, 0, delta);
          delete data.v;
        } else if (step !== null) {
          data.v = (srcVals[0] as number) + step * offset;
        } else if (textMatch[k]) {
          const m = textMatch[k]!;
          const next = Number(m[2]) + delta / len;
          data.v = `${m[1]}${String(Math.max(0, Math.round(next))).padStart(m[2]!.length, '0')}`;
        }
      }
      const row = vertical ? pos : lane;
      const col = vertical ? lane : pos;
      cells.push({ row, col, data });
    }
  }
  wb.apply({ type: 'setCells', sheet: s.name, cells }, options);
}

// ---------------------------------------------------------------------------
// Sorting
// ---------------------------------------------------------------------------

export interface SortKey {
  /** Absolute column index (0-based). */
  col: number;
  descending?: boolean;
}

export function sortRange(
  wb: Workbook,
  sheet: string | number,
  range: RangeAddress,
  keys: SortKey[],
  options: ApplyOptions & { hasHeader?: boolean } = {},
): void {
  const s = wb.requireSheet(sheet);
  const r = clipToUsed(s, range);
  if (!r) return;
  const firstRow = r.start.row + (options.hasHeader ? 1 : 0);
  if (!keys.length || keys.some((k) => !Number.isInteger(k.col) || k.col < r.start.col || k.col > r.end.col)) throw new Error('Sort columns must be inside the selected range');
  if ((r.end.row - firstRow + 1) * (r.end.col - r.start.col + 1) > 1_000_000) throw new Error('Sort range exceeds the cell limit');
  if (s.merges.some((merge) => rangesIntersect(merge, r))) throw new Error('Cannot sort a range containing merged cells');
  for (const [key, cell] of s.cells) {
    const p = keyToAddress(key);
    if (p.row >= r.start.row && p.row <= r.end.row && p.col >= r.start.col && p.col <= r.end.col && (cell.spillParent !== undefined || cell.spillRange)) throw new Error('Cannot sort a spilled array. Use the SORT formula instead.');
  }
  const rows: { orig: number; cells: (CellData | null)[]; keys: Scalar[] }[] = [];
  for (let row = firstRow; row <= r.end.row; row++) {
    const cells: (CellData | null)[] = [];
    for (let col = r.start.col; col <= r.end.col; col++) cells.push(Sheet.cellToData(s.getCell(row, col)));
    rows.push({ orig: row, cells, keys: keys.map((k) => s.getValue(row, k.col)) });
  }
  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const x = a.keys[i]!;
      const y = b.keys[i]!;
      // Empty cells always sort last.
      if (x === null || y === null) {
        if (x === y) continue;
        return x === null ? 1 : -1;
      }
      const c = isError(x) || isError(y) ? 0 : compareScalars(x, y);
      if (c !== 0) return keys[i]!.descending ? -c : c;
    }
    return a.orig - b.orig;
  });
  const cells: CellPatch[] = [];
  rows.forEach((entry, i) => {
    const row = firstRow + i;
    entry.cells.forEach((d, j) => {
      let data = d;
      if (data?.f !== undefined) data = { ...data, f: shiftFormula(data.f, row - entry.orig, 0) };
      cells.push({ row, col: r.start.col + j, data });
    });
  });
  wb.apply({ type: 'setCells', sheet: s.name, cells, reordersCells: true }, options);
}

// ---------------------------------------------------------------------------
// Rows / columns / sizes
// ---------------------------------------------------------------------------

export function insertRows(wb: Workbook, sheet: string | number, index: number, count = 1, options: ApplyOptions = {}): void {
  wb.apply({ type: 'insertRows', sheet: sheetName(wb, sheet), index, count }, options);
}
export function deleteRows(wb: Workbook, sheet: string | number, index: number, count = 1, options: ApplyOptions = {}): void {
  wb.apply({ type: 'deleteRows', sheet: sheetName(wb, sheet), index, count }, options);
}
export function insertCols(wb: Workbook, sheet: string | number, index: number, count = 1, options: ApplyOptions = {}): void {
  wb.apply({ type: 'insertCols', sheet: sheetName(wb, sheet), index, count }, options);
}
export function deleteCols(wb: Workbook, sheet: string | number, index: number, count = 1, options: ApplyOptions = {}): void {
  wb.apply({ type: 'deleteCols', sheet: sheetName(wb, sheet), index, count }, options);
}

export function setColumnWidth(wb: Workbook, sheet: string | number, cols: number[], width: number | null, options: ApplyOptions = {}): void {
  wb.apply({ type: 'setColWidths', sheet: sheetName(wb, sheet), sizes: cols.map((index) => ({ index, size: width })) }, options);
}

export function setRowHeight(wb: Workbook, sheet: string | number, rows: number[], height: number | null, options: ApplyOptions = {}): void {
  wb.apply({ type: 'setRowHeights', sheet: sheetName(wb, sheet), sizes: rows.map((index) => ({ index, size: height })) }, options);
}

/** Hides rows (a hidden row has height 0, as in the XLSX model). */
export function hideRows(wb: Workbook, sheet: string | number, first: number, last = first, options: ApplyOptions = {}): void {
  const indices = Array.from({ length: last - first + 1 }, (_, i) => first + i);
  wb.apply({ type: 'setRowHeights', sheet: sheetName(wb, sheet), sizes: indices.map((index) => ({ index, size: 0 })) }, options);
}

/** Shows hidden rows in the range again at the default height. */
export function unhideRows(wb: Workbook, sheet: string | number, first: number, last = first, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const sizes = [...s.rowHeights].filter(([i, h]) => h === 0 && i >= first && i <= last).map(([index]) => ({ index, size: null }));
  if (sizes.length) wb.apply({ type: 'setRowHeights', sheet: s.name, sizes }, options);
}

export function hideCols(wb: Workbook, sheet: string | number, first: number, last = first, options: ApplyOptions = {}): void {
  const indices = Array.from({ length: last - first + 1 }, (_, i) => first + i);
  wb.apply({ type: 'setColWidths', sheet: sheetName(wb, sheet), sizes: indices.map((index) => ({ index, size: 0 })) }, options);
}

export function unhideCols(wb: Workbook, sheet: string | number, first: number, last = first, options: ApplyOptions = {}): void {
  const s = wb.requireSheet(sheet);
  const sizes = [...s.colWidths].filter(([i, w]) => w === 0 && i >= first && i <= last).map(([index]) => ({ index, size: null }));
  if (sizes.length) wb.apply({ type: 'setColWidths', sheet: s.name, sizes }, options);
}

/** True if any row (or column) in the range is hidden manually. */
export function hasHidden(sheet: Sheet, axis: 'row' | 'col', first: number, last: number): boolean {
  for (const [i, size] of axis === 'row' ? sheet.rowHeights : sheet.colWidths) if (size === 0 && i >= first && i <= last) return true;
  return false;
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------

export function addChart(wb: Workbook, sheet: string | number, spec: ChartSpec, options: ApplyOptions = {}): ChartSpec {
  const s = wb.requireSheet(sheet);
  wb.apply({ type: 'setCharts', sheet: s.name, charts: [...s.charts, spec] }, options);
  return spec;
}

export function insertChartFromRange(
  wb: Workbook,
  sheet: string | number,
  range: RangeAddress,
  type: ChartType,
  opts: { title?: string; anchor?: ChartAnchor; seriesInRows?: boolean } & ApplyOptions = {},
): ChartSpec {
  const s = wb.requireSheet(sheet);
  const spec = chartFromRange(wb, s.name, range, type, opts);
  return addChart(wb, s.name, spec, opts);
}

export function updateChart(wb: Workbook, id: string, patch: Partial<Omit<ChartSpec, 'id'>>, options: ApplyOptions = {}): ChartSpec {
  const found = wb.getChart(id);
  if (!found) throw new Error(`Chart not found: ${id}`);
  const updated: ChartSpec = { ...found.chart, ...patch, id };
  for (const k of Object.keys(patch) as (keyof typeof patch)[]) if (patch[k] === undefined) delete updated[k];
  wb.apply({ type: 'setCharts', sheet: found.sheet.name, charts: found.sheet.charts.map((c) => (c.id === id ? updated : c)) }, options);
  return updated;
}

export function removeChart(wb: Workbook, id: string, options: ApplyOptions = {}): void {
  const found = wb.getChart(id);
  if (!found) throw new Error(`Chart not found: ${id}`);
  wb.apply({ type: 'setCharts', sheet: found.sheet.name, charts: found.sheet.charts.filter((c) => c.id !== id) }, options);
}

// ---------------------------------------------------------------------------
// Search
// ---------------------------------------------------------------------------

export interface FindOptions {
  sheet?: string;
  /** Match the whole cell content. */
  wholeCell?: boolean;
  caseSensitive?: boolean;
  regex?: boolean;
  /** Search formulas instead of displayed values. */
  inFormulas?: boolean;
  /** Restrict to a range of the searched sheet(s). */
  range?: RangeAddress;
  limit?: number;
}

export interface FindResult {
  sheet: string;
  address: string;
  row: number;
  col: number;
  value: Scalar;
  display: string;
  formula?: string;
}

function searchPattern(query: string, opts: FindOptions, global = false): RegExp {
  const flags = (opts.caseSensitive ? '' : 'i') + (global ? 'g' : '');
  const source = opts.regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp(opts.wholeCell ? `^(?:${source})$` : source, flags);
}

function inSearchRange(opts: FindOptions, row: number, col: number): boolean {
  const r = opts.range;
  return !r || (row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col);
}

export function findCells(wb: Workbook, query: string, opts: FindOptions = {}): FindResult[] {
  const re = searchPattern(query, opts);
  const out: FindResult[] = [];
  const limit = opts.limit ?? 1000;
  for (const sheet of wb.sheets) {
    if (opts.sheet && sheet.name.toUpperCase() !== opts.sheet.toUpperCase()) continue;
    const keys = [...sheet.cells.keys()].sort((a, b) => a - b);
    for (const key of keys) {
      const cell = sheet.cells.get(key)!;
      if (opts.range) {
        const { row, col } = keyToAddress(key);
        if (!inSearchRange(opts, row, col)) continue;
      }
      const display = formatValue(cell.value, cell.style?.numFmt);
      const haystack = opts.inFormulas ? (cell.formula !== undefined ? `=${cell.formula}` : display) : display;
      if (!re.test(haystack)) continue;
      const { row, col } = keyToAddress(key);
      const res: FindResult = { sheet: sheet.name, address: formatAddress(row, col), row, col, value: cell.value, display };
      if (cell.formula !== undefined) res.formula = `=${cell.formula}`;
      out.push(res);
      if (out.length >= limit) return out;
    }
  }
  return out;
}

export interface ReplaceOptions extends Omit<FindOptions, 'limit'> {
  /** Replace only in these cells (e.g. the current match); defaults to all matches. */
  cells?: { sheet: string; row: number; col: number }[];
}

/** Editable content of a constant cell, as it would be typed. */
function constantText(v: Scalar): string | null {
  if (v === null || isError(v)) return null;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

/**
 * Replaces matches like Excel's Replace dialog: formulas are changed in their
 * text (only when searching formulas), constants in their typed content.
 * Returns the number of changed cells; all changes form one undo step.
 */
export function replaceCells(wb: Workbook, query: string, replacement: string, opts: ReplaceOptions = {}, options: ApplyOptions = {}): number {
  if (query === '') return 0;
  const re = searchPattern(query, opts, true);
  const only = opts.cells && new Set(opts.cells.map((c) => `${c.sheet.toUpperCase()}!${c.row}:${c.col}`));
  const patches = new Map<string, CellPatch[]>();
  let count = 0;
  for (const sheet of wb.sheets) {
    if (opts.sheet && sheet.name.toUpperCase() !== opts.sheet.toUpperCase()) continue;
    for (const [key, cell] of sheet.cells) {
      if (cell.spillParent !== undefined) continue;
      const { row, col } = keyToAddress(key);
      if (!inSearchRange(opts, row, col) || (only && !only.has(`${sheet.name.toUpperCase()}!${row}:${col}`))) continue;
      const content = cell.formula !== undefined ? (opts.inFormulas ? `=${cell.formula}` : null) : constantText(cell.value);
      if (content === null) continue;
      re.lastIndex = 0;
      if (!re.test(content)) continue;
      re.lastIndex = 0;
      const next = content.replace(re, opts.regex ? replacement : replacement.replace(/\$/g, '$$$$'));
      if (next === content) continue;
      const list = patches.get(sheet.name) ?? [];
      list.push({ row, col, data: inputToCellData(next, cell.style) });
      patches.set(sheet.name, list);
      count++;
    }
  }
  if (count) wb.apply([...patches].map(([sheet, cells]) => ({ type: 'setCells' as const, sheet, cells })), options);
  return count;
}

export type { Operation };
