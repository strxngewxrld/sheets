import { cellKey, formatRange, keyToAddress, parseAddress, parseRange, formatAddress, type RangeAddress } from './address.js';
import type { ChartSpec } from './chart.js';
import type { CellStyle } from './style.js';
import { scalarFromJson, scalarToJson, type JsonScalar, type Scalar } from './values.js';

export interface Cell {
  /** Current (computed) value. */
  value: Scalar;
  /** Formula without the leading "=". */
  formula?: string;
  style?: CellStyle;
  /** Value came from a file cache and must be kept if the formula cannot be evaluated. */
  cached?: boolean;
}

/** Serializable cell content. */
export interface CellData {
  v?: JsonScalar;
  f?: string;
  s?: CellStyle;
}

export interface SheetData {
  name: string;
  /** Cells keyed by A1 address. */
  cells: Record<string, CellData>;
  colWidths?: Record<number, number>;
  rowHeights?: Record<number, number>;
  merges?: string[];
  frozen?: { rows: number; cols: number };
  charts?: ChartSpec[];
  defaultColWidth?: number;
  defaultRowHeight?: number;
}

export const DEFAULT_COL_WIDTH = 64;
export const DEFAULT_ROW_HEIGHT = 20;

let nextSheetId = 1;

export class Sheet {
  readonly id = nextSheetId++;
  name: string;
  readonly cells = new Map<number, Cell>();
  readonly colWidths = new Map<number, number>();
  readonly rowHeights = new Map<number, number>();
  merges: RangeAddress[] = [];
  frozen = { rows: 0, cols: 0 };
  charts: ChartSpec[] = [];
  defaultColWidth = DEFAULT_COL_WIDTH;
  defaultRowHeight = DEFAULT_ROW_HEIGHT;
  private usedCache: RangeAddress | null | undefined;

  constructor(name: string) {
    this.name = name;
  }

  getCell(row: number, col: number): Cell | undefined {
    return this.cells.get(cellKey(row, col));
  }

  getValue(row: number, col: number): Scalar {
    return this.cells.get(cellKey(row, col))?.value ?? null;
  }

  /** Low-level write without recalculation or history; prefer Workbook.apply(). */
  putCell(row: number, col: number, cell: Cell | undefined): void {
    const key = cellKey(row, col);
    if (cell && (cell.value !== null || cell.formula !== undefined || cell.style)) this.cells.set(key, cell);
    else this.cells.delete(key);
    this.usedCache = undefined;
  }

  invalidate(): void {
    this.usedCache = undefined;
  }

  /** Bounding box of all non-empty cells (and charts are ignored). */
  usedRange(): RangeAddress | null {
    if (this.usedCache !== undefined) return this.usedCache;
    let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
    for (const key of this.cells.keys()) {
      const { row, col } = keyToAddress(key);
      if (row < minR) minR = row;
      if (row > maxR) maxR = row;
      if (col < minC) minC = col;
      if (col > maxC) maxC = col;
    }
    this.usedCache = maxR < 0 ? null : { start: { row: minR, col: minC }, end: { row: maxR, col: maxC } };
    return this.usedCache;
  }

  colWidth(col: number): number {
    return this.colWidths.get(col) ?? this.defaultColWidth;
  }

  rowHeight(row: number): number {
    return this.rowHeights.get(row) ?? this.defaultRowHeight;
  }

  /** Returns the merge containing the cell, if any. */
  mergeAt(row: number, col: number): RangeAddress | undefined {
    return this.merges.find((m) => row >= m.start.row && row <= m.end.row && col >= m.start.col && col <= m.end.col);
  }

  static cellToData(cell: Cell | undefined): CellData | null {
    if (!cell) return null;
    const d: CellData = {};
    if (cell.formula !== undefined) d.f = cell.formula;
    if (cell.value !== null) d.v = scalarToJson(cell.value);
    if (cell.style) d.s = cell.style;
    return d;
  }

  static dataToCell(d: CellData): Cell {
    const cell: Cell = { value: scalarFromJson(d.v) };
    if (d.f !== undefined) {
      cell.formula = d.f.startsWith('=') ? d.f.slice(1) : d.f;
      // A stored result is kept if the formula cannot be evaluated (e.g. unsupported function).
      if (d.v !== undefined) cell.cached = true;
    }
    if (d.s) cell.style = d.s;
    return cell;
  }

  toJSON(): SheetData {
    const cells: Record<string, CellData> = {};
    for (const [key, cell] of this.cells) {
      const { row, col } = keyToAddress(key);
      cells[formatAddress(row, col)] = Sheet.cellToData(cell)!;
    }
    const data: SheetData = { name: this.name, cells };
    if (this.colWidths.size) data.colWidths = Object.fromEntries(this.colWidths);
    if (this.rowHeights.size) data.rowHeights = Object.fromEntries(this.rowHeights);
    if (this.merges.length) data.merges = this.merges.map((m) => formatRange(m));
    if (this.frozen.rows || this.frozen.cols) data.frozen = { ...this.frozen };
    if (this.charts.length) data.charts = structuredClone(this.charts);
    if (this.defaultColWidth !== DEFAULT_COL_WIDTH) data.defaultColWidth = this.defaultColWidth;
    if (this.defaultRowHeight !== DEFAULT_ROW_HEIGHT) data.defaultRowHeight = this.defaultRowHeight;
    return data;
  }

  static fromJSON(data: SheetData): Sheet {
    const sheet = new Sheet(data.name);
    for (const [addr, cd] of Object.entries(data.cells ?? {})) {
      const a = parseAddress(addr);
      if (a) sheet.cells.set(cellKey(a.row, a.col), Sheet.dataToCell(cd));
    }
    for (const [k, v] of Object.entries(data.colWidths ?? {})) sheet.colWidths.set(Number(k), v);
    for (const [k, v] of Object.entries(data.rowHeights ?? {})) sheet.rowHeights.set(Number(k), v);
    sheet.merges = (data.merges ?? []).map((m) => parseRange(m)?.range).filter((r): r is RangeAddress => !!r);
    if (data.frozen) sheet.frozen = { ...data.frozen };
    sheet.charts = structuredClone(data.charts ?? []);
    if (data.defaultColWidth) sheet.defaultColWidth = data.defaultColWidth;
    if (data.defaultRowHeight) sheet.defaultRowHeight = data.defaultRowHeight;
    return sheet;
  }
}
