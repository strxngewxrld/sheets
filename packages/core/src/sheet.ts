import { cellKey, formatRange, keyToAddress, parseAddress, parseRange, formatAddress, type RangeAddress } from './address.js';
import type { ChartSpec } from './chart.js';
import type { ConditionalFormatRule } from './conditional.js';
import type { SheetTable } from './table.js';
import type { CellStyle } from './style.js';
import type { DataValidation } from './validation.js';
import { scalarFromJson, scalarToJson, type JsonScalar, type Scalar } from './values.js';

export interface Cell {
  /** Current (computed) value. */
  value: Scalar;
  /** Formula without the leading "=". */
  formula?: string;
  style?: CellStyle;
  /** Value came from a file cache and must be kept if the formula cannot be evaluated. */
  cached?: boolean;
  /** Derived array cell: key of its formula anchor. Never stored as user input. */
  spillParent?: number;
  /** Current output rectangle of an array formula. */
  spillRange?: RangeAddress;
}

/** A cell note (Excel comment), shown read-only. */
export interface CellNote {
  author?: string;
  text: string;
}

/** A hyperlink stored on a cell: an external target or a location in the workbook. */
export interface CellLink {
  /** External address, e.g. "https://example.com" or "mailto:a@b.c". */
  target?: string;
  /** Location in the workbook, e.g. "Sheet2!A1" or a defined name. */
  location?: string;
  tooltip?: string;
}

/** Serializable cell content. */
export interface CellData {
  v?: JsonScalar;
  f?: string;
  s?: CellStyle;
  /** Cached array extent (primarily used by XLSX import). */
  a?: string;
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
  conditionalFormats?: ConditionalFormatRule[];
  tables?: SheetTable[];
  validations?: DataValidation[];
  /** Notes keyed by A1 address. */
  notes?: Record<string, CellNote>;
  /** Hyperlinks keyed by A1 address. */
  links?: Record<string, CellLink>;
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
  conditionalFormats: ConditionalFormatRule[] = [];
  tables: SheetTable[] = [];
  validations: DataValidation[] = [];
  /** Notes and hyperlinks keyed by cellKey(row, col). */
  readonly notes = new Map<number, CellNote>();
  readonly links = new Map<number, CellLink>();
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
    for (const table of this.tables) {
      const range = parseRange(table.range)?.range;
      if (!range) continue;
      minR = Math.min(minR, range.start.row); minC = Math.min(minC, range.start.col);
      maxR = Math.max(maxR, range.end.row); maxC = Math.max(maxC, range.end.col);
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
      if (cell.spillParent !== undefined) {
        if (cell.style) cells[formatAddress(row, col)] = { s: cell.style };
      } else cells[formatAddress(row, col)] = Sheet.cellToData(cell)!;
    }
    const data: SheetData = { name: this.name, cells };
    if (this.colWidths.size) data.colWidths = Object.fromEntries(this.colWidths);
    if (this.rowHeights.size) data.rowHeights = Object.fromEntries(this.rowHeights);
    if (this.merges.length) data.merges = this.merges.map((m) => formatRange(m));
    if (this.frozen.rows || this.frozen.cols) data.frozen = { ...this.frozen };
    if (this.charts.length) data.charts = structuredClone(this.charts);
    if (this.conditionalFormats.length) data.conditionalFormats = structuredClone(this.conditionalFormats);
    if (this.tables.length) data.tables = structuredClone(this.tables);
    if (this.validations.length) data.validations = structuredClone(this.validations);
    const byAddress = <T>(map: Map<number, T>) => Object.fromEntries([...map].map(([key, v]) => { const a = keyToAddress(key); return [formatAddress(a.row, a.col), structuredClone(v)]; }));
    if (this.notes.size) data.notes = byAddress(this.notes);
    if (this.links.size) data.links = byAddress(this.links);
    if (this.defaultColWidth !== DEFAULT_COL_WIDTH) data.defaultColWidth = this.defaultColWidth;
    if (this.defaultRowHeight !== DEFAULT_ROW_HEIGHT) data.defaultRowHeight = this.defaultRowHeight;
    return data;
  }

  static fromJSON(data: SheetData): Sheet {
    const sheet = new Sheet(data.name);
    const cells = data.cells ?? {};
    const anchors: [string, string][] = [];
    for (const addr in cells) {
      const cd = cells[addr]!;
      const a = parseAddress(addr);
      if (a) sheet.cells.set(cellKey(a.row, a.col), Sheet.dataToCell(cd));
      if (cd.a) anchors.push([addr, cd.a]);
    }
    for (const [addr, extent] of anchors) {
      const anchor = parseAddress(addr);
      const range = parseRange(extent)?.range;
      if (!anchor || !range || !sheet.getCell(anchor.row, anchor.col)?.formula) continue;
      const parent = cellKey(anchor.row, anchor.col);
      sheet.cells.get(parent)!.spillRange = range;
      const mark = (key: number, cell: Cell) => { if (key !== parent && !cell.formula) cell.spillParent = parent; };
      const area = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
      if (area <= sheet.cells.size) {
        for (let r = range.start.row; r <= range.end.row; r++) for (let c = range.start.col; c <= range.end.col; c++) {
          const key = cellKey(r, c), cell = sheet.cells.get(key);
          if (cell) mark(key, cell);
        }
      } else {
        for (const [key, cell] of sheet.cells) {
          const p = keyToAddress(key);
          if (p.row >= range.start.row && p.row <= range.end.row && p.col >= range.start.col && p.col <= range.end.col) mark(key, cell);
        }
      }
    }
    for (const [k, v] of Object.entries(data.colWidths ?? {})) sheet.colWidths.set(Number(k), v);
    for (const [k, v] of Object.entries(data.rowHeights ?? {})) sheet.rowHeights.set(Number(k), v);
    sheet.merges = (data.merges ?? []).map((m) => parseRange(m)?.range).filter((r): r is RangeAddress => !!r);
    if (data.frozen) sheet.frozen = { ...data.frozen };
    sheet.charts = structuredClone(data.charts ?? []);
    sheet.conditionalFormats = structuredClone(data.conditionalFormats ?? []);
    sheet.tables = structuredClone(data.tables ?? []);
    sheet.validations = structuredClone(data.validations ?? []);
    for (const [map, entries] of [[sheet.notes, data.notes], [sheet.links, data.links]] as const) {
      for (const [addr, v] of Object.entries(entries ?? {})) {
        const a = parseAddress(addr);
        if (a) (map as Map<number, unknown>).set(cellKey(a.row, a.col), structuredClone(v));
      }
    }
    if (data.defaultColWidth) sheet.defaultColWidth = data.defaultColWidth;
    if (data.defaultRowHeight) sheet.defaultRowHeight = data.defaultRowHeight;
    return sheet;
  }
}
