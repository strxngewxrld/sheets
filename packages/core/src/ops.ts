import type { ChartSpec } from './chart.js';
import type { CellData, SheetData } from './sheet.js';

/** A single cell write. `data: null` deletes the cell (value, formula and style). */
export interface CellPatch {
  row: number;
  col: number;
  data: CellData | null;
}

export interface SizePatch {
  index: number;
  /** Size in pixels, `null` resets to the default. */
  size: number | null;
}

/**
 * Primitive, JSON-serializable workbook mutations. Every change to a
 * workbook goes through these, which makes them usable for undo/redo,
 * collaboration, logging and remote control (e.g. by AI agents).
 */
export type Operation =
  | { type: 'setCells'; sheet: string; cells: CellPatch[] }
  | { type: 'insertRows'; sheet: string; index: number; count: number }
  | { type: 'deleteRows'; sheet: string; index: number; count: number }
  | { type: 'insertCols'; sheet: string; index: number; count: number }
  | { type: 'deleteCols'; sheet: string; index: number; count: number }
  | { type: 'addSheet'; name: string; index?: number; data?: SheetData }
  | { type: 'removeSheet'; name: string }
  | { type: 'renameSheet'; from: string; to: string }
  | { type: 'moveSheet'; name: string; index: number }
  | { type: 'setColWidths'; sheet: string; sizes: SizePatch[] }
  | { type: 'setRowHeights'; sheet: string; sizes: SizePatch[] }
  | { type: 'setMerges'; sheet: string; merges: string[] }
  | { type: 'setFrozen'; sheet: string; rows: number; cols: number }
  | { type: 'setCharts'; sheet: string; charts: ChartSpec[] }
  | { type: 'setName'; name: string; ref: string | null }
  | { type: 'restore'; data: WorkbookData };

export type OperationType = Operation['type'];

export interface WorkbookData {
  sheets: SheetData[];
  /** Defined names: name → reference formula (e.g. "Sheet1!$A$1:$A$10"). */
  names?: Record<string, string>;
  activeSheet?: number;
}
