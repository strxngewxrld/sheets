import { clipToUsed, keyToAddress, MAX_COLS, MAX_ROWS, parseRange, tableAt, type CellAddress, type RangeAddress, type Sheet, type SheetTable } from '@sheets/core';

/** The contiguous non-empty block around a cell (Excel's current region). */
export function currentRegion(sheet: Sheet, cell: CellAddress): RangeAddress {
  // One pass over the sheet builds sorted occupancy lists, so every "is this row/column filled
  // within [a, b]" probe is a binary search instead of a scan (large tables used to be quadratic).
  const byRow = new Map<number, number[]>();
  const byCol = new Map<number, number[]>();
  for (const [key, data] of sheet.cells) {
    if ((data.value ?? null) === null) continue;
    const p = keyToAddress(key);
    (byRow.get(p.row) ?? byRow.set(p.row, []).get(p.row)!).push(p.col);
    (byCol.get(p.col) ?? byCol.set(p.col, []).get(p.col)!).push(p.row);
  }
  for (const list of byRow.values()) list.sort((a, b) => a - b);
  for (const list of byCol.values()) list.sort((a, b) => a - b);
  const anyIn = (list: number[] | undefined, first: number, last: number) => {
    if (!list || first > last) return false;
    let lo = 0;
    let hi = list.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (list[mid]! < first) lo = mid + 1;
      else hi = mid;
    }
    return lo < list.length && list[lo]! <= last;
  };
  const range = { start: { ...cell }, end: { ...cell } };
  const filledRow = (row: number, first: number, last: number) => anyIn(byRow.get(row), first, Math.min(last, MAX_COLS - 1));
  const filledCol = (col: number, first: number, last: number) => anyIn(byCol.get(col), first, Math.min(last, MAX_ROWS - 1));
  let changed = true;
  while (changed) {
    changed = false;
    if (range.start.row > 0 && filledRow(range.start.row - 1, Math.max(0, range.start.col - 1), range.end.col + 1)) { range.start.row--; changed = true; }
    if (range.end.row < MAX_ROWS - 1 && filledRow(range.end.row + 1, Math.max(0, range.start.col - 1), range.end.col + 1)) { range.end.row++; changed = true; }
    if (range.start.col > 0 && filledCol(range.start.col - 1, Math.max(0, range.start.row - 1), range.end.row + 1)) { range.start.col--; changed = true; }
    if (range.end.col < MAX_COLS - 1 && filledCol(range.end.col + 1, Math.max(0, range.start.row - 1), range.end.row + 1)) { range.end.col++; changed = true; }
  }
  return range;
}

export interface SortContext { range: RangeAddress; hasHeader: boolean; table?: SheetTable }

export function resolveSortContext(sheet: Sheet, selection: RangeAddress, active: CellAddress): SortContext {
  const table = tableAt(sheet, active.row, active.col);
  if (table) {
    const range = parseRange(table.range)!.range;
    if (selection.start.row >= range.start.row && selection.end.row <= range.end.row && selection.start.col >= range.start.col && selection.end.col <= range.end.col) return { range, hasHeader: table.headerRow, table };
  }
  let range = clipToUsed(sheet, selection) ?? selection;
  if (range.start.col === range.end.col || (range.start.row === range.end.row && range.start.col === range.end.col)) {
    let seed = active;
    if (sheet.getValue(seed.row, seed.col) === null) {
      const used = sheet.usedRange();
      for (let row = used?.start.row ?? 0; row <= (used?.end.row ?? 0); row++) {
        if (sheet.getValue(row, active.col) !== null) { seed = { row, col: active.col }; break; }
      }
    }
    range = currentRegion(sheet, seed);
    const within = tableAt(sheet, range.start.row, range.start.col);
    if (within) return { range: parseRange(within.range)!.range, hasHeader: within.headerRow, table: within };
  }
  const headers = Array.from({ length: range.end.col - range.start.col + 1 }, (_, i) => sheet.getValue(range.start.row, range.start.col + i));
  const hasHeader = range.end.row > range.start.row && headers.every((v) => typeof v === 'string' && v.trim() !== '') && new Set(headers.map((v) => String(v).toLowerCase())).size === headers.length;
  return { range, hasHeader };
}
