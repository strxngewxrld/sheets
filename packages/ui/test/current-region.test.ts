import { describe, expect, it } from 'vitest';
import { MAX_COLS, MAX_ROWS, setValues, Sheet, Workbook, type CellAddress, type RangeAddress } from '@sheets/core';
import { currentRegion } from '../src/sheet-actions.js';

// The original scanning implementation, kept as the reference for equivalence checks.
function reference(sheet: Sheet, cell: CellAddress): RangeAddress {
  const range = { start: { ...cell }, end: { ...cell } };
  const filledRow = (row: number, first: number, last: number) => { for (let c = first; c <= Math.min(last, MAX_COLS - 1); c++) if (sheet.getValue(row, c) !== null) return true; return false; };
  const filledCol = (col: number, first: number, last: number) => { for (let r = first; r <= Math.min(last, MAX_ROWS - 1); r++) if (sheet.getValue(r, col) !== null) return true; return false; };
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

describe('currentRegion', () => {
  it('matches the reference on random sheets', () => {
    let seed = 12345;
    const rnd = () => (seed = (seed * 1664525 + 1013904223) % 4294967296) / 4294967296;
    for (let round = 0; round < 300; round++) {
      const wb = new Workbook();
      const rows = 3 + Math.floor(rnd() * 25), cols = 2 + Math.floor(rnd() * 12), density = 0.1 + rnd() * 0.6;
      const m: (number | null)[][] = Array.from({ length: rows }, () => Array.from({ length: cols }, () => (rnd() < density ? 1 : null)));
      setValues(wb, 0, 0, 0, m);
      const s = wb.requireSheet(0);
      for (let k = 0; k < 4; k++) {
        const cell = { row: Math.floor(rnd() * (rows + 3)), col: Math.floor(rnd() * (cols + 3)) };
        expect(currentRegion(s, cell), JSON.stringify({ round, cell })).toEqual(reference(s, cell));
      }
    }
  });

  it('handles a 200k row table quickly and does not truncate it', () => {
    const n = 200_000;
    // Built directly: the test measures currentRegion, not CSV parsing.
    const s = new Sheet('big');
    for (let r = 0; r <= n; r++) for (let c = 0; c < 3; c++) s.putCell(r, c, { value: r * 3 + c });
    const t = Date.now();
    const r = currentRegion(s, { row: n, col: 0 });
    expect(r).toEqual({ start: { row: 0, col: 0 }, end: { row: n, col: 2 } });
    expect(currentRegion(s, { row: 5, col: 6 })).toEqual({ start: { row: 5, col: 6 }, end: { row: 5, col: 6 } });
    // The previous scan needed minutes here; a generous bound keeps the test stable on slow machines.
    expect(Date.now() - t).toBeLessThan(15_000);
  }, 60_000);
});
