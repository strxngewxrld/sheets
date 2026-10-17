import { describe, expect, it } from 'vitest';
import { findCells, setValues, Workbook } from '@sheets/core';
import { nextMatchIndex } from '../src/find-dialog.js';

describe('find navigation', () => {
  const wb = new Workbook();
  setValues(wb, 0, 0, 0, [['x', null, 'x'], [null, 'x', null]]);
  wb.addSheet('Two');
  setValues(wb, 'Two', 0, 0, [['x']]);
  const matches = findCells(wb, 'x');

  it('continues after the active cell and wraps across sheets', () => {
    expect(matches.map((m) => `${m.sheet}!${m.address}`)).toEqual(['Sheet1!A1', 'Sheet1!C1', 'Sheet1!B2', 'Two!A1']);
    expect(nextMatchIndex(wb, matches, { sheet: 'Sheet1', row: 0, col: 0 }, false)).toBe(1);
    expect(nextMatchIndex(wb, matches, { sheet: 'Sheet1', row: 0, col: 1 }, false)).toBe(1);
    expect(nextMatchIndex(wb, matches, { sheet: 'Sheet1', row: 5, col: 0 }, false)).toBe(3);
    expect(nextMatchIndex(wb, matches, { sheet: 'Two', row: 0, col: 0 }, false)).toBe(0);
  });

  it('goes backwards and wraps to the last match', () => {
    expect(nextMatchIndex(wb, matches, { sheet: 'Sheet1', row: 1, col: 1 }, true)).toBe(1);
    expect(nextMatchIndex(wb, matches, { sheet: 'Sheet1', row: 0, col: 0 }, true)).toBe(3);
    expect(nextMatchIndex(wb, [], { sheet: 'Sheet1', row: 0, col: 0 }, true)).toBe(-1);
  });
});
