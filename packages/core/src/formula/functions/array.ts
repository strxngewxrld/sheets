import { MAX_COLS, MAX_ROWS } from '../../address.js';
import { CellError, Err, isError, type Scalar } from '../../values.js';
import { compareScalars, toBoolean, toMatrix, type Matrix } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { boolArg, def, numArg } from './helpers.js';

export const MAX_ARRAY_CELLS = 1_000_000;

function transpose(matrix: Matrix): Matrix {
  return Array.from({ length: matrix[0]?.length ?? 0 }, (_, c) => matrix.map((row) => row[c] ?? null));
}

export const arrayFunctions: FunctionDef[] = [
  def('SEQUENCE', 'lookup', 'SEQUENCE(rows, [columns], [start], [step])', 'Returns a rectangular sequence of numbers.', ([r, c, s, d], ctx) => {
    const rows = Math.trunc(numArg(r, ctx));
    const cols = Math.trunc(numArg(c, ctx, 1));
    const start = numArg(s, ctx, 1);
    const step = numArg(d, ctx, 1);
    if (!(rows >= 1 && rows <= MAX_ROWS && cols >= 1 && cols <= MAX_COLS)) return Err.VALUE;
    if (rows * cols > MAX_ARRAY_CELLS) return new CellError('#CALC!', 'Array exceeds the cell limit');
    return Array.from({ length: rows }, (_, row) => Array.from({ length: cols }, (_, col) => {
      const n = start + (row * cols + col) * step;
      return Number.isFinite(n) ? n : Err.NUM;
    }));
  }, { minArgs: 1, maxArgs: 4 }),
  def('FILTER', 'lookup', 'FILTER(array, include, [if_empty])', 'Filters rows or columns using a Boolean include vector.', ([a, i, fallback], ctx) => {
    const matrix = toMatrix(a!, ctx);
    const mask = toMatrix(i!, ctx);
    const rows = matrix.length;
    const cols = matrix[0]?.length ?? 0;
    const vertical = mask.length === rows && mask.every((row) => row.length === 1);
    const horizontal = mask.length === 1 && mask[0]!.length === cols;
    if (!vertical && !horizontal) return Err.VALUE;
    const flags = (vertical ? mask.map((row) => row[0] ?? null) : mask[0]!).map(toBoolean);
    const error = flags.find(isError);
    if (error) return error;
    const result = vertical ? matrix.filter((_, row) => flags[row]) : matrix.map((row) => row.filter((_, col) => flags[col]));
    return !result.length || !result[0]?.length ? fallback ?? new CellError('#CALC!', 'No matching values') : result;
  }, { minArgs: 2, maxArgs: 3 }),
  def('SORT', 'lookup', 'SORT(array, [sort_index], [sort_order], [by_col])', 'Sorts rows or columns by an index; order is 1 or -1.', ([a, i, o, b], ctx) => {
    const matrix = toMatrix(a!, ctx);
    const index = numArg(i, ctx, 1);
    const order = numArg(o, ctx, 1);
    const byCol = boolArg(b, ctx);
    const rows = byCol ? transpose(matrix) : matrix;
    if (!Number.isInteger(index) || index < 1 || index > (rows[0]?.length ?? 0) || (order !== 1 && order !== -1)) return Err.VALUE;
    const sorted = [...rows].sort((a, b) => {
      const x = a[index - 1] ?? null;
      const y = b[index - 1] ?? null;
      if (x === null || y === null) return x === y ? 0 : x === null ? 1 : -1;
      return order * compareScalars(x, y);
    });
    return byCol ? transpose(sorted) : sorted;
  }, { minArgs: 1, maxArgs: 4 }),
  def('UNIQUE', 'lookup', 'UNIQUE(array, [by_col], [exactly_once])', 'Returns distinct rows or columns, optionally only those occurring once.', ([a, b, e], ctx) => {
    const matrix = toMatrix(a!, ctx);
    const byCol = boolArg(b, ctx);
    const once = boolArg(e, ctx);
    const rows = byCol ? transpose(matrix) : matrix;
    const key = (row: Scalar[]) => JSON.stringify(row.map((v) => isError(v) ? ['error', v.code] : [typeof (v ?? 0), typeof v === 'string' ? v.toLowerCase() : v ?? 0]));
    const counts = new Map<string, number>();
    for (const row of rows) { const k = key(row); counts.set(k, (counts.get(k) ?? 0) + 1); }
    const seen = new Set<string>();
    const result = rows.filter((row) => {
      const k = key(row);
      if (seen.has(k) || (once && counts.get(k) !== 1)) return false;
      seen.add(k);
      return true;
    });
    if (!result.length) return new CellError('#CALC!', 'No unique values');
    return byCol ? transpose(result) : result;
  }, { minArgs: 1, maxArgs: 3 }),
];
