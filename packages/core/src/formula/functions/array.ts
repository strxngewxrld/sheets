import { MAX_COLS, MAX_ROWS } from '../../address.js';
import { CellError, Err, isError, type Scalar } from '../../values.js';
import { compareScalars, toBoolean, toMatrix, type EvalContext, type EvalValue, type Matrix } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { checkMatrixSize as limit } from '../limits.js';
import { bail, boolArg, def, numArg, scalarArg } from './helpers.js';

export { MAX_ARRAY_CELLS } from '../limits.js';

function transpose(matrix: Matrix): Matrix {
  return Array.from({ length: matrix[0]?.length ?? 0 }, (_, c) => matrix.map((row) => row[c] ?? null));
}

function width(m: Matrix): number {
  return m[0]?.length ?? 0;
}

/** Rows (or columns) selected by TAKE/DROP; negative counts work from the end. */
function slice<T>(items: T[], count: number, take: boolean): T[] {
  if (take) return count >= 0 ? items.slice(0, count) : items.slice(Math.max(0, items.length + count));
  return count >= 0 ? items.slice(count) : items.slice(0, Math.max(0, items.length + count));
}

function nonEmpty(m: Matrix): Matrix {
  return m.length && width(m) ? m : bail(new CellError('#CALC!', 'Empty array'));
}

/** One-based indexes (negative from the end) into a list of length n. */
function pick(indexArgs: EvalValue[], n: number, ctx: EvalContext): number[] {
  const out: number[] = [];
  for (const arg of indexArgs) for (const row of toMatrix(arg, ctx)) for (const v of row) {
    const i = Math.trunc(numArg(v, ctx));
    if (i === 0 || Math.abs(i) > n) bail(Err.VALUE);
    out.push(i > 0 ? i - 1 : n + i);
  }
  return out;
}

function padValue(v: EvalValue | undefined, ctx: EvalContext): Scalar {
  return v === undefined ? Err.NA : scalarArg(v, ctx);
}

function flattenTo(m: Matrix, ignore: number, byColumn: boolean): Scalar[] {
  const source = byColumn ? transpose(m) : m;
  return source.flat().filter((v) => !((ignore & 1 && v === null) || (ignore & 2 && isError(v))));
}

function wrap(values: Scalar[], count: number, pad: Scalar): Matrix {
  limit(Math.ceil(values.length / count), count);
  const out: Matrix = [];
  for (let i = 0; i < values.length; i += count) {
    const row = values.slice(i, i + count);
    while (row.length < count) row.push(pad);
    out.push(row);
  }
  return out;
}

function vector(m: Matrix): Scalar[] {
  if (m.length !== 1 && width(m) !== 1) bail(Err.VALUE);
  return m.flat();
}

export const arrayFunctions: FunctionDef[] = [
  def('SEQUENCE', 'lookup', 'SEQUENCE(rows, [columns], [start], [step])', 'Returns a rectangular sequence of numbers.', ([r, c, s, d], ctx) => {
    const rows = Math.trunc(numArg(r, ctx));
    const cols = Math.trunc(numArg(c, ctx, 1));
    const start = numArg(s, ctx, 1);
    const step = numArg(d, ctx, 1);
    if (!(rows >= 1 && rows <= MAX_ROWS && cols >= 1 && cols <= MAX_COLS)) return Err.VALUE;
    limit(rows, cols);
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
  def('VSTACK', 'lookup', 'VSTACK(array1, [array2], ...)', 'Stacks arrays vertically; shorter rows are padded with #N/A.', (args, ctx) => {
    const arrays = args.map((a) => toMatrix(a, ctx));
    const cols = Math.max(...arrays.map(width));
    limit(arrays.reduce((n, m) => n + m.length, 0), cols);
    return arrays.flatMap((m) => m.map((row) => Array.from({ length: cols }, (_, c) => (c < row.length ? row[c]! : Err.NA))));
  }, { minArgs: 1 }),
  def('HSTACK', 'lookup', 'HSTACK(array1, [array2], ...)', 'Stacks arrays horizontally; shorter columns are padded with #N/A.', (args, ctx) => {
    const arrays = args.map((a) => toMatrix(a, ctx));
    const rows = Math.max(...arrays.map((m) => m.length));
    limit(rows, arrays.reduce((n, m) => n + width(m), 0));
    return Array.from({ length: rows }, (_, r) => arrays.flatMap((m) => m[r] ?? Array.from({ length: width(m) }, () => Err.NA)));
  }, { minArgs: 1 }),
  def('TAKE', 'lookup', 'TAKE(array, rows, [columns])', 'Returns rows or columns from the start (positive) or end (negative) of an array.', ([a, r, c], ctx) => {
    const m = toMatrix(a!, ctx);
    const rows = r === undefined || scalarArg(r, ctx) === null ? m.length : Math.trunc(numArg(r, ctx));
    const cols = c === undefined || scalarArg(c, ctx) === null ? width(m) : Math.trunc(numArg(c, ctx));
    if (rows === 0 || cols === 0) return new CellError('#CALC!', 'Empty array');
    return nonEmpty(slice(m, rows, true).map((row) => slice(row, cols, true)));
  }, { minArgs: 2, maxArgs: 3 }),
  def('DROP', 'lookup', 'DROP(array, rows, [columns])', 'Removes rows or columns from the start (positive) or end (negative) of an array.', ([a, r, c], ctx) => {
    const m = toMatrix(a!, ctx);
    const rows = Math.trunc(numArg(r, ctx, 0));
    const cols = Math.trunc(numArg(c, ctx, 0));
    return nonEmpty(slice(m, rows, false).map((row) => slice(row, cols, false)));
  }, { minArgs: 2, maxArgs: 3 }),
  def('CHOOSEROWS', 'lookup', 'CHOOSEROWS(array, row_num1, [row_num2], ...)', 'Returns the given rows of an array; negative numbers count from the end.', ([a, ...idx], ctx) => {
    const m = toMatrix(a!, ctx);
    const rows = pick(idx, m.length, ctx);
    limit(rows.length, width(m));
    return rows.map((i) => m[i]!);
  }, { minArgs: 2 }),
  def('CHOOSECOLS', 'lookup', 'CHOOSECOLS(array, col_num1, [col_num2], ...)', 'Returns the given columns of an array; negative numbers count from the end.', ([a, ...idx], ctx) => {
    const m = toMatrix(a!, ctx);
    const cols = pick(idx, width(m), ctx);
    limit(m.length, cols.length);
    return m.map((row) => cols.map((c) => row[c] ?? null));
  }, { minArgs: 2 }),
  def('TOCOL', 'lookup', 'TOCOL(array, [ignore], [scan_by_column])', 'Returns an array as one column; ignore 1 = blanks, 2 = errors, 3 = both.', ([a, i, b], ctx) => {
    const ignore = Math.trunc(numArg(i, ctx, 0));
    if (ignore < 0 || ignore > 3) return Err.VALUE;
    const values = flattenTo(toMatrix(a!, ctx), ignore, boolArg(b, ctx));
    return values.length ? values.map((v) => [v]) : new CellError('#CALC!', 'Empty array');
  }, { minArgs: 1, maxArgs: 3 }),
  def('TOROW', 'lookup', 'TOROW(array, [ignore], [scan_by_column])', 'Returns an array as one row; ignore 1 = blanks, 2 = errors, 3 = both.', ([a, i, b], ctx) => {
    const ignore = Math.trunc(numArg(i, ctx, 0));
    if (ignore < 0 || ignore > 3) return Err.VALUE;
    const values = flattenTo(toMatrix(a!, ctx), ignore, boolArg(b, ctx));
    return values.length ? [values] : new CellError('#CALC!', 'Empty array');
  }, { minArgs: 1, maxArgs: 3 }),
  def('WRAPROWS', 'lookup', 'WRAPROWS(vector, wrap_count, [pad_with])', 'Wraps a row or column into rows of the given length.', ([v, n, p], ctx) => {
    const count = Math.trunc(numArg(n, ctx));
    if (count < 1) return Err.NUM;
    return wrap(vector(toMatrix(v!, ctx)), count, padValue(p, ctx));
  }, { minArgs: 2, maxArgs: 3 }),
  def('WRAPCOLS', 'lookup', 'WRAPCOLS(vector, wrap_count, [pad_with])', 'Wraps a row or column into columns of the given length.', ([v, n, p], ctx) => {
    const count = Math.trunc(numArg(n, ctx));
    if (count < 1) return Err.NUM;
    return transpose(wrap(vector(toMatrix(v!, ctx)), count, padValue(p, ctx)));
  }, { minArgs: 2, maxArgs: 3 }),
  def('EXPAND', 'lookup', 'EXPAND(array, rows, [columns], [pad_with])', 'Pads an array to the given size.', ([a, r, c, p], ctx) => {
    const m = toMatrix(a!, ctx);
    const rows = r === undefined || scalarArg(r, ctx) === null ? m.length : Math.trunc(numArg(r, ctx));
    const cols = c === undefined || scalarArg(c, ctx) === null ? width(m) : Math.trunc(numArg(c, ctx));
    if (rows < m.length || cols < width(m)) return Err.VALUE;
    limit(rows, cols);
    const pad = padValue(p, ctx);
    return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => (m[i] && j < m[i]!.length ? m[i]![j]! : pad)));
  }, { minArgs: 2, maxArgs: 4 }),
  def('SORTBY', 'lookup', 'SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)', 'Sorts an array by the values of other arrays.', ([a, ...rest], ctx) => {
    const m = toMatrix(a!, ctx);
    const keys: { values: Scalar[]; order: number; byCol: boolean }[] = [];
    for (let i = 0; i < rest.length; i += 2) {
      const by = toMatrix(rest[i]!, ctx);
      const order = numArg(rest[i + 1], ctx, 1);
      if (order !== 1 && order !== -1) return Err.VALUE;
      const byCol = by.length === 1 && width(by) > 1;
      if (byCol ? width(by) !== width(m) : by.length !== m.length || width(by) !== 1) return Err.VALUE;
      keys.push({ values: by.flat(), order, byCol });
    }
    if (!keys.length || keys.some((k) => k.byCol !== keys[0]!.byCol)) return Err.VALUE;
    const byCol = keys[0]!.byCol;
    const items = byCol ? transpose(m) : m;
    const order = items.map((_, i) => i).sort((x, y) => {
      for (const k of keys) {
        const p = k.values[x] ?? null;
        const q = k.values[y] ?? null;
        if (p === null || q === null) { if (p !== q) return p === null ? 1 : -1; continue; }
        const c = compareScalars(p, q);
        if (c) return k.order * c;
      }
      return x - y;
    });
    const sorted = order.map((i) => items[i]!);
    return byCol ? transpose(sorted) : sorted;
  }, { minArgs: 2 }),
];
