import { colToName, formatAddress, MAX_COLS, MAX_ROWS, normalizeRange, quoteSheetName } from '../../address.js';
import { Err, isError, type Scalar } from '../../values.js';
import { compareScalars, isRef, toMatrix, toScalar, type EvalValue, type FunctionContext, type Matrix } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { bail, boolArg, def, numArg, rangeCols, rangeRows, scalarArg, textArg, wildcardToRegExp } from './helpers.js';

function cellAt(m: Matrix, r: number, c: number): Scalar {
  return m[r]?.[c] ?? null;
}

function vector(arg: EvalValue, ctx: FunctionContext): Scalar[] {
  const m = toMatrix(arg, ctx);
  const rows = rangeRows(arg, m);
  const cols = rangeCols(arg, m);
  const out: Scalar[] = [];
  if (rows === 1) for (let c = 0; c < cols; c++) out.push(cellAt(m, 0, c));
  else if (cols === 1) for (let r = 0; r < rows; r++) out.push(cellAt(m, r, 0));
  else bail(Err.NA);
  return out;
}

function exactMatcher(needle: Scalar, wildcards: boolean): (v: Scalar) => boolean {
  if (wildcards && typeof needle === 'string' && /[*?~]/.test(needle)) {
    const re = wildcardToRegExp(needle);
    return (v) => typeof v === 'string' && re.test(v);
  }
  return (v) => v !== null && typeof v === typeof needle && compareScalars(v, needle) === 0;
}

/** Binary search for the largest value <= needle (ascending data). */
function approxAscending(values: Scalar[], needle: Scalar): number {
  let lo = 0;
  let hi = values.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = values[mid]!;
    if (v === null || typeof v !== typeof needle) {
      // Skip mismatched types linearly (Excel behaves similarly).
      hi = mid - 1;
      continue;
    }
    if (compareScalars(v, needle) <= 0) {
      found = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return found;
}

function matchIndex(needle: Scalar, values: Scalar[], mode: number): number {
  if (mode === 0) return values.findIndex(exactMatcher(needle, true));
  if (mode === 1) return approxAscending(values, needle);
  // -1: smallest value >= needle, data descending
  let found = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i]!;
    if (v === null || typeof v !== typeof needle) continue;
    if (compareScalars(v, needle) >= 0) found = i;
    else break;
  }
  return found;
}

export const lookupFunctions: FunctionDef[] = [
  def('VLOOKUP', 'lookup', 'VLOOKUP(lookup_value, table, col_index, [range_lookup])', 'Looks up a value in the first column of a table.', ([needle, table, colIdx, approx], ctx) => {
    const v = scalarArg(needle, ctx);
    if (isError(v)) return v;
    const m = toMatrix(table!, ctx);
    const col = Math.trunc(numArg(colIdx, ctx));
    if (col < 1) return Err.VALUE;
    if (col > rangeCols(table!, m)) return Err.REF;
    const first = m.map((row) => row[0] ?? null);
    const idx = boolArg(approx, ctx, true) ? approxAscending(first, v) : first.findIndex(exactMatcher(v, true));
    return idx < 0 ? Err.NA : cellAt(m, idx, col - 1);
  }, { minArgs: 3, maxArgs: 4 }),
  def('HLOOKUP', 'lookup', 'HLOOKUP(lookup_value, table, row_index, [range_lookup])', 'Looks up a value in the first row of a table.', ([needle, table, rowIdx, approx], ctx) => {
    const v = scalarArg(needle, ctx);
    if (isError(v)) return v;
    const m = toMatrix(table!, ctx);
    const row = Math.trunc(numArg(rowIdx, ctx));
    if (row < 1) return Err.VALUE;
    if (row > rangeRows(table!, m)) return Err.REF;
    const first = m[0] ?? [];
    const idx = boolArg(approx, ctx, true) ? approxAscending(first, v) : first.findIndex(exactMatcher(v, true));
    return idx < 0 ? Err.NA : cellAt(m, row - 1, idx);
  }, { minArgs: 3, maxArgs: 4 }),
  def('MATCH', 'lookup', 'MATCH(lookup_value, lookup_array, [match_type])', 'Position of a value in a range (1-based).', ([needle, arr, type], ctx) => {
    const v = scalarArg(needle, ctx);
    if (isError(v)) return v;
    const idx = matchIndex(v, vector(arr!, ctx), Math.sign(numArg(type, ctx, 1)));
    return idx < 0 ? Err.NA : idx + 1;
  }, { minArgs: 2, maxArgs: 3 }),
  def('XMATCH', 'lookup', 'XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])', 'Position of a value (exact by default).', ([needle, arr, mode, search], ctx) => {
    const v = scalarArg(needle, ctx);
    if (isError(v)) return v;
    const values = vector(arr!, ctx);
    const idx = xIndex(v, values, numArg(mode, ctx, 0), numArg(search, ctx, 1));
    return idx < 0 ? Err.NA : idx + 1;
  }, { minArgs: 2, maxArgs: 4 }),
  def('XLOOKUP', 'lookup', 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])', 'Modern lookup in any direction.', ([needle, arr, ret, notFound, mode, search], ctx) => {
    const v = scalarArg(needle, ctx);
    if (isError(v)) return v;
    const values = vector(arr!, ctx);
    const idx = xIndex(v, values, numArg(mode, ctx, 0), numArg(search, ctx, 1));
    if (idx < 0) return notFound !== undefined && notFound !== null ? notFound : Err.NA;
    const rm = toMatrix(ret!, ctx);
    const vertical = rangeCols(arr!, toMatrix(arr!, ctx)) === 1;
    if (vertical) {
      const row = rm[idx] ?? [];
      const width = rangeCols(ret!, rm);
      if (width === 1) return row[0] ?? null;
      return [Array.from({ length: width }, (_, c) => row[c] ?? null)];
    }
    const height = rangeRows(ret!, rm);
    if (height === 1) return cellAt(rm, 0, idx);
    return Array.from({ length: height }, (_, r) => [cellAt(rm, r, idx)]);
  }, { minArgs: 3, maxArgs: 6 }),
  def('INDEX', 'lookup', 'INDEX(array, row_num, [column_num])', 'Value at a given position of a range.', ([arr, r, c], ctx) => {
    const m = toMatrix(arr!, ctx);
    const rows = rangeRows(arr!, m);
    const cols = rangeCols(arr!, m);
    let row = Math.trunc(numArg(r, ctx, 0));
    let col = Math.trunc(numArg(c, ctx, 0));
    if (c === undefined && rows === 1) [row, col] = [1, row];
    if (row < 0 || col < 0 || row > rows || col > cols) return Err.REF;
    if (isRef(arr!)) {
      const { start } = arr.range;
      if (row === 0 && col === 0) return arr;
      if (row === 0) return { kind: 'ref', sheet: arr.sheet, range: normalizeRange({ row: start.row, col: start.col + col - 1 }, { row: arr.range.end.row, col: start.col + col - 1 }) };
      if (col === 0) {
        if (cols === 1) return { kind: 'ref', sheet: arr.sheet, range: normalizeRange({ row: start.row + row - 1, col: start.col }, { row: start.row + row - 1, col: start.col }) };
        return { kind: 'ref', sheet: arr.sheet, range: normalizeRange({ row: start.row + row - 1, col: start.col }, { row: start.row + row - 1, col: arr.range.end.col }) };
      }
      return { kind: 'ref', sheet: arr.sheet, range: normalizeRange({ row: start.row + row - 1, col: start.col + col - 1 }, { row: start.row + row - 1, col: start.col + col - 1 }) };
    }
    if (row === 0) return m.map((rr) => [rr[Math.max(col, 1) - 1] ?? null]);
    if (col === 0) return cols === 1 ? cellAt(m, row - 1, 0) : [m[row - 1] ?? []];
    return cellAt(m, row - 1, col - 1);
  }, { minArgs: 2, maxArgs: 3 }),
  def('ROW', 'lookup', 'ROW([reference])', 'Row number of a reference.', ([ref], ctx) => (ref === undefined ? ctx.row + 1 : isRef(ref) ? ref.range.start.row + 1 : Err.VALUE), { maxArgs: 1 }),
  def('COLUMN', 'lookup', 'COLUMN([reference])', 'Column number of a reference.', ([ref], ctx) => (ref === undefined ? ctx.col + 1 : isRef(ref) ? ref.range.start.col + 1 : Err.VALUE), { maxArgs: 1 }),
  def('ROWS', 'lookup', 'ROWS(array)', 'Number of rows in a range.', ([a], ctx) => rangeRows(a!, toMatrix(a!, ctx)), { minArgs: 1, maxArgs: 1 }),
  def('COLUMNS', 'lookup', 'COLUMNS(array)', 'Number of columns in a range.', ([a], ctx) => rangeCols(a!, toMatrix(a!, ctx)), { minArgs: 1, maxArgs: 1 }),
  def('OFFSET', 'lookup', 'OFFSET(reference, rows, cols, [height], [width])', 'Reference shifted by rows/cols.', ([ref, r, c, h, w], ctx) => {
    if (!isRef(ref!)) return Err.VALUE;
    const height = numArg(h, ctx, ref.range.end.row - ref.range.start.row + 1);
    const width = numArg(w, ctx, ref.range.end.col - ref.range.start.col + 1);
    const sr = ref.range.start.row + Math.trunc(numArg(r, ctx));
    const sc = ref.range.start.col + Math.trunc(numArg(c, ctx));
    if (sr < 0 || sc < 0 || height < 1 || width < 1) return Err.REF;
    return { kind: 'ref', sheet: ref.sheet, range: { start: { row: sr, col: sc }, end: { row: sr + height - 1, col: sc + width - 1 } } };
  }, { minArgs: 3, maxArgs: 5, volatile: true }),
  def('INDIRECT', 'lookup', 'INDIRECT(ref_text, [a1])', 'Reference from a text string (A1 or, with a1 = FALSE, R1C1 style).', ([t, a1], ctx) => {
    const text = textArg(t, ctx);
    const ref = a1 === undefined || boolArg(a1, ctx, true) ? text : r1c1ToA1(text, ctx.row, ctx.col);
    const node = ref === null ? undefined : ctx.parseRef?.(ref);
    if (!node || (node.type !== 'cell' && node.type !== 'range')) return Err.REF;
    return ctx.evaluate(node);
  }, { minArgs: 1, maxArgs: 2, volatile: true }),
  def('ADDRESS', 'lookup', 'ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])', 'Cell address as text (abs_num 1 = $A$1, 2 = A$1, 3 = $A1, 4 = A1).', ([r, c, abs, a1, sheet], ctx) => {
    const row = Math.trunc(numArg(r, ctx));
    const col = Math.trunc(numArg(c, ctx));
    const mode = Math.trunc(numArg(abs, ctx, 1));
    if (row < 1 || col < 1 || row > MAX_ROWS || col > MAX_COLS || mode < 1 || mode > 4) return Err.VALUE;
    const absRow = mode === 1 || mode === 2, absCol = mode === 1 || mode === 3;
    const ref = boolArg(a1, ctx, true)
      ? formatAddress(row - 1, col - 1, absRow, absCol)
      : `R${absRow ? row : `[${row}]`}C${absCol ? col : `[${col}]`}`;
    if (sheet === undefined || sheet === null) return ref;
    const name = textArg(sheet, ctx);
    return `${name === '' ? '' : quoteSheetName(name)}!${ref}`;
  }, { minArgs: 2, maxArgs: 5 }),
  def('AREAS', 'lookup', 'AREAS(reference)', 'Number of areas in a reference.', ([ref]) => (isRef(ref!) ? 1 : Err.VALUE), { minArgs: 1, maxArgs: 1 }),
  def('SHEET', 'lookup', 'SHEET([value])', 'Sheet number of a reference or sheet name.', ([v], ctx) => {
    const names = ctx.sheetNames?.() ?? [ctx.sheet];
    const target = v === undefined ? ctx.sheet : isRef(v) ? v.sheet : scalarArg(v, ctx);
    if (isError(target)) return target;
    const index = typeof target === 'string' ? names.findIndex((n) => n.toUpperCase() === target.toUpperCase()) : -1;
    return index < 0 ? Err.NA : index + 1;
  }, { maxArgs: 1, volatile: true }),
  def('SHEETS', 'lookup', 'SHEETS([reference])', 'Number of sheets in the workbook or in a reference.', ([ref], ctx) => {
    if (ref === undefined) return (ctx.sheetNames?.() ?? [ctx.sheet]).length;
    return isRef(ref) ? 1 : Err.VALUE;
  }, { maxArgs: 1, volatile: true }),
  def('TRANSPOSE', 'lookup', 'TRANSPOSE(array)', 'Swaps rows and columns.', ([a], ctx) => {
    const m = toMatrix(a!, ctx);
    const cols = m[0]?.length ?? 0;
    return Array.from({ length: cols }, (_, c) => m.map((row) => row[c] ?? null));
  }, { minArgs: 1, maxArgs: 1 }),
  def('LOOKUP', 'lookup', 'LOOKUP(lookup_value, lookup_vector, [result_vector])', 'Approximate lookup in a sorted vector.', ([needle, vec, res], ctx) => {
    const v = toScalar(needle!, ctx);
    const values = vector(vec!, ctx);
    const idx = approxAscending(values, v);
    if (idx < 0) return Err.NA;
    return (res ? vector(res, ctx) : values)[idx] ?? Err.NA;
  }, { minArgs: 2, maxArgs: 3 }),
];

/**
 * Converts an R1C1 reference (R2C3, R[1]C[-1], R2C3:R4C5, R2, C3, with an optional
 * sheet prefix) to A1 style. Relative parts count from the formula's cell (0-based
 * row and col). Returns null for text that is not an R1C1 reference.
 */
export function r1c1ToA1(text: string, row: number, col: number): string | null {
  const bang = text.lastIndexOf('!');
  const sheet = bang >= 0 ? text.slice(0, bang + 1) : '';
  // R2 is absolute (1-based), R[-1] relative to the formula's cell, a bare R the same row.
  const index = (spec: string | undefined, current: number): number | null => {
    const value = !spec ? current : spec.startsWith('[') ? current + Number(spec.slice(1, -1)) : Number(spec) - 1;
    return Number.isInteger(value) && value >= 0 ? value : null;
  };
  const N = '(\\[-?\\d+\\]|\\d+)?';
  const cell = new RegExp(`^R${N}C${N}$`, 'i'), rowOnly = new RegExp(`^R${N}$`, 'i'), colOnly = new RegExp(`^C${N}$`, 'i');
  const parts = text.slice(bang + 1).split(':').map((p) => p.trim());
  if (parts.length > 2) return null;
  const parsed = parts.map((p) => {
    let m = cell.exec(p);
    if (m) return { r: index(m[1], row), c: index(m[2], col) };
    if ((m = rowOnly.exec(p))) return { r: index(m[1], row), c: undefined };
    if ((m = colOnly.exec(p))) return { r: undefined, c: index(m[1], col) };
    return null;
  });
  if (parsed.some((p) => p === null || p.r === null || p.c === null)) return null;
  const [a, b = a] = parsed as { r?: number; c?: number }[];
  if (a!.r !== undefined && a!.c !== undefined && b!.r !== undefined && b!.c !== undefined) {
    const first = formatAddress(a!.r, a!.c);
    return sheet + (parts.length === 1 ? first : `${first}:${formatAddress(b!.r, b!.c)}`);
  }
  if (a!.c === undefined && b!.c === undefined) return `${sheet}${a!.r! + 1}:${b!.r! + 1}`;
  if (a!.r === undefined && b!.r === undefined) return `${sheet}${colToName(a!.c!)}:${colToName(b!.c!)}`;
  return null;
}

function xIndex(needle: Scalar, values: Scalar[], mode: number, search: number): number {
  const order = search < 0 ? [...values.keys()].reverse() : [...values.keys()];
  if (mode === 2 || mode === 0) {
    const m = exactMatcher(needle, mode === 2);
    for (const i of order) if (m(values[i]!)) return i;
    if (mode === 0) return -1;
  }
  // -1: exact or next smaller, 1: exact or next larger
  let best = -1;
  for (const i of order) {
    const v = values[i]!;
    if (v === null || typeof v !== typeof needle) continue;
    const c = compareScalars(v, needle);
    if (c === 0) return i;
    if (mode === -1 && c < 0 && (best < 0 || compareScalars(v, values[best]!) > 0)) best = i;
    if (mode === 1 && c > 0 && (best < 0 || compareScalars(v, values[best]!) < 0)) best = i;
  }
  return best;
}
