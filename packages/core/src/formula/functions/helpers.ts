import { Err, isError, type CellError, type Scalar } from '../../values.js';
import { compileSafeRegex, RE2JS } from '../../regex.js';
import {
  broadcast,
  checkNumber,
  compareScalars,
  deref,
  isMatrix,
  isRef,
  toMatrix,
  toNumber,
  type EvalContext,
  type EvalValue,
  type Matrix,
} from '../evaluator.js';
import type { EagerFunctionDef, FunctionCategory } from '../registry.js';

/** Throws a CellError so it can be returned from deep inside a function implementation. */
export function bail(e: CellError): never {
  throw e;
}

export function num(v: Scalar): number {
  const n = toNumber(v);
  if (isError(n)) bail(n);
  return n;
}

/**
 * Collects numbers following Excel's aggregate rules: direct arguments are
 * coerced, values inside ranges are only counted when numeric.
 */
export function numbersOf(args: EvalValue[], ctx: EvalContext): number[] {
  const out: number[] = [];
  for (const a of args) {
    const v = deref(a, ctx);
    if (isMatrix(v) || isRef(a)) {
      for (const row of toMatrix(v, ctx)) {
        for (const s of row) {
          if (typeof s === 'number') out.push(s);
          else if (isError(s)) bail(s);
        }
      }
    } else if (v !== null) {
      out.push(num(v));
    }
  }
  return out;
}

/** Iterates every scalar (including empty cells) of all arguments. */
export function* scalarsOf(args: EvalValue[], ctx: EvalContext): Generator<Scalar> {
  for (const a of args) {
    if (isRef(a) || isMatrix(a)) {
      for (const row of toMatrix(a, ctx)) yield* row;
    } else yield a;
  }
}

export function flatten(m: Matrix): Scalar[] {
  return m.flat();
}

export function scalarArg(v: EvalValue | undefined, ctx: EvalContext): Scalar {
  if (v === undefined) return null;
  const d = deref(v, ctx);
  if (isMatrix(d)) return d[0]?.[0] ?? null;
  return d;
}

export function numArg(v: EvalValue | undefined, ctx: EvalContext, fallback?: number): number {
  const s = scalarArg(v, ctx);
  if (s === null && fallback !== undefined) return fallback;
  return num(s);
}

export function textArg(v: EvalValue | undefined, ctx: EvalContext): string {
  const s = scalarArg(v, ctx);
  if (isError(s)) bail(s);
  if (s === null) return '';
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE';
  if (typeof s === 'number') return String(Number(s.toPrecision(15)));
  return s;
}

export function boolArg(v: EvalValue | undefined, ctx: EvalContext, fallback = false): boolean {
  const s = scalarArg(v, ctx);
  if (s === null) return fallback;
  if (isError(s)) bail(s);
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  const u = s.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return bail(Err.VALUE);
}

export function roundHalfAway(x: number, digits: number): number {
  const f = Math.pow(10, digits);
  const v = Math.abs(x) * f;
  const r = Math.round(Number(v.toPrecision(15)));
  return (Math.sign(x) * r) / f;
}

// ---------------------------------------------------------------------------
// Criteria (COUNTIF, SUMIF, ...)
// ---------------------------------------------------------------------------

export function wildcardToRegExp(pattern: string, unanchored = false): RE2JS {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!;
    if (ch === '~' && i + 1 < pattern.length) re += escapeRe(pattern[++i]!);
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += escapeRe(ch);
  }
  return compileSafeRegex(unanchored ? re : `^${re}$`, RE2JS.CASE_INSENSITIVE | RE2JS.DOTALL);
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function makeCriteria(crit: Scalar): (v: Scalar) => boolean {
  if (typeof crit === 'number') return (v) => (typeof v === 'number' ? v === crit : typeof v === 'string' && v.trim() !== '' && Number(v) === crit);
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => v === null || v === '';
  if (isError(crit)) return (v) => isError(v) && v.code === crit.code;
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(crit)!;
  const op = m[1] ?? '=';
  const rest = m[2]!;
  const n = rest.trim() !== '' ? Number(rest) : NaN;
  if (!Number.isNaN(n)) {
    return (v) => {
      if (typeof v !== 'number') return op === '<>' ? true : op === '=' && typeof v === 'string' && Number(v) === n && v.trim() !== '';
      switch (op) {
        case '=': return v === n;
        case '<>': return v !== n;
        case '<': return v < n;
        case '>': return v > n;
        case '<=': return v <= n;
        default: return v >= n;
      }
    };
  }
  if (rest === '') {
    if (op === '=') return (v) => v === null || v === '';
    if (op === '<>') return (v) => v !== null && v !== '';
  }
  const upper = rest.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    return (v) => (op === '<>' ? v !== b : v === b);
  }
  if (op === '=' || op === '<>') {
    const re = /[*?~]/.test(rest) ? wildcardToRegExp(rest) : null;
    const lower = rest.toLowerCase();
    const match = (v: Scalar) => typeof v === 'string' && (re ? re.test(v) : v.toLowerCase() === lower);
    return op === '=' ? match : (v) => !match(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const c = compareScalars(v, rest);
    return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
  };
}

/** Evaluates (range, criteria) pairs and returns a mask of matching positions. */
export function criteriaMask(pairs: [EvalValue, EvalValue][], ctx: EvalContext): { mask: boolean[]; rows: number; cols: number } {
  let rows = -1;
  let cols = -1;
  let mask: boolean[] = [];
  for (const [rangeArg, critArg] of pairs) {
    const m = toMatrix(rangeArg, ctx);
    const r = rangeRows(rangeArg, m);
    const c = rangeCols(rangeArg, m);
    if (rows < 0) {
      rows = r;
      cols = c;
      mask = new Array(rows * cols).fill(true);
    } else if (r !== rows || c !== cols) bail(Err.VALUE);
    const pred = makeCriteria(scalarArg(critArg, ctx));
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const idx = i * cols + j;
        if (mask[idx] && !pred(m[i]?.[j] ?? null)) mask[idx] = false;
      }
    }
  }
  return { mask, rows, cols };
}

/** Logical size of a range argument (refs keep their declared size even if clipped). */
export function rangeRows(arg: EvalValue, m: Matrix): number {
  if (isRef(arg)) return arg.range.end.row - arg.range.start.row + 1;
  return m.length;
}

export function rangeCols(arg: EvalValue, m: Matrix): number {
  if (isRef(arg)) return arg.range.end.col - arg.range.start.col + 1;
  return m[0]?.length ?? 0;
}

// ---------------------------------------------------------------------------
// Definition builders
// ---------------------------------------------------------------------------

export function def(
  name: string,
  category: FunctionCategory,
  syntax: string,
  description: string,
  impl: EagerFunctionDef['impl'],
  extra: Partial<EagerFunctionDef> = {},
): EagerFunctionDef {
  return { name, category, syntax, description, impl, ...extra };
}

/** Numeric function of one argument, broadcast over arrays. */
export function unary(name: string, description: string, f: (x: number) => number | CellError, category: FunctionCategory = 'math'): EagerFunctionDef {
  return def(name, category, `${name}(number)`, description, ([a], ctx) =>
    broadcast([deref(a!, ctx)], (s) => {
      const n = toNumber(s);
      if (isError(n)) return n;
      const r = f(n);
      return isError(r) ? r : checkNumber(r);
    }),
    { minArgs: 1, maxArgs: 1 },
  );
}

/** Numeric function of two arguments, broadcast over arrays. */
export function binary(
  name: string,
  syntax: string,
  description: string,
  f: (x: number, y: number) => number | CellError,
  opts: { defaultY?: number; category?: FunctionCategory } = {},
): EagerFunctionDef {
  return def(name, opts.category ?? 'math', syntax, description, ([a, b], ctx) =>
    broadcast([deref(a!, ctx), b === undefined ? (opts.defaultY ?? null) : deref(b, ctx)], (x, y) => {
      const n = toNumber(x);
      if (isError(n)) return n;
      const m = y === null && opts.defaultY !== undefined ? opts.defaultY : toNumber(y);
      if (isError(m)) return m;
      const r = f(n, m);
      return isError(r) ? r : checkNumber(r);
    }),
    { minArgs: opts.defaultY !== undefined ? 1 : 2, maxArgs: 2 },
  );
}
