import { normalizeRange, singleCellRange, type RangeAddress } from '../address.js';
import { formatGeneral } from '../numfmt.js';
import { CellError, Err, isError, type Scalar } from '../values.js';
import type { Node, StructuredRef } from './ast.js';
import { getFunction } from './registry.js';

export type Matrix = Scalar[][];

export interface RefValue {
  kind: 'ref';
  sheet: string;
  range: RangeAddress;
}

/** Intermediate value produced while evaluating a formula. */
export type EvalValue = Scalar | Matrix | RefValue;

/** Services the evaluator needs from its host (the workbook). */
export interface EvalContext {
  /** Sheet the formula lives on. */
  readonly sheet: string;
  readonly row: number;
  readonly col: number;
  getCell(sheet: string, row: number, col: number): Scalar;
  /** Returns the values of a range. Implementations may clip trailing empty rows/cols. */
  getRange(sheet: string, range: RangeAddress): Matrix;
  sheetExists(sheet: string): boolean;
  /** Resolves a defined name to its parsed formula. */
  resolveName(name: string): Node | undefined;
  /** Parses a reference string (for INDIRECT). */
  parseRef?(ref: string): Node | undefined;
  getSpillRange?(sheet: string, row: number, col: number): RangeAddress | CellError;
  /** Resolves a structured table reference for the formula cell. */
  resolveStructured?(ref: StructuredRef, sheet: string, row: number, col: number): RefValue | CellError;
  /** Why a row is hidden: by a table filter, manually (height 0), or not at all. */
  rowHidden?(sheet: string, row: number): 'filtered' | 'manual' | false;
  /** Names bound by LET or LAMBDA parameters (upper-case keys). */
  readonly locals?: ReadonlyMap<string, EvalValue>;
}

export interface FunctionContext extends EvalContext {
  evaluate(node: Node): EvalValue;
}

/**
 * A function value created by LAMBDA. It is a #CALC! error when it ends up in
 * a cell, as in Excel, but can be passed to MAP, BYROW, REDUCE, ... or called.
 */
export class LambdaValue extends CellError {
  readonly params: readonly string[];
  readonly body: Node;
  readonly scope: EvalContext;
  constructor(params: readonly string[], body: Node, scope: EvalContext) {
    super('#CALC!', 'A LAMBDA function needs to be called');
    this.params = params;
    this.body = body;
    this.scope = scope;
  }
}

export function isLambda(v: unknown): v is LambdaValue {
  return v instanceof LambdaValue;
}

/** Returns a context that additionally binds the given names. */
export function withLocals(ctx: EvalContext, bindings: Iterable<[string, EvalValue]>): EvalContext {
  const locals = new Map(ctx.locals ?? []);
  for (const [name, value] of bindings) locals.set(name.toUpperCase(), value);
  return Object.assign(Object.create(ctx) as EvalContext, { locals });
}

/**
 * Nested LAMBDA calls before a formula counts as runaway recursion. Failing here
 * takes milliseconds; running into the engine's stack limit took about half a
 * second per evaluation and froze the editor while typing.
 */
const MAX_LAMBDA_DEPTH = 1000;
let lambdaDepth = 0;

/** Calls a LAMBDA with already evaluated arguments. */
export function callLambda(fn: LambdaValue, args: EvalValue[], ctx: EvalContext): EvalValue {
  if (args.length > fn.params.length) return Err.VALUE;
  // A RangeError is reported like a stack overflow: #CALC! "Recursion too deep".
  if (lambdaDepth >= MAX_LAMBDA_DEPTH) throw new RangeError('LAMBDA recursion too deep');
  const caller = Object.assign(Object.create(fn.scope) as EvalContext, { row: ctx.row, col: ctx.col });
  lambdaDepth++;
  try {
    return evaluate(fn.body, withLocals(caller, fn.params.map((p, i) => [p, args[i] ?? null])));
  } finally {
    lambdaDepth--;
  }
}

export function isMatrix(v: EvalValue): v is Matrix {
  return Array.isArray(v);
}

export function isRef(v: EvalValue): v is RefValue {
  return v !== null && typeof v === 'object' && !Array.isArray(v) && !(v instanceof CellError) && (v as RefValue).kind === 'ref';
}

// ---------------------------------------------------------------------------
// Coercion helpers
// ---------------------------------------------------------------------------

export function deref(v: EvalValue, ctx: EvalContext): Scalar | Matrix {
  if (!isRef(v)) return v;
  const { start, end } = v.range;
  if (start.row === end.row && start.col === end.col) return ctx.getCell(v.sheet, start.row, start.col);
  return ctx.getRange(v.sheet, v.range);
}

export function toMatrix(v: EvalValue, ctx: EvalContext): Matrix {
  if (isRef(v)) return ctx.getRange(v.sheet, v.range);
  if (isMatrix(v)) return v;
  return [[v]];
}

/** Reduces any value to a scalar using Excel's implicit intersection rules. */
export function toScalar(v: EvalValue, ctx: EvalContext): Scalar {
  if (isRef(v)) {
    const { start, end } = v.range;
    if (start.row === end.row && start.col === end.col) return ctx.getCell(v.sheet, start.row, start.col);
    if (start.col === end.col && ctx.row >= start.row && ctx.row <= end.row) return ctx.getCell(v.sheet, ctx.row, start.col);
    if (start.row === end.row && ctx.col >= start.col && ctx.col <= end.col) return ctx.getCell(v.sheet, start.row, ctx.col);
    return Err.VALUE;
  }
  if (isMatrix(v)) return v[0]?.[0] ?? null;
  return v;
}

export function toNumber(v: Scalar): number | CellError {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isError(v)) return v;
  const s = v.trim();
  if (s === '') return Err.VALUE;
  const pct = s.endsWith('%');
  const n = Number((pct ? s.slice(0, -1) : s).replace(/,/g, ''));
  if (Number.isFinite(n)) return pct ? n / 100 : n;
  return Err.VALUE;
}

export function toText(v: Scalar): string | CellError {
  if (typeof v === 'string') return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return v;
  return formatGeneral(v);
}

export function toBoolean(v: Scalar): boolean | CellError {
  if (typeof v === 'boolean') return v;
  if (v === null) return false;
  if (typeof v === 'number') return v !== 0;
  if (isError(v)) return v;
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return Err.VALUE;
}

function typeRank(v: Scalar): number {
  if (typeof v === 'number') return 0;
  if (typeof v === 'string') return 1;
  if (typeof v === 'boolean') return 2;
  return 3;
}

/** Excel-style comparison: numbers < text < booleans, text case-insensitive. */
export function compareScalars(a: Scalar, b: Scalar): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : a < b ? -1 : 1;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x === y ? 0 : x < y ? -1 : 1;
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  return 0;
}

export function checkNumber(n: number): number | CellError {
  return Number.isFinite(n) ? n : Err.NUM;
}

/** Applies a scalar function element-wise if any argument is a matrix (array broadcasting). */
export function broadcast(values: (Scalar | Matrix)[], fn: (...s: Scalar[]) => Scalar): Scalar | Matrix {
  const matrices = values.filter(isMatrix) as Matrix[];
  if (matrices.length === 0) return fn(...(values as Scalar[]));
  const rows = Math.max(...matrices.map((m) => m.length));
  const cols = Math.max(...matrices.map((m) => m[0]?.length ?? 0));
  const out: Matrix = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) {
      const args = values.map((v) => {
        if (!isMatrix(v)) return v;
        const rr = v.length === 1 ? 0 : r;
        const cc = (v[0]?.length ?? 0) === 1 ? 0 : c;
        const row = v[rr];
        return row && cc < row.length ? row[cc]! : Err.NA;
      });
      row.push(fn(...args));
    }
    out.push(row);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

function arith(op: string, a: Scalar, b: Scalar): Scalar {
  const x = toNumber(a);
  if (isError(x)) return x;
  const y = toNumber(b);
  if (isError(y)) return y;
  switch (op) {
    case '+': return checkNumber(x + y);
    case '-': return checkNumber(x - y);
    case '*': return checkNumber(x * y);
    case '/': return y === 0 ? Err.DIV0 : checkNumber(x / y);
    case '^': {
      if (x === 0 && y === 0) return Err.NUM;
      return checkNumber(Math.pow(x, y));
    }
  }
  return Err.VALUE;
}

function binaryScalar(op: string, a: Scalar, b: Scalar): Scalar {
  switch (op) {
    case '&': {
      const x = toText(a);
      if (isError(x)) return x;
      const y = toText(b);
      if (isError(y)) return y;
      return x + y;
    }
    case '=': case '<>': case '<': case '>': case '<=': case '>=': {
      if (isError(a)) return a;
      if (isError(b)) return b;
      const c = compareScalars(a, b);
      switch (op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '<': return c < 0;
        case '>': return c > 0;
        case '<=': return c <= 0;
        default: return c >= 0;
      }
    }
    default:
      return arith(op, a, b);
  }
}

export function evaluate(node: Node, ctx: EvalContext): EvalValue {
  switch (node.type) {
    case 'number':
      return node.value;
    case 'string':
      return node.value;
    case 'bool':
      return node.value;
    case 'error':
      return new CellError(node.code);
    case 'missing':
      return null;
    case 'cell': {
      const sheet = node.sheet ?? ctx.sheet;
      if (!ctx.sheetExists(sheet)) return Err.REF;
      return { kind: 'ref', sheet, range: singleCellRange(node.ref.row, node.ref.col) };
    }
    case 'range': {
      const sheet = node.sheet ?? ctx.sheet;
      if (!ctx.sheetExists(sheet)) return Err.REF;
      return { kind: 'ref', sheet, range: normalizeRange(node.start, node.end) };
    }
    case 'structured':
      return ctx.resolveStructured ? ctx.resolveStructured(node, ctx.sheet, ctx.row, ctx.col) : Err.REF;
    case 'name': {
      const local = ctx.locals?.get(node.name.toUpperCase());
      if (local !== undefined) return local;
      const target = ctx.resolveName(node.name);
      if (!target) return Err.NAME;
      return evaluate(target, ctx);
    }
    case 'array':
      return node.rows.map((r) => r.map((c) => toScalar(evaluate(c, ctx), ctx)));
    case 'unary': {
      const v = deref(evaluate(node.arg, ctx), ctx);
      return broadcast([v], (s) => {
        const n = toNumber(s);
        if (isError(n)) return n;
        return node.op === '-' ? -n : n;
      });
    }
    case 'percent': {
      const v = deref(evaluate(node.arg, ctx), ctx);
      return broadcast([v], (s) => {
        const n = toNumber(s);
        return isError(n) ? n : n / 100;
      });
    }
    case 'spill': {
      const ref = evaluate(node.arg, ctx);
      if (!isRef(ref) || ref.range.start.row !== ref.range.end.row || ref.range.start.col !== ref.range.end.col) return Err.REF;
      const range = ctx.getSpillRange?.(ref.sheet, ref.range.start.row, ref.range.start.col);
      return !range ? Err.REF : isError(range) ? range : { kind: 'ref', sheet: ref.sheet, range };
    }
    case 'binary': {
      const left = evaluate(node.left, ctx);
      const right = evaluate(node.right, ctx);
      if (node.op === ':') {
        if (isRef(left) && isRef(right) && left.sheet === right.sheet) {
          const r = normalizeRange(left.range.start, right.range.end);
          const r2 = normalizeRange(left.range.end, right.range.start);
          return {
            kind: 'ref',
            sheet: left.sheet,
            range: {
              start: { row: Math.min(r.start.row, r2.start.row), col: Math.min(r.start.col, r2.start.col) },
              end: { row: Math.max(r.end.row, r2.end.row), col: Math.max(r.end.col, r2.end.col) },
            },
          };
        }
        return Err.VALUE;
      }
      return broadcast([deref(left, ctx), deref(right, ctx)], (a, b) => binaryScalar(node.op, a, b));
    }
    case 'func': {
      const def = getFunction(node.name);
      if (!def) {
        const fn = ctx.locals?.get(node.name.toUpperCase()) ?? (() => {
          const target = ctx.resolveName(node.name);
          return target ? evaluate(target, ctx) : undefined;
        })();
        if (!isLambda(fn)) return Err.NAME;
        return callLambda(fn, node.args.map((a) => (a.type === 'missing' ? null : evaluate(a, ctx))), ctx);
      }
      const argc = node.args.length;
      if ((def.minArgs !== undefined && argc < def.minArgs) || (def.maxArgs !== undefined && argc > def.maxArgs)) {
        return Err.VALUE;
      }
      const fctx: FunctionContext = Object.assign(Object.create(ctx) as EvalContext, {
        evaluate: (n: Node) => evaluate(n, ctx),
      }) as FunctionContext;
      try {
        if (def.lazy) return def.impl(node.args, fctx);
        return def.impl(node.args.map((a) => evaluate(a, ctx)), fctx);
      } catch (e) {
        if (e instanceof CellError) return e;
        throw e;
      }
    }
  }
}

/** Evaluates a formula AST to the final value stored in a cell. */
export function evaluateCell(node: Node, ctx: EvalContext): Scalar {
  const v = evaluate(node, ctx);
  const s = toScalar(v, ctx);
  return s === null ? 0 : s;
}
