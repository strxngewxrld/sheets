import { CellError, Err, isError, type Scalar } from '../../values.js';
import { broadcast, compareScalars, deref, isRef, toBoolean, toScalar, type EvalValue, type FunctionContext } from '../evaluator.js';
import type { Node } from '../ast.js';
import type { FunctionDef, LazyFunctionDef } from '../registry.js';
import { def, scalarsOf } from './helpers.js';

function lazy(name: string, syntax: string, description: string, impl: LazyFunctionDef['impl'], extra: Partial<LazyFunctionDef> = {}): LazyFunctionDef {
  return { name, category: 'logical', syntax, description, lazy: true, impl, ...extra };
}

function evalArg(node: Node | undefined, ctx: FunctionContext, fallback: EvalValue): EvalValue {
  if (!node) return fallback;
  if (node.type === 'missing') return 0;
  return ctx.evaluate(node);
}

function condition(node: Node, ctx: FunctionContext): boolean | CellError {
  return toBoolean(toScalar(ctx.evaluate(node), ctx));
}

function logicalValues(args: EvalValue[], ctx: FunctionContext): boolean[] | CellError {
  const out: boolean[] = [];
  for (const a of args) {
    if (isRef(a) || Array.isArray(a)) {
      for (const v of scalarsOf([a], ctx)) {
        if (isError(v)) return v;
        if (typeof v === 'boolean') out.push(v);
        else if (typeof v === 'number') out.push(v !== 0);
      }
    } else {
      const b = toBoolean(a as Scalar);
      if (isError(b)) return b;
      out.push(b);
    }
  }
  return out.length ? out : Err.VALUE;
}

/** Element-wise predicate over a (possibly array) argument. */
function is(name: string, description: string, pred: (v: Scalar) => boolean): FunctionDef {
  return def(name, 'info', `${name}(value)`, description, ([a], ctx) => broadcast([deref(a!, ctx)], (v) => pred(v)), { minArgs: 1, maxArgs: 1 });
}

export const logicalFunctions: FunctionDef[] = [
  lazy('IF', 'IF(condition, value_if_true, [value_if_false])', 'Returns one value if a condition is true and another if false.', ([c, t, f], ctx) => {
    const cond = condition(c!, ctx);
    if (isError(cond)) return cond;
    return cond ? evalArg(t, ctx, true) : evalArg(f, ctx, false);
  }, { minArgs: 2, maxArgs: 3 }),
  lazy('IFS', 'IFS(condition1, value1, [condition2, value2], ...)', 'Returns the value of the first true condition.', (args, ctx) => {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const cond = condition(args[i]!, ctx);
      if (isError(cond)) return cond;
      if (cond) return ctx.evaluate(args[i + 1]!);
    }
    return Err.NA;
  }, { minArgs: 2 }),
  lazy('IFERROR', 'IFERROR(value, value_if_error)', 'Returns a fallback if the value is an error.', ([v, alt], ctx) => {
    const r = deref(ctx.evaluate(v!), ctx);
    if (Array.isArray(r)) {
      const fallback = toScalar(evalArg(alt, ctx, 0), ctx);
      return r.map((row) => row.map((x) => (isError(x) ? fallback : x)));
    }
    return isError(r) ? evalArg(alt, ctx, 0) : r;
  }, { minArgs: 2, maxArgs: 2 }),
  lazy('IFNA', 'IFNA(value, value_if_na)', 'Returns a fallback if the value is #N/A.', ([v, alt], ctx) => {
    const r = ctx.evaluate(v!);
    const s = toScalar(r, ctx);
    return isError(s) && s.code === '#N/A' ? evalArg(alt, ctx, 0) : r;
  }, { minArgs: 2, maxArgs: 2 }),
  lazy('SWITCH', 'SWITCH(expression, value1, result1, [default_or_value2, result2], ...)', 'Matches an expression against values.', ([expr, ...rest], ctx) => {
    const v = toScalar(ctx.evaluate(expr!), ctx);
    if (isError(v)) return v;
    for (let i = 0; i + 1 < rest.length; i += 2) {
      const candidate = toScalar(ctx.evaluate(rest[i]!), ctx);
      if (compareScalars(v, candidate) === 0 && typeof v === typeof candidate) return ctx.evaluate(rest[i + 1]!);
    }
    return rest.length % 2 === 1 ? ctx.evaluate(rest.at(-1)!) : Err.NA;
  }, { minArgs: 3 }),
  lazy('CHOOSE', 'CHOOSE(index, value1, [value2], ...)', 'Picks a value by index.', ([idx, ...rest], ctx) => {
    const i = toScalar(ctx.evaluate(idx!), ctx);
    const n = typeof i === 'number' ? Math.trunc(i) : NaN;
    if (!(n >= 1 && n <= rest.length)) return Err.VALUE;
    return ctx.evaluate(rest[n - 1]!);
  }, { minArgs: 2 }),
  def('AND', 'logical', 'AND(logical1, [logical2], ...)', 'TRUE if all arguments are TRUE.', (args, ctx) => {
    const vs = logicalValues(args, ctx);
    return isError(vs) ? vs : vs.every(Boolean);
  }, { minArgs: 1 }),
  def('OR', 'logical', 'OR(logical1, [logical2], ...)', 'TRUE if any argument is TRUE.', (args, ctx) => {
    const vs = logicalValues(args, ctx);
    return isError(vs) ? vs : vs.some(Boolean);
  }, { minArgs: 1 }),
  def('XOR', 'logical', 'XOR(logical1, [logical2], ...)', 'TRUE if an odd number of arguments are TRUE.', (args, ctx) => {
    const vs = logicalValues(args, ctx);
    return isError(vs) ? vs : vs.filter(Boolean).length % 2 === 1;
  }, { minArgs: 1 }),
  def('NOT', 'logical', 'NOT(logical)', 'Reverses a logical value.', ([a], ctx) => broadcast([deref(a!, ctx)], (v) => {
    const b = toBoolean(v);
    return isError(b) ? b : !b;
  }), { minArgs: 1, maxArgs: 1 }),
  def('TRUE', 'logical', 'TRUE()', 'The logical value TRUE.', () => true, { maxArgs: 0 }),
  def('FALSE', 'logical', 'FALSE()', 'The logical value FALSE.', () => false, { maxArgs: 0 }),

  is('ISBLANK', 'TRUE if the cell is empty.', (v) => v === null),
  is('ISNUMBER', 'TRUE if the value is a number.', (v) => typeof v === 'number'),
  is('ISTEXT', 'TRUE if the value is text.', (v) => typeof v === 'string'),
  is('ISNONTEXT', 'TRUE if the value is not text.', (v) => typeof v !== 'string'),
  is('ISLOGICAL', 'TRUE if the value is a boolean.', (v) => typeof v === 'boolean'),
  is('ISERROR', 'TRUE if the value is any error.', (v) => isError(v)),
  is('ISERR', 'TRUE if the value is an error other than #N/A.', (v) => isError(v) && v.code !== '#N/A'),
  is('ISNA', 'TRUE if the value is #N/A.', (v) => isError(v) && v.code === '#N/A'),
  is('ISEVEN', 'TRUE if the number is even.', (v) => typeof v === 'number' && Math.trunc(v) % 2 === 0),
  is('ISODD', 'TRUE if the number is odd.', (v) => typeof v === 'number' && Math.abs(Math.trunc(v) % 2) === 1),
  lazy('ISREF', 'ISREF(value)', 'TRUE if the argument is a reference.', ([a], ctx) => isRef(ctx.evaluate(a!)), { minArgs: 1, maxArgs: 1 }),
  lazy('ISFORMULA', 'ISFORMULA(reference)', 'TRUE if the referenced cell contains a formula.', ([a], ctx) => {
    const r = ctx.evaluate(a!);
    if (!isRef(r)) return Err.VALUE;
    const host = ctx as FunctionContext & { hasFormula?: (s: string, r: number, c: number) => boolean };
    return host.hasFormula?.(r.sheet, r.range.start.row, r.range.start.col) ?? false;
  }, { minArgs: 1, maxArgs: 1 }),
  def('NA', 'info', 'NA()', 'Returns the #N/A error.', () => Err.NA, { maxArgs: 0 }),
  def('N', 'info', 'N(value)', 'Converts a value to a number.', ([a], ctx) => {
    const v = toScalar(a!, ctx);
    return typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : isError(v) ? v : 0;
  }, { minArgs: 1, maxArgs: 1 }),
  def('TYPE', 'info', 'TYPE(value)', 'Type code: 1 number, 2 text, 4 boolean, 16 error, 64 array.', ([a], ctx) => {
    if (Array.isArray(a) || (isRef(a!) && (a.range.start.row !== a.range.end.row || a.range.start.col !== a.range.end.col))) return 64;
    const v = toScalar(a!, ctx);
    if (typeof v === 'number' || v === null) return 1;
    if (typeof v === 'string') return 2;
    if (typeof v === 'boolean') return 4;
    return 16;
  }, { minArgs: 1, maxArgs: 1 }),
  def('ERROR.TYPE', 'info', 'ERROR.TYPE(error_val)', 'Number identifying an error type.', ([a], ctx) => {
    const v = toScalar(a!, ctx);
    if (!(v instanceof CellError)) return Err.NA;
    const codes = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
    const i = codes.indexOf(v.code);
    return i >= 0 ? i + 1 : Err.NA;
  }, { minArgs: 1, maxArgs: 1 }),
];
