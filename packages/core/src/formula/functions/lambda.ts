import { CellError, Err, type Scalar } from '../../values.js';
import type { Node } from '../ast.js';
import {
  callLambda,
  evaluate,
  isLambda,
  isMatrix,
  LambdaValue,
  toMatrix,
  toScalar,
  withLocals,
  type EvalContext,
  type EvalValue,
  type FunctionContext,
  type Matrix,
} from '../evaluator.js';
import type { FunctionDef, LazyFunctionDef } from '../registry.js';
import { MAX_ARRAY_CELLS } from './array.js';
import { bail, def, numArg } from './helpers.js';

const NAME_RE = /^(?:_xlpm\.)?[A-Za-z_\\À-￿][\w.À-￿]*$/u;

function paramName(node: Node): string {
  if (node.type !== 'name' || !NAME_RE.test(node.name)) bail(Err.VALUE);
  return node.name.replace(/^_xlpm\./i, '');
}

function lambdaArg(v: EvalValue | undefined, arity: number): LambdaValue {
  if (!isLambda(v)) bail(Err.VALUE);
  if (v.params.length < arity) bail(Err.VALUE);
  return v;
}

/** Calls a lambda and reduces the result to one cell value. */
function callScalar(fn: LambdaValue, args: EvalValue[], ctx: EvalContext): Scalar {
  const result = callLambda(fn, args, ctx);
  if (isMatrix(result) && (result.length !== 1 || result[0]!.length !== 1)) return new CellError('#CALC!', 'Nested arrays are not supported');
  return toScalar(result, ctx);
}

function checkSize(rows: number, cols: number): void {
  if (rows < 1 || cols < 1) bail(Err.VALUE);
  if (rows * cols > MAX_ARRAY_CELLS) bail(new CellError('#CALC!', 'Array exceeds the cell limit'));
}

function lazy(name: string, syntax: string, description: string, impl: LazyFunctionDef['impl'], extra: Partial<LazyFunctionDef> = {}): LazyFunctionDef {
  return { name, category: 'logical', syntax, description, lazy: true, impl, ...extra };
}

export const lambdaFunctions: FunctionDef[] = [
  lazy('LET', 'LET(name1, value1, [name2, value2, ...], calculation)', 'Assigns names to intermediate results and returns the final calculation.', (args, ctx: FunctionContext) => {
    if (args.length < 3 || args.length % 2 === 0) return Err.VALUE;
    let scope: EvalContext = ctx;
    for (let i = 0; i < args.length - 1; i += 2) {
      const name = paramName(args[i]!);
      scope = withLocals(scope, [[name, evaluate(args[i + 1]!, scope)]]);
    }
    return evaluate(args.at(-1)!, scope);
  }, { minArgs: 3 }),
  lazy('LAMBDA', 'LAMBDA([parameter1, ...], calculation)', 'Creates a reusable function, e.g. for MAP, REDUCE or a defined name.', (args, ctx) => {
    const params = args.slice(0, -1).map(paramName);
    if (new Set(params.map((p) => p.toUpperCase())).size !== params.length) return Err.VALUE;
    return new LambdaValue(params, args.at(-1)!, ctx);
  }, { minArgs: 1, maxArgs: 254 }),
  def('MAP', 'logical', 'MAP(array1, [array2, ...], lambda)', 'Applies a LAMBDA to each value of one or more arrays.', (args, ctx) => {
    const fn = lambdaArg(args.at(-1), args.length - 1);
    const arrays = args.slice(0, -1).map((a) => toMatrix(a, ctx));
    if (!arrays.length) return Err.VALUE;
    const rows = Math.max(...arrays.map((m) => m.length));
    const cols = Math.max(...arrays.map((m) => m[0]?.length ?? 0));
    checkSize(rows, cols);
    return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => {
      const values = arrays.map((m) => {
        const rr = m.length === 1 ? 0 : r;
        const cc = (m[0]?.length ?? 0) === 1 ? 0 : c;
        return m[rr]?.[cc] ?? Err.NA;
      });
      return callScalar(fn, values, ctx);
    }));
  }, { minArgs: 2 }),
  def('BYROW', 'logical', 'BYROW(array, lambda)', 'Applies a LAMBDA to each row and returns one value per row.', ([a, l], ctx) => {
    const fn = lambdaArg(l, 1);
    const m = toMatrix(a!, ctx);
    checkSize(m.length, 1);
    return m.map((row) => [callScalar(fn, [[row]], ctx)]);
  }, { minArgs: 2, maxArgs: 2 }),
  def('BYCOL', 'logical', 'BYCOL(array, lambda)', 'Applies a LAMBDA to each column and returns one value per column.', ([a, l], ctx) => {
    const fn = lambdaArg(l, 1);
    const m = toMatrix(a!, ctx);
    const cols = m[0]?.length ?? 0;
    checkSize(1, cols);
    return [Array.from({ length: cols }, (_, c) => callScalar(fn, [m.map((row) => [row[c] ?? null])], ctx))];
  }, { minArgs: 2, maxArgs: 2 }),
  def('REDUCE', 'logical', 'REDUCE([initial_value], array, lambda)', 'Reduces an array to one value by applying LAMBDA(accumulator, value).', ([init, a, l], ctx) => {
    const fn = lambdaArg(l, 2);
    let acc: EvalValue = init ?? null;
    for (const row of toMatrix(a!, ctx)) for (const v of row) acc = callLambda(fn, [acc, v], ctx);
    return acc;
  }, { minArgs: 3, maxArgs: 3 }),
  def('SCAN', 'logical', 'SCAN([initial_value], array, lambda)', 'Like REDUCE, but returns every intermediate accumulator value.', ([init, a, l], ctx) => {
    const fn = lambdaArg(l, 2);
    let acc: EvalValue = init ?? null;
    return toMatrix(a!, ctx).map((row) => row.map((v) => {
      const next = callScalar(fn, [acc, v], ctx);
      acc = next;
      return next;
    }));
  }, { minArgs: 3, maxArgs: 3 }),
  def('MAKEARRAY', 'logical', 'MAKEARRAY(rows, columns, lambda)', 'Builds an array by calling LAMBDA(row, column) for every position.', ([r, c, l], ctx) => {
    const rows = Math.trunc(numArg(r, ctx));
    const cols = Math.trunc(numArg(c, ctx));
    const fn = lambdaArg(l, 2);
    checkSize(rows, cols);
    const out: Matrix = [];
    for (let i = 0; i < rows; i++) out.push(Array.from({ length: cols }, (_, j) => callScalar(fn, [i + 1, j + 1], ctx)));
    return out;
  }, { minArgs: 3, maxArgs: 3 }),
  def('ISOMITTED', 'info', 'ISOMITTED(argument)', 'Returns TRUE if a LAMBDA argument was not supplied.', ([a]) => a === null || a === undefined, { minArgs: 1, maxArgs: 1 }),
];

