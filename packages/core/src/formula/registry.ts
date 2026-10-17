import type { Node } from './ast.js';
import type { EvalValue, FunctionContext } from './evaluator.js';

export type FunctionCategory = 'math' | 'statistical' | 'logical' | 'lookup' | 'text' | 'date' | 'info';

interface FunctionDefBase {
  name: string;
  category: FunctionCategory;
  /** Short signature for autocomplete and agents, e.g. "SUM(number1, [number2], ...)". */
  syntax: string;
  description: string;
  minArgs?: number;
  maxArgs?: number;
  /** Result changes on every recalculation (NOW, RAND, ...). */
  volatile?: boolean;
}

export interface EagerFunctionDef extends FunctionDefBase {
  lazy?: false;
  impl: (args: EvalValue[], ctx: FunctionContext) => EvalValue;
}

/** Lazy functions receive unevaluated AST nodes (IF, IFERROR, CHOOSE, ...). */
export interface LazyFunctionDef extends FunctionDefBase {
  lazy: true;
  impl: (args: Node[], ctx: FunctionContext) => EvalValue;
}

export type FunctionDef = EagerFunctionDef | LazyFunctionDef;

const registry = new Map<string, FunctionDef>();

/** Registers (or overrides) a spreadsheet function. Names are case-insensitive. */
export function registerFunction(def: FunctionDef): void {
  registry.set(def.name.toUpperCase(), def);
}

export function registerFunctions(defs: FunctionDef[]): void {
  defs.forEach(registerFunction);
}

export function getFunction(name: string): FunctionDef | undefined {
  return registry.get(name.toUpperCase());
}

export function listFunctions(): FunctionDef[] {
  return [...registry.values()].sort((a, b) => a.name.localeCompare(b.name));
}
