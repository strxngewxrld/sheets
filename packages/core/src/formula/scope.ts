import { INVOKE, type Node } from './ast.js';
import { walk } from './parser.js';
import { getFunction } from './registry.js';

/** Upper-case names introduced by LET or as LAMBDA parameters anywhere in the formula. */
export function boundNames(node: Node): Set<string> {
  const names = new Set<string>();
  walk(node, (n) => {
    if (n.type !== 'func') return;
    if (n.name === 'LET') n.args.forEach((a, i) => { if (i % 2 === 0 && i < n.args.length - 1 && a.type === 'name') names.add(a.name.toUpperCase()); });
    else if (n.name === 'LAMBDA') n.args.slice(0, -1).forEach((a) => { if (a.type === 'name') names.add(a.name.toUpperCase()); });
  });
  return names;
}

/**
 * True when the formula calls a function that is neither built in nor bound by
 * LET/LAMBDA. `isDefinedName` lets callers accept LAMBDAs stored as workbook names.
 */
export function hasUnknownFunctions(node: Node, isDefinedName: (name: string) => boolean = () => false): boolean {
  const bound = boundNames(node);
  let unknown = false;
  walk(node, (n) => {
    if (n.type === 'func' && n.name !== INVOKE && !getFunction(n.name) && !bound.has(n.name.toUpperCase()) && !isDefinedName(n.name)) unknown = true;
  });
  return unknown;
}
