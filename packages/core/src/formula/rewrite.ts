import { colToName, MAX_COLS, MAX_ROWS, quoteSheetName, splitSheetPrefix } from '../address.js';
import type { RefPoint } from './ast.js';
import { parseFormula, parseRefPoint } from './parser.js';
import { boundNames } from './scope.js';
import { tokenize } from './tokenizer.js';

export interface RefToken {
  sheet?: string;
  kind: 'cell' | 'cells' | 'cols' | 'rows';
  start: RefPoint;
  end: RefPoint;
}

export function parseRefToken(text: string): RefToken {
  const { sheet, ref } = splitSheetPrefix(text);
  const parts = ref.split(':');
  const start = parseRefPoint(parts[0]!);
  const end = parts[1] ? parseRefPoint(parts[1]) : { ...start };
  let kind: RefToken['kind'] = parts.length === 1 ? 'cell' : 'cells';
  if (parts.length === 2 && /^\$?[A-Za-z]+$/.test(parts[0]!)) kind = 'cols';
  else if (parts.length === 2 && /^\$?\d+$/.test(parts[0]!)) kind = 'rows';
  return sheet !== undefined ? { sheet, kind, start, end } : { kind, start, end };
}

function pointText(p: RefPoint, kind: RefToken['kind']): string {
  const c = `${p.absCol ? '$' : ''}${colToName(p.col)}`;
  const r = `${p.absRow ? '$' : ''}${p.row + 1}`;
  if (kind === 'cols') return c;
  if (kind === 'rows') return r;
  return c + r;
}

export function formatRefToken(t: RefToken): string {
  const prefix = t.sheet !== undefined ? `${quoteSheetName(t.sheet)}!` : '';
  if (t.kind === 'cell') return prefix + pointText(t.start, 'cell');
  return `${prefix}${pointText(t.start, t.kind)}:${pointText(t.end, t.kind)}`;
}

/**
 * Rewrites every reference token in a formula. Returning `null` from the
 * callback replaces the reference with #REF!.
 */
export function rewriteRefs(formula: string, fn: (ref: RefToken) => RefToken | null): string {
  let out = '';
  for (const tok of tokenize(formula)) {
    if (tok.type !== 'ref') {
      out += tok.text;
      continue;
    }
    const res = fn(parseRefToken(tok.text));
    out += res ? formatRefToken(res) : '#REF!';
  }
  return out;
}

function inBounds(p: RefPoint): boolean {
  return p.row >= 0 && p.row < MAX_ROWS && p.col >= 0 && p.col < MAX_COLS;
}

/** Moves relative references by the given offset (used for copy/paste and fill). */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  return rewriteRefs(formula, (ref) => {
    const move = (p: RefPoint): RefPoint => ({
      ...p,
      row: p.absRow || ref.kind === 'cols' ? p.row : p.row + dRow,
      col: p.absCol || ref.kind === 'rows' ? p.col : p.col + dCol,
    });
    const start = move(ref.start);
    const end = move(ref.end);
    if (!inBounds(start) || !inBounds(end)) return null;
    return { ...ref, start, end };
  });
}

export interface StructuralChange {
  sheet: string;
  axis: 'row' | 'col';
  index: number;
  /** Positive = insert, negative = delete. */
  count: number;
}

/**
 * Adjusts references after rows/columns were inserted or deleted.
 * @param formulaSheet the sheet the formula lives on (for unqualified refs)
 */
export function adjustFormulaForStructure(formula: string, formulaSheet: string, change: StructuralChange): string {
  const target = change.sheet.toUpperCase();
  return rewriteRefs(formula, (ref) => {
    const refSheet = (ref.sheet ?? formulaSheet).toUpperCase();
    if (refSheet !== target) return ref;
    const key = change.axis === 'row' ? 'row' : 'col';
    if ((change.axis === 'row' && ref.kind === 'cols') || (change.axis === 'col' && ref.kind === 'rows')) return ref;
    const start = { ...ref.start };
    const end = { ...ref.end };
    const { index, count } = change;
    if (count > 0) {
      if (start[key] >= index) start[key] += count;
      if (end[key] >= index) end[key] += count;
      if (end[key] >= (key === 'row' ? MAX_ROWS : MAX_COLS)) end[key] = (key === 'row' ? MAX_ROWS : MAX_COLS) - 1;
      if (start[key] > end[key]) return null;
      return { ...ref, start, end };
    }
    const n = -count;
    const delEnd = index + n - 1;
    const isRange = ref.kind !== 'cell';
    if (!isRange) {
      if (start[key] >= index && start[key] <= delEnd) return null;
      if (start[key] > delEnd) start[key] -= n;
      return { ...ref, start, end: { ...start } };
    }
    if (start[key] >= index && end[key] <= delEnd) return null; // fully deleted
    if (start[key] > delEnd) start[key] -= n;
    else if (start[key] >= index) start[key] = index;
    if (end[key] > delEnd) end[key] -= n;
    else if (end[key] >= index) end[key] = index - 1;
    return { ...ref, start, end };
  });
}

export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const old = oldName.toUpperCase();
  return rewriteRefs(formula, (ref) => (ref.sheet?.toUpperCase() === old ? { ...ref, sheet: newName } : ref));
}

export function invalidateSheetRefs(formula: string, sheetName: string): string {
  const name = sheetName.toUpperCase();
  return rewriteRefs(formula, (ref) => (ref.sheet?.toUpperCase() === name ? null : ref));
}

/** Functions that Excel stores with the "_xlfn." prefix in files. */
const FUTURE_FUNCTIONS = new Set([
  'XLOOKUP', 'XMATCH', 'IFS', 'SWITCH', 'CONCAT', 'TEXTJOIN', 'MAXIFS', 'MINIFS', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P',
  'IFNA', 'DAYS', 'FILTER', 'SORT', 'UNIQUE', 'SEQUENCE', 'XOR', 'CEILING.MATH', 'FLOOR.MATH', 'RANK.EQ', 'MODE.SNGL',
  'PERCENTILE.INC', 'QUARTILE.INC', 'NORM.DIST', 'LET', 'ISOWEEKNUM', 'LAMBDA', 'MAP', 'BYROW', 'BYCOL', 'REDUCE', 'SCAN',
  'MAKEARRAY', 'ISOMITTED',
]);

function lambdaParams(formula: string): Set<string> {
  if (!/\b(?:LET|LAMBDA)\s*\(/i.test(formula)) return new Set();
  try { return boundNames(parseFormula(formula)); } catch { return new Set(); }
}

export function addFutureFunctionPrefixes(formula: string): string {
  const tokens = tokenize(formula);
  const params = lambdaParams(formula);
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i]!;
    let next = i + 1;
    while (tokens[next]?.type === 'ws') next++;
    if (t.type === 'ref' && tokens[next]?.text === '#') {
      out += `_xlfn.ANCHORARRAY(${t.text})`;
      i = next;
    } else if ((t.type === 'name' || t.type === 'func') && params.has(t.text.toUpperCase())) {
      out += `_xlpm.${t.text}`;
    } else if (t.type === 'func' && FUTURE_FUNCTIONS.has(t.text.toUpperCase())) {
      out += `_xlfn.${['FILTER', 'SORT'].includes(t.text.toUpperCase()) ? '_xlws.' : ''}${t.text}`;
    } else out += t.text;
  }
  return out;
}

export function stripFutureFunctionPrefixes(formula: string): string {
  return tokenize(formula)
    .map((t) => (t.type === 'func' ? t.text.replace(/^(_xlfn\.)?(_xlws\.)?/i, '').replace(/^_xlpm\./i, '') : t.type === 'name' ? t.text.replace(/^_xlpm\./i, '') : t.text))
    .join('');
}
