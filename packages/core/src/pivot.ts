import { cellKey, formatAddress, formatRange, keyToAddress, MAX_COLS, MAX_ROWS, parseAddress, parseRange, rangesIntersect } from './address.js';
import { adjustFormulaForStructure, type StructuralChange } from './formula/rewrite.js';
import { makeCriteria } from './formula/functions/helpers.js';
import type { Matrix } from './formula/evaluator.js';
import type { CellPatch } from './ops.js';
import { Err, isError, scalarFromJson, scalarToJson, type JsonScalar, type Scalar } from './values.js';
import type { ApplyOptions, Workbook } from './workbook.js';

export const PIVOT_AGGREGATES = ['sum', 'count', 'average', 'min', 'max'] as const;
export type PivotAggregate = typeof PIVOT_AGGREGATES[number];
export type PivotField = string | number;

export interface PivotTableSpec {
  id: string;
  source: { sheet: string; range: string };
  target: { sheet: string; cell: string };
  /** Header names or zero-based column offsets within the source range. */
  rows: PivotField[];
  columns?: PivotField[];
  values: { field: PivotField; aggregate: PivotAggregate; label?: string }[];
  filters?: { field: PivotField; criteria: JsonScalar }[];
  /** Include an overall total row (default true). */
  grandTotalRow?: boolean;
  /** Include totals over all column groups (default true). */
  grandTotalColumn?: boolean;
  /** Last materialized output rectangle, managed by the framework. */
  outputRange?: string;
}

interface Aggregate {
  count: number;
  numeric: number;
  sum: number;
  min: number;
  max: number;
  error?: Scalar;
}

function add(state: Aggregate, value: Scalar): void {
  if (value !== null && value !== '') state.count++;
  if (isError(value)) state.error ??= value;
  if (typeof value === 'number') {
    state.numeric++; state.sum += value;
    state.min = Math.min(state.min, value); state.max = Math.max(state.max, value);
  }
}

function result(state: Aggregate | undefined, type: PivotAggregate, total: boolean): Scalar {
  // Like Excel, a row/column combination without records stays empty; totals of an empty source are 0.
  if (!state) return total && type !== 'average' ? 0 : null;
  if (type === 'count') return state.count;
  if (state.error) return state.error;
  switch (type) {
    case 'sum': return Number.isFinite(state.sum) ? state.sum : Err.NUM;
    case 'average': return state.numeric ? (Number.isFinite(state.sum / state.numeric) ? state.sum / state.numeric : Err.NUM) : null;
    case 'min': return state.numeric ? state.min : 0;
    case 'max': return state.numeric ? state.max : 0;
  }
}

const tupleKey = (tuple: Scalar[]) => JSON.stringify(tuple.map((v) => isError(v) ? { error: v.code } : typeof v === 'string' ? v.toLowerCase() : v));
const label = (v: Scalar) => v === null || v === '' ? '(blank)' : String(v);

/** Builds an in-memory pivot report without changing the workbook. */
export function evaluatePivot(wb: Workbook, spec: PivotTableSpec): Matrix {
  const sheet = wb.requireSheet(spec.source.sheet);
  const parsed = parseRange(spec.source.range);
  if (!parsed || parsed.sheet !== undefined) throw new Error('Pivot source must be an unqualified A1 range');
  const range = parsed.range;
  const used = sheet.usedRange();
  const lastRow = Math.min(range.end.row, used?.end.row ?? range.start.row);
  const width = range.end.col - range.start.col + 1;
  if (width > 16384 || (lastRow - range.start.row + 1) * width > 1_000_000) throw new Error('Pivot source exceeds the cell limit');
  const headers = Array.from({ length: width }, (_, c) => label(sheet.getValue(range.start.row, range.start.col + c)));
  const fieldIndex = (field: PivotField): number => {
    if (typeof field === 'number' && Number.isInteger(field) && field >= 0 && field < width) return field;
    if (typeof field === 'string') {
      const matches = headers.map((h, i) => h.toLowerCase() === field.toLowerCase() ? i : -1).filter((i) => i >= 0);
      if (matches.length === 1) return matches[0]!;
    }
    throw new Error(`Pivot field is missing, ambiguous or out of bounds: ${String(field)}`);
  };
  if (!Array.isArray(spec.rows) || !Array.isArray(spec.values) || !spec.values.length) throw new Error('Pivot requires row fields and at least one value field');
  const rowFields = spec.rows.map(fieldIndex);
  const colFields = (spec.columns ?? []).map(fieldIndex);
  if (new Set([...rowFields, ...colFields]).size !== rowFields.length + colFields.length) throw new Error('Pivot grouping fields must be unique');
  const measures = spec.values.map((value) => {
    if (!PIVOT_AGGREGATES.includes(value.aggregate)) throw new Error('Invalid pivot aggregation');
    return { ...value, index: fieldIndex(value.field) };
  });
  const filters = (spec.filters ?? []).map((f) => ({ index: fieldIndex(f.field), matches: makeCriteria(scalarFromJson(f.criteria)) }));
  const rowGroups = new Map<string, Scalar[]>();
  const colGroups = new Map<string, Scalar[]>();
  const states = new Map<string, Aggregate[]>();
  const accumulate = (r: string | null, c: string | null, values: Scalar[]) => {
    const key = JSON.stringify([r, c]);
    let group = states.get(key);
    if (!group) {
      group = measures.map(() => ({ count: 0, numeric: 0, sum: 0, min: Infinity, max: -Infinity }));
      states.set(key, group);
    }
    measures.forEach((m, i) => add(group![i]!, values[m.index] ?? null));
  };
  for (let r = range.start.row + 1; r <= lastRow; r++) {
    const values = Array.from({ length: width }, (_, c) => sheet.getValue(r, range.start.col + c));
    if (values.every((v) => v === null) || filters.some((f) => !f.matches(values[f.index] ?? null))) continue;
    const row = rowFields.map((c) => values[c] ?? null);
    const col = colFields.map((c) => values[c] ?? null);
    const rk = tupleKey(row), ck = tupleKey(col);
    rowGroups.set(rk, rowGroups.get(rk) ?? row);
    colGroups.set(ck, colGroups.get(ck) ?? col);
    accumulate(rk, ck, values);
    accumulate(rk, null, values);
    accumulate(null, ck, values);
    accumulate(null, null, values);
  }
  if (!colGroups.size && !colFields.length) colGroups.set('[]', []);
  const columnGroups: { key: string | null; prefix: string }[] = [...colGroups].map(([key, tuple]) => ({ key, prefix: tuple.map(label).join(' / ') }));
  if (colFields.length && spec.grandTotalColumn !== false) columnGroups.push({ key: null, prefix: 'Grand Total' });
  const outputWidth = rowFields.length + columnGroups.length * measures.length;
  const outputHeight = 1 + rowGroups.size + (spec.grandTotalRow !== false && rowFields.length ? 1 : 0);
  if (outputWidth > MAX_COLS || outputHeight > MAX_ROWS || outputHeight * outputWidth > 1_000_000) throw new Error('Pivot output exceeds the cell limit');
  const header: Scalar[] = [...rowFields.map((c) => headers[c]!), ...columnGroups.flatMap((g) => measures.map((m) => [g.prefix, m.label ?? `${m.aggregate.toUpperCase()} ${headers[m.index]}`].filter(Boolean).join(' / ')))];
  const output: Matrix = [header];
  const rowFor = (key: string | null, labels: Scalar[]) => [...labels, ...columnGroups.flatMap((g) => measures.map((m, i) => result(states.get(JSON.stringify([key, g.key]))?.[i], m.aggregate, key === null || g.key === null)))];
  for (const [key, tuple] of rowGroups) output.push(rowFor(key, tuple.map((v) => v ?? '(blank)')));
  if (spec.grandTotalRow !== false && rowFields.length) output.push(rowFor(null, rowFields.map((_, i) => i === 0 ? 'Grand Total' : null)));
  return output;
}

/** Keeps field offsets attached to their original source columns during edits. */
export function adjustPivotForStructure(wb: Workbook, original: PivotTableSpec, change: StructuralChange): PivotTableSpec | null {
  const pivot = structuredClone(original);
  if (pivot.source.sheet.toUpperCase() === change.sheet.toUpperCase()) {
    const source = parseRange(pivot.source.range)!.range;
    const ref = adjustFormulaForStructure(pivot.source.range, pivot.source.sheet, change);
    if (ref === '#REF!') return null;
    const nextRange = parseRange(ref)!.range;
    if (change.axis === 'row' && change.count < 0 && source.start.row >= change.index && source.start.row < change.index - change.count) return null;
    if (change.axis === 'col') {
      let deleted = false;
      const sheet = wb.requireSheet(pivot.source.sheet);
      const field = (f: PivotField): PivotField => {
        let col = typeof f === 'number' ? source.start.col + f : undefined;
        if (typeof f === 'string') {
          for (let c = source.start.col; c <= source.end.col; c++) {
            if (String(sheet.getValue(source.start.row, c)).toLowerCase() === f.toLowerCase()) { col = c; break; }
          }
        }
        if (col === undefined) return f;
        if (change.count < 0 && col >= change.index && col < change.index - change.count) deleted = true;
        const shifted = col < change.index ? col : col + change.count;
        return typeof f === 'number' ? shifted - nextRange.start.col : f;
      };
      pivot.rows = pivot.rows.map(field);
      pivot.columns = pivot.columns?.map(field);
      pivot.values = pivot.values.map((v) => ({ ...v, field: field(v.field) }));
      pivot.filters = pivot.filters?.map((v) => ({ ...v, field: field(v.field) }));
      if (deleted) return null;
    }
    pivot.source.range = ref;
  }
  if (pivot.target.sheet.toUpperCase() === change.sheet.toUpperCase()) {
    const cell = adjustFormulaForStructure(pivot.target.cell, pivot.target.sheet, change);
    if (cell === '#REF!') return null;
    pivot.target.cell = cell;
    if (pivot.outputRange) {
      const ref = adjustFormulaForStructure(pivot.outputRange, pivot.target.sheet, change);
      if (ref === '#REF!') return null;
      pivot.outputRange = ref;
    }
  }
  return pivot;
}

/** Creates or refreshes a pivot as one undoable change. Existing unrelated cells are protected. */
export function createPivotTable(wb: Workbook, spec: PivotTableSpec, options: ApplyOptions = {}): PivotTableSpec {
  if (!spec.id || typeof spec.id !== 'string') throw new Error('Pivot requires a nonempty ID');
  const existing = wb.pivotTables.find((p) => p.id === spec.id);
  if (existing && (existing.target.sheet !== spec.target.sheet || existing.target.cell !== spec.target.cell)) throw new Error('Remove a pivot before moving its target');
  const target = wb.requireSheet(spec.target.sheet);
  const anchor = parseAddress(spec.target.cell);
  if (!anchor) throw new Error('Pivot target must be a single A1 cell');
  const matrix = evaluatePivot(wb, spec);
  const range = { start: anchor, end: { row: anchor.row + matrix.length - 1, col: anchor.col + matrix[0]!.length - 1 } };
  if (range.end.row >= MAX_ROWS || range.end.col >= MAX_COLS) throw new Error('Pivot output exceeds sheet bounds');
  if (target === wb.requireSheet(spec.source.sheet) && rangesIntersect(range, parseRange(spec.source.range)!.range)) throw new Error('Pivot output cannot overlap its source');
  const oldRange = existing?.outputRange ? parseRange(existing.outputRange)?.range : undefined;
  for (const other of wb.pivotTables) {
    if (other.id !== spec.id && wb.getSheet(other.target.sheet) === target && other.outputRange && rangesIntersect(range, parseRange(other.outputRange)!.range)) throw new Error('Pivot output overlaps another pivot');
  }
  const cells = new Map<number, CellPatch>();
  if (oldRange) {
    for (const [key, cell] of target.cells) {
      const p = keyToAddress(key);
      if (p.row >= oldRange.start.row && p.row <= oldRange.end.row && p.col >= oldRange.start.col && p.col <= oldRange.end.col) cells.set(key, { ...p, data: cell.style ? { s: cell.style } : null });
    }
  }
  matrix.forEach((line, r) => line.forEach((v, c) => {
    const row = anchor.row + r, col = anchor.col + c;
    const cell = target.getCell(row, col);
    const owned = oldRange && row >= oldRange.start.row && row <= oldRange.end.row && col >= oldRange.start.col && col <= oldRange.end.col;
    if (target.mergeAt(row, col) || cell?.spillParent !== undefined || (!owned && cell && (cell.formula !== undefined || cell.value !== null))) throw new Error(`Pivot output is blocked at ${formatAddress(row, col)}`);
    cells.set(cellKey(row, col), { row, col, data: { v: scalarToJson(v), ...(cell?.style ? { s: cell.style } : {}) } });
  }));
  const updated = structuredClone({ ...spec, source: { ...spec.source, sheet: wb.requireSheet(spec.source.sheet).name }, target: { ...spec.target, sheet: target.name }, outputRange: formatRange(range) });
  // Widen output columns that are too narrow for their headers or values (never narrower, at most 320 px).
  const sizes = matrix[0]!.map((_, c) => {
    const chars = Math.max(...matrix.map((line) => { const v = line[c]; return typeof v === 'number' ? String(Math.round(v * 100) / 100).length : v === null ? 0 : String(isError(v) ? v.code : v).length; }));
    const col = anchor.col + c, needed = Math.min(320, chars * 7 + 12);
    return needed > target.colWidth(col) && target.colWidth(col) > 0 ? { index: col, size: needed } : null;
  }).filter((size): size is { index: number; size: number } => size !== null);
  wb.apply([
    ...(sizes.length ? [{ type: 'setColWidths' as const, sheet: target.name, sizes }] : []),
    { type: 'setCells', sheet: target.name, cells: [...cells.values()] },
    { type: 'setPivotTables', tables: [...wb.pivotTables.filter((p) => p.id !== spec.id), updated] },
  ], options);
  return updated;
}

export function refreshPivotTable(wb: Workbook, id: string, options: ApplyOptions = {}): PivotTableSpec {
  const spec = wb.pivotTables.find((p) => p.id === id);
  if (!spec) throw new Error(`Pivot not found: ${id}`);
  return createPivotTable(wb, spec, options);
}

export function removePivotTable(wb: Workbook, id: string, options: ApplyOptions = {}): void {
  const spec = wb.pivotTables.find((p) => p.id === id);
  if (!spec) throw new Error(`Pivot not found: ${id}`);
  const target = wb.requireSheet(spec.target.sheet);
  const range = spec.outputRange ? parseRange(spec.outputRange)?.range : undefined;
  const cells: CellPatch[] = [];
  if (range) for (const [key, cell] of target.cells) {
    const p = keyToAddress(key);
    if (p.row >= range.start.row && p.row <= range.end.row && p.col >= range.start.col && p.col <= range.end.col) cells.push({ ...p, data: cell.style ? { s: cell.style } : null });
  }
  wb.apply([{ type: 'setCells', sheet: target.name, cells }, { type: 'setPivotTables', tables: wb.pivotTables.filter((p) => p.id !== id) }], options);
}
