import { parseRange, type RangeAddress } from './address.js';
import { toMatrix, type EvalValue } from './formula/evaluator.js';
import { parseFormula } from './formula/parser.js';
import { shiftFormula } from './formula/rewrite.js';
import type { Sheet } from './sheet.js';
import { isError, type Scalar } from './values.js';
import type { ApplyOptions, Workbook } from './workbook.js';

export const VALIDATION_TYPES = ['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom'] as const;
export type ValidationType = typeof VALIDATION_TYPES[number];

export const VALIDATION_OPERATORS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const;
export type ValidationOperator = typeof VALIDATION_OPERATORS[number];

/** A data validation rule (Excel "Data Validation"). */
export interface DataValidation {
  /** One or more unqualified A1 ranges separated by spaces, e.g. "A2:A100 C2". */
  range: string;
  type: ValidationType;
  /** Comparison for number, date, time and text length rules (default "between"). */
  operator?: ValidationOperator;
  /**
   * Formula without "=", relative to the top-left cell of the first range.
   * Lists use a quoted literal ("\"Yes,No\""), a range ("$D$1:$D$5") or a name.
   */
  formula1?: string;
  formula2?: string;
  /** Empty cells are valid (default true). */
  allowBlank?: boolean;
  /** Show the in-cell dropdown for lists (default true). */
  dropdown?: boolean;
  /** How invalid input is handled (default "stop"). */
  errorStyle?: 'stop' | 'warning' | 'information';
  /** Show an error alert for invalid input (default true). */
  showError?: boolean;
  errorTitle?: string;
  error?: string;
  /** Show the input message while the cell is selected (default true). */
  showPrompt?: boolean;
  promptTitle?: string;
  prompt?: string;
}

export function validationRanges(rule: DataValidation): RangeAddress[] {
  return rule.range.split(/\s+/).filter(Boolean).map((part) => parseRange(part)?.range).filter((r): r is RangeAddress => !!r);
}

function contains(r: RangeAddress, row: number, col: number): boolean {
  return row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;
}

/** The rule applying to a cell (the last matching rule wins, as in Excel). */
export function validationAt(sheet: Sheet, row: number, col: number): DataValidation | undefined {
  for (let i = sheet.validations.length - 1; i >= 0; i--) {
    const rule = sheet.validations[i]!;
    if (validationRanges(rule).some((r) => contains(r, row, col))) return rule;
  }
  return undefined;
}

export function validateValidations(rules: DataValidation[]): void {
  for (const rule of rules) {
    const parts = rule.range.split(/\s+/).filter(Boolean);
    if (!parts.length || parts.some((p) => { const r = parseRange(p); return !r || r.sheet !== undefined; })) throw new Error('Data validation ranges must be unqualified A1 ranges');
    if (!VALIDATION_TYPES.includes(rule.type)) throw new Error(`Unsupported data validation type "${rule.type}"`);
    if (rule.operator !== undefined && !VALIDATION_OPERATORS.includes(rule.operator)) throw new Error(`Invalid data validation operator "${rule.operator}"`);
    if (!rule.formula1?.trim()) throw new Error('Data validation requires a value, source or formula');
    const needsTwo = rule.type !== 'list' && rule.type !== 'custom' && ['between', 'notBetween'].includes(rule.operator ?? 'between');
    if (needsTwo && !rule.formula2?.trim()) throw new Error('This comparison needs a minimum and a maximum');
    for (const f of [rule.formula1, rule.formula2]) if (f !== undefined && f.trim()) parseFormula(f);
  }
}

/** Replaces all data validation rules of a sheet (undoable). */
export function setDataValidations(wb: Workbook, sheet: string | number, rules: DataValidation[], options: ApplyOptions = {}): void {
  wb.apply({ type: 'setValidations', sheet: wb.requireSheet(sheet).name, rules }, options);
}

/** Formula of a rule as seen from a cell (relative references are moved). */
function formulaAt(rule: DataValidation, formula: string, row: number, col: number): string {
  const origin = validationRanges(rule)[0]!.start;
  return shiftFormula(formula, row - origin.row, col - origin.col);
}

function evaluateAt(wb: Workbook, sheet: Sheet, rule: DataValidation, formula: string, row: number, col: number): EvalValue {
  return wb.calc.evaluateRaw(formulaAt(rule, formula, row, col), sheet, row, col);
}

/** Values offered by a list rule for a cell. */
export function validationListValues(wb: Workbook, sheet: string | number | Sheet, rule: DataValidation, row: number, col: number): Scalar[] {
  if (rule.type !== 'list' || !rule.formula1) return [];
  const f = rule.formula1.trim();
  if (/^".*"$/s.test(f)) return f.slice(1, -1).replace(/""/g, '"').split(',').map((v) => v.trim()).filter((v) => v !== '');
  const s = typeof sheet === 'object' ? sheet : wb.requireSheet(sheet);
  const value = evaluateAt(wb, s, rule, f, row, col);
  if (isError(value)) return [];
  const ctx = wb.calc.contextFor(s, row, col);
  return toMatrix(value, ctx).flat().filter((v) => v !== null && v !== '' && !isError(v));
}

function compare(op: ValidationOperator, x: number, a: number, b: number): boolean {
  switch (op) {
    case 'between': return x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween': return x < Math.min(a, b) || x > Math.max(a, b);
    case 'equal': return x === a;
    case 'notEqual': return x !== a;
    case 'greaterThan': return x > a;
    case 'lessThan': return x < a;
    case 'greaterThanOrEqual': return x >= a;
    default: return x <= a;
  }
}

function numberOf(v: EvalValue, wb: Workbook, sheet: Sheet, row: number, col: number): number | null {
  const m = toMatrix(v, wb.calc.contextFor(sheet, row, col));
  const s = m[0]?.[0] ?? null;
  return typeof s === 'number' ? s : typeof s === 'string' && s.trim() !== '' && Number.isFinite(Number(s)) ? Number(s) : null;
}

/**
 * Checks a value against a rule as if it were entered in the cell. For custom
 * rules the formula is evaluated with the current workbook state, so callers
 * check after writing the value (see `isCellValid`).
 */
export function isValidValue(wb: Workbook, sheet: string | number | Sheet, rule: DataValidation, row: number, col: number, value: Scalar): boolean {
  const s = typeof sheet === 'object' ? sheet : wb.requireSheet(sheet);
  if (value === null || value === '') return rule.allowBlank !== false;
  if (isError(value)) return false;
  switch (rule.type) {
    case 'list': {
      const options = validationListValues(wb, s, rule, row, col);
      const text = String(value).toLowerCase();
      return options.some((o) => (typeof o === 'number' && typeof value === 'number' ? o === value : String(o).toLowerCase() === text));
    }
    case 'custom': {
      const result = evaluateAt(wb, s, rule, rule.formula1 ?? 'FALSE', row, col);
      const m = toMatrix(result, wb.calc.contextFor(s, row, col));
      const v = m[0]?.[0] ?? null;
      return v === true || (typeof v === 'number' && v !== 0);
    }
    default: {
      let x: number;
      if (rule.type === 'textLength') x = String(value).length;
      else if (typeof value === 'number') x = value;
      else return false;
      if (rule.type === 'whole' && !Number.isInteger(x)) return false;
      if (rule.type === 'time' && (x < 0 || x >= 1)) return false;
      const a = numberOf(evaluateAt(wb, s, rule, rule.formula1 ?? '0', row, col), wb, s, row, col);
      const b = rule.formula2 ? numberOf(evaluateAt(wb, s, rule, rule.formula2, row, col), wb, s, row, col) : a;
      if (a === null || b === null) return false;
      return compare(rule.operator ?? 'between', x, a, b);
    }
  }
}

/** Whether a cell's current value satisfies its validation rule (true without a rule). */
export function isCellValid(wb: Workbook, sheet: string | number, row: number, col: number): boolean {
  const s = wb.requireSheet(sheet);
  const rule = validationAt(s, row, col);
  return !rule || isValidValue(wb, s, rule, row, col, s.getValue(row, col));
}

/** Cells in the sheet's validation ranges that currently hold invalid values. */
export function invalidCells(wb: Workbook, sheet: string | number, limit = 10_000): { row: number; col: number }[] {
  const s = wb.requireSheet(sheet);
  const out: { row: number; col: number }[] = [];
  const used = s.usedRange();
  if (!used) return out;
  const seen = new Set<string>();
  for (const rule of s.validations) {
    for (const r of validationRanges(rule)) {
      for (let row = r.start.row; row <= Math.min(r.end.row, used.end.row); row++) {
        for (let col = r.start.col; col <= Math.min(r.end.col, used.end.col); col++) {
          const key = `${row}:${col}`;
          if (seen.has(key) || s.getValue(row, col) === null) continue;
          seen.add(key);
          if (validationAt(s, row, col) === rule && !isValidValue(wb, s, rule, row, col, s.getValue(row, col))) out.push({ row, col });
          if (out.length >= limit) return out;
        }
      }
    }
  }
  return out;
}
