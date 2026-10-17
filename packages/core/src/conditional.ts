import { keyToAddress, parseRange, type RangeAddress } from './address.js';
import { compareScalars, toBoolean } from './formula/evaluator.js';
import { shiftFormula } from './formula/rewrite.js';
import { parseFormula } from './formula/parser.js';
import type { Sheet } from './sheet.js';
import type { CellStyle } from './style.js';
import { isError } from './values.js';
import type { ApplyOptions, Workbook } from './workbook.js';
import { tableCellStyle } from './table.js';

export const CONDITIONAL_OPERATORS = ['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual', 'between', 'notBetween'] as const;
export type ConditionalOperator = typeof CONDITIONAL_OPERATORS[number];

interface RuleBase {
  id: string;
  range: string;
  stopIfTrue?: boolean;
}

export type ConditionalFormatRule = RuleBase & (
  | { type: 'cellIs'; operator: ConditionalOperator; formulas: string[]; style: CellStyle }
  | { type: 'expression'; formulas: string[]; style: CellStyle }
  | { type: 'colorScale'; colors: string[] }
);

function contains(range: RangeAddress, row: number, col: number): boolean {
  return row >= range.start.row && row <= range.end.row && col >= range.start.col && col <= range.end.col;
}

export function validateConditionalFormats(rules: ConditionalFormatRule[]): void {
  const ids = new Set<string>();
  for (const rule of rules) {
    const parsed = parseRange(rule.range);
    if (!rule.id || ids.has(rule.id)) throw new Error('Conditional format IDs must be nonempty and unique');
    ids.add(rule.id);
    if (!parsed || parsed.sheet !== undefined) throw new Error('Conditional format range must be an unqualified A1 range');
    if (rule.type === 'colorScale') {
      if (![2, 3].includes(rule.colors?.length) || rule.colors.some((c) => !/^#[\da-f]{6}$/i.test(c))) throw new Error('A color scale needs two or three #RRGGBB colors');
    } else if (rule.type === 'cellIs' || rule.type === 'expression') {
      const required = rule.type === 'cellIs' && ['between', 'notBetween'].includes(rule.operator) ? 2 : 1;
      if (rule.type === 'cellIs' && !CONDITIONAL_OPERATORS.includes(rule.operator)) throw new Error('Invalid conditional comparison operator');
      if (!Array.isArray(rule.formulas) || rule.formulas.length !== required || rule.formulas.some((f) => typeof f !== 'string' || !f.trim())) throw new Error(`Conditional rule requires ${required} formula(s)`);
      rule.formulas.forEach(parseFormula);
      if (!rule.style || typeof rule.style !== 'object') throw new Error('Conditional rule requires a style');
    } else throw new Error('Unsupported conditional format type');
  }
}

/** Earlier rules have higher priority. A resolver can be reused for one render/read. */
export class ConditionalFormatEvaluator {
  private scales = new Map<ConditionalFormatRule, { min: number; max: number }>();
  constructor(private wb: Workbook) {}

  style(sheet: Sheet, row: number, col: number): CellStyle | undefined {
    const direct = sheet.getCell(row, col)?.style;
    const inherited = tableCellStyle(sheet, row, col);
    const base = inherited ? { ...inherited, ...direct } : direct;
    if (!sheet.conditionalFormats.length) return base;
    let applied: CellStyle = {};
    const value = sheet.getValue(row, col);
    for (const rule of sheet.conditionalFormats) {
      const range = parseRange(rule.range)!.range;
      if (!contains(range, row, col)) continue;
      let patch: CellStyle | undefined;
      if (rule.type === 'colorScale') {
        if (typeof value !== 'number') continue;
        let scale = this.scales.get(rule);
        if (!scale) {
          let min = Infinity, max = -Infinity;
          for (const [key, cell] of sheet.cells) {
            const p = keyToAddress(key);
            if (typeof cell.value === 'number' && contains(range, p.row, p.col)) {
              min = Math.min(min, cell.value); max = Math.max(max, cell.value);
            }
          }
          scale = { min, max };
          this.scales.set(rule, scale);
        }
        const fraction = scale.min === scale.max ? 0 : Math.max(0, Math.min(1, (value - scale.min) / (scale.max - scale.min)));
        const segment = Math.min(rule.colors.length - 2, Math.floor(fraction * (rule.colors.length - 1)));
        const part = fraction * (rule.colors.length - 1) - segment;
        const from = rule.colors[segment]!, to = rule.colors[segment + 1]!;
        const rgb = [1, 3, 5].map((i) => Math.round(parseInt(from.slice(i, i + 2), 16) * (1 - part) + parseInt(to.slice(i, i + 2), 16) * part).toString(16).padStart(2, '0')).join('');
        patch = { fill: `#${rgb.toUpperCase()}` };
      } else {
        const args = rule.formulas.map((f) => this.wb.calc.evaluateFormula(shiftFormula(f, row - range.start.row, col - range.start.col), sheet, row, col));
        if (args.some(isError)) continue;
        let matches = false;
        if (rule.type === 'expression') matches = toBoolean(args[0]!) === true;
        else if (!isError(value)) {
          const comparison = compareScalars(value, args[0]!);
          switch (rule.operator) {
            case 'equal': matches = comparison === 0; break;
            case 'notEqual': matches = comparison !== 0; break;
            case 'greaterThan': matches = comparison > 0; break;
            case 'lessThan': matches = comparison < 0; break;
            case 'greaterThanOrEqual': matches = comparison >= 0; break;
            case 'lessThanOrEqual': matches = comparison <= 0; break;
            case 'between': matches = comparison >= 0 && compareScalars(value, args[1]!) <= 0; break;
            case 'notBetween': matches = comparison < 0 || compareScalars(value, args[1]!) > 0; break;
          }
        }
        if (matches) patch = rule.style;
      }
      if (patch) {
        const border = patch.border || applied.border ? { ...patch.border, ...applied.border } : undefined;
        applied = { ...patch, ...applied, ...(border ? { border } : {}) };
        if (rule.stopIfTrue) break;
      }
    }
    return Object.keys(applied).length ? { ...base, ...applied, ...(base?.border || applied.border ? { border: { ...base?.border, ...applied.border } } : {}) } : base;
  }
}

export function setConditionalFormats(wb: Workbook, sheet: string | number, rules: ConditionalFormatRule[], options: ApplyOptions = {}): void {
  wb.apply({ type: 'setConditionalFormats', sheet: wb.requireSheet(sheet).name, rules }, options);
}
