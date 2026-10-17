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

export const ICON_SETS = [
  '3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '3Stars', '3Triangles',
  '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '5Boxes',
] as const;
export type IconSetName = typeof ICON_SETS[number];

/** A data bar or icon threshold (Excel "cfvo"). */
export interface Threshold {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile';
  value?: number;
}

export type ConditionalFormatRule = RuleBase & (
  | { type: 'cellIs'; operator: ConditionalOperator; formulas: string[]; style: CellStyle }
  | { type: 'expression'; formulas: string[]; style: CellStyle }
  | { type: 'colorScale'; colors: string[] }
  | { type: 'dataBar'; color: string; min?: Threshold; max?: Threshold; showValue?: boolean }
  /** `thresholds` holds one entry per icon, lowest first (the first is the lower bound). */
  | { type: 'iconSet'; iconSet: IconSetName; thresholds?: Threshold[]; reverse?: boolean; showValue?: boolean }
);

/** Rules that compare cell values with formulas (and carry a style). */
export function hasFormulas(rule: ConditionalFormatRule): rule is Extract<ConditionalFormatRule, { formulas: string[] }> {
  return rule.type === 'cellIs' || rule.type === 'expression';
}

/** Number of icons in a set ("3Arrows" → 3). */
export function iconCount(set: IconSetName): number {
  return Number(set[0]);
}

/** Default icon thresholds: equal percent steps. */
export function defaultIconThresholds(set: IconSetName): Threshold[] {
  const n = iconCount(set);
  return Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((i * 100) / n) }));
}

/** Data bar or icon to draw in a cell. */
export interface CellVisual {
  bar?: { fraction: number; color: string };
  icon?: { set: IconSetName; index: number };
  /** The rule hides the cell value (showValue = false). */
  hideValue?: boolean;
}

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
    if (rule.type === 'dataBar') {
      if (!/^#[\da-f]{6}$/i.test(rule.color ?? '')) throw new Error('A data bar needs a #RRGGBB color');
      for (const t of [rule.min, rule.max]) validateThreshold(t);
    } else if (rule.type === 'iconSet') {
      if (!ICON_SETS.includes(rule.iconSet)) throw new Error(`Unknown icon set "${rule.iconSet}"`);
      if (rule.thresholds && rule.thresholds.length !== iconCount(rule.iconSet)) throw new Error(`Icon set ${rule.iconSet} needs ${iconCount(rule.iconSet)} thresholds`);
      rule.thresholds?.forEach(validateThreshold);
    } else if (rule.type === 'colorScale') {
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

function validateThreshold(t: Threshold | undefined): void {
  if (!t) return;
  if (!['min', 'max', 'num', 'percent', 'percentile'].includes(t.type)) throw new Error(`Invalid threshold type "${t.type}"`);
  if (['num', 'percent', 'percentile'].includes(t.type) && !Number.isFinite(t.value)) throw new Error('Threshold needs a numeric value');
}

/** Earlier rules have higher priority. A resolver can be reused for one render/read. */
export class ConditionalFormatEvaluator {
  private scales = new Map<ConditionalFormatRule, { min: number; max: number }>();
  private stats = new Map<ConditionalFormatRule, number[]>();
  constructor(private wb: Workbook) {}

  /** Sorted numeric values of a rule's range (cached for this evaluator). */
  private numbers(sheet: Sheet, rule: ConditionalFormatRule, range: RangeAddress): number[] {
    let xs = this.stats.get(rule);
    if (!xs) {
      xs = [];
      for (const [key, cell] of sheet.cells) {
        const p = keyToAddress(key);
        if (typeof cell.value === 'number' && contains(range, p.row, p.col)) xs.push(cell.value);
      }
      xs.sort((a, b) => a - b);
      this.stats.set(rule, xs);
    }
    return xs;
  }

  private threshold(t: Threshold, xs: number[]): number {
    const min = xs[0] ?? 0, max = xs.at(-1) ?? 0;
    switch (t.type) {
      case 'min': return min;
      case 'max': return max;
      case 'num': return t.value ?? 0;
      case 'percent': return min + ((max - min) * (t.value ?? 0)) / 100;
      case 'percentile': {
        if (!xs.length) return 0;
        const pos = (xs.length - 1) * Math.max(0, Math.min(100, t.value ?? 0)) / 100;
        const lo = Math.floor(pos), hi = Math.ceil(pos);
        return xs[lo]! + (xs[hi]! - xs[lo]!) * (pos - lo);
      }
    }
  }

  /** Data bar and icon of a cell (the first matching rule of each kind wins). */
  visual(sheet: Sheet, row: number, col: number): CellVisual | undefined {
    const value = sheet.getValue(row, col);
    if (typeof value !== 'number' || !sheet.conditionalFormats.length) return undefined;
    let out: CellVisual | undefined;
    for (const rule of sheet.conditionalFormats) {
      if (rule.type !== 'dataBar' && rule.type !== 'iconSet') continue;
      const range = parseRange(rule.range)!.range;
      if (!contains(range, row, col)) continue;
      const xs = this.numbers(sheet, rule, range);
      out ??= {};
      if (rule.type === 'dataBar' && !out.bar) {
        const lo = this.threshold(rule.min ?? { type: 'min' }, xs);
        const hi = this.threshold(rule.max ?? { type: 'max' }, xs);
        const fraction = hi === lo ? (value >= hi ? 1 : 0) : Math.max(0, Math.min(1, (value - lo) / (hi - lo)));
        out.bar = { fraction, color: rule.color };
        if (rule.showValue === false) out.hideValue = true;
      } else if (rule.type === 'iconSet' && !out.icon) {
        const thresholds = rule.thresholds ?? defaultIconThresholds(rule.iconSet);
        let index = 0;
        for (let i = 1; i < thresholds.length; i++) if (value >= this.threshold(thresholds[i]!, xs)) index = i;
        out.icon = { set: rule.iconSet, index: rule.reverse ? thresholds.length - 1 - index : index };
        if (rule.showValue === false) out.hideValue = true;
      }
      if (rule.stopIfTrue) break;
    }
    return out;
  }

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
      if (rule.type === 'dataBar' || rule.type === 'iconSet') continue;
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
