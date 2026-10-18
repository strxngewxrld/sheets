import type { ConditionalFormatRule, ConditionalOperator } from '@sheets/core';
import type { Strings } from './i18n.js';

const SYMBOLS: Partial<Record<ConditionalOperator, string>> = { equal: '=', notEqual: '≠', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '≥', lessThanOrEqual: '≤' };

const fill = (template: string, ...values: string[]) => template.replace(/\{(\d)\}/g, (_, i: string) => values[Number(i)] ?? '');

/** Readable, localized summary of a conditional formatting rule (for rule lists). */
export function describeConditionalRule(rule: ConditionalFormatRule, s: Strings): string {
  const formula = (i: number) => ('formulas' in rule ? rule.formulas[i] ?? '' : '').replace(/^=/, '');
  switch (rule.type) {
    case 'cellIs':
      if (rule.operator === 'between') return fill(s.ruleBetween, formula(0), formula(1));
      if (rule.operator === 'notBetween') return fill(s.ruleNotBetween, formula(0), formula(1));
      return fill(s.ruleCellValue, SYMBOLS[rule.operator] ?? rule.operator, formula(0));
    case 'expression':
      return fill(s.ruleFormula, `=${formula(0)}`);
    case 'colorScale':
      return fill(s.ruleColorScale, String(rule.colors.length));
    case 'dataBar':
      return s.ruleDataBar;
    case 'iconSet':
      return fill(s.ruleIconSet, s.iconSetNames[rule.iconSet.slice(1)] ?? rule.iconSet);
  }
}

/** List entry: range and rule summary. */
export function conditionalRuleLabel(rule: ConditionalFormatRule, s: Strings): string {
  return `${rule.range} · ${describeConditionalRule(rule, s)}`;
}
