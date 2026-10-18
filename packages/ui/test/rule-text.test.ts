import { describe, expect, it } from 'vitest';
import type { ConditionalFormatRule } from '@sheets/core';
import { getStrings } from '../src/i18n.js';
import { conditionalRuleLabel } from '../src/rule-text.js';

const rules: ConditionalFormatRule[] = [
  { id: '1', range: 'C2:C6', type: 'cellIs', operator: 'greaterThan', formulas: ['6'], style: { fill: '#FF0000' } },
  { id: '2', range: 'C2:C6', type: 'cellIs', operator: 'lessThanOrEqual', formulas: ['"x"'], style: {} },
  { id: '3', range: 'A1:A9', type: 'cellIs', operator: 'between', formulas: ['1', '$B$1'], style: {} },
  { id: '4', range: 'A1:A9', type: 'cellIs', operator: 'notBetween', formulas: ['1', '5'], style: {} },
  { id: '5', range: 'A1:D9', type: 'expression', formulas: ['$B1<0'], style: {} },
  { id: '6', range: 'B2:B9', type: 'colorScale', colors: ['#FF0000', '#FFFF00', '#00FF00'] },
  { id: '7', range: 'B2:B9', type: 'dataBar', color: '#638EC6' },
  { id: '8', range: 'B2:B9', type: 'iconSet', iconSet: '3TrafficLights1' },
];

describe('conditional rule labels', () => {
  it('describe every rule type in German', () => {
    expect(rules.map((r) => conditionalRuleLabel(r, getStrings('de')))).toEqual([
      'C2:C6 · Zellwert > 6', 'C2:C6 · Zellwert ≤ "x"', 'A1:A9 · Zellwert zwischen 1 und $B$1', 'A1:A9 · Zellwert nicht zwischen 1 und 5',
      'A1:D9 · Formel: =$B1<0', 'B2:B9 · 3-Farben-Skala', 'B2:B9 · Datenbalken', 'B2:B9 · Symbole: Ampeln',
    ]);
  });
  it('describe every rule type in English without internal type names', () => {
    const labels = rules.map((r) => conditionalRuleLabel(r, getStrings('en')));
    expect(labels[0]).toBe('C2:C6 · Cell value > 6');
    expect(labels[5]).toBe('B2:B9 · 3-color scale');
    expect(labels[7]).toBe('B2:B9 · Icons: traffic lights');
    for (const label of labels) expect(label).not.toMatch(/cellIs|expression|colorScale|dataBar|iconSet/);
  });
});
