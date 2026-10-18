import { describe, expect, it } from 'vitest';
import { numberFormatCategory } from '../src/format-names.js';

describe('number format categories', () => {
  it('names custom formats like the toolbar presets', () => {
    const cases: [string, string | undefined][] = [
      ['0.00 "€"', 'currency'], ['#,##0.00 [$€-407]', 'currency'], ['"$"#,##0', 'currency'], ['#,##0 "EUR";[Red]-#,##0 "EUR"', 'currency'],
      ['0.0%', 'percent'], ['0.000E+00', 'scientific'], ['d. mmmm yyyy', 'date'], ['mmm-yy', 'date'], ['[h]:mm:ss', 'time'], ['hh:mm AM/PM', 'time'],
      ['#,##0', 'thousands'], ['0.000', 'number'], ['0.000 "kg"', 'number'], ['@', 'text'], ['General', 'general'], ['', 'general'], ['"Text only"', undefined],
    ];
    for (const [code, category] of cases) expect(numberFormatCategory(code), code).toBe(category);
  });
});
