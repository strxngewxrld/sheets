export type NumberFormatCategory = 'general' | 'number' | 'thousands' | 'currency' | 'percent' | 'date' | 'time' | 'scientific' | 'text';

/** The toolbar category of a number format code, e.g. `0.00 "€"` → currency. Unknown codes return undefined. */
export function numberFormatCategory(code: string | undefined): NumberFormatCategory | undefined {
  if (!code || code === 'General') return 'general';
  const section = code.split(';')[0]!;
  if (/^@$/.test(section.trim())) return 'text';
  if (/[€$£¥₹]|\[\$[^\]]*\]|"(?:EUR|USD|CHF|GBP)"/i.test(section)) return 'currency';
  // Ignore literal text and colors/conditions when looking at the format tokens.
  const tokens = section.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
  if (tokens.includes('%')) return 'percent';
  if (/E[+-]/i.test(tokens)) return 'scientific';
  if (/[dy]/i.test(tokens)) return 'date';
  if (/[hs]/i.test(tokens) || /m/i.test(tokens)) return /[hs]/i.test(tokens) ? 'time' : 'date';
  if (/#,##0|0,000/.test(tokens)) return 'thousands';
  if (/[0#]/.test(tokens)) return 'number';
  return undefined;
}
