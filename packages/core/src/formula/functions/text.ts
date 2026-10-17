import { formatValue, parseNumber } from '../../numfmt.js';
import { Err, isError } from '../../values.js';
import { broadcast, deref, toText } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { boolArg, def, numArg, scalarArg, scalarsOf, textArg, wildcardToRegExp } from './helpers.js';

/** Text function of one argument, broadcast over arrays. */
function text1(name: string, description: string, f: (s: string) => string | number): FunctionDef {
  return def(name, 'text', `${name}(text)`, description, ([a], ctx) =>
    broadcast([deref(a!, ctx)], (v) => {
      const s = toText(v);
      return isError(s) ? s : f(s);
    }),
    { minArgs: 1, maxArgs: 1 },
  );
}

export const textFunctions: FunctionDef[] = [
  def('CONCATENATE', 'text', 'CONCATENATE(text1, [text2], ...)', 'Joins text values.', (args, ctx) => args.map((a) => textArg(a, ctx)).join(''), { minArgs: 1 }),
  def('CONCAT', 'text', 'CONCAT(text1, [text2], ...)', 'Joins text values and ranges.', (args, ctx) => {
    let out = '';
    for (const v of scalarsOf(args, ctx)) {
      const s = toText(v);
      if (isError(s)) return s;
      out += s;
    }
    return out;
  }, { minArgs: 1 }),
  def('TEXTJOIN', 'text', 'TEXTJOIN(delimiter, ignore_empty, text1, ...)', 'Joins text with a delimiter.', ([d, ignore, ...rest], ctx) => {
    const delim = textArg(d, ctx);
    const skip = boolArg(ignore, ctx, true);
    const parts: string[] = [];
    for (const v of scalarsOf(rest, ctx)) {
      const s = toText(v);
      if (isError(s)) return s;
      if (skip && s === '') continue;
      parts.push(s);
    }
    return parts.join(delim);
  }, { minArgs: 3 }),
  def('LEFT', 'text', 'LEFT(text, [num_chars])', 'Leftmost characters.', ([t, n], ctx) => {
    const k = numArg(n, ctx, 1);
    return k < 0 ? Err.VALUE : textArg(t, ctx).slice(0, k);
  }, { minArgs: 1, maxArgs: 2 }),
  def('RIGHT', 'text', 'RIGHT(text, [num_chars])', 'Rightmost characters.', ([t, n], ctx) => {
    const k = numArg(n, ctx, 1);
    if (k < 0) return Err.VALUE;
    const s = textArg(t, ctx);
    return k === 0 ? '' : s.slice(-k);
  }, { minArgs: 1, maxArgs: 2 }),
  def('MID', 'text', 'MID(text, start_num, num_chars)', 'Characters from the middle of a text.', ([t, s, n], ctx) => {
    const start = numArg(s, ctx);
    const len = numArg(n, ctx);
    if (start < 1 || len < 0) return Err.VALUE;
    return textArg(t, ctx).substr(start - 1, len);
  }, { minArgs: 3, maxArgs: 3 }),
  text1('LEN', 'Number of characters.', (s) => s.length),
  text1('UPPER', 'Converts to upper case.', (s) => s.toUpperCase()),
  text1('LOWER', 'Converts to lower case.', (s) => s.toLowerCase()),
  text1('PROPER', 'Capitalizes each word.', (s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_, p: string, c: string) => p + c.toUpperCase())),
  text1('TRIM', 'Removes extra spaces.', (s) => s.trim().replace(/ {2,}/g, ' ')),
  text1('CLEAN', 'Removes non-printable characters.', (s) => s.replace(/[\x00-\x1f]/g, '')),
  def('SUBSTITUTE', 'text', 'SUBSTITUTE(text, old_text, new_text, [instance_num])', 'Replaces occurrences of text.', ([t, o, n, inst], ctx) => {
    const s = textArg(t, ctx);
    const old = textArg(o, ctx);
    const rep = textArg(n, ctx);
    if (!old) return s;
    if (inst === undefined) return s.split(old).join(rep);
    const k = numArg(inst, ctx);
    if (k < 1) return Err.VALUE;
    let idx = -1;
    for (let i = 0; i < k; i++) {
      idx = s.indexOf(old, idx + 1);
      if (idx < 0) return s;
    }
    return s.slice(0, idx) + rep + s.slice(idx + old.length);
  }, { minArgs: 3, maxArgs: 4 }),
  def('REPLACE', 'text', 'REPLACE(old_text, start_num, num_chars, new_text)', 'Replaces part of a text by position.', ([t, s, n, r], ctx) => {
    const str = textArg(t, ctx);
    const start = numArg(s, ctx);
    const len = numArg(n, ctx);
    if (start < 1 || len < 0) return Err.VALUE;
    return str.slice(0, start - 1) + textArg(r, ctx) + str.slice(start - 1 + len);
  }, { minArgs: 4, maxArgs: 4 }),
  def('FIND', 'text', 'FIND(find_text, within_text, [start_num])', 'Case-sensitive position of text (1-based).', ([f, w, s], ctx) => {
    const start = numArg(s, ctx, 1);
    const idx = textArg(w, ctx).indexOf(textArg(f, ctx), start - 1);
    return idx < 0 || start < 1 ? Err.VALUE : idx + 1;
  }, { minArgs: 2, maxArgs: 3 }),
  def('SEARCH', 'text', 'SEARCH(find_text, within_text, [start_num])', 'Case-insensitive position of text with wildcards.', ([f, w, s], ctx) => {
    const start = numArg(s, ctx, 1);
    const within = textArg(w, ctx);
    const re = wildcardToRegExp(textArg(f, ctx));
    const pattern = new RegExp(re.source.slice(1, -1), 'is');
    const m = pattern.exec(within.slice(start - 1));
    return !m || start < 1 ? Err.VALUE : m.index + start;
  }, { minArgs: 2, maxArgs: 3 }),
  def('TEXT', 'text', 'TEXT(value, format_text)', 'Formats a number using a format code.', ([v, f], ctx) => {
    const val = scalarArg(v, ctx);
    if (isError(val)) return val;
    return formatValue(val, textArg(f, ctx));
  }, { minArgs: 2, maxArgs: 2 }),
  def('VALUE', 'text', 'VALUE(text)', 'Converts text to a number.', ([t], ctx) => {
    const v = scalarArg(t, ctx);
    if (typeof v === 'number') return v;
    const n = parseNumber(textArg(t, ctx));
    return n === null ? Err.VALUE : n;
  }, { minArgs: 1, maxArgs: 1 }),
  def('REPT', 'text', 'REPT(text, number_times)', 'Repeats text.', ([t, n], ctx) => {
    const k = numArg(n, ctx);
    return k < 0 ? Err.VALUE : textArg(t, ctx).repeat(Math.trunc(k));
  }, { minArgs: 2, maxArgs: 2 }),
  def('EXACT', 'text', 'EXACT(text1, text2)', 'Case-sensitive equality.', ([a, b], ctx) => textArg(a, ctx) === textArg(b, ctx), { minArgs: 2, maxArgs: 2 }),
  def('CHAR', 'text', 'CHAR(number)', 'Character for a code.', ([n], ctx) => String.fromCharCode(numArg(n, ctx)), { minArgs: 1, maxArgs: 1 }),
  def('CODE', 'text', 'CODE(text)', 'Code of the first character.', ([t], ctx) => {
    const s = textArg(t, ctx);
    return s ? s.charCodeAt(0) : Err.VALUE;
  }, { minArgs: 1, maxArgs: 1 }),
  def('T', 'text', 'T(value)', 'Returns text or empty string.', ([v], ctx) => {
    const s = scalarArg(v, ctx);
    return typeof s === 'string' ? s : isError(s) ? s : '';
  }, { minArgs: 1, maxArgs: 1 }),
  def('FIXED', 'text', 'FIXED(number, [decimals], [no_commas])', 'Formats a number with fixed decimals.', ([n, d, nc], ctx) => {
    const dec = Math.trunc(numArg(d, ctx, 2));
    const fmt = (boolArg(nc, ctx) ? '0' : '#,##0') + (dec > 0 ? '.' + '0'.repeat(dec) : '');
    return formatValue(numArg(n, ctx), fmt);
  }, { minArgs: 1, maxArgs: 3 }),
];
