import { Err, isError, type CellError, type Scalar } from '../../values.js';
import { broadcast, deref, toNumber, toText, type EvalValue, type FunctionContext } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { binary, def, numArg, textArg, unary } from './helpers.js';

// Excel's binary, octal and hexadecimal numbers have at most 10 digits in two's complement.
const SYSTEMS = { BIN: { radix: 2, bits: 10 }, OCT: { radix: 8, bits: 30 }, DEC: { radix: 10, bits: 0 }, HEX: { radix: 16, bits: 40 } } as const;
type System = keyof typeof SYSTEMS;

function digitsRe(radix: number): RegExp {
  const digits = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'.slice(0, radix);
  return new RegExp(`^[${digits}]*$`, 'i');
}

function fromSystem(v: Scalar, from: System): number | CellError {
  if (from === 'DEC') {
    const n = toNumber(v);
    return isError(n) ? n : Math.trunc(n);
  }
  const text = toText(v);
  if (isError(text)) return text;
  const { radix, bits } = SYSTEMS[from];
  if (text.length > 10 || !digitsRe(radix).test(text)) return Err.NUM;
  const value = text ? parseInt(text, radix) : 0;
  return text.length === 10 && value >= 2 ** (bits - 1) ? value - 2 ** bits : value;
}

function toSystem(value: number, to: System, places: Scalar): Scalar {
  if (to === 'DEC') return value;
  const { radix, bits } = SYSTEMS[to];
  if (value < -(2 ** (bits - 1)) || value >= 2 ** (bits - 1)) return Err.NUM;
  // Negative numbers always use all 10 digits; places is ignored for them.
  if (value < 0) return (value + 2 ** bits).toString(radix).toUpperCase();
  const text = value.toString(radix).toUpperCase();
  if (places === null) return text;
  const p = toNumber(places);
  if (isError(p)) return p;
  const width = Math.trunc(p);
  if (width < 1 || width > 10 || text.length > width) return Err.NUM;
  return text.padStart(width, '0');
}

function conversion(from: System, to: System): FunctionDef {
  const name = `${from}2${to}`;
  const places = to !== 'DEC';
  return def(name, 'math', places ? `${name}(number, [places])` : `${name}(number)`, `Converts a ${from.toLowerCase()} number to ${to.toLowerCase()}.`, ([a, b], ctx) =>
    broadcast([deref(a!, ctx), b === undefined ? null : deref(b, ctx)], (v, p) => {
      const value = fromSystem(v, from);
      return isError(value) ? value : toSystem(value, to, p);
    }),
    { minArgs: 1, maxArgs: places ? 2 : 1 },
  );
}

const PAIRS: [System, System][] = [
  ['BIN', 'DEC'], ['BIN', 'HEX'], ['BIN', 'OCT'], ['DEC', 'BIN'], ['DEC', 'HEX'], ['DEC', 'OCT'],
  ['HEX', 'BIN'], ['HEX', 'DEC'], ['HEX', 'OCT'], ['OCT', 'BIN'], ['OCT', 'DEC'], ['OCT', 'HEX'],
];

const ROMAN_CHARS = ['M', 'D', 'C', 'L', 'X', 'V', 'I'];
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/** Roman numeral; forms 1-4 are Excel's increasingly concise variants (499 = CDXCIX, LDVLIV, XDIX, VDIV, ID). */
function roman(value: number, form: number): string {
  let out = '';
  let rest = value;
  const last = ROMAN_VALUES.length - 1;
  for (let i = 0; i <= last / 2; i++) {
    let index = 2 * i;
    const digit = Math.floor(rest / ROMAN_VALUES[index]!);
    if (digit % 5 === 4) {
      const index2 = digit === 4 ? index - 1 : index - 2;
      // A more concise form subtracts smaller numerals where the value allows it.
      for (let steps = 0; steps < form && index < last; steps++) {
        if (ROMAN_VALUES[index2]! - ROMAN_VALUES[index + 1]! <= rest) index++;
        else break;
      }
      out += ROMAN_CHARS[index]! + ROMAN_CHARS[index2]!;
      rest += ROMAN_VALUES[index]! - ROMAN_VALUES[index2]!;
    } else {
      if (digit > 4) out += ROMAN_CHARS[index - 1]!;
      out += ROMAN_CHARS[index]!.repeat(digit % 5);
      rest %= ROMAN_VALUES[index]!;
    }
  }
  return out;
}

function arabic(text: string): number | CellError {
  let s = text.trim().toUpperCase();
  if (s.length > 255) return Err.VALUE;
  const sign = s.startsWith('-') ? -1 : 1;
  if (sign < 0) s = s.slice(1);
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const v = ROMAN_VALUES[ROMAN_CHARS.indexOf(s[i]!)];
    if (v === undefined) return Err.VALUE;
    const next = ROMAN_VALUES[ROMAN_CHARS.indexOf(s[i + 1] ?? '')] ?? 0;
    total += v < next ? -v : v;
  }
  return sign * total;
}

function romanForm(arg: EvalValue | undefined, ctx: FunctionContext): number | CellError {
  if (arg === undefined) return 0;
  const v = deref(arg, ctx);
  if (typeof v === 'boolean') return v ? 0 : 4;
  const n = toNumber(v as Scalar);
  if (isError(n)) return n;
  const form = Math.trunc(n);
  return form < 0 || form > 4 ? Err.VALUE : form;
}

export const engineeringFunctions: FunctionDef[] = [
  ...PAIRS.map(([from, to]) => conversion(from, to)),
  def('BASE', 'math', 'BASE(number, radix, [min_length])', 'Converts a number to text in another base (2-36).', ([n, r, l], ctx) => {
    const value = Math.trunc(numArg(n, ctx));
    const radix = Math.trunc(numArg(r, ctx));
    const length = Math.trunc(numArg(l, ctx, 0));
    if (value < 0 || value >= 2 ** 53 || radix < 2 || radix > 36 || length < 0 || length > 255) return Err.NUM;
    return value.toString(radix).toUpperCase().padStart(length, '0');
  }, { minArgs: 2, maxArgs: 3 }),
  def('DECIMAL', 'math', 'DECIMAL(text, radix)', 'Converts text in another base (2-36) to a number.', ([t, r], ctx) => {
    const text = textArg(t, ctx);
    const radix = Math.trunc(numArg(r, ctx));
    if (radix < 2 || radix > 36 || text.length > 255 || !digitsRe(radix).test(text)) return Err.NUM;
    let value = 0;
    for (const ch of text) value = value * radix + parseInt(ch, radix);
    return value;
  }, { minArgs: 2, maxArgs: 2 }),
  def('ROMAN', 'text', 'ROMAN(number, [form])', 'Roman numeral as text; form 0 (classic) to 4 (most concise).', ([n, f], ctx) => {
    const value = Math.trunc(numArg(n, ctx));
    const form = romanForm(f, ctx);
    if (isError(form)) return form;
    return value < 0 || value > 3999 ? Err.VALUE : roman(value, form);
  }, { minArgs: 1, maxArgs: 2 }),
  def('ARABIC', 'text', 'ARABIC(text)', 'Converts a Roman numeral to a number.', ([t], ctx) => arabic(textArg(t, ctx)), { minArgs: 1, maxArgs: 1 }),
  unary('SINH', 'Hyperbolic sine.', Math.sinh),
  unary('COSH', 'Hyperbolic cosine.', Math.cosh),
  unary('TANH', 'Hyperbolic tangent.', Math.tanh),
  unary('ASINH', 'Inverse hyperbolic sine.', Math.asinh),
  unary('ACOSH', 'Inverse hyperbolic cosine.', (x) => (x < 1 ? Err.NUM : Math.acosh(x))),
  unary('ATANH', 'Inverse hyperbolic tangent.', (x) => (Math.abs(x) >= 1 ? Err.NUM : Math.atanh(x))),
  binary('COMBIN', 'COMBIN(number, number_chosen)', 'Number of combinations without repetition.', (a, b) => {
    const n = Math.trunc(a), k = Math.trunc(b);
    if (n < 0 || k < 0 || n < k) return Err.NUM;
    let r = 1;
    // Stop at the overflow, which comes within about a thousand steps.
    for (let i = 1; i <= Math.min(k, n - k) && Number.isFinite(r); i++) r = (r * (n - Math.min(k, n - k) + i)) / i;
    return !Number.isFinite(r) ? Err.NUM : r < 2 ** 53 ? Math.round(r) : r;
  }),
  binary('PERMUT', 'PERMUT(number, number_chosen)', 'Number of ordered selections without repetition.', (a, b) => {
    const n = Math.trunc(a), k = Math.trunc(b);
    if (n <= 0 || k < 0 || n < k) return Err.NUM;
    let r = 1;
    for (let i = 0; i < k && Number.isFinite(r); i++) r *= n - i;
    return Number.isFinite(r) ? r : Err.NUM;
  }),
];
