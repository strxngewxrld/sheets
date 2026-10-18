import { Err, isError, type CellError, type Scalar } from '../../values.js';
import { broadcast, deref, isMatrix, isRef, toMatrix, toNumber, type EvalValue, type FunctionContext } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import {
  bail,
  binary,
  criteriaMask,
  def,
  numArg,
  numbersOf,
  roundHalfAway,
  scalarsOf,
  unary,
} from './helpers.js';

export const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);

// Rounding to multiples works on 15 significant digits like Excel: in binary floating
// point 0.3 / 0.1 is 2.9999999999999996 (FLOOR would give 0.2) and 3 * 0.1 is 0.30000000000000004.
const ratio = (x: number, s: number) => Number((x / s).toPrecision(15));
const times = (q: number, s: number) => Number((q * s).toPrecision(15));

/** CEILING.MATH / FLOOR.MATH(number, [significance], [mode]): the sign of significance is ignored. */
function toMultiple(name: string, description: string, round: (q: number) => number): FunctionDef {
  return def(name, 'math', `${name}(number, [significance], [mode])`, description, ([a, b, c], ctx) =>
    broadcast([deref(a!, ctx), b === undefined ? null : deref(b, ctx), c === undefined ? null : deref(c, ctx)], (x, y, z) => {
      const n = toNumber(x);
      if (isError(n)) return n;
      const sig = y === null ? 1 : toNumber(y);
      if (isError(sig)) return sig;
      const mode = z === null ? 0 : toNumber(z);
      if (isError(mode)) return mode;
      const s = Math.abs(sig);
      if (s === 0) return 0;
      // With a mode, negative numbers are rounded as their absolute value and negated.
      return n < 0 && mode !== 0 ? -times(round(ratio(-n, s)), s) : times(round(ratio(n, s)), s);
    }),
    { minArgs: 1, maxArgs: 3 },
  );
}
export const mean = (xs: number[]) => (xs.length ? sum(xs) / xs.length : bail(Err.DIV0));

export function variance(xs: number[], sample: boolean): number {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) bail(Err.DIV0);
  const m = sum(xs) / n;
  return xs.reduce((a, x) => a + (x - m) ** 2, 0) / (sample ? n - 1 : n);
}

export function percentileInc(xs: number[], k: number): number {
  if (!xs.length || k < 0 || k > 1) bail(Err.NUM);
  const s = [...xs].sort((a, b) => a - b);
  const pos = (s.length - 1) * k;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return s[lo]! + (s[hi]! - s[lo]!) * (pos - lo);
}

export function percentileExc(xs: number[], k: number): number {
  const n = xs.length;
  if (!n || k <= 0 || k >= 1) bail(Err.NUM);
  const pos = k * (n + 1) - 1;
  if (pos < 0 || pos > n - 1) bail(Err.NUM);
  const s = [...xs].sort((a, b) => a - b);
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return s[lo]! + (s[hi]! - s[lo]!) * (pos - lo);
}

export function modeOf(xs: number[]): number | CellError {
  const counts = new Map<number, number>();
  let best: number | null = null;
  let bestCount = 1;
  for (const x of xs) {
    const c = (counts.get(x) ?? 0) + 1;
    counts.set(x, c);
    if (c > bestCount) {
      bestCount = c;
      best = x;
    }
  }
  return best ?? Err.NA;
}

function median(xs: number[]): number {
  if (!xs.length) bail(Err.NUM);
  return percentileInc(xs, 0.5);
}

function gcd(a: number, b: number): number {
  return b === 0 ? a : gcd(b, a % b);
}

/** Collects values from the "sum range" positions where the mask is true. */
function maskedNumbers(arg: EvalValue, mask: boolean[], cols: number, ctx: FunctionContext): number[] {
  const m = toMatrix(arg, ctx);
  const out: number[] = [];
  mask.forEach((ok, idx) => {
    if (!ok) return;
    const v = m[Math.floor(idx / cols)]?.[idx % cols] ?? null;
    if (typeof v === 'number') out.push(v);
    else if (isError(v)) bail(v);
  });
  return out;
}

function pairs(args: EvalValue[]): [EvalValue, EvalValue][] {
  if (args.length % 2 !== 0) bail(Err.VALUE);
  const out: [EvalValue, EvalValue][] = [];
  for (let i = 0; i < args.length; i += 2) out.push([args[i]!, args[i + 1]!]);
  return out;
}

const AGGREGATES: Record<number, (xs: number[], all: Scalar[]) => number> = {
  1: mean,
  2: (xs) => xs.length,
  3: (_xs, all) => all.filter((v) => v !== null && v !== '').length,
  4: (xs) => (xs.length ? Math.max(...xs) : 0),
  5: (xs) => (xs.length ? Math.min(...xs) : 0),
  6: (xs) => xs.reduce((a, b) => a * b, 1),
  7: (xs) => Math.sqrt(variance(xs, true)),
  8: (xs) => Math.sqrt(variance(xs, false)),
  9: sum,
  10: (xs) => variance(xs, true),
  11: (xs) => variance(xs, false),
};

/**
 * Replaces range arguments by their values without hidden rows. Rows hidden
 * by a filter are always skipped; manually hidden rows only if `manual`.
 */
function withoutHiddenRows(args: EvalValue[], ctx: FunctionContext, manual: boolean): EvalValue[] {
  if (!ctx.rowHidden) return args;
  return args.map((a) => {
    if (!isRef(a)) return a;
    const m = toMatrix(a, ctx);
    return m.filter((_, i) => {
      const why = ctx.rowHidden!(a.sheet, a.range.start.row + i);
      return !(why === 'filtered' || (why === 'manual' && manual));
    });
  });
}

export const mathFunctions: FunctionDef[] = [
  def('SUM', 'math', 'SUM(number1, [number2], ...)', 'Adds all numbers.', (args, ctx) => sum(numbersOf(args, ctx)), { minArgs: 1 }),
  def('PRODUCT', 'math', 'PRODUCT(number1, [number2], ...)', 'Multiplies all numbers.', (args, ctx) => numbersOf(args, ctx).reduce((a, b) => a * b, 1), { minArgs: 1 }),
  def('AVERAGE', 'statistical', 'AVERAGE(number1, [number2], ...)', 'Arithmetic mean.', (args, ctx) => mean(numbersOf(args, ctx)), { minArgs: 1 }),
  def('MIN', 'statistical', 'MIN(number1, [number2], ...)', 'Smallest number.', (args, ctx) => {
    const xs = numbersOf(args, ctx);
    return xs.length ? Math.min(...xs) : 0;
  }, { minArgs: 1 }),
  def('MAX', 'statistical', 'MAX(number1, [number2], ...)', 'Largest number.', (args, ctx) => {
    const xs = numbersOf(args, ctx);
    return xs.length ? Math.max(...xs) : 0;
  }, { minArgs: 1 }),
  def('COUNT', 'statistical', 'COUNT(value1, [value2], ...)', 'Counts numeric values.', (args, ctx) => {
    let n = 0;
    for (const a of args) {
      if (isRef(a) || isMatrix(a)) {
        for (const row of toMatrix(a, ctx)) for (const v of row) if (typeof v === 'number') n++;
      } else if (a !== null && !isError(toNumber(a as Scalar))) n++;
    }
    return n;
  }, { minArgs: 1 }),
  def('COUNTA', 'statistical', 'COUNTA(value1, [value2], ...)', 'Counts non-empty values.', (args, ctx) => {
    let n = 0;
    for (const v of scalarsOf(args, ctx)) if (v !== null) n++;
    return n;
  }, { minArgs: 1 }),
  def('COUNTBLANK', 'statistical', 'COUNTBLANK(range)', 'Counts empty cells.', ([a], ctx) => {
    const m = toMatrix(a!, ctx);
    let filled = 0;
    for (const row of m) for (const v of row) if (v !== null && v !== '') filled++;
    const total = isRef(a!) ? (a.range.end.row - a.range.start.row + 1) * (a.range.end.col - a.range.start.col + 1) : m.length * (m[0]?.length ?? 0);
    return total - filled;
  }, { minArgs: 1, maxArgs: 1 }),

  binary('ROUND', 'ROUND(number, digits)', 'Rounds to the given number of digits.', (x, d) => roundHalfAway(x, Math.trunc(d))),
  binary('ROUNDUP', 'ROUNDUP(number, digits)', 'Rounds away from zero.', (x, d) => {
    const f = 10 ** Math.trunc(d);
    return (Math.sign(x) * Math.ceil(Number((Math.abs(x) * f).toPrecision(15)))) / f;
  }),
  binary('ROUNDDOWN', 'ROUNDDOWN(number, digits)', 'Rounds towards zero.', (x, d) => {
    const f = 10 ** Math.trunc(d);
    return (Math.sign(x) * Math.floor(Number((Math.abs(x) * f).toPrecision(15)))) / f;
  }),
  binary('TRUNC', 'TRUNC(number, [digits])', 'Truncates a number.', (x, d) => {
    const f = 10 ** Math.trunc(d);
    return Math.trunc(x * f) / f;
  }, { defaultY: 0 }),
  binary('MROUND', 'MROUND(number, multiple)', 'Rounds to the nearest multiple.', (x, m) => (m === 0 ? 0 : x * m < 0 ? Err.NUM : times(roundHalfAway(ratio(x, m), 0), m))),
  // A positive number with a negative significance is #NUM! in Excel.
  binary('CEILING', 'CEILING(number, significance)', 'Rounds up to a multiple.', (x, s) => (s === 0 ? 0 : x > 0 && s < 0 ? Err.NUM : times(Math.ceil(ratio(x, s)), s)), { defaultY: 1 }),
  toMultiple('CEILING.MATH', 'Rounds up to a multiple; negative numbers away from zero when mode is not 0.', Math.ceil),
  binary('FLOOR', 'FLOOR(number, significance)', 'Rounds down to a multiple.', (x, s) => (s === 0 ? Err.DIV0 : x > 0 && s < 0 ? Err.NUM : times(Math.floor(ratio(x, s)), s)), { defaultY: 1 }),
  toMultiple('FLOOR.MATH', 'Rounds down to a multiple; negative numbers towards zero when mode is not 0.', Math.floor),
  binary('MOD', 'MOD(number, divisor)', 'Remainder after division (sign of divisor).', (x, y) => (y === 0 ? Err.DIV0 : x - y * Math.floor(x / y))),
  binary('POWER', 'POWER(number, power)', 'Raises a number to a power.', (x, y) => Math.pow(x, y)),
  binary('LOG', 'LOG(number, [base])', 'Logarithm to the given base.', (x, b) => (x <= 0 || b <= 0 || b === 1 ? Err.NUM : Math.log(x) / Math.log(b)), { defaultY: 10 }),
  binary('ATAN2', 'ATAN2(x, y)', 'Arc tangent of x/y coordinates.', (x, y) => (x === 0 && y === 0 ? Err.DIV0 : Math.atan2(y, x))),
  binary('QUOTIENT', 'QUOTIENT(numerator, denominator)', 'Integer part of a division.', (x, y) => (y === 0 ? Err.DIV0 : Math.trunc(x / y))),
  unary('INT', 'Rounds down to the nearest integer.', Math.floor),
  unary('ABS', 'Absolute value.', Math.abs),
  unary('SIGN', 'Sign of a number (-1, 0, 1).', Math.sign),
  unary('SQRT', 'Square root.', (x) => (x < 0 ? Err.NUM : Math.sqrt(x))),
  unary('EXP', 'e raised to a power.', Math.exp),
  unary('LN', 'Natural logarithm.', (x) => (x <= 0 ? Err.NUM : Math.log(x))),
  unary('LOG10', 'Base-10 logarithm.', (x) => (x <= 0 ? Err.NUM : Math.log10(x))),
  unary('SIN', 'Sine (radians).', Math.sin),
  unary('COS', 'Cosine (radians).', Math.cos),
  unary('TAN', 'Tangent (radians).', Math.tan),
  unary('ASIN', 'Arc sine.', (x) => (Math.abs(x) > 1 ? Err.NUM : Math.asin(x))),
  unary('ACOS', 'Arc cosine.', (x) => (Math.abs(x) > 1 ? Err.NUM : Math.acos(x))),
  unary('ATAN', 'Arc tangent.', Math.atan),
  unary('DEGREES', 'Radians to degrees.', (x) => (x * 180) / Math.PI),
  unary('RADIANS', 'Degrees to radians.', (x) => (x * Math.PI) / 180),
  unary('EVEN', 'Rounds up to the next even integer.', (x) => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2),
  unary('ODD', 'Rounds up to the next odd integer.', (x) => {
    const v = Math.ceil(Math.abs(x));
    return Math.sign(x || 1) * (v % 2 === 1 ? v : v + 1);
  }),
  unary('FACT', 'Factorial.', (x) => {
    if (x < 0) return Err.NUM;
    let r = 1;
    for (let i = 2; i <= Math.floor(x); i++) r *= i;
    return r;
  }),
  def('PI', 'math', 'PI()', 'The number π.', () => Math.PI, { maxArgs: 0 }),
  def('RAND', 'math', 'RAND()', 'Random number between 0 and 1.', () => Math.random(), { maxArgs: 0, volatile: true }),
  def('RANDBETWEEN', 'math', 'RANDBETWEEN(bottom, top)', 'Random integer in a range.', ([a, b], ctx) => {
    const lo = Math.ceil(numArg(a, ctx));
    const hi = Math.floor(numArg(b, ctx));
    if (lo > hi) return Err.NUM;
    return lo + Math.floor(Math.random() * (hi - lo + 1));
  }, { minArgs: 2, maxArgs: 2, volatile: true }),
  def('GCD', 'math', 'GCD(number1, [number2], ...)', 'Greatest common divisor.', (args, ctx) => numbersOf(args, ctx).map(Math.trunc).reduce((a, b) => gcd(a, b), 0), { minArgs: 1 }),
  def('LCM', 'math', 'LCM(number1, [number2], ...)', 'Least common multiple.', (args, ctx) => numbersOf(args, ctx).map(Math.trunc).reduce((a, b) => (a === 0 || b === 0 ? 0 : (a * b) / gcd(a, b)), 1), { minArgs: 1 }),
  def('SUMSQ', 'math', 'SUMSQ(number1, [number2], ...)', 'Sum of squares.', (args, ctx) => sum(numbersOf(args, ctx).map((x) => x * x)), { minArgs: 1 }),

  // Conditional aggregates
  def('SUMIF', 'math', 'SUMIF(range, criteria, [sum_range])', 'Sums cells matching a condition.', ([range, crit, sumRange], ctx) => {
    const { mask, cols } = criteriaMask([[range!, crit!]], ctx);
    return sum(maskedNumbers(sumRange ?? range!, mask, cols, ctx));
  }, { minArgs: 2, maxArgs: 3 }),
  def('SUMIFS', 'math', 'SUMIFS(sum_range, criteria_range1, criteria1, ...)', 'Sums cells matching all conditions.', ([sumRange, ...rest], ctx) => {
    const { mask, cols } = criteriaMask(pairs(rest), ctx);
    return sum(maskedNumbers(sumRange!, mask, cols, ctx));
  }, { minArgs: 3 }),
  def('COUNTIF', 'statistical', 'COUNTIF(range, criteria)', 'Counts cells matching a condition.', ([range, crit], ctx) => criteriaMask([[range!, crit!]], ctx).mask.filter(Boolean).length, { minArgs: 2, maxArgs: 2 }),
  def('COUNTIFS', 'statistical', 'COUNTIFS(criteria_range1, criteria1, ...)', 'Counts cells matching all conditions.', (args, ctx) => criteriaMask(pairs(args), ctx).mask.filter(Boolean).length, { minArgs: 2 }),
  def('AVERAGEIF', 'statistical', 'AVERAGEIF(range, criteria, [average_range])', 'Average of cells matching a condition.', ([range, crit, avgRange], ctx) => {
    const { mask, cols } = criteriaMask([[range!, crit!]], ctx);
    return mean(maskedNumbers(avgRange ?? range!, mask, cols, ctx));
  }, { minArgs: 2, maxArgs: 3 }),
  def('AVERAGEIFS', 'statistical', 'AVERAGEIFS(average_range, criteria_range1, criteria1, ...)', 'Average of cells matching all conditions.', ([avgRange, ...rest], ctx) => {
    const { mask, cols } = criteriaMask(pairs(rest), ctx);
    return mean(maskedNumbers(avgRange!, mask, cols, ctx));
  }, { minArgs: 3 }),
  def('MAXIFS', 'statistical', 'MAXIFS(max_range, criteria_range1, criteria1, ...)', 'Maximum of cells matching all conditions.', ([r, ...rest], ctx) => {
    const { mask, cols } = criteriaMask(pairs(rest), ctx);
    const xs = maskedNumbers(r!, mask, cols, ctx);
    return xs.length ? Math.max(...xs) : 0;
  }, { minArgs: 3 }),
  def('MINIFS', 'statistical', 'MINIFS(min_range, criteria_range1, criteria1, ...)', 'Minimum of cells matching all conditions.', ([r, ...rest], ctx) => {
    const { mask, cols } = criteriaMask(pairs(rest), ctx);
    const xs = maskedNumbers(r!, mask, cols, ctx);
    return xs.length ? Math.min(...xs) : 0;
  }, { minArgs: 3 }),
  def('SUMPRODUCT', 'math', 'SUMPRODUCT(array1, [array2], ...)', 'Sum of element-wise products.', (args, ctx) => {
    const ms = args.map((a) => toMatrix(deref(a, ctx) as EvalValue, ctx));
    const rows = ms[0]!.length;
    const cols = ms[0]![0]?.length ?? 0;
    if (ms.some((m) => m.length !== rows || (m[0]?.length ?? 0) !== cols)) return Err.VALUE;
    let total = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let p = 1;
        for (const m of ms) {
          const v = m[r]![c]!;
          if (isError(v)) return v;
          p *= typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : 0;
        }
        total += p;
      }
    }
    return total;
  }, { minArgs: 1 }),
  def('SUBTOTAL', 'math', 'SUBTOTAL(function_num, ref1, ...)', 'Aggregate of visible rows (1=AVERAGE, 9=SUM, ...; 101-111 also skip manually hidden rows).', ([fnum, ...refs], ctx) => {
    const raw = Math.trunc(numArg(fnum, ctx));
    const agg = AGGREGATES[raw % 100];
    if (!agg || (raw > 11 && raw < 101) || raw > 111) return Err.VALUE;
    const visible = withoutHiddenRows(refs, ctx, raw > 100);
    return agg(numbersOf(visible, ctx), [...scalarsOf(visible, ctx)]);
  }, { minArgs: 2, volatile: true }),

  def('AGGREGATE', 'math', 'AGGREGATE(function_num, options, ref1, [ref2 | k])', 'Aggregate that can ignore errors (options 2, 3, 6, 7); 1=AVERAGE ... 19=QUARTILE.EXC.', ([fnum, opt, ...refs], ctx) => {
    const code = Math.trunc(numArg(fnum, ctx));
    const options = Math.trunc(numArg(opt, ctx, 0));
    if (code < 1 || code > 19 || options < 0 || options > 7) return Err.VALUE;
    const skipErrors = (options & 2) !== 0;
    const withK = code >= 14;
    if (withK && refs.length !== 2) return Err.VALUE;
    const data = (options & 1) ? withoutHiddenRows(withK ? refs.slice(0, 1) : refs, ctx, true) : withK ? refs.slice(0, 1) : refs;
    const xs: number[] = [];
    const all: Scalar[] = [];
    for (const v of scalarsOf(data, ctx)) {
      if (isError(v)) { if (skipErrors) continue; return v; }
      all.push(v);
      if (typeof v === 'number') xs.push(v);
    }
    if (code <= 11) return AGGREGATES[code]!(xs, all);
    if (code === 12) return median(xs);
    if (code === 13) return modeOf(xs);
    const k = numArg(refs[1], ctx);
    const sorted = [...xs].sort((a, b) => a - b);
    switch (code) {
      case 14: case 15: {
        const i = Math.ceil(k);
        if (i < 1 || i > sorted.length) return Err.NUM;
        return code === 14 ? sorted[sorted.length - i]! : sorted[i - 1]!;
      }
      case 16: return percentileInc(xs, k);
      case 17: return Number.isInteger(k) && k >= 0 && k <= 4 ? percentileInc(xs, k / 4) : Err.NUM;
      case 18: return percentileExc(xs, k);
      default: return Number.isInteger(k) && k >= 1 && k <= 3 ? percentileExc(xs, k / 4) : Err.NUM;
    }
  }, { minArgs: 3, volatile: true }),

  // Statistics
  def('MEDIAN', 'statistical', 'MEDIAN(number1, [number2], ...)', 'Median.', (args, ctx) => median(numbersOf(args, ctx)), { minArgs: 1 }),
  def('MODE', 'statistical', 'MODE(number1, [number2], ...)', 'Most frequent value.', (args, ctx) => modeOf(numbersOf(args, ctx)), { minArgs: 1 }),
  ...(['STDEV', 'STDEV.S'] as const).map((n) => def(n, 'statistical', `${n}(number1, ...)`, 'Sample standard deviation.', (args, ctx) => Math.sqrt(variance(numbersOf(args, ctx), true)), { minArgs: 1 })),
  ...(['STDEVP', 'STDEV.P'] as const).map((n) => def(n, 'statistical', `${n}(number1, ...)`, 'Population standard deviation.', (args, ctx) => Math.sqrt(variance(numbersOf(args, ctx), false)), { minArgs: 1 })),
  ...(['VAR', 'VAR.S'] as const).map((n) => def(n, 'statistical', `${n}(number1, ...)`, 'Sample variance.', (args, ctx) => variance(numbersOf(args, ctx), true), { minArgs: 1 })),
  ...(['VARP', 'VAR.P'] as const).map((n) => def(n, 'statistical', `${n}(number1, ...)`, 'Population variance.', (args, ctx) => variance(numbersOf(args, ctx), false), { minArgs: 1 })),
  def('LARGE', 'statistical', 'LARGE(array, k)', 'k-th largest value.', ([a, k], ctx) => {
    const xs = numbersOf([a!], ctx).sort((x, y) => y - x);
    const i = Math.ceil(numArg(k, ctx));
    return i >= 1 && i <= xs.length ? xs[i - 1]! : Err.NUM;
  }, { minArgs: 2, maxArgs: 2 }),
  def('SMALL', 'statistical', 'SMALL(array, k)', 'k-th smallest value.', ([a, k], ctx) => {
    const xs = numbersOf([a!], ctx).sort((x, y) => x - y);
    const i = Math.ceil(numArg(k, ctx));
    return i >= 1 && i <= xs.length ? xs[i - 1]! : Err.NUM;
  }, { minArgs: 2, maxArgs: 2 }),
  ...(['RANK', 'RANK.EQ'] as const).map((n) =>
    def(n, 'statistical', `${n}(number, ref, [order])`, 'Rank of a number in a list.', ([x, ref, order], ctx) => {
      const v = numArg(x, ctx);
      const xs = numbersOf([ref!], ctx);
      const asc = order !== undefined && numArg(order, ctx, 0) !== 0;
      if (!xs.includes(v)) return Err.NA;
      return 1 + xs.filter((y) => (asc ? y < v : y > v)).length;
    }, { minArgs: 2, maxArgs: 3 }),
  ),
  ...(['PERCENTILE', 'PERCENTILE.INC'] as const).map((n) =>
    def(n, 'statistical', `${n}(array, k)`, 'k-th percentile (0..1).', ([a, k], ctx) => percentileInc(numbersOf([a!], ctx), numArg(k, ctx)), { minArgs: 2, maxArgs: 2 }),
  ),
  ...(['QUARTILE', 'QUARTILE.INC'] as const).map((n) =>
    def(n, 'statistical', `${n}(array, quart)`, 'Quartile (0..4).', ([a, q], ctx) => {
      const qq = Math.trunc(numArg(q, ctx));
      if (qq < 0 || qq > 4) return Err.NUM;
      return percentileInc(numbersOf([a!], ctx), qq / 4);
    }, { minArgs: 2, maxArgs: 2 }),
  ),
  def('CORREL', 'statistical', 'CORREL(array1, array2)', 'Correlation coefficient.', ([a, b], ctx) => {
    const xs = toMatrix(a!, ctx).flat();
    const ys = toMatrix(b!, ctx).flat();
    if (xs.length !== ys.length) return Err.NA;
    const px: number[] = [];
    const py: number[] = [];
    xs.forEach((x, i) => {
      const y = ys[i];
      if (typeof x === 'number' && typeof y === 'number') {
        px.push(x);
        py.push(y);
      }
    });
    if (px.length < 2) return Err.DIV0;
    const mx = mean(px);
    const my = mean(py);
    let sxy = 0, sxx = 0, syy = 0;
    px.forEach((x, i) => {
      sxy += (x - mx) * (py[i]! - my);
      sxx += (x - mx) ** 2;
      syy += (py[i]! - my) ** 2;
    });
    return sxx === 0 || syy === 0 ? Err.DIV0 : sxy / Math.sqrt(sxx * syy);
  }, { minArgs: 2, maxArgs: 2 }),
  def('AVERAGEA', 'statistical', 'AVERAGEA(value1, ...)', 'Average counting text as 0 and TRUE as 1.', (args, ctx) => {
    const xs: number[] = [];
    for (const v of scalarsOf(args, ctx)) {
      if (v === null) continue;
      if (isError(v)) return v;
      xs.push(typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : 0);
    }
    return mean(xs);
  }, { minArgs: 1 }),
  def('SEQUENCE', 'math', 'SEQUENCE(rows, [columns], [start], [step])', 'Array of sequential numbers.', ([r, c, s, st], ctx) => {
    const rows = numArg(r, ctx);
    const cols = numArg(c, ctx, 1);
    const start = numArg(s, ctx, 1);
    const step = numArg(st, ctx, 1);
    const out: Scalar[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: Scalar[] = [];
      for (let j = 0; j < cols; j++) row.push(start + (i * cols + j) * step);
      out.push(row);
    }
    return out;
  }, { minArgs: 1, maxArgs: 4 }),
];
