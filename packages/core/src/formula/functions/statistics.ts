import { Err, isError, type CellError, type Scalar } from '../../values.js';
import { broadcast, checkNumber, deref, isMatrix, isRef, toMatrix, toNumber, type EvalContext, type EvalValue, type FunctionContext, type Matrix } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { MAX_ARRAY_CELLS } from './array.js';
import { bail, boolArg, def, num, numArg, numbersOf } from './helpers.js';
import { mean, modeOf, percentileExc, sum, variance } from './math.js';

/** Numeric function of several arguments, broadcast over arrays; omitted optional arguments are null. */
function numeric(name: string, syntax: string, description: string, minArgs: number, maxArgs: number, f: (...xs: (number | null)[]) => number | CellError): FunctionDef {
  return def(name, 'statistical', syntax, description, (args, ctx) =>
    broadcast(Array.from({ length: maxArgs }, (_, i) => (args[i] === undefined ? null : deref(args[i]!, ctx))), (...values) => {
      const xs: (number | null)[] = [];
      for (let i = 0; i < values.length; i++) {
        if (values[i] === null && i >= minArgs) { xs.push(null); continue; }
        const n = toNumber(values[i]!);
        if (isError(n)) return n;
        xs.push(n);
      }
      const r = f(...xs);
      return isError(r) ? r : checkNumber(r);
    }),
    { minArgs, maxArgs },
  );
}

/** Values for MAXA/MINA: inside ranges text counts as 0 and logical values as 0/1. */
function valuesA(args: EvalValue[], ctx: EvalContext): number[] {
  const out: number[] = [];
  for (const a of args) {
    const v = deref(a, ctx);
    if (isMatrix(v) || isRef(a)) {
      for (const s of toMatrix(v, ctx).flat()) {
        if (isError(s)) bail(s);
        if (s !== null) out.push(typeof s === 'number' ? s : typeof s === 'boolean' ? Number(s) : 0);
      }
    } else if (v !== null) out.push(num(v));
  }
  return out;
}

/** Positions where both arrays hold numbers; arrays of different size are #N/A. */
function paired(a: EvalValue, b: EvalValue, ctx: EvalContext): { xs: number[]; ys: number[] } {
  const as = toMatrix(a, ctx).flat();
  const bs = toMatrix(b, ctx).flat();
  if (as.length !== bs.length) bail(Err.NA);
  const xs: number[] = [];
  const ys: number[] = [];
  as.forEach((x, i) => {
    const y = bs[i]!;
    if (isError(x)) bail(x);
    if (isError(y)) bail(y);
    if (typeof x === 'number' && typeof y === 'number') { xs.push(x); ys.push(y); }
  });
  return { xs, ys };
}

/** Sums of squares and products around the means of paired data (x = first array). */
function moments(xs: number[], ys: number[]): { n: number; mx: number; my: number; sxx: number; syy: number; sxy: number } {
  const n = xs.length;
  if (!n) bail(Err.DIV0);
  const mx = mean(xs), my = mean(ys);
  let sxx = 0, syy = 0, sxy = 0;
  xs.forEach((x, i) => {
    sxx += (x - mx) ** 2;
    syy += (ys[i]! - my) ** 2;
    sxy += (x - mx) * (ys[i]! - my);
  });
  return { n, mx, my, sxx, syy, sxy };
}

/** Least-squares fit for SLOPE, INTERCEPT and FORECAST: (known_y, known_x). */
function line(knownY: EvalValue, knownX: EvalValue, ctx: EvalContext): { slope: number; intercept: number } {
  const { xs, ys } = paired(knownY, knownX, ctx);
  const { mx, my, sxx, sxy } = moments(ys, xs);
  if (sxx === 0) bail(Err.DIV0);
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx };
}

// Normal distribution. erfc is evaluated with a series near zero and a continued fraction
// in the tails, both accurate to about 1e-15.
function erfc(x: number): number {
  if (x < 0) return 2 - erfc(-x);
  if (x < 2.5) {
    let term = x, total = x;
    for (let n = 1; n < 200 && Math.abs(term) > 1e-17 * total; n++) total += (term *= (2 * x * x) / (2 * n + 1));
    return 1 - (2 / Math.sqrt(Math.PI)) * Math.exp(-x * x) * total;
  }
  let f = x;
  for (let n = 100; n >= 1; n--) f = x + n / 2 / f;
  return Math.exp(-x * x) / Math.sqrt(Math.PI) / f;
}

const normCdf = (z: number) => 0.5 * erfc(-z / Math.SQRT2);
const normPdf = (z: number) => Math.exp(-z * z / 2) / Math.sqrt(2 * Math.PI);

/** Inverse of the standard normal distribution: Acklam's approximation refined by Halley steps. */
function normInv(p: number): number | CellError {
  if (!(p > 0 && p < 1)) return Err.NUM;
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const tail = (q: number) => (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) / ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1);
  let x: number;
  if (p < 0.02425) x = tail(Math.sqrt(-2 * Math.log(p)));
  else if (p > 1 - 0.02425) x = -tail(Math.sqrt(-2 * Math.log(1 - p)));
  else {
    const q = p - 0.5, r = q * q;
    x = (((((a[0]! * r + a[1]!) * r + a[2]!) * r + a[3]!) * r + a[4]!) * r + a[5]!) * q / (((((b[0]! * r + b[1]!) * r + b[2]!) * r + b[3]!) * r + b[4]!) * r + 1);
  }
  for (let i = 0; i < 3; i++) {
    const u = (normCdf(x) - p) * Math.sqrt(2 * Math.PI) * Math.exp(x * x / 2);
    x -= u / (1 + x * u / 2);
  }
  return x;
}

/** Solves A·x = b with Gaussian elimination and partial pivoting; singular systems are #NUM!. */
function solve(A: number[][], b: number[]): number[] {
  const n = b.length;
  const m = A.map((row, i) => [...row, b[i]!]);
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(m[r]![col]!) > Math.abs(m[pivot]![col]!)) pivot = r;
    if (Math.abs(m[pivot]![col]!) < 1e-12) bail(Err.NUM);
    [m[col], m[pivot]] = [m[pivot]!, m[col]!];
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const factor = m[r]![col]! / m[col]![col]!;
      for (let c = col; c <= n; c++) m[r]![c]! -= factor * m[col]![c]!;
    }
  }
  return m.map((row, i) => row[n]! / row[i]!);
}

function numbersIn(m: Matrix): number[] {
  return m.flat().map((v) => {
    if (isError(v)) bail(v);
    if (typeof v !== 'number') bail(Err.VALUE);
    return v;
  });
}

const absent = (v: EvalValue | undefined) => v === undefined || v === null;

/**
 * TREND and GROWTH: least squares on one or more x variables. A y column takes the
 * variables from the columns of known_x, a y row from its rows.
 */
function trend(args: EvalValue[], ctx: FunctionContext, growth: boolean): Matrix {
  const [knownY, knownX, newX, constant] = args;
  const yMatrix = toMatrix(knownY!, ctx);
  const rows = yMatrix.length, cols = yMatrix[0]?.length ?? 0;
  let ys = numbersIn(yMatrix);
  if (growth) ys = ys.map((y) => (y > 0 ? Math.log(y) : bail(Err.NUM)));
  const n = ys.length;
  const xMatrix = absent(knownX) ? yMatrix.map((row, r) => row.map((_, c) => r * cols + c + 1)) : toMatrix(knownX!, ctx);
  const xRows = xMatrix.length, xCols = xMatrix[0]?.length ?? 0;
  numbersIn(xMatrix);
  // observations[i][j] = value of variable j in observation i
  let observations: number[][];
  let shape: 'same' | 'columns' | 'rows';
  if (xRows === rows && xCols === cols) { shape = 'same'; observations = numbersIn(xMatrix).map((x) => [x]); }
  else if (cols === 1 && xRows === rows) { shape = 'columns'; observations = xMatrix.map((row) => row as number[]); }
  else if (rows === 1 && xCols === cols) { shape = 'rows'; observations = Array.from({ length: cols }, (_, c) => xMatrix.map((row) => row[c] as number)); }
  else return bail(Err.REF);
  const k = observations[0]!.length;
  const withConst = absent(constant) || boolArg(constant, ctx, true);
  const design = observations.map((x) => (withConst ? [1, ...x] : x));
  const p = design[0]!.length;
  if (n < p) bail(Err.NUM);
  const XtX = Array.from({ length: p }, (_, i) => Array.from({ length: p }, (_, j) => design.reduce((a, row) => a + row[i]! * row[j]!, 0)));
  const Xty = Array.from({ length: p }, (_, i) => design.reduce((a, row, r) => a + row[i]! * ys[r]!, 0));
  const coef = solve(XtX, Xty);
  const predict = (x: number[]) => {
    const v = (withConst ? coef[0]! : 0) + x.reduce((a, xi, j) => a + xi * coef[j + (withConst ? 1 : 0)]!, 0);
    return checkNumber(growth ? Math.exp(v) : v);
  };
  const target = absent(newX) ? xMatrix : toMatrix(newX!, ctx);
  numbersIn(target);
  if (shape === 'same') return target.map((row) => row.map((x) => predict([x as number])));
  if (shape === 'columns') {
    if ((target[0]?.length ?? 0) !== k) bail(Err.REF);
    return target.map((row) => [predict(row as number[])]);
  }
  if (target.length !== k) bail(Err.REF);
  return [Array.from({ length: target[0]?.length ?? 0 }, (_, c) => predict(target.map((row) => row[c] as number)))];
}

/** Position of x in sorted data, interpolated between neighbours; outside the data #N/A. */
function rankPosition(sorted: number[], x: number): number {
  const n = sorted.length;
  if (!n || x < sorted[0]! || x > sorted[n - 1]!) bail(Err.NA);
  const exact = sorted.indexOf(x);
  if (exact >= 0) return exact;
  let i = 0;
  while (sorted[i + 1]! < x) i++;
  return i + (x - sorted[i]!) / (sorted[i + 1]! - sorted[i]!);
}

function percentRank(name: string, exclusive: boolean): FunctionDef {
  return def(name, 'statistical', `${name}(array, x, [significance])`, `Rank of a value as a percentage of the data set (${exclusive ? 'exclusive' : 'inclusive'}).`, ([a, x, s], ctx) => {
    const sorted = numbersOf([a!], ctx).sort((p, q) => p - q);
    const digits = Math.trunc(numArg(s, ctx, 3));
    if (digits < 1) return Err.NUM;
    const n = sorted.length;
    if (!n) return Err.NUM;
    const pos = rankPosition(sorted, numArg(x, ctx));
    const rank = exclusive ? (pos + 1) / (n + 1) : n === 1 ? 1 : pos / (n - 1);
    const f = 10 ** digits;
    return Math.floor(Number((rank * f).toPrecision(15))) / f;
  }, { minArgs: 2, maxArgs: 3 });
}

function covariance(name: string, sample: boolean): FunctionDef {
  return def(name, 'statistical', `${name}(array1, array2)`, `${sample ? 'Sample' : 'Population'} covariance.`, ([a, b], ctx) => {
    const { xs, ys } = paired(a!, b!, ctx);
    const { n, sxy } = moments(xs, ys);
    if (sample && n < 2) return Err.DIV0;
    return sxy / (sample ? n - 1 : n);
  }, { minArgs: 2, maxArgs: 2 });
}

const positive = (xs: number[]) => (xs.length ? xs.map((x) => (x > 0 ? x : bail(Err.NUM))) : bail(Err.NUM));

function standardized(xs: number[], minimum: number): { n: number; z: number[] } {
  const n = xs.length;
  if (n < minimum) bail(Err.DIV0);
  const s = Math.sqrt(variance(xs, true));
  if (s === 0) bail(Err.DIV0);
  const m = mean(xs);
  return { n, z: xs.map((x) => (x - m) / s) };
}

const normDist = (x: number, m: number, s: number, cumulative: number) =>
  s <= 0 ? Err.NUM : cumulative ? normCdf((x - m) / s) : normPdf((x - m) / s) / s;
const normInverse = (p: number, m: number, s: number) => {
  if (s <= 0) return Err.NUM;
  const z = normInv(p);
  return isError(z) ? z : m + s * z;
};

export const statisticsFunctions: FunctionDef[] = [
  def('MAXA', 'statistical', 'MAXA(value1, [value2], ...)', 'Largest value; text in ranges counts as 0, TRUE as 1.', (args, ctx) => {
    const xs = valuesA(args, ctx);
    return xs.length ? Math.max(...xs) : 0;
  }, { minArgs: 1 }),
  def('MINA', 'statistical', 'MINA(value1, [value2], ...)', 'Smallest value; text in ranges counts as 0, TRUE as 1.', (args, ctx) => {
    const xs = valuesA(args, ctx);
    return xs.length ? Math.min(...xs) : 0;
  }, { minArgs: 1 }),
  def('GEOMEAN', 'statistical', 'GEOMEAN(number1, [number2], ...)', 'Geometric mean of positive numbers.', (args, ctx) => {
    const xs = positive(numbersOf(args, ctx));
    return Math.exp(sum(xs.map(Math.log)) / xs.length);
  }, { minArgs: 1 }),
  def('HARMEAN', 'statistical', 'HARMEAN(number1, [number2], ...)', 'Harmonic mean of positive numbers.', (args, ctx) => {
    const xs = positive(numbersOf(args, ctx));
    return xs.length / sum(xs.map((x) => 1 / x));
  }, { minArgs: 1 }),
  def('AVEDEV', 'statistical', 'AVEDEV(number1, [number2], ...)', 'Average absolute deviation from the mean.', (args, ctx) => {
    const xs = numbersOf(args, ctx);
    if (!xs.length) return Err.NUM;
    const m = mean(xs);
    return mean(xs.map((x) => Math.abs(x - m)));
  }, { minArgs: 1 }),
  def('DEVSQ', 'statistical', 'DEVSQ(number1, [number2], ...)', 'Sum of squared deviations from the mean.', (args, ctx) => {
    const xs = numbersOf(args, ctx);
    if (!xs.length) return Err.NUM;
    const m = mean(xs);
    return sum(xs.map((x) => (x - m) ** 2));
  }, { minArgs: 1 }),
  def('SKEW', 'statistical', 'SKEW(number1, [number2], ...)', 'Skewness of a distribution (sample).', (args, ctx) => {
    const { n, z } = standardized(numbersOf(args, ctx), 3);
    return (n / ((n - 1) * (n - 2))) * sum(z.map((v) => v ** 3));
  }, { minArgs: 1 }),
  def('KURT', 'statistical', 'KURT(number1, [number2], ...)', 'Excess kurtosis of a distribution (sample).', (args, ctx) => {
    const { n, z } = standardized(numbersOf(args, ctx), 4);
    return (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3)) * sum(z.map((v) => v ** 4)) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  }, { minArgs: 1 }),
  def('MODE.SNGL', 'statistical', 'MODE.SNGL(number1, [number2], ...)', 'Most frequent value.', (args, ctx) => modeOf(numbersOf(args, ctx)), { minArgs: 1 }),
  def('MODE.MULT', 'statistical', 'MODE.MULT(number1, [number2], ...)', 'Vertical array of all most frequent values.', (args, ctx) => {
    const counts = new Map<number, number>();
    for (const x of numbersOf(args, ctx)) counts.set(x, (counts.get(x) ?? 0) + 1);
    const best = Math.max(0, ...counts.values());
    if (best < 2) return Err.NA;
    return [...counts].filter(([, c]) => c === best).map(([x]) => [x]);
  }, { minArgs: 1 }),
  def('RANK.AVG', 'statistical', 'RANK.AVG(number, ref, [order])', 'Rank of a number; ties get their average rank.', ([x, ref, order], ctx) => {
    const v = numArg(x, ctx);
    const xs = numbersOf([ref!], ctx);
    const asc = numArg(order, ctx, 0) !== 0;
    const ties = xs.filter((y) => y === v).length;
    if (!ties) return Err.NA;
    return 1 + xs.filter((y) => (asc ? y < v : y > v)).length + (ties - 1) / 2;
  }, { minArgs: 2, maxArgs: 3 }),
  def('PERCENTILE.EXC', 'statistical', 'PERCENTILE.EXC(array, k)', 'k-th percentile, k strictly between 0 and 1.', ([a, k], ctx) => percentileExc(numbersOf([a!], ctx), numArg(k, ctx)), { minArgs: 2, maxArgs: 2 }),
  def('QUARTILE.EXC', 'statistical', 'QUARTILE.EXC(array, quart)', 'Quartile (1..3) based on exclusive percentiles.', ([a, q], ctx) => {
    const qq = Math.trunc(numArg(q, ctx));
    return qq < 1 || qq > 3 ? Err.NUM : percentileExc(numbersOf([a!], ctx), qq / 4);
  }, { minArgs: 2, maxArgs: 2 }),
  percentRank('PERCENTRANK', false),
  percentRank('PERCENTRANK.INC', false),
  percentRank('PERCENTRANK.EXC', true),
  def('SLOPE', 'statistical', 'SLOPE(known_y, known_x)', 'Slope of the linear regression line.', ([y, x], ctx) => line(y!, x!, ctx).slope, { minArgs: 2, maxArgs: 2 }),
  def('INTERCEPT', 'statistical', 'INTERCEPT(known_y, known_x)', 'Intercept of the linear regression line.', ([y, x], ctx) => line(y!, x!, ctx).intercept, { minArgs: 2, maxArgs: 2 }),
  ...(['FORECAST', 'FORECAST.LINEAR'] as const).map((n) =>
    def(n, 'statistical', `${n}(x, known_y, known_x)`, 'Value on the linear regression line at x.', ([x, y, xs], ctx) => {
      const at = numArg(x, ctx);
      const { slope, intercept } = line(y!, xs!, ctx);
      return intercept + slope * at;
    }, { minArgs: 3, maxArgs: 3 }),
  ),
  def('RSQ', 'statistical', 'RSQ(known_y, known_x)', 'Square of the correlation coefficient.', ([y, x], ctx) => {
    const { xs, ys } = paired(y!, x!, ctx);
    const { sxx, syy, sxy } = moments(xs, ys);
    return sxx === 0 || syy === 0 ? Err.DIV0 : (sxy * sxy) / (sxx * syy);
  }, { minArgs: 2, maxArgs: 2 }),
  def('PEARSON', 'statistical', 'PEARSON(array1, array2)', 'Pearson correlation coefficient.', ([a, b], ctx) => {
    const { xs, ys } = paired(a!, b!, ctx);
    const { sxx, syy, sxy } = moments(xs, ys);
    return sxx === 0 || syy === 0 ? Err.DIV0 : sxy / Math.sqrt(sxx * syy);
  }, { minArgs: 2, maxArgs: 2 }),
  def('STEYX', 'statistical', 'STEYX(known_y, known_x)', 'Standard error of the predicted y values.', ([y, x], ctx) => {
    const { xs, ys } = paired(y!, x!, ctx);
    const { n, sxx, syy, sxy } = moments(ys, xs);
    if (n < 3 || sxx === 0) return Err.DIV0;
    return Math.sqrt((syy - (sxy * sxy) / sxx) / (n - 2));
  }, { minArgs: 2, maxArgs: 2 }),
  covariance('COVAR', false),
  covariance('COVARIANCE.P', false),
  covariance('COVARIANCE.S', true),
  def('TREND', 'statistical', 'TREND(known_y, [known_x], [new_x], [const])', 'Values on the least-squares line (or plane) through the data.', (args, ctx) => trend(args, ctx, false), { minArgs: 1, maxArgs: 4 }),
  def('GROWTH', 'statistical', 'GROWTH(known_y, [known_x], [new_x], [const])', 'Values on the exponential growth curve through the data.', (args, ctx) => trend(args, ctx, true), { minArgs: 1, maxArgs: 4 }),
  def('FREQUENCY', 'statistical', 'FREQUENCY(data_array, bins_array)', 'Vertical array counting the values in each bin, plus one for values above the last bin.', ([data, bins], ctx) => {
    const xs = numbersOf([data!], ctx);
    const edges = numbersOf([bins!], ctx);
    const order = edges.map((_, i) => i).sort((a, b) => edges[a]! - edges[b]! || a - b);
    const counts = new Array<number>(edges.length + 1).fill(0);
    for (const x of xs) {
      const slot = order.find((i) => x <= edges[i]!);
      counts[slot ?? edges.length]!++;
    }
    return counts.map((c) => [c]);
  }, { minArgs: 2, maxArgs: 2 }),
  def('RANDARRAY', 'math', 'RANDARRAY([rows], [columns], [min], [max], [integer])', 'Array of random numbers.', ([r, c, lo, hi, int], ctx) => {
    const rows = Math.trunc(numArg(r, ctx, 1));
    const cols = Math.trunc(numArg(c, ctx, 1));
    const min = numArg(lo, ctx, 0);
    const max = numArg(hi, ctx, 1);
    const integer = boolArg(int, ctx, false);
    if (rows < 1 || cols < 1 || min > max) return Err.VALUE;
    if (rows * cols > MAX_ARRAY_CELLS) return Err.NUM;
    if (integer && (!Number.isInteger(min) || !Number.isInteger(max))) return Err.VALUE;
    const draw = integer ? () => min + Math.floor(Math.random() * (max - min + 1)) : () => min + Math.random() * (max - min);
    return Array.from({ length: rows }, () => Array.from({ length: cols }, draw));
  }, { maxArgs: 5, volatile: true }),

  numeric('NORM.DIST', 'NORM.DIST(x, mean, standard_dev, cumulative)', 'Normal distribution (density or cumulative).', 4, 4, (x, m, s, c) => normDist(x!, m!, s!, c!)),
  numeric('NORMDIST', 'NORMDIST(x, mean, standard_dev, cumulative)', 'Normal distribution (density or cumulative).', 4, 4, (x, m, s, c) => normDist(x!, m!, s!, c!)),
  numeric('NORM.S.DIST', 'NORM.S.DIST(z, cumulative)', 'Standard normal distribution (density or cumulative).', 2, 2, (z, c) => normDist(z!, 0, 1, c!)),
  numeric('NORMSDIST', 'NORMSDIST(z)', 'Cumulative standard normal distribution.', 1, 1, (z) => normCdf(z!)),
  numeric('NORM.INV', 'NORM.INV(probability, mean, standard_dev)', 'Inverse of the cumulative normal distribution.', 3, 3, (p, m, s) => normInverse(p!, m!, s!)),
  numeric('NORMINV', 'NORMINV(probability, mean, standard_dev)', 'Inverse of the cumulative normal distribution.', 3, 3, (p, m, s) => normInverse(p!, m!, s!)),
  numeric('NORM.S.INV', 'NORM.S.INV(probability)', 'Inverse of the cumulative standard normal distribution.', 1, 1, (p) => normInv(p!)),
  numeric('NORMSINV', 'NORMSINV(probability)', 'Inverse of the cumulative standard normal distribution.', 1, 1, (p) => normInv(p!)),
  numeric('STANDARDIZE', 'STANDARDIZE(x, mean, standard_dev)', 'Normalized value (x - mean) / standard_dev.', 3, 3, (x, m, s) => (s! <= 0 ? Err.NUM : (x! - m!) / s!)),
];
