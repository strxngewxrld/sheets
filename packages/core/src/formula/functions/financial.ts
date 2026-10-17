import { Err, isError } from '../../values.js';
import { checkNumber, toMatrix, type EvalValue, type FunctionContext } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { bail, def, numArg, scalarsOf } from './helpers.js';

function paymentType(arg: EvalValue | undefined, ctx: FunctionContext): number {
  const type = numArg(arg, ctx, 0);
  if (type !== 0 && type !== 1) bail(Err.NUM);
  return type;
}

/** Growth and accumulated payments; expm1 avoids cancellation near zero interest. */
function factors(rate: number, periods: number): { growth: number; annuity: number } {
  if (rate === 0) return { growth: 1, annuity: periods };
  if (rate > -1) {
    const exponent = periods * Math.log1p(rate);
    return { growth: Math.exp(exponent), annuity: Math.expm1(exponent) / rate };
  }
  const growth = Math.pow(1 + rate, periods);
  return { growth, annuity: (growth - 1) / rate };
}

function payment(rate: number, periods: number, present: number, future: number, type: number): number {
  const { growth, annuity } = factors(rate, periods);
  const divisor = (1 + rate * type) * annuity;
  if (divisor === 0) bail(Err.DIV0);
  return -(present * growth + future) / divisor;
}

function futureValue(rate: number, periods: number, pmt: number, present: number, type: number): number {
  const { growth, annuity } = factors(rate, periods);
  return -(present * growth + pmt * (1 + rate * type) * annuity);
}

function periodPayment(name: 'IPMT' | 'PPMT'): FunctionDef {
  return def(name, 'financial', `${name}(rate, per, nper, pv, [fv], [type])`,
    name === 'IPMT' ? 'Interest portion of a payment for a given period.' : 'Principal portion of a payment for a given period.',
    ([r, p, n, pv, fv, t], ctx) => {
      const rate = numArg(r, ctx);
      const per = numArg(p, ctx);
      const periods = numArg(n, ctx);
      const present = numArg(pv, ctx);
      const future = numArg(fv, ctx, 0);
      const type = paymentType(t, ctx);
      if (per < 1 || per > periods) return Err.NUM;
      const pmt = payment(rate, periods, present, future, type);
      let interest = 0;
      if (rate !== 0 && !(type === 1 && per === 1)) {
        interest = futureValue(rate, per - 1, pmt, present, type) * rate;
        if (type === 1) interest /= 1 + rate;
      }
      return checkNumber(name === 'IPMT' ? interest : pmt - interest);
    }, { minArgs: 4, maxArgs: 6 });
}

/**
 * Finds a root of f near `guess` with Newton's method and falls back to
 * bisection over a bracket when Newton does not converge (Excel's IRR/RATE
 * likewise iterate to 1e-7 within a bounded number of steps).
 */
function solveRate(f: (r: number) => number, df: (r: number) => number, guess: number): number {
  let r = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(r);
    if (!Number.isFinite(y)) break;
    if (Math.abs(y) < 1e-10) return r;
    const d = df(r);
    if (!d || !Number.isFinite(d)) break;
    const next = r - y / d;
    if (next <= -1) break;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  // Bisection: look for a sign change between -0.99999 and growing upper bounds.
  let lo = -0.99999999;
  let flo = f(lo);
  for (let hi = 0.1; hi <= 1e6; hi *= 2) {
    const fhi = f(hi);
    if (Number.isFinite(flo) && Number.isFinite(fhi) && Math.sign(flo) !== Math.sign(fhi)) {
      for (let i = 0; i < 300; i++) {
        const mid = (lo + hi) / 2;
        const fm = f(mid);
        if (Math.abs(fm) < 1e-10 || hi - lo < 1e-15) return mid;
        if (Math.sign(fm) === Math.sign(flo)) { lo = mid; flo = fm; } else hi = mid;
      }
      return (lo + hi) / 2;
    }
  }
  return bail(Err.NUM);
}

function cashFlows(arg: EvalValue, ctx: FunctionContext): number[] {
  const out: number[] = [];
  for (const row of toMatrix(arg, ctx)) for (const v of row) {
    if (isError(v)) bail(v);
    if (typeof v === 'number') out.push(v);
  }
  return out;
}

export const financialFunctions: FunctionDef[] = [
  def('IRR', 'financial', 'IRR(values, [guess])', 'Internal rate of return of periodic cash flows.', ([v, g], ctx) => {
    const flows = cashFlows(v!, ctx);
    if (!flows.some((x) => x > 0) || !flows.some((x) => x < 0)) return Err.NUM;
    const npv = (r: number) => flows.reduce((a, x, i) => a + x / Math.pow(1 + r, i), 0);
    const dnpv = (r: number) => flows.reduce((a, x, i) => a - (i * x) / Math.pow(1 + r, i + 1), 0);
    return checkNumber(solveRate(npv, dnpv, numArg(g, ctx, 0.1)));
  }, { minArgs: 1, maxArgs: 2 }),
  def('XIRR', 'financial', 'XIRR(values, dates, [guess])', 'Internal rate of return of cash flows on specific dates.', ([v, d, g], ctx) => {
    const values = toMatrix(v!, ctx).flat();
    const dates = toMatrix(d!, ctx).flat();
    if (values.length !== dates.length || values.length < 2) return Err.NUM;
    const flows: number[] = [];
    const days: number[] = [];
    for (let i = 0; i < values.length; i++) {
      const x = values[i]!;
      const t = dates[i]!;
      if (isError(x)) return x;
      if (isError(t)) return t;
      if (typeof x !== 'number' || typeof t !== 'number') return Err.VALUE;
      if (Math.trunc(t) < Math.trunc(dates[0] as number)) return Err.NUM;
      flows.push(x);
      days.push((Math.trunc(t) - Math.trunc(dates[0] as number)) / 365);
    }
    if (!flows.some((x) => x > 0) || !flows.some((x) => x < 0)) return Err.NUM;
    const f = (r: number) => flows.reduce((a, x, i) => a + x / Math.pow(1 + r, days[i]!), 0);
    const df = (r: number) => flows.reduce((a, x, i) => a - (days[i]! * x) / Math.pow(1 + r, days[i]! + 1), 0);
    return checkNumber(solveRate(f, df, numArg(g, ctx, 0.1)));
  }, { minArgs: 2, maxArgs: 3 }),
  def('RATE', 'financial', 'RATE(nper, pmt, pv, [fv], [type], [guess])', 'Interest rate per period of an annuity.', ([n, p, pv, fv, t, g], ctx) => {
    const periods = numArg(n, ctx);
    const pmt = numArg(p, ctx);
    const present = numArg(pv, ctx);
    const future = numArg(fv, ctx, 0);
    const type = paymentType(t, ctx);
    if (periods <= 0) return Err.NUM;
    const f = (r: number) => {
      const { growth, annuity } = factors(r, periods);
      return present * growth + pmt * (1 + r * type) * annuity + future;
    };
    const df = (r: number) => (f(r + 1e-7) - f(r - 1e-7)) / 2e-7;
    return checkNumber(solveRate(f, df, numArg(g, ctx, 0.1)));
  }, { minArgs: 3, maxArgs: 6 }),
  def('PMT', 'financial', 'PMT(rate, nper, pv, [fv], [type])', 'Constant periodic payment for a loan or investment.', ([r, n, pv, fv, t], ctx) =>
    checkNumber(payment(numArg(r, ctx), numArg(n, ctx), numArg(pv, ctx), numArg(fv, ctx, 0), paymentType(t, ctx))),
  { minArgs: 3, maxArgs: 5 }),
  def('FV', 'financial', 'FV(rate, nper, pmt, [pv], [type])', 'Future value of constant periodic payments.', ([r, n, p, pv, t], ctx) =>
    checkNumber(futureValue(numArg(r, ctx), numArg(n, ctx), numArg(p, ctx), numArg(pv, ctx, 0), paymentType(t, ctx))),
  { minArgs: 3, maxArgs: 5 }),
  def('PV', 'financial', 'PV(rate, nper, pmt, [fv], [type])', 'Present value of constant periodic payments.', ([r, n, p, fv, t], ctx) => {
    const rate = numArg(r, ctx);
    const periods = numArg(n, ctx);
    const pmt = numArg(p, ctx);
    const future = numArg(fv, ctx, 0);
    const type = paymentType(t, ctx);
    const { growth, annuity } = factors(rate, periods);
    if (growth === 0) return Err.DIV0;
    return checkNumber(-(future + pmt * (1 + rate * type) * annuity) / growth);
  }, { minArgs: 3, maxArgs: 5 }),
  def('NPER', 'financial', 'NPER(rate, pmt, pv, [fv], [type])', 'Number of periods for constant payments and interest.', ([r, p, pv, fv, t], ctx) => {
    const rate = numArg(r, ctx);
    const pmt = numArg(p, ctx);
    const present = numArg(pv, ctx);
    const future = numArg(fv, ctx, 0);
    const type = paymentType(t, ctx);
    if (rate === 0) return pmt === 0 ? Err.DIV0 : checkNumber(-(present + future) / pmt);
    if (rate <= -1) return Err.NUM;
    const adjusted = pmt * (1 + rate * type);
    const denominator = adjusted + present * rate;
    const numerator = adjusted - future * rate;
    if (denominator === 0 || numerator / denominator <= 0) return Err.NUM;
    const delta = -(present + future) * rate / denominator;
    const logarithm = Math.abs(delta) < 0.5 ? Math.log1p(delta) : Math.log(numerator / denominator);
    return checkNumber(logarithm / Math.log1p(rate));
  }, { minArgs: 3, maxArgs: 5 }),
  def('NPV', 'financial', 'NPV(rate, value1, [value2], ...)', 'Net present value of cash flows at the end of each period. Add an initial cash flow separately.', ([r, ...values], ctx) => {
    const rate = numArg(r, ctx);
    if (rate === -1) return Err.DIV0;
    let total = 0;
    let period = 0;
    // Only numeric cash flows occupy a period; ignore all other value types.
    for (const flow of scalarsOf(values, ctx)) {
      if (typeof flow === 'number') total += flow / Math.pow(1 + rate, ++period);
    }
    return checkNumber(total);
  }, { minArgs: 2 }),
  periodPayment('IPMT'),
  periodPayment('PPMT'),
];
