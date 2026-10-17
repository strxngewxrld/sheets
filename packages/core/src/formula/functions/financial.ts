import { Err } from '../../values.js';
import { checkNumber, type EvalValue, type FunctionContext } from '../evaluator.js';
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

export const financialFunctions: FunctionDef[] = [
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
