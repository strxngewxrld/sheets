import { dateToSerial, jsDateToSerial, serialToParts, timeToSerial } from '../../dates.js';
import { parseInput } from '../../numfmt.js';
import { Err } from '../../values.js';
import { toMatrix } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { def, numArg, scalarArg, textArg, unary } from './helpers.js';

function dateArg(v: Parameters<typeof scalarArg>[0], ctx: Parameters<typeof scalarArg>[1]): number {
  const s = scalarArg(v, ctx);
  if (typeof s === 'string') {
    const parsed = parseInput(s);
    if (typeof parsed.value === 'number') return parsed.value;
  }
  return numArg(v, ctx);
}

function addMonths(serial: number, months: number): { y: number; m: number; d: number } {
  const p = serialToParts(serial);
  const total = p.year * 12 + (p.month - 1) + months;
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  const lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return { y, m, d: Math.min(p.day, lastDay) };
}

function isWorkday(serial: number, holidays: Set<number>): boolean {
  const wd = serialToParts(serial).weekday;
  return wd !== 0 && wd !== 6 && !holidays.has(serial);
}

export const dateFunctions: FunctionDef[] = [
  def('TODAY', 'date', 'TODAY()', "Today's date.", () => Math.floor(jsDateToSerial(new Date())), { maxArgs: 0, volatile: true }),
  def('NOW', 'date', 'NOW()', 'Current date and time.', () => jsDateToSerial(new Date()), { maxArgs: 0, volatile: true }),
  def('DATE', 'date', 'DATE(year, month, day)', 'Serial number of a date.', ([y, m, d], ctx) => {
    let year = Math.trunc(numArg(y, ctx));
    if (year < 1900) year += 1900;
    const s = dateToSerial(year, Math.trunc(numArg(m, ctx)), Math.trunc(numArg(d, ctx)));
    return s < 0 ? Err.NUM : s;
  }, { minArgs: 3, maxArgs: 3 }),
  def('TIME', 'date', 'TIME(hour, minute, second)', 'Serial number of a time.', ([h, m, s], ctx) => {
    const t = timeToSerial(numArg(h, ctx), numArg(m, ctx), numArg(s, ctx));
    return t < 0 ? Err.NUM : t - Math.floor(t);
  }, { minArgs: 3, maxArgs: 3 }),
  def('DATEVALUE', 'date', 'DATEVALUE(date_text)', 'Converts date text to a serial number.', ([t], ctx) => {
    const p = parseInput(textArg(t, ctx));
    return typeof p.value === 'number' && p.numFmt ? Math.floor(p.value) : Err.VALUE;
  }, { minArgs: 1, maxArgs: 1 }),
  def('TIMEVALUE', 'date', 'TIMEVALUE(time_text)', 'Converts time text to a serial number.', ([t], ctx) => {
    const p = parseInput(textArg(t, ctx));
    return typeof p.value === 'number' && p.numFmt ? p.value - Math.floor(p.value) : Err.VALUE;
  }, { minArgs: 1, maxArgs: 1 }),
  def('YEAR', 'date', 'YEAR(date)', 'Year of a date.', ([d], ctx) => serialToParts(dateArg(d, ctx)).year, { minArgs: 1, maxArgs: 1 }),
  def('MONTH', 'date', 'MONTH(date)', 'Month of a date (1-12).', ([d], ctx) => serialToParts(dateArg(d, ctx)).month, { minArgs: 1, maxArgs: 1 }),
  def('DAY', 'date', 'DAY(date)', 'Day of the month.', ([d], ctx) => serialToParts(dateArg(d, ctx)).day, { minArgs: 1, maxArgs: 1 }),
  unary('HOUR', 'Hour of a time value.', (x) => serialToParts(x).hours, 'date'),
  unary('MINUTE', 'Minute of a time value.', (x) => serialToParts(x).minutes, 'date'),
  unary('SECOND', 'Second of a time value.', (x) => serialToParts(x).seconds, 'date'),
  def('WEEKDAY', 'date', 'WEEKDAY(date, [return_type])', 'Day of the week (1=Sunday by default).', ([d, t], ctx) => {
    const wd = serialToParts(dateArg(d, ctx)).weekday;
    const type = numArg(t, ctx, 1);
    if (type === 1) return wd + 1;
    if (type === 2) return wd === 0 ? 7 : wd;
    if (type === 3) return wd === 0 ? 6 : wd - 1;
    return Err.NUM;
  }, { minArgs: 1, maxArgs: 2 }),
  def('WEEKNUM', 'date', 'WEEKNUM(date, [return_type])', 'Week number of the year.', ([d, t], ctx) => {
    const serial = dateArg(d, ctx);
    const p = serialToParts(serial);
    const jan1 = dateToSerial(p.year, 1, 1);
    const startDay = numArg(t, ctx, 1) === 2 ? 1 : 0;
    const offset = (serialToParts(jan1).weekday - startDay + 7) % 7;
    return Math.floor((serial - jan1 + offset) / 7) + 1;
  }, { minArgs: 1, maxArgs: 2 }),
  def('ISOWEEKNUM', 'date', 'ISOWEEKNUM(date)', 'ISO 8601 week number.', ([d], ctx) => {
    const p = serialToParts(dateArg(d, ctx));
    const date = new Date(Date.UTC(p.year, p.month - 1, p.day));
    const day = date.getUTCDay() || 7;
    date.setUTCDate(date.getUTCDate() + 4 - day);
    const yearStart = Date.UTC(date.getUTCFullYear(), 0, 1);
    return Math.ceil(((date.getTime() - yearStart) / 86_400_000 + 1) / 7);
  }, { minArgs: 1, maxArgs: 1 }),
  def('EDATE', 'date', 'EDATE(start_date, months)', 'Date a number of months before/after.', ([d, m], ctx) => {
    const r = addMonths(dateArg(d, ctx), Math.trunc(numArg(m, ctx)));
    return dateToSerial(r.y, r.m, r.d);
  }, { minArgs: 2, maxArgs: 2 }),
  def('EOMONTH', 'date', 'EOMONTH(start_date, months)', 'Last day of the month, months before/after.', ([d, m], ctx) => {
    const r = addMonths(dateArg(d, ctx), Math.trunc(numArg(m, ctx)));
    return dateToSerial(r.y, r.m + 1, 0);
  }, { minArgs: 2, maxArgs: 2 }),
  def('DAYS', 'date', 'DAYS(end_date, start_date)', 'Days between two dates.', ([e, s], ctx) => Math.trunc(dateArg(e, ctx)) - Math.trunc(dateArg(s, ctx)), { minArgs: 2, maxArgs: 2 }),
  def('DATEDIF', 'date', 'DATEDIF(start_date, end_date, unit)', 'Difference between dates in "Y", "M", "D", "MD", "YM", "YD".', ([s, e, u], ctx) => {
    const start = Math.trunc(dateArg(s, ctx));
    const end = Math.trunc(dateArg(e, ctx));
    if (start > end) return Err.NUM;
    const a = serialToParts(start);
    const b = serialToParts(end);
    let months = (b.year - a.year) * 12 + (b.month - a.month);
    if (b.day < a.day) months--;
    switch (textArg(u, ctx).toUpperCase()) {
      case 'Y': return Math.floor(months / 12);
      case 'M': return months;
      case 'D': return end - start;
      case 'YM': return months % 12;
      case 'MD': {
        const r = addMonths(start, months);
        return end - dateToSerial(r.y, r.m, r.d);
      }
      case 'YD': {
        const r = addMonths(start, Math.floor(months / 12) * 12);
        return end - dateToSerial(r.y, r.m, r.d);
      }
      default: return Err.NUM;
    }
  }, { minArgs: 3, maxArgs: 3 }),
  def('NETWORKDAYS', 'date', 'NETWORKDAYS(start_date, end_date, [holidays])', 'Working days between two dates.', ([s, e, h], ctx) => {
    let start = Math.trunc(dateArg(s, ctx));
    let end = Math.trunc(dateArg(e, ctx));
    const sign = start <= end ? 1 : -1;
    if (sign < 0) [start, end] = [end, start];
    const holidays = new Set<number>();
    if (h) for (const v of toMatrix(h, ctx).flat()) if (typeof v === 'number') holidays.add(Math.trunc(v));
    let n = 0;
    for (let d = start; d <= end; d++) if (isWorkday(d, holidays)) n++;
    return n * sign;
  }, { minArgs: 2, maxArgs: 3 }),
  def('WORKDAY', 'date', 'WORKDAY(start_date, days, [holidays])', 'Date after a number of working days.', ([s, n, h], ctx) => {
    let d = Math.trunc(dateArg(s, ctx));
    let days = Math.trunc(numArg(n, ctx));
    const holidays = new Set<number>();
    if (h) for (const v of toMatrix(h, ctx).flat()) if (typeof v === 'number') holidays.add(Math.trunc(v));
    const step = days >= 0 ? 1 : -1;
    while (days !== 0) {
      d += step;
      if (isWorkday(d, holidays)) days -= step;
    }
    return d;
  }, { minArgs: 2, maxArgs: 3 }),
];
