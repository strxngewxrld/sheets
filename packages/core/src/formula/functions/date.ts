import { dateToSerial, jsDateToSerial, serialToParts, timeToSerial } from '../../dates.js';
import { parseInput } from '../../numfmt.js';
import { Err, isError } from '../../values.js';
import { toMatrix, type EvalValue, type FunctionContext } from '../evaluator.js';
import type { FunctionDef } from '../registry.js';
import { bail, boolArg, def, numArg, scalarArg, textArg, unary } from './helpers.js';

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

// Weekend days indexed by weekday (0 = Sunday).
const SAT_SUN = [true, false, false, false, false, false, true];
const MAX_DATE_SERIAL = 2_958_465; // 9999-12-31, Excel's last supported date.

function checkDate(serial: number): void {
  if (!Number.isSafeInteger(serial) || serial < 0 || serial > MAX_DATE_SERIAL) bail(Err.NUM);
}

function holidaySet(arg: EvalValue | undefined, ctx: FunctionContext): Set<number> {
  const holidays = new Set<number>();
  if (arg !== undefined && arg !== null) for (const v of toMatrix(arg, ctx).flat()) {
    if (isError(v)) bail(v);
    // Dates given as text, e.g. {"2006/1/2"}, count like in Excel.
    const serial = typeof v === 'string' ? parseInput(v).value : v;
    if (typeof v === 'string' && typeof serial !== 'number') bail(Err.VALUE);
    if (typeof serial === 'number') {
      checkDate(Math.trunc(serial));
      holidays.add(Math.trunc(serial));
    }
  }
  return holidays;
}

/**
 * The weekend of NETWORKDAYS.INTL and WORKDAY.INTL: 1-7 for two days (1 = Saturday
 * and Sunday, 2 = Sunday and Monday, ...), 11-17 for one day (11 = Sunday, ...) or
 * a mask of seven 0/1 characters starting on Monday.
 */
function weekendArg(arg: EvalValue | undefined, ctx: FunctionContext): boolean[] {
  const v = scalarArg(arg, ctx);
  if (v === null) return SAT_SUN;
  if (isError(v)) bail(v);
  if (typeof v === 'string') {
    if (!/^[01]{7}$/.test(v)) bail(Err.VALUE);
    return [v[6] === '1', ...[...v.slice(0, 6)].map((c) => c === '1')];
  }
  const code = Math.trunc(numArg(arg, ctx));
  const days = new Array<boolean>(7).fill(false);
  if (code >= 1 && code <= 7) {
    days[(code + 5) % 7] = true;
    days[(code + 6) % 7] = true;
  } else if (code >= 11 && code <= 17) days[code - 11] = true;
  else bail(Err.NUM);
  return days;
}

/** At most six remainder days; split at Excel's fictitious leap day. */
function weekdays(start: number, end: number, weekend: boolean[]): number {
  if (start > end) return 0;
  if (start <= 60 && end > 60) return weekdays(start, 60, weekend) + weekdays(61, end, weekend);
  const length = end - start + 1;
  let n = Math.floor(length / 7) * weekend.filter((day) => !day).length;
  const first = serialToParts(start).weekday;
  for (let i = 0; i < length % 7; i++) if (!weekend[(first + i) % 7]) n++;
  return n;
}

function holidayIndex(holidays: Set<number>, weekend: boolean[]): number[] {
  return [...holidays].filter((day) => !weekend[serialToParts(day).weekday]).sort((a, b) => a - b);
}

/** Index of the first holiday on or after day. */
function lowerBound(holidays: number[], day: number): number {
  let lo = 0, hi = holidays.length;
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    if (holidays[mid]! < day) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function countWorkdays(start: number, end: number, holidays: number[], weekend: boolean[]): number {
  return weekdays(start, end, weekend) - (lowerBound(holidays, end + 1) - lowerBound(holidays, start));
}

function networkDays(start: number, end: number, holidays: Set<number>, weekend = SAT_SUN): number {
  checkDate(start);
  checkDate(end);
  const sign = start <= end ? 1 : -1;
  if (sign < 0) [start, end] = [end, start];
  return countWorkdays(start, end, holidayIndex(holidays, weekend), weekend) * sign;
}

function workday(start: number, days: number, holidays: Set<number>, weekend = SAT_SUN): number {
  checkDate(start);
  if (!Number.isSafeInteger(days) || Math.abs(days) > MAX_DATE_SERIAL) bail(Err.NUM);
  if (days === 0) return start;
  const step = days >= 0 ? 1 : -1;
  const sorted = holidayIndex(holidays, weekend);
  const count = (distance: number) => step > 0
    ? countWorkdays(start + 1, start + distance, sorted, weekend)
    : countWorkdays(start - distance, start - 1, sorted, weekend);
  let lo = 1, hi = step > 0 ? MAX_DATE_SERIAL - start : start;
  const target = Math.abs(days);
  if (hi < 1 || count(hi) < target) bail(Err.NUM);
  // Find the first distance containing the requested number of workdays.
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    if (count(mid) < target) lo = mid + 1;
    else hi = mid;
  }
  return start + step * lo;
}

const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const lastDay = (y: number, m: number) => new Date(Date.UTC(y, m, 0)).getUTCDate();

/** DAYS360: the US (NASD) method or, with european = TRUE, the European method. */
function days360(start: number, end: number, european: boolean): number {
  const a = serialToParts(start), b = serialToParts(end);
  let d1 = a.day, d2 = b.day, m2 = b.month, y2 = b.year;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    if (d1 === lastDay(a.year, a.month)) d1 = 30;
    if (d2 === 31) {
      if (d1 < 30) {
        d2 = 1;
        if (++m2 > 12) { m2 = 1; y2++; }
      } else d2 = 30;
    }
  }
  return (y2 - a.year) * 360 + (m2 - a.month) * 30 + (d2 - d1);
}

/** YEARFRAC with the day count bases 0 (US 30/360), 1 (actual/actual), 2 (actual/360), 3 (actual/365), 4 (European 30/360). */
function yearFrac(start: number, end: number, basis: number): number {
  if (start > end) [start, end] = [end, start];
  const a = serialToParts(start), b = serialToParts(end);
  let d1 = a.day, d2 = b.day;
  const thirty = () => (d2 + b.month * 30 + b.year * 360 - (d1 + a.month * 30 + a.year * 360)) / 360;
  switch (basis) {
    case 0: {
      const lastOfFeb = (p: typeof a) => p.month === 2 && p.day === lastDay(p.year, 2);
      if (d1 === 31 && d2 === 31) { d1 = 30; d2 = 30; }
      else if (d1 === 31) d1 = 30;
      else if (d1 === 30 && d2 === 31) d2 = 30;
      else if (lastOfFeb(a) && lastOfFeb(b)) { d1 = 30; d2 = 30; }
      else if (lastOfFeb(a)) d1 = 30;
      return thirty();
    }
    case 1: {
      const days = end - start;
      if (a.year === b.year || (b.year === a.year + 1 && (a.month > b.month || (a.month === b.month && a.day >= b.day)))) {
        // Up to one year: 366 days if the span contains a 29 February.
        const feb29 = (y: number) => isLeap(y) && start <= dateToSerial(y, 2, 29) && dateToSerial(y, 2, 29) <= end;
        const leap = a.year === b.year ? isLeap(a.year) : feb29(a.year) || feb29(b.year);
        return days / (leap ? 366 : 365);
      }
      // Longer spans: the average length of the calendar years touched.
      return days / ((dateToSerial(b.year + 1, 1, 1) - dateToSerial(a.year, 1, 1)) / (b.year - a.year + 1));
    }
    case 2: return (end - start) / 360;
    case 3: return (end - start) / 365;
    default:
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
      return thirty();
  }
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
  def('NETWORKDAYS', 'date', 'NETWORKDAYS(start_date, end_date, [holidays])', 'Working days between two dates.', ([s, e, h], ctx) =>
    networkDays(Math.trunc(dateArg(s, ctx)), Math.trunc(dateArg(e, ctx)), holidaySet(h, ctx)), { minArgs: 2, maxArgs: 3 }),
  def('NETWORKDAYS.INTL', 'date', 'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])', 'Working days between two dates with a custom weekend.', ([s, e, w, h], ctx) =>
    networkDays(Math.trunc(dateArg(s, ctx)), Math.trunc(dateArg(e, ctx)), holidaySet(h, ctx), weekendArg(w, ctx)), { minArgs: 2, maxArgs: 4 }),
  def('WORKDAY', 'date', 'WORKDAY(start_date, days, [holidays])', 'Date after a number of working days.', ([s, n, h], ctx) =>
    workday(Math.trunc(dateArg(s, ctx)), Math.trunc(numArg(n, ctx)), holidaySet(h, ctx)), { minArgs: 2, maxArgs: 3 }),
  def('WORKDAY.INTL', 'date', 'WORKDAY.INTL(start_date, days, [weekend], [holidays])', 'Date after a number of working days with a custom weekend.', ([s, n, w, h], ctx) => {
    const weekend = weekendArg(w, ctx);
    if (weekend.every(Boolean)) return Err.VALUE;
    return workday(Math.trunc(dateArg(s, ctx)), Math.trunc(numArg(n, ctx)), holidaySet(h, ctx), weekend);
  }, { minArgs: 2, maxArgs: 4 }),
  def('DAYS360', 'date', 'DAYS360(start_date, end_date, [method])', 'Days between two dates in a 360-day year (method TRUE: European).', ([s, e, m], ctx) =>
    days360(Math.trunc(dateArg(s, ctx)), Math.trunc(dateArg(e, ctx)), boolArg(m, ctx, false)), { minArgs: 2, maxArgs: 3 }),
  def('YEARFRAC', 'date', 'YEARFRAC(start_date, end_date, [basis])', 'Fraction of a year between two dates (basis 0-4).', ([s, e, b], ctx) => {
    const basis = Math.trunc(numArg(b, ctx, 0));
    const start = Math.trunc(dateArg(s, ctx)), end = Math.trunc(dateArg(e, ctx));
    if (basis < 0 || basis > 4 || start < 0 || end < 0) return Err.NUM;
    return yearFrac(start, end, basis);
  }, { minArgs: 2, maxArgs: 3 }),
];
