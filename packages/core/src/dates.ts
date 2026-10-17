/**
 * Excel serial date helpers (1900 date system, including Excel's fictitious 1900-02-29).
 */
const MS_PER_DAY = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 30);

export function dateToSerial(year: number, month: number, day: number): number {
  // Date.UTC handles month/day overflow like Excel's DATE() does.
  const diff = Math.round((Date.UTC(year, month - 1, day) - EPOCH) / MS_PER_DAY);
  return diff <= 60 ? diff - 1 : diff;
}

export function timeToSerial(h: number, m: number, s: number): number {
  return (h * 3600 + m * 60 + s) / 86_400;
}

export interface DateParts {
  year: number;
  month: number; // 1-12
  day: number;
  hours: number;
  minutes: number;
  seconds: number;
  milliseconds: number;
  /** 0 = Sunday */
  weekday: number;
}

export function serialToParts(serial: number): DateParts {
  let days = Math.floor(serial);
  let frac = serial - days;
  let ms = Math.round(frac * MS_PER_DAY);
  if (ms >= MS_PER_DAY) {
    days += 1;
    ms -= MS_PER_DAY;
  }
  frac = ms;
  let year: number, month: number, day: number, weekday: number;
  if (days === 60) {
    [year, month, day, weekday] = [1900, 2, 29, 4];
  } else {
    const d = new Date(EPOCH + (days < 60 ? days + 1 : days) * MS_PER_DAY);
    year = d.getUTCFullYear();
    month = d.getUTCMonth() + 1;
    day = d.getUTCDate();
    weekday = d.getUTCDay();
  }
  const totalSec = Math.floor(frac / 1000);
  return {
    year,
    month,
    day,
    hours: Math.floor(totalSec / 3600),
    minutes: Math.floor((totalSec % 3600) / 60),
    seconds: totalSec % 60,
    milliseconds: frac % 1000,
    weekday,
  };
}

export function jsDateToSerial(d: Date): number {
  return (
    dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) +
    timeToSerial(d.getHours(), d.getMinutes(), d.getSeconds() + d.getMilliseconds() / 1000)
  );
}

export function serialToJsDate(serial: number): Date {
  const p = serialToParts(serial);
  return new Date(p.year, p.month - 1, p.day, p.hours, p.minutes, p.seconds, p.milliseconds);
}
