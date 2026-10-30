import { dateToSerial, serialToParts, timeToSerial } from './dates.js';
import { CellError, ERROR_CODES, isError, type Scalar } from './values.js';

export interface FormatLocale {
  decimal: string;
  group: string;
  monthNames: string[];
  dayNames: string[];
  trueLabel: string;
  falseLabel: string;
}

export const LOCALE_EN: FormatLocale = {
  decimal: '.',
  group: ',',
  monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  dayNames: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  trueLabel: 'TRUE',
  falseLabel: 'FALSE',
};

export const LOCALE_DE: FormatLocale = {
  decimal: ',',
  group: '.',
  monthNames: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
  dayNames: ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'],
  trueLabel: 'WAHR',
  falseLabel: 'FALSCH',
};

let defaultLocale: FormatLocale = LOCALE_EN;

export function setDefaultLocale(locale: FormatLocale): void {
  defaultLocale = locale;
}

export function getDefaultLocale(): FormatLocale {
  return defaultLocale;
}

export interface FormattedValue {
  text: string;
  color?: string;
}

// ---------------------------------------------------------------------------
// General format
// ---------------------------------------------------------------------------

export function formatGeneral(n: number, loc: FormatLocale = defaultLocale): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const abs = Math.abs(n);
  let s: string;
  if (abs >= 1e11 || abs < 1e-9) {
    const [mant, exp] = n.toExponential(5).split('e') as [string, string];
    const m = mant.includes('.') ? mant.replace(/0+$/, '').replace(/\.$/, '') : mant;
    const e = Number(exp);
    s = `${m}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  } else {
    s = String(parseFloat(n.toPrecision(11)));
    if (s.includes('e')) s = n.toFixed(10).replace(/0+$/, '').replace(/\.$/, '');
  }
  return loc.decimal === '.' ? s : s.replace('.', loc.decimal);
}

// ---------------------------------------------------------------------------
// Format code parsing
// ---------------------------------------------------------------------------

const COLORS: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#008000', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
};

interface Section {
  raw: string;
  color?: string;
  kind: 'general' | 'number' | 'date' | 'text';
}

const sectionCache = new Map<string, Section[]>();

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i]!;
    if (ch === '\\' && !inQuote) {
      cur += ch + (fmt[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"') inQuote = !inQuote;
    else if (!inQuote && ch === '[') inBracket = true;
    else if (!inQuote && ch === ']') inBracket = false;
    if (ch === ';' && !inQuote && !inBracket) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

function stripOutsideQuotes(s: string): string {
  return s.replace(/"[^"]*"/g, '').replace(/\\./g, '');
}

function parseSections(fmt: string): Section[] {
  let cached = sectionCache.get(fmt);
  if (cached) return cached;
  cached = splitSections(fmt).map((raw) => {
    let color: string | undefined;
    let body = raw.replace(/\[([^\]]*)\]/g, (m, inner: string) => {
      const lower = inner.toLowerCase();
      if (COLORS[lower]) {
        color = COLORS[lower];
        return '';
      }
      if (inner.startsWith('$')) {
        const sym = inner.slice(1).split('-')[0] ?? '';
        return sym ? `"${sym}"` : '';
      }
      if (/^h+$|^m+$|^s+$/i.test(inner)) return m; // elapsed time
      return ''; // conditions, locales etc.
    });
    const plain = stripOutsideQuotes(body);
    let kind: Section['kind'];
    if (/^\s*general\s*$/i.test(plain)) kind = 'general';
    else if (/[0#?]/.test(plain.replace(/\[[hms]+\]/gi, '').replace(/\.0+/g, '')) || !/[ymdhs]/i.test(plain)) {
      kind = plain.includes('@') && !/[0#?]/.test(plain) ? 'text' : 'number';
    } else kind = 'date';
    if (kind === 'general') body = 'General';
    return color ? { raw: body, color, kind } : { raw: body, kind };
  });
  sectionCache.set(fmt, cached);
  return cached;
}

/** Converts the literal parts of a format section to display text. */
function literalText(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      out += s.slice(i + 1, end < 0 ? undefined : end);
      i = end < 0 ? s.length : end;
    } else if (ch === '\\') {
      out += s[i + 1] ?? '';
      i++;
    } else if (ch === '_') {
      out += ' ';
      i++;
    } else if (ch === '*') {
      i++;
    } else if (ch === '@') {
      out += '@';
    } else out += ch;
  }
  return out;
}

// ---------------------------------------------------------------------------
// Number formatting
// ---------------------------------------------------------------------------

function groupDigits(int: string, sep: string): string {
  let out = '';
  for (let i = 0; i < int.length; i++) {
    if (i > 0 && (int.length - i) % 3 === 0) out += sep;
    out += int[i];
  }
  return out;
}

function formatNumberSection(n: number, raw: string, loc: FormatLocale): string {
  // Locate numeric placeholder region (outside quotes).
  let first = -1;
  let last = -1;
  let inQuote = false;
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i]!;
    if (ch === '"') inQuote = !inQuote;
    else if (ch === '\\' || ch === '_' || ch === '*') i++;
    else if (!inQuote && /[0#?]/.test(ch)) {
      if (first < 0) first = i;
      last = i;
    } else if (!inQuote && ch === '.' && first < 0 && /[0#?]/.test(raw[i + 1] ?? '')) first = i;
  }
  if (first < 0) return literalText(raw).replace('@', '');
  let end = last + 1;
  // Include exponent and trailing commas (scaling) in the pattern.
  const expMatch = /^[eE][+-][0#?]+/.exec(raw.slice(end));
  if (expMatch) end += expMatch[0].length;
  while (raw[end] === ',') end++;
  const prefix = literalText(raw.slice(0, first));
  const suffix = literalText(raw.slice(end));
  let pattern = raw.slice(first, end);

  const percent = (stripOutsideQuotes(raw).match(/%/g) ?? []).length;
  let value = Math.abs(n) * Math.pow(100, percent);

  let expPattern = '';
  const eIdx = pattern.search(/[eE][+-]/);
  if (eIdx >= 0) {
    expPattern = pattern.slice(eIdx);
    pattern = pattern.slice(0, eIdx);
  }
  let [intPat = '', decPat = ''] = pattern.split('.');
  let scale = 0;
  while (intPat.endsWith(',') && !decPat) {
    intPat = intPat.slice(0, -1);
    scale++;
  }
  value /= Math.pow(1000, scale);
  const grouping = intPat.includes(',');
  const minInt = (intPat.match(/0/g) ?? []).length;
  const maxDec = (decPat.match(/[0#?]/g) ?? []).length;
  const minDec = (decPat.match(/0/g) ?? []).length;

  let expText = '';
  if (expPattern) {
    let exp = value === 0 ? 0 : Math.floor(Math.log10(value));
    value = value / Math.pow(10, exp);
    if (Number(value.toFixed(maxDec)) >= 10) {
      value /= 10;
      exp++;
    }
    const digits = (expPattern.match(/0/g) ?? []).length;
    const sign = exp < 0 ? '-' : expPattern[1] === '+' ? '+' : '';
    expText = `E${sign}${String(Math.abs(exp)).padStart(digits, '0')}`;
  }

  // toFixed accepts at most 100 digits; further format digits can only be zeros.
  let [int, dec = ''] = value.toFixed(Math.min(maxDec, 100)).split('.') as [string, string?];
  dec = dec.slice(0, maxDec).padEnd(minDec, '0');
  while (dec.length > minDec && dec.endsWith('0')) dec = dec.slice(0, -1);
  if (int === '0' && minInt === 0) int = '';
  int = int.padStart(minInt, '0');
  if (grouping) int = groupDigits(int, loc.group);
  return prefix + int + (dec ? loc.decimal + dec : '') + expText + suffix;
}

// ---------------------------------------------------------------------------
// Date formatting
// ---------------------------------------------------------------------------

const DATE_TOKEN_RE = /\[h+\]|\[m+\]|\[s+\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|\.0+|"[^"]*"|\\.|./gi;

function formatDateSection(serial: number, raw: string, loc: FormatLocale): string {
  if (serial < 0) return '#'.repeat(8);
  const p = serialToParts(serial);
  const tokens = raw.match(DATE_TOKEN_RE) ?? [];
  const lower = tokens.map((t) => t.toLowerCase());
  const ampm = lower.some((t) => t === 'am/pm' || t === 'a/p');
  // Resolve m/mm → minutes when following an hour token or preceding a seconds token.
  const isMinute = new Set<number>();
  for (let i = 0; i < lower.length; i++) {
    if (lower[i] !== 'm' && lower[i] !== 'mm') continue;
    for (let j = i - 1; j >= 0; j--) {
      const t = lower[j]!;
      if (/^(h|hh|\[h+\])$/.test(t)) isMinute.add(i);
      if (/^[ymdhs]|^\[/.test(t)) break;
    }
    for (let j = i + 1; j < lower.length; j++) {
      const t = lower[j]!;
      if (/^(s|ss)$/.test(t)) isMinute.add(i);
      if (/^[ymdhs]|^\[/.test(t)) break;
    }
  }
  const hours12 = p.hours % 12 === 0 ? 12 : p.hours % 12;
  const pad = (v: number, len = 2) => String(v).padStart(len, '0');
  let out = '';
  tokens.forEach((tok, i) => {
    const t = lower[i]!;
    switch (t) {
      case 'yyyy': out += pad(p.year, 4); break;
      case 'yy': out += pad(p.year % 100); break;
      case 'mmmmm': out += loc.monthNames[p.month - 1]!.charAt(0); break;
      case 'mmmm': out += loc.monthNames[p.month - 1]; break;
      case 'mmm': out += loc.monthNames[p.month - 1]!.slice(0, 3); break;
      case 'mm': out += isMinute.has(i) ? pad(p.minutes) : pad(p.month); break;
      case 'm': out += isMinute.has(i) ? String(p.minutes) : String(p.month); break;
      case 'dddd': out += loc.dayNames[p.weekday]; break;
      case 'ddd': out += loc.dayNames[p.weekday]!.slice(0, 3); break;
      case 'dd': out += pad(p.day); break;
      case 'd': out += String(p.day); break;
      case 'hh': out += pad(ampm ? hours12 : p.hours); break;
      case 'h': out += String(ampm ? hours12 : p.hours); break;
      case 'ss': out += pad(p.seconds); break;
      case 's': out += String(p.seconds); break;
      case 'am/pm': out += p.hours < 12 ? 'AM' : 'PM'; break;
      case 'a/p': out += p.hours < 12 ? 'A' : 'P'; break;
      default:
        if (t.startsWith('[h')) out += pad(Math.floor(serial * 24), t.length - 2);
        else if (t.startsWith('[m')) out += pad(Math.floor(serial * 1440), t.length - 2);
        else if (t.startsWith('[s')) out += pad(Math.floor(serial * 86400), t.length - 2);
        else if (t.startsWith('.0')) out += loc.decimal + pad(Math.floor(p.milliseconds / Math.pow(10, 3 - (t.length - 1))), t.length - 1).slice(0, t.length - 1);
        else out += literalText(tok);
    }
  });
  return out;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false;
  return parseSections(fmt)[0]?.kind === 'date';
}

export function formatValueEx(v: Scalar, fmt?: string, loc: FormatLocale = defaultLocale): FormattedValue {
  if (v === null) return { text: '' };
  if (isError(v)) return { text: v.code };
  if (typeof v === 'boolean') return { text: v ? loc.trueLabel : loc.falseLabel };
  if (!fmt || /^general$/i.test(fmt)) {
    return { text: typeof v === 'number' ? formatGeneral(v, loc) : v };
  }
  const sections = parseSections(fmt);
  if (typeof v === 'string') {
    const textSec = sections.length >= 4 ? sections[3] : sections.find((s) => s.kind === 'text');
    if (!textSec) return { text: v };
    const res: FormattedValue = { text: literalText(textSec.raw).replace('@', v) };
    if (textSec.color) res.color = textSec.color;
    return res;
  }
  let sec = sections[0]!;
  let n = v;
  let autoMinus = v < 0;
  if (v < 0 && sections.length >= 2 && sections[1]!.raw.trim() !== '') {
    sec = sections[1]!;
    autoMinus = false;
    n = -v;
  } else if (v === 0 && sections.length >= 3 && sections[2]!.raw.trim() !== '') {
    sec = sections[2]!;
  }
  let text: string;
  switch (sec.kind) {
    case 'general':
      text = formatGeneral(Math.abs(n), loc);
      break;
    case 'date':
      text = formatDateSection(n, sec.raw, loc);
      autoMinus = false;
      break;
    case 'text':
      text = literalText(sec.raw).replace('@', formatGeneral(n, loc));
      autoMinus = false;
      break;
    default:
      text = formatNumberSection(n, sec.raw, loc);
  }
  if (autoMinus && /[1-9]/.test(text)) text = '-' + text;
  const res: FormattedValue = { text };
  if (sec.color) res.color = sec.color;
  return res;
}

export function formatValue(v: Scalar, fmt?: string, loc: FormatLocale = defaultLocale): string {
  return formatValueEx(v, fmt, loc).text;
}

// ---------------------------------------------------------------------------
// Input parsing (what the user typed into a cell)
// ---------------------------------------------------------------------------

export interface ParsedInput {
  value: Scalar;
  formula?: string;
  /** Suggested number format inferred from the input (e.g. "0%" for "12%"). */
  numFmt?: string;
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

const localeNumberRes = new Map<string, RegExp>();

/** Number pattern with the locale's group and decimal separators (cached). */
function localeNumberRe(loc: FormatLocale): RegExp {
  const key = `${loc.group}\u0000${loc.decimal}`;
  let re = localeNumberRes.get(key);
  if (!re) {
    const g = escapeRe(loc.group);
    const d = escapeRe(loc.decimal);
    re = new RegExp(`^[+-]?(\\d{1,3}(${g}\\d{3})+|\\d*)(${d}\\d+)?([eE][+-]?\\d+)?$`);
    localeNumberRes.set(key, re);
  }
  return re;
}

export function parseNumber(text: string, loc: FormatLocale = defaultLocale): number | null {
  const s = text.trim().replace(/\s/g, '');
  if (!s) return null;
  const localeRe = localeNumberRe(loc);
  if (localeRe.test(s) && /\d/.test(s)) {
    const n = Number(s.split(loc.group).join('').replace(loc.decimal, '.'));
    if (Number.isFinite(n)) return n;
  }
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) return n;
  }
  return null;
}

export function parseInput(text: string, loc: FormatLocale = defaultLocale): ParsedInput {
  if (text === '') return { value: null };
  if (text.startsWith('=') && text.length > 1) return { value: null, formula: text.slice(1) };
  if (text.startsWith("'")) return { value: text.slice(1) };
  const t = text.trim();
  const upper = t.toUpperCase();
  if (upper === 'TRUE' || upper === loc.trueLabel.toUpperCase()) return { value: true };
  if (upper === 'FALSE' || upper === loc.falseLabel.toUpperCase()) return { value: false };
  if ((ERROR_CODES as readonly string[]).includes(upper)) return { value: new CellError(upper as CellError['code']) };

  // Percent
  if (t.endsWith('%')) {
    const n = parseNumber(t.slice(0, -1), loc);
    if (n !== null) {
      const decimals = t.slice(0, -1).split(loc.decimal)[1]?.length ?? 0;
      return { value: n / 100, numFmt: decimals ? `0.${'0'.repeat(decimals)}%` : '0%' };
    }
  }
  // Currency
  const cur = /^([€$£])\s*(.+)$|^(.+?)\s*([€$£])$/.exec(t);
  if (cur) {
    const sym = cur[1] ?? cur[4]!;
    const n = parseNumber(cur[2] ?? cur[3]!, loc);
    if (n !== null) return { value: n, numFmt: cur[1] ? `"${sym}"#,##0.00` : `#,##0.00 "${sym}"` };
  }
  const n = parseNumber(t, loc);
  if (n !== null) {
    const hasGroup = new RegExp(`\\d${escapeRe(loc.group)}\\d{3}`).test(t);
    return hasGroup ? { value: n, numFmt: t.includes(loc.decimal) ? '#,##0.00' : '#,##0' } : { value: n };
  }
  // Dates & times
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  if (m) {
    const date = dateToSerial(+m[1]!, +m[2]!, +m[3]!);
    if (m[4] !== undefined) return { value: date + timeToSerial(+m[4], +m[5]!, +(m[6] ?? 0)), numFmt: 'yyyy-mm-dd hh:mm' };
    return { value: date, numFmt: 'yyyy-mm-dd' };
  }
  // Year first with slashes is a date in every Excel locale.
  m = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(t);
  if (m && +m[2]! >= 1 && +m[2]! <= 12 && +m[3]! >= 1 && +m[3]! <= 31) return { value: dateToSerial(+m[1]!, +m[2]!, +m[3]!), numFmt: 'yyyy/m/d' };
  m = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/.exec(t);
  if (m && +m[2]! <= 12 && +m[1]! <= 31) {
    const y = m[3]!.length === 2 ? 2000 + +m[3]! : +m[3]!;
    return { value: dateToSerial(y, +m[2]!, +m[1]!), numFmt: 'dd.mm.yyyy' };
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (m && +m[1]! <= 12 && +m[2]! <= 31) return { value: dateToSerial(+m[3]!, +m[1]!, +m[2]!), numFmt: 'm/d/yyyy' };
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m && +m[1]! < 24 && +m[2]! < 60) {
    return { value: timeToSerial(+m[1]!, +m[2]!, +(m[3] ?? 0)), numFmt: m[3] ? 'hh:mm:ss' : 'hh:mm' };
  }
  return { value: text };
}
