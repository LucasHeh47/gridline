// Interprets what the user types into a cell.
import { config } from './config';
import { dateToSerial, MONTHS, timeToSerial } from './dates';
import { parseErrorCode, ErrorValue, type Scalar } from './values';

export interface ParsedInput {
  /** Literal value (when not a formula). */
  value?: Scalar;
  /** Formula text without '='. */
  formula?: string;
  /** Number format implied by the input (e.g. "$1,200" -> currency). */
  numFmt?: string;
}

const MONTH_ABBR = MONTHS.map((m) => m.slice(0, 3).toLowerCase());

function monthIndex(s: string): number {
  const l = s.toLowerCase();
  let i = MONTH_ABBR.indexOf(l.slice(0, 3));
  if (i < 0) return -1;
  if (l.length > 3 && !MONTHS[i].toLowerCase().startsWith(l) && l !== 'sept') return -1;
  return i + 1;
}

const NUM_RE = /^([+-])?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?([eE][+-]?\d+)?$/;

/**
 * Parses text that looks like a number (with optional currency, thousands separators, percent,
 * parentheses for negatives). Returns null if it isn't one. Used for both input and coercion.
 */
export function parseNumberLike(raw: string): number | null {
  const r = parseNumberWithFormat(raw);
  if (r) return r.value;
  const d = parseDateTime(raw.trim());
  return d ? d.value : null;
}

function parseNumberWithFormat(raw: string): { value: number; numFmt?: string } | null {
  let s = raw.trim();
  if (!s) return null;
  let neg = false;
  let currency = '';
  let percent = false;
  if (s.startsWith('(') && s.endsWith(')')) {
    neg = true;
    s = s.slice(1, -1).trim();
  }
  if (s.startsWith('-')) {
    neg = !neg;
    s = s.slice(1).trim();
  } else if (s.startsWith('+')) s = s.slice(1).trim();
  const cm = /^([$€£¥])\s*/.exec(s);
  if (cm) {
    currency = cm[1];
    s = s.slice(cm[0].length);
    if (s.startsWith('-')) {
      neg = !neg;
      s = s.slice(1);
    }
  }
  if (s.endsWith('%')) {
    percent = true;
    s = s.slice(0, -1).trim();
  }
  if (!s || !NUM_RE.test(s) || s === '.') return null;
  const hasThousands = s.includes(',');
  const n = parseFloat(s.replace(/,/g, ''));
  if (isNaN(n)) return null;
  let value = neg ? -n : n;
  if (percent) value /= 100;
  const dot = s.indexOf('.');
  const decimals = dot >= 0 && !/[eE]/.test(s) ? s.length - dot - 1 : 0;
  let numFmt: string | undefined;
  if (currency) numFmt = `"${currency}"#,##0${decimals > 0 ? '.00' : ''}`;
  else if (percent) numFmt = decimals > 0 ? '0.' + '0'.repeat(Math.min(decimals, 10)) + '%' : '0%';
  else if (hasThousands) numFmt = decimals > 0 ? '#,##0.' + '0'.repeat(decimals) : '#,##0';
  else if (/[eE]/.test(s)) numFmt = '0.00E+00';
  return { value, numFmt };
}

/** Parses dates and times: 3/14/2025, 2025-03-14, Mar 14 2025, 14-Mar-25, 3:30 PM, 2025-03-14 13:05 */
export function parseDateTime(s: string): { value: number; numFmt: string } | null {
  const t = s.trim();
  if (!t || t.length > 40) return null;
  // Time only
  const tm = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm]?)?$/.exec(t);
  if (tm) {
    const time = parseTimeMatch(tm);
    if (time === null) return null;
    const fmt = tm[4] ? (tm[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : tm[3] ? 'h:mm:ss' : 'h:mm';
    return { value: time, numFmt: fmt };
  }
  // Split off a trailing time
  let datePart = t;
  let time = 0;
  let hasTime = false;
  const tt = /^(.*?)[\sT]+(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm]?)?$/.exec(t);
  if (tt) {
    const tv = parseTimeMatch([tt[0], tt[2], tt[3], tt[4], tt[5]] as unknown as RegExpExecArray);
    if (tv === null) return null;
    datePart = tt[1];
    time = tv;
    hasTime = true;
  }
  const d = parseDateOnly(datePart);
  if (!d) return null;
  return { value: d.value + time, numFmt: hasTime ? d.fmt + ' h:mm' : d.fmt };
}

function parseTimeMatch(m: RegExpExecArray): number | null {
  let h = parseInt(m[1], 10);
  const mi = parseInt(m[2], 10);
  const sec = m[3] ? parseFloat(m[3]) : 0;
  const ap = m[4]?.toLowerCase();
  if (mi > 59 || sec >= 60) return null;
  if (ap) {
    if (h < 1 || h > 12) return null;
    if (ap.startsWith('p') && h !== 12) h += 12;
    if (ap.startsWith('a') && h === 12) h = 0;
  } else if (h > 23) {
    // Allow durations like 25:30
    return timeToSerial(h, mi, sec);
  }
  return timeToSerial(h, mi, sec);
}

function validYmd(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false;
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return d <= dim;
}

function fixYear(y: number, len: number): number {
  if (len <= 2) return y < 30 ? 2000 + y : 1900 + y;
  return y;
}

function parseDateOnly(s: string): { value: number; fmt: string } | null {
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s);
  if (m) {
    const y = +m[1];
    const mo = +m[2];
    const d = +m[3];
    if (!validYmd(y, mo, d)) return null;
    return { value: dateToSerial(y, mo, d), fmt: 'yyyy-mm-dd' };
  }
  const dmy = config.dateOrder === 'DMY';
  m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})$/.exec(s);
  if (m) {
    const mo = dmy ? +m[2] : +m[1];
    const d = dmy ? +m[1] : +m[2];
    const y = fixYear(+m[3], m[3].length);
    if (!validYmd(y, mo, d)) return null;
    return { value: dateToSerial(y, mo, d), fmt: dmy ? 'd/m/yyyy' : 'm/d/yyyy' };
  }
  m = /^(\d{1,2})\/(\d{1,2})$/.exec(s);
  if (m) {
    const mo = dmy ? +m[2] : +m[1];
    const d = dmy ? +m[1] : +m[2];
    const y = new Date().getFullYear();
    if (!validYmd(y, mo, d)) return null;
    return { value: dateToSerial(y, mo, d), fmt: 'd-mmm' };
  }
  // 14-Mar-2025, 14 Mar 2025, 14-Mar
  m = /^(\d{1,2})[\s-]([A-Za-z]{3,9})\.?(?:[\s,-]+(\d{2}|\d{4}))?$/.exec(s);
  if (m) {
    const mo = monthIndex(m[2]);
    const d = +m[1];
    const y = m[3] ? fixYear(+m[3], m[3].length) : new Date().getFullYear();
    if (mo < 0 || !validYmd(y, mo, d)) return null;
    return { value: dateToSerial(y, mo, d), fmt: m[3] ? 'd-mmm-yy' : 'd-mmm' };
  }
  // Mar 14, 2025 / March 14 2025 / Mar 14
  m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}|\d{2}))?$/.exec(s);
  if (m) {
    const mo = monthIndex(m[1]);
    const d = +m[2];
    const y = m[3] ? fixYear(+m[3], m[3].length) : new Date().getFullYear();
    if (mo < 0 || !validYmd(y, mo, d)) return null;
    return { value: dateToSerial(y, mo, d), fmt: m[3] ? 'mmm d, yyyy' : 'mmm d' };
  }
  // Mar 2025 / March-2025
  m = /^([A-Za-z]{3,9})[\s-](\d{4})$/.exec(s);
  if (m) {
    const mo = monthIndex(m[1]);
    const y = +m[2];
    if (mo < 0 || !validYmd(y, mo, 1)) return null;
    return { value: dateToSerial(y, mo, 1), fmt: 'mmm-yy' };
  }
  return null;
}

/**
 * Interprets user input. Improvements over Excel:
 *  - Numbers with leading zeros ("00123") stay as text, so IDs and ZIP codes survive.
 *  - Long digit strings (> 15 digits, e.g. card or account numbers) stay as text instead of losing precision.
 */
export function parseInput(raw: string, currentFmt?: string): ParsedInput {
  if (raw === '') return { value: null };
  if (raw.startsWith('=') && raw.length > 1) return { formula: raw.slice(1) };
  if (raw.startsWith("'")) return { value: raw.slice(1) };
  if (currentFmt === '@') return { value: raw };
  const t = raw.trim();
  if (/^(TRUE|FALSE)$/i.test(t)) return { value: t.toUpperCase() === 'TRUE' };
  if (t.startsWith('#')) {
    const code = parseErrorCode(t);
    if (code) return { value: new ErrorValue(code, 'This error was typed directly into the cell.') };
  }
  // Leading zeros: keep as text (except "0", "0.5", "-0.5")
  if (config.keepLeadingZeros && /^[+-]?0\d+$/.test(t)) return { value: raw };
  const digits = t.replace(/[^0-9]/g, '');
  if (config.longDigitsAsText && /^[+-]?\d+$/.test(t) && digits.length > 15) return { value: raw };
  const num = parseNumberWithFormat(t);
  if (num) {
    const keepFmt = currentFmt && currentFmt !== 'General';
    return { value: num.value, numFmt: keepFmt ? undefined : num.numFmt };
  }
  const dt = config.recognizeDates ? parseDateTime(t) : null;
  if (dt) {
    const keepFmt = currentFmt && currentFmt !== 'General' && isDateLike(currentFmt);
    return { value: dt.value, numFmt: keepFmt ? undefined : dt.numFmt };
  }
  return { value: raw };
}

function isDateLike(fmt: string): boolean {
  return /[ymdhs]/i.test(fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''));
}
