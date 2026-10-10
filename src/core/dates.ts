// Excel date serial numbers (1900 date system, including Excel's fake 1900-02-29).

const MS_PER_DAY = 86_400_000;
const EPOCH_AFTER_LEAP = Date.UTC(1899, 11, 30); // serial 0 for serials >= 61
const EPOCH_BEFORE_LEAP = Date.UTC(1899, 11, 31); // serial 0 for serials < 60

export interface DateParts {
  y: number;
  m: number; // 1-12
  d: number;
  H: number;
  M: number;
  S: number;
  ms: number;
  dow: number; // 0 = Sunday
}

export function serialToParts(serial: number): DateParts {
  let days = Math.floor(serial);
  let frac = serial - days;
  // Round to the nearest millisecond to avoid 59.9999 seconds artifacts.
  let msInDay = Math.round(frac * MS_PER_DAY);
  if (msInDay >= MS_PER_DAY) {
    days += 1;
    msInDay -= MS_PER_DAY;
  }
  if (days === 60) {
    // The non-existent 1900-02-29 that Excel keeps for Lotus compatibility.
    return withTime({ y: 1900, m: 2, d: 29, H: 0, M: 0, S: 0, ms: 0, dow: 3 }, msInDay);
  }
  const base = days < 60 ? EPOCH_BEFORE_LEAP : EPOCH_AFTER_LEAP;
  const dt = new Date(base + days * MS_PER_DAY);
  return withTime(
    { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), H: 0, M: 0, S: 0, ms: 0, dow: dt.getUTCDay() },
    msInDay,
  );
}

function withTime(p: DateParts, msInDay: number): DateParts {
  p.H = Math.floor(msInDay / 3_600_000);
  p.M = Math.floor((msInDay % 3_600_000) / 60_000);
  p.S = Math.floor((msInDay % 60_000) / 1000);
  p.ms = msInDay % 1000;
  return p;
}

/** Builds a serial from a (possibly out-of-range) year/month/day, like Excel's DATE(). */
export function dateToSerial(y: number, m: number, d: number): number {
  if (y >= 0 && y < 1900) y += 1900;
  const t = Date.UTC(y, m - 1, d);
  const serial = Math.round((t - EPOCH_AFTER_LEAP) / MS_PER_DAY);
  if (serial < 61) {
    // Dates before March 1, 1900 use the other epoch.
    return Math.round((t - EPOCH_BEFORE_LEAP) / MS_PER_DAY);
  }
  return serial;
}

export function timeToSerial(h: number, m: number, s: number): number {
  return (h * 3600 + m * 60 + s) / 86400;
}

export function jsDateToSerial(date: Date): number {
  // Treat the JS date's UTC fields as the wall-clock value (ExcelJS reads dates as UTC).
  const t = date.getTime();
  const serial = (t - EPOCH_AFTER_LEAP) / MS_PER_DAY;
  if (serial < 61) return (t - EPOCH_BEFORE_LEAP) / MS_PER_DAY;
  return serial;
}

export function serialToJsDate(serial: number): Date {
  const p = serialToParts(serial);
  return new Date(Date.UTC(p.y, p.m - 1, p.d, p.H, p.M, p.S, p.ms));
}

export function todaySerial(): number {
  const now = new Date();
  return dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

export function nowSerial(): number {
  const now = new Date();
  return (
    dateToSerial(now.getFullYear(), now.getMonth() + 1, now.getDate()) +
    timeToSerial(now.getHours(), now.getMinutes(), now.getSeconds() + now.getMilliseconds() / 1000)
  );
}

export function daysInMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

export function isLeapYear(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

export const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
export const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
