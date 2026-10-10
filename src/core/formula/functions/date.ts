import { dateToSerial, daysInMonth, nowSerial, serialToParts, timeToSerial, todaySerial } from '../../dates';
import { parseDateTime } from '../../input';
import { Errors, ErrorValue, isError } from '../../values';
import { argInt, argNum, argText, collectNumbers } from '../coerce';
import type { Arg, Ctx, Value } from '../types';
import { define } from './registry';

function serialArg(ctx: Ctx, a: Arg): number | ErrorValue {
  const n = argNum(ctx, a);
  if (isError(n)) return n;
  if (n < 0) return Errors.num('Dates can\'t be negative.');
  if (n > 2958465) return Errors.num('That date is after the year 9999.');
  return n;
}

function part(f: (p: ReturnType<typeof serialToParts>) => number) {
  return (ctx: Ctx, [a]: Arg[]): Value => {
    const s = serialArg(ctx, a);
    return isError(s) ? s : f(serialToParts(s));
  };
}

define('TODAY', 'Date & time', '', 'Returns today\'s date. Updates every time the sheet recalculates.', () => todaySerial(), { volatile: true });
define('NOW', 'Date & time', '', 'Returns the current date and time. Updates every time the sheet recalculates.', () => nowSerial(), { volatile: true });

define(
  'DATE',
  'Date & time',
  'year, month, day',
  'Builds a date from year, month and day. Months and days can overflow: DATE(2025, 14, 1) is Feb 1, 2026.',
  (ctx, [a, b, c]) => {
    const y = argInt(ctx, a);
    if (isError(y)) return y;
    const m = argInt(ctx, b);
    if (isError(m)) return m;
    const d = argInt(ctx, c);
    if (isError(d)) return d;
    if (y < 0 || y > 9999) return Errors.num('The year must be between 0 and 9999.');
    const s = dateToSerial(y, m, d);
    if (s < 0) return Errors.num('That date is before January 1, 1900.');
    return s;
  },
  { lift: true },
);
define(
  'TIME',
  'Date & time',
  'hour, minute, second',
  'Builds a time from hours, minutes and seconds.',
  (ctx, [a, b, c]) => {
    const h = argInt(ctx, a);
    if (isError(h)) return h;
    const m = argInt(ctx, b);
    if (isError(m)) return m;
    const s = argInt(ctx, c);
    if (isError(s)) return s;
    const t = timeToSerial(h, m, s);
    if (t < 0) return Errors.num('The time can\'t be negative.');
    return t - Math.floor(t);
  },
  { lift: true },
);

define('YEAR', 'Date & time', 'serial_number', 'Returns the year of a date.', part((p) => p.y), { lift: true });
define('MONTH', 'Date & time', 'serial_number', 'Returns the month of a date (1-12).', part((p) => p.m), { lift: true });
define('DAY', 'Date & time', 'serial_number', 'Returns the day of the month (1-31).', part((p) => p.d), { lift: true });
define('HOUR', 'Date & time', 'serial_number', 'Returns the hour (0-23).', part((p) => p.H), { lift: true });
define('MINUTE', 'Date & time', 'serial_number', 'Returns the minute (0-59).', part((p) => p.M), { lift: true });
define('SECOND', 'Date & time', 'serial_number', 'Returns the second (0-59).', part((p) => p.S + (p.ms >= 500 ? 1 : 0)), { lift: true });

define(
  'WEEKDAY',
  'Date & time',
  'serial_number, [return_type]',
  'Returns the day of the week. Default: 1 (Sunday) to 7 (Saturday); return_type 2 gives 1 (Monday) to 7 (Sunday).',
  (ctx, [a, b]) => {
    const s = serialArg(ctx, a);
    if (isError(s)) return s;
    const t = argInt(ctx, b, 1);
    if (isError(t)) return t;
    const dow = serialToParts(s).dow; // 0 = Sunday
    switch (t) {
      case 1:
      case 17:
        return dow + 1;
      case 2:
      case 11:
        return ((dow + 6) % 7) + 1;
      case 3:
        return (dow + 6) % 7;
      case 12:
      case 13:
      case 14:
      case 15:
      case 16: {
        const start = t - 10; // 12 -> Tuesday(2)
        return ((dow - start + 7) % 7) + 1;
      }
    }
    return Errors.num('return_type must be 1, 2, 3 or 11-17.');
  },
  { lift: true },
);

function isoWeek(serial: number): number {
  const p = serialToParts(serial);
  const d = new Date(Date.UTC(p.y, p.m - 1, p.d));
  const day = (d.getUTCDay() + 6) % 7;
  d.setUTCDate(d.getUTCDate() - day + 3);
  const firstThursday = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((d.getTime() - firstThursday.getTime()) / 86400000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
}

define('ISOWEEKNUM', 'Date & time', 'date', 'Returns the ISO week number of the year.', (ctx, [a]) => {
  const s = serialArg(ctx, a);
  return isError(s) ? s : isoWeek(s);
}, { lift: true });
define(
  'WEEKNUM',
  'Date & time',
  'serial_number, [return_type]',
  'Returns the week number of the year (1 = week containing Jan 1, weeks start Sunday; 2 = Monday; 21 = ISO).',
  (ctx, [a, b]) => {
    const s = serialArg(ctx, a);
    if (isError(s)) return s;
    const t = argInt(ctx, b, 1);
    if (isError(t)) return t;
    if (t === 21) return isoWeek(s);
    const startDow = t === 1 || t === 17 ? 0 : t === 2 || t === 11 ? 1 : t >= 12 && t <= 16 ? t - 10 : -1;
    if (startDow < 0) return Errors.num();
    const p = serialToParts(s);
    const jan1 = dateToSerial(p.y, 1, 1);
    const jan1Dow = serialToParts(jan1).dow;
    const offset = (jan1Dow - startDow + 7) % 7;
    return Math.floor((Math.floor(s) - jan1 + offset) / 7) + 1;
  },
  { lift: true },
);

function addMonths(serial: number, months: number, endOfMonth: boolean): number {
  const p = serialToParts(serial);
  let y = p.y;
  let m = p.m + months;
  y += Math.floor((m - 1) / 12);
  m = ((((m - 1) % 12) + 12) % 12) + 1;
  const dim = daysInMonth(y, m);
  return dateToSerial(y, m, endOfMonth ? dim : Math.min(p.d, dim));
}

define('EDATE', 'Date & time', 'start_date, months', 'Returns the date a number of months before or after start_date.', (ctx, [a, b]) => {
  const s = serialArg(ctx, a);
  if (isError(s)) return s;
  const m = argInt(ctx, b);
  if (isError(m)) return m;
  return addMonths(s, m, false);
}, { lift: true });
define('EOMONTH', 'Date & time', 'start_date, months', 'Returns the last day of the month, a number of months away.', (ctx, [a, b]) => {
  const s = serialArg(ctx, a);
  if (isError(s)) return s;
  const m = argInt(ctx, b);
  if (isError(m)) return m;
  return addMonths(s, m, true);
}, { lift: true });

define(
  'DATEDIF',
  'Date & time',
  'start_date, end_date, unit',
  'Calculates the difference between two dates in "Y" years, "M" months, "D" days, "MD", "YM" or "YD".',
  (ctx, [a, b, c]) => {
    const s = serialArg(ctx, a);
    if (isError(s)) return s;
    const e = serialArg(ctx, b);
    if (isError(e)) return e;
    const unit = argText(ctx, c);
    if (isError(unit)) return unit;
    if (s > e) return Errors.num('start_date must be on or before end_date.');
    const ps = serialToParts(s);
    const pe = serialToParts(e);
    let months = (pe.y - ps.y) * 12 + (pe.m - ps.m);
    if (pe.d < ps.d) months--;
    switch (unit.toUpperCase()) {
      case 'Y':
        return Math.floor(months / 12);
      case 'M':
        return months;
      case 'D':
        return Math.floor(e) - Math.floor(s);
      case 'MD': {
        let d = pe.d - ps.d;
        if (d < 0) {
          const pm = pe.m === 1 ? 12 : pe.m - 1;
          const py = pe.m === 1 ? pe.y - 1 : pe.y;
          d += daysInMonth(py, pm);
        }
        return d;
      }
      case 'YM':
        return ((months % 12) + 12) % 12;
      case 'YD': {
        let start = dateToSerial(pe.y, ps.m, ps.d);
        if (start > Math.floor(e)) start = dateToSerial(pe.y - 1, ps.m, ps.d);
        return Math.floor(e) - start;
      }
    }
    return Errors.num('unit must be "Y", "M", "D", "MD", "YM" or "YD".');
  },
);

define('DAYS', 'Date & time', 'end_date, start_date', 'Returns the number of days between two dates.', (ctx, [a, b]) => {
  const e = argNum(ctx, a);
  if (isError(e)) return e;
  const s = argNum(ctx, b);
  if (isError(s)) return s;
  return Math.floor(e) - Math.floor(s);
}, { lift: true });

define('DAYS360', 'Date & time', 'start_date, end_date, [method]', 'Days between two dates on a 360-day year.', (ctx, [a, b, c]) => {
  const s = serialArg(ctx, a);
  if (isError(s)) return s;
  const e = serialArg(ctx, b);
  if (isError(e)) return e;
  const eu = c !== undefined && !!argNum(ctx, c);
  const ps = serialToParts(s);
  const pe = serialToParts(e);
  let sd = ps.d;
  let ed = pe.d;
  if (eu) {
    if (sd === 31) sd = 30;
    if (ed === 31) ed = 30;
  } else {
    const lastFeb = (p: typeof ps) => p.m === 2 && p.d === daysInMonth(p.y, 2);
    if (lastFeb(ps)) sd = 30;
    if (sd >= 30 && ed === 31) ed = 30;
    if (sd === 31) sd = 30;
  }
  return (pe.y - ps.y) * 360 + (pe.m - ps.m) * 30 + (ed - sd);
});

function weekendMask(ctx: Ctx, w: Arg): boolean[] | ErrorValue {
  // index by dow (0 = Sunday)
  if (w === undefined) return [true, false, false, false, false, false, true];
  const v = argText(ctx, w);
  if (isError(v)) return v;
  if (/^[01]{7}$/.test(v)) {
    // string starts on Monday
    const mask = new Array(7).fill(false);
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = v[i] === '1';
    return mask;
  }
  const n = parseInt(v, 10);
  const pairs: Record<number, number[]> = { 1: [6, 0], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6] };
  const singles: Record<number, number> = { 11: 0, 12: 1, 13: 2, 14: 3, 15: 4, 16: 5, 17: 6 };
  const mask = new Array(7).fill(false);
  if (pairs[n]) for (const d of pairs[n]) mask[d] = true;
  else if (singles[n] !== undefined) mask[singles[n]] = true;
  else return Errors.num('That weekend code isn\'t valid.');
  return mask;
}

function holidaySet(ctx: Ctx, h: Arg): Set<number> | ErrorValue {
  if (h === undefined) return new Set();
  const nums = collectNumbers(ctx, [h]);
  if (isError(nums)) return nums;
  return new Set(nums.map(Math.floor));
}

function networkdays(intl: boolean) {
  return (ctx: Ctx, args: Arg[]): Value => {
    const s = serialArg(ctx, args[0]);
    if (isError(s)) return s;
    const e = serialArg(ctx, args[1]);
    if (isError(e)) return e;
    const mask = weekendMask(ctx, intl ? args[2] : undefined);
    if (isError(mask)) return mask;
    const hol = holidaySet(ctx, intl ? args[3] : args[2]);
    if (isError(hol)) return hol;
    let a = Math.floor(s);
    let b = Math.floor(e);
    const sign = a <= b ? 1 : -1;
    if (sign < 0) [a, b] = [b, a];
    let count = 0;
    for (let d = a; d <= b; d++) {
      const dow = serialToParts(d).dow;
      if (!mask[dow] && !hol.has(d)) count++;
    }
    return count * sign;
  };
}
define('NETWORKDAYS', 'Date & time', 'start_date, end_date, [holidays]', 'Counts working days (Mon-Fri) between two dates, excluding holidays.', networkdays(false));
define('NETWORKDAYS.INTL', 'Date & time', 'start_date, end_date, [weekend], [holidays]', 'Counts working days with custom weekends.', networkdays(true));

function workday(intl: boolean) {
  return (ctx: Ctx, args: Arg[]): Value => {
    const s = serialArg(ctx, args[0]);
    if (isError(s)) return s;
    const n = argInt(ctx, args[1]);
    if (isError(n)) return n;
    const mask = weekendMask(ctx, intl ? args[2] : undefined);
    if (isError(mask)) return mask;
    if (mask.every(Boolean)) return Errors.value();
    const hol = holidaySet(ctx, intl ? args[3] : args[2]);
    if (isError(hol)) return hol;
    let d = Math.floor(s);
    let left = Math.abs(n);
    const step = n >= 0 ? 1 : -1;
    while (left > 0) {
      d += step;
      if (!mask[serialToParts(d).dow] && !hol.has(d)) left--;
    }
    return d;
  };
}
define('WORKDAY', 'Date & time', 'start_date, days, [holidays]', 'Returns the date a number of working days away.', workday(false));
define('WORKDAY.INTL', 'Date & time', 'start_date, days, [weekend], [holidays]', 'Returns a working-day offset date with custom weekends.', workday(true));

define('DATEVALUE', 'Date & time', 'date_text', 'Converts a date written as text into a date number.', (ctx, [a]) => {
  const t = argText(ctx, a);
  if (isError(t)) return t;
  const d = parseDateTime(t);
  if (!d) return Errors.value(`"${t}" isn't a date Gridline recognizes. Try a format like 2025-03-14 or 3/14/2025.`);
  return Math.floor(d.value);
}, { lift: true });
define('TIMEVALUE', 'Date & time', 'time_text', 'Converts a time written as text into a time number.', (ctx, [a]) => {
  const t = argText(ctx, a);
  if (isError(t)) return t;
  const d = parseDateTime(t);
  if (!d) return Errors.value(`"${t}" isn't a time Gridline recognizes. Try 14:30 or 2:30 PM.`);
  return d.value - Math.floor(d.value);
}, { lift: true });

define('YEARFRAC', 'Date & time', 'start_date, end_date, [basis]', 'Returns the fraction of a year between two dates.', (ctx, [a, b, c]) => {
  let s = serialArg(ctx, a);
  if (isError(s)) return s;
  let e = serialArg(ctx, b);
  if (isError(e)) return e;
  const basis = argInt(ctx, c, 0);
  if (isError(basis)) return basis;
  s = Math.floor(s);
  e = Math.floor(e);
  if (s > e) [s, e] = [e, s];
  const ps = serialToParts(s);
  const pe = serialToParts(e);
  switch (basis) {
    case 0: {
      let sd = ps.d;
      let ed = pe.d;
      const lastFeb = (p: typeof ps) => p.m === 2 && p.d === daysInMonth(p.y, 2);
      if (lastFeb(ps) && lastFeb(pe)) ed = 30;
      if (lastFeb(ps)) sd = 30;
      if (ed === 31 && sd >= 30) ed = 30;
      if (sd === 31) sd = 30;
      return ((pe.y - ps.y) * 360 + (pe.m - ps.m) * 30 + (ed - sd)) / 360;
    }
    case 1: {
      const days = e - s;
      if (ps.y === pe.y) return days / (dateToSerial(ps.y + 1, 1, 1) - dateToSerial(ps.y, 1, 1));
      let total = 0;
      for (let y = ps.y; y <= pe.y; y++) total += dateToSerial(y + 1, 1, 1) - dateToSerial(y, 1, 1);
      return days / (total / (pe.y - ps.y + 1));
    }
    case 2:
      return (e - s) / 360;
    case 3:
      return (e - s) / 365;
    case 4: {
      const sd = Math.min(ps.d, 30);
      const ed = Math.min(pe.d, 30);
      return ((pe.y - ps.y) * 360 + (pe.m - ps.m) * 30 + (ed - sd)) / 360;
    }
  }
  return Errors.num('basis must be 0-4.');
});
