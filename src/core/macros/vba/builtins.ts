// Built-in VBA functions and constants. Anything that would reach outside the workbook (files,
// Shell, the registry, Windows API calls) is deliberately missing and reports a clear error.
import { dateToSerial, serialToParts } from '../../dates';
import { formatValue } from '../../numfmt';
import { bankersRound, compare, EMPTY, isObject, NULL, toBool, toDate, toInt, toNum, toStr, typeName, VbArray, VbDate, VbError, type Value } from './values';

export const MISSING = Symbol('Missing');

export interface BuiltinCtx {
  compareText: boolean;
  msgBox(prompt: string, buttons: number, title: string): Promise<number>;
  inputBox(prompt: string, title: string, dflt: string): Promise<string | null>;
  createObject(progId: string): Value;
  optionBase: number;
}

type Fn = (a: Value[], ctx: BuiltinCtx) => Value | Promise<Value>;

const arg = (a: Value[], i: number): Value | undefined => (a[i] === (MISSING as unknown) ? undefined : a[i]);
const has = (a: Value[], i: number) => arg(a, i) !== undefined;
const str = (a: Value[], i: number, d = '') => (has(a, i) ? toStr(a[i]) : d);
const num = (a: Value[], i: number, d = 0) => (has(a, i) ? toNum(a[i]) : d);

function need(a: Value[], n: number, name: string): void {
  for (let i = 0; i < n; i++) if (!has(a, i)) throw new VbError(450, `${name} needs ${n} argument${n === 1 ? '' : 's'}.`);
}

function now(): VbDate {
  const d = new Date();
  return new VbDate(dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400);
}

/** VBA format names and codes, translated to spreadsheet number formats. */
function vbaFormat(v: Value, fmt: string): string {
  const named: Record<string, string> = {
    'general number': 'General',
    currency: '"$"#,##0.00',
    fixed: '0.00',
    standard: '#,##0.00',
    percent: '0.00%',
    scientific: '0.00E+00',
    'short date': 'm/d/yyyy',
    'medium date': 'd-mmm-yy',
    'long date': 'dddd, mmmm d, yyyy',
    'short time': 'hh:mm',
    'medium time': 'h:mm AM/PM',
    'long time': 'h:mm:ss AM/PM',
    'general date': 'm/d/yyyy h:mm:ss AM/PM',
  };
  const lower = fmt.toLowerCase();
  if (lower === 'yes/no') return toBool(v) ? 'Yes' : 'No';
  if (lower === 'true/false') return toBool(v) ? 'True' : 'False';
  if (lower === 'on/off') return toBool(v) ? 'On' : 'Off';
  let f = named[lower] ?? fmt;
  // VBA uses "n" for minutes
  f = f.replace(/n/g, (m, i) => (/"/.test(f.slice(0, i).replace(/"[^"]*"/g, '')) ? m : 'm'));
  if (v instanceof VbDate) return formatValue(v.serial, f).text;
  if (typeof v === 'string') {
    const n = Number(v);
    if (!Number.isNaN(n) && v.trim()) return formatValue(n, f).text;
    try {
      return formatValue(toDate(v).serial, f).text;
    } catch {
      return v;
    }
  }
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (v === EMPTY || v === undefined) return '';
  return formatValue(toNum(v), f).text;
}

function dateAdd(interval: string, n: number, d: VbDate): VbDate {
  const p = serialToParts(Math.floor(d.serial));
  const t = d.serial - Math.floor(d.serial);
  const addMonths = (m: number) => {
    const total = p.y * 12 + (p.m - 1) + m;
    const y = Math.floor(total / 12);
    const mo = (total % 12) + 1;
    const dim = new Date(Date.UTC(y, mo, 0)).getUTCDate();
    return new VbDate(dateToSerial(y, mo, Math.min(p.d, dim)) + t);
  };
  switch (interval.toLowerCase()) {
    case 'yyyy':
      return addMonths(12 * n);
    case 'q':
      return addMonths(3 * n);
    case 'm':
      return addMonths(n);
    case 'y':
    case 'd':
    case 'w':
      return new VbDate(d.serial + n);
    case 'ww':
      return new VbDate(d.serial + 7 * n);
    case 'h':
      return new VbDate(d.serial + n / 24);
    case 'n':
      return new VbDate(d.serial + n / 1440);
    case 's':
      return new VbDate(d.serial + n / 86400);
  }
  throw new VbError(5, `"${interval}" isn't a valid interval (use yyyy, q, m, d, ww, h, n or s).`);
}

function dateDiff(interval: string, a: VbDate, b: VbDate): number {
  const pa = serialToParts(Math.floor(a.serial));
  const pb = serialToParts(Math.floor(b.serial));
  switch (interval.toLowerCase()) {
    case 'yyyy':
      return pb.y - pa.y;
    case 'q':
      return pb.y * 4 + Math.floor((pb.m - 1) / 3) - (pa.y * 4 + Math.floor((pa.m - 1) / 3));
    case 'm':
      return pb.y * 12 + pb.m - (pa.y * 12 + pa.m);
    case 'y':
    case 'd':
      return Math.floor(b.serial) - Math.floor(a.serial);
    case 'w':
    case 'ww':
      return Math.trunc((Math.floor(b.serial) - Math.floor(a.serial)) / 7);
    case 'h':
      return Math.trunc((b.serial - a.serial) * 24);
    case 'n':
      return Math.trunc((b.serial - a.serial) * 1440);
    case 's':
      return Math.round((b.serial - a.serial) * 86400);
  }
  throw new VbError(5, `"${interval}" isn't a valid interval.`);
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function weekday(serial: number, first = 1): number {
  // Serial 1 (1900-01-01) is a Sunday in Excel's calendar
  const dow = (((Math.floor(serial) - 1) % 7) + 7) % 7; // 0 = Sunday
  return ((dow - (first - 1) + 7) % 7) + 1;
}

function toArray1(v: Value): VbArray {
  if (v instanceof VbArray) return v;
  throw new VbError(13, 'Type mismatch: expected an array.');
}

function blocked(what: string): Fn {
  return () => {
    throw new VbError(70, `${what} isn't available: Gridline macros can only work with the workbook, not files or programs on your computer.`);
  };
}

export const BUILTINS: Record<string, Fn> = {
  // ----- text -----
  len: (a) => (a[0] === NULL ? NULL : str(a, 0).length),
  left: (a) => str(a, 0).slice(0, Math.max(0, num(a, 1))),
  right: (a) => {
    const s = str(a, 0);
    const n = Math.max(0, num(a, 1));
    return n === 0 ? '' : s.slice(-n);
  },
  mid: (a) => {
    const s = str(a, 0);
    const start = Math.max(1, num(a, 1, 1));
    return has(a, 2) ? s.substr(start - 1, Math.max(0, num(a, 2))) : s.slice(start - 1);
  },
  lcase: (a) => str(a, 0).toLowerCase(),
  ucase: (a) => str(a, 0).toUpperCase(),
  trim: (a) => str(a, 0).trim(),
  ltrim: (a) => str(a, 0).replace(/^\s+/, ''),
  rtrim: (a) => str(a, 0).replace(/\s+$/, ''),
  space: (a) => ' '.repeat(Math.max(0, num(a, 0))),
  string: (a) => {
    const c = typeof a[1] === 'number' ? String.fromCharCode(a[1]) : str(a, 1).charAt(0);
    return c.repeat(Math.max(0, num(a, 0)));
  },
  chr: (a) => String.fromCharCode(num(a, 0)),
  chrw: (a) => String.fromCharCode(num(a, 0)),
  asc: (a) => {
    const s = str(a, 0);
    if (!s) throw new VbError(5, 'Asc needs a non-empty text.');
    return s.charCodeAt(0);
  },
  ascw: (a) => str(a, 0).charCodeAt(0),
  instr: (a, ctx) => {
    // InStr([start], text, find, [compare])
    let start = 1;
    let i = 0;
    if (a.length >= 3 || typeof a[0] === 'number') {
      start = num(a, 0, 1);
      i = 1;
    }
    const s = str(a, i);
    const f = str(a, i + 1);
    const text = has(a, i + 2) ? num(a, i + 2) === 1 : ctx.compareText;
    if (start > s.length + 1) return 0;
    return (text ? s.toLowerCase() : s).indexOf(text ? f.toLowerCase() : f, start - 1) + 1;
  },
  instrrev: (a, ctx) => {
    const s = str(a, 0);
    const f = str(a, 1);
    const start = has(a, 2) && num(a, 2) > 0 ? num(a, 2) : s.length;
    const text = has(a, 3) ? num(a, 3) === 1 : ctx.compareText;
    return (text ? s.toLowerCase() : s).lastIndexOf(text ? f.toLowerCase() : f, start - f.length) + 1;
  },
  replace: (a, ctx) => {
    const s = str(a, 0);
    const f = str(a, 1);
    const r = str(a, 2);
    if (!f) return s;
    const start = num(a, 3, 1);
    const count = num(a, 4, -1);
    const text = has(a, 5) ? num(a, 5) === 1 : ctx.compareText;
    let out = '';
    let rest = s.slice(start - 1);
    let n = 0;
    for (;;) {
      const idx = (text ? rest.toLowerCase() : rest).indexOf(text ? f.toLowerCase() : f);
      if (idx < 0 || (count >= 0 && n >= count)) break;
      out += rest.slice(0, idx) + r;
      rest = rest.slice(idx + f.length);
      n++;
    }
    return out + rest;
  },
  strreverse: (a) => [...str(a, 0)].reverse().join(''),
  strcomp: (a, ctx) => {
    const text = has(a, 2) ? num(a, 2) === 1 : ctx.compareText;
    return compare(str(a, 0), str(a, 1), text);
  },
  strconv: (a) => {
    const s = str(a, 0);
    const mode = num(a, 1);
    if (mode === 1) return s.toUpperCase();
    if (mode === 2) return s.toLowerCase();
    if (mode === 3) return s.toLowerCase().replace(/(^|\s)\S/g, (m) => m.toUpperCase());
    return s;
  },
  split: (a) => {
    const s = str(a, 0);
    const d = has(a, 1) ? str(a, 1) : ' ';
    let parts = s === '' ? [] : d === '' ? [s] : s.split(d);
    const limit = num(a, 2, -1);
    if (limit > 0 && parts.length > limit) parts = [...parts.slice(0, limit - 1), parts.slice(limit - 1).join(d)];
    const arr = new VbArray([{ lb: 0, ub: parts.length - 1 }]);
    arr.data = parts;
    return arr;
  },
  join: (a) => toArray1(a[0]).data.map((v) => toStr(v)).join(has(a, 1) ? str(a, 1) : ' '),
  format: (a) => (has(a, 1) ? vbaFormat(a[0], str(a, 1)) : a[0] instanceof VbDate ? toStr(a[0]) : toStr(a[0] ?? EMPTY)),
  formatnumber: (a) => formatValue(toNum(a[0]), `#,##0${num(a, 1, 2) > 0 ? '.' + '0'.repeat(num(a, 1, 2)) : ''}`).text,
  formatcurrency: (a) => formatValue(toNum(a[0]), `"$"#,##0${num(a, 1, 2) > 0 ? '.' + '0'.repeat(num(a, 1, 2)) : ''}`).text,
  formatpercent: (a) => formatValue(toNum(a[0]), `0${num(a, 1, 2) > 0 ? '.' + '0'.repeat(num(a, 1, 2)) : ''}%`).text,
  formatdatetime: (a) => vbaFormat(a[0], ['general date', 'long date', 'short date', 'long time', 'short time'][num(a, 1, 0)] ?? 'general date'),
  // ----- conversion -----
  cstr: (a) => str(a, 0),
  cint: (a) => toInt(a[0], 32767),
  clng: (a) => toInt(a[0]),
  clnglng: (a) => toInt(a[0], Number.MAX_SAFE_INTEGER),
  cbyte: (a) => {
    const n = toInt(a[0]);
    if (n < 0 || n > 255) throw new VbError(6, 'Overflow: a Byte must be 0 to 255.');
    return n;
  },
  cdbl: (a) => toNum(a[0]),
  csng: (a) => Math.fround(toNum(a[0])),
  ccur: (a) => Math.round(toNum(a[0]) * 10000) / 10000,
  cdec: (a) => toNum(a[0]),
  cbool: (a) => toBool(a[0]),
  cdate: (a) => toDate(a[0]),
  cvdate: (a) => toDate(a[0]),
  cvar: (a) => a[0],
  val: (a) => {
    const m = /^\s*[+-]?(\d[\d\s]*\.?\d*|\.\d+)(e[+-]?\d+)?/i.exec(str(a, 0));
    return m ? parseFloat(m[0].replace(/\s/g, '')) : 0;
  },
  str: (a) => {
    const n = toNum(a[0]);
    return (n >= 0 ? ' ' : '') + toStr(n);
  },
  hex: (a) => toInt(a[0]).toString(16).toUpperCase(),
  oct: (a) => toInt(a[0]).toString(8),
  // ----- math -----
  abs: (a) => Math.abs(toNum(a[0])),
  int: (a) => Math.floor(toNum(a[0])),
  fix: (a) => Math.trunc(toNum(a[0])),
  round: (a) => bankersRound(toNum(a[0]), num(a, 1)),
  sqr: (a) => {
    const n = toNum(a[0]);
    if (n < 0) throw new VbError(5, "Sqr can't take a negative number.");
    return Math.sqrt(n);
  },
  exp: (a) => Math.exp(toNum(a[0])),
  log: (a) => {
    const n = toNum(a[0]);
    if (n <= 0) throw new VbError(5, 'Log needs a positive number.');
    return Math.log(n);
  },
  sin: (a) => Math.sin(toNum(a[0])),
  cos: (a) => Math.cos(toNum(a[0])),
  tan: (a) => Math.tan(toNum(a[0])),
  atn: (a) => Math.atan(toNum(a[0])),
  sgn: (a) => Math.sign(toNum(a[0])),
  rnd: () => Math.random(),
  randomize: () => EMPTY,
  // ----- dates -----
  now: () => now(),
  date: () => new VbDate(Math.floor(now().serial)),
  time: () => new VbDate(now().serial - Math.floor(now().serial)),
  timer: () => {
    const d = new Date();
    return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000;
  },
  year: (a) => serialToParts(Math.floor(toDate(a[0]).serial)).y,
  month: (a) => serialToParts(Math.floor(toDate(a[0]).serial)).m,
  day: (a) => serialToParts(Math.floor(toDate(a[0]).serial)).d,
  hour: (a) => Math.floor((toDate(a[0]).serial % 1) * 24 + 1e-9),
  minute: (a) => Math.floor((((toDate(a[0]).serial % 1) * 1440) % 60) + 1e-9),
  second: (a) => Math.round((toDate(a[0]).serial % 1) * 86400) % 60,
  weekday: (a) => weekday(toDate(a[0]).serial, num(a, 1, 1)),
  weekdayname: (a) => {
    const n = num(a, 0);
    const first = num(a, 2, 1);
    const name = DAYS[(n - 1 + first - 1) % 7] ?? '';
    return a[1] !== undefined && toBool(a[1]) ? name.slice(0, 3) : name;
  },
  monthname: (a) => {
    const name = MONTHS[num(a, 0) - 1];
    if (!name) throw new VbError(5, 'Month must be 1 to 12.');
    return has(a, 1) && toBool(a[1]) ? name.slice(0, 3) : name;
  },
  dateserial: (a) => {
    need(a, 3, 'DateSerial');
    const y = num(a, 0);
    const m = num(a, 1);
    const d = num(a, 2);
    const base = Date.UTC(y < 100 ? (y < 30 ? 2000 + y : 1900 + y) : y, m - 1, d);
    return new VbDate(Math.round((base - Date.UTC(1899, 11, 30)) / 86400000));
  },
  timeserial: (a) => new VbDate((num(a, 0) * 3600 + num(a, 1) * 60 + num(a, 2)) / 86400),
  datevalue: (a) => new VbDate(Math.floor(toDate(a[0]).serial)),
  timevalue: (a) => {
    const s = toDate(a[0]).serial;
    return new VbDate(s - Math.floor(s));
  },
  dateadd: (a) => dateAdd(str(a, 0), num(a, 1), toDate(a[2])),
  datediff: (a) => dateDiff(str(a, 0), toDate(a[1]), toDate(a[2])),
  datepart: (a) => {
    const d = toDate(a[1]);
    const p = serialToParts(Math.floor(d.serial));
    switch (str(a, 0).toLowerCase()) {
      case 'yyyy':
        return p.y;
      case 'q':
        return Math.floor((p.m - 1) / 3) + 1;
      case 'm':
        return p.m;
      case 'd':
        return p.d;
      case 'w':
        return weekday(d.serial);
      case 'y':
        return Math.floor(d.serial) - dateToSerial(p.y, 1, 1) + 1;
      case 'h':
        return Math.floor((d.serial % 1) * 24 + 1e-9);
      case 'n':
        return Math.floor((((d.serial % 1) * 1440) % 60) + 1e-9);
      case 's':
        return Math.round((d.serial % 1) * 86400) % 60;
    }
    throw new VbError(5, 'Invalid interval.');
  },
  // ----- arrays -----
  array: (a) => {
    const arr = new VbArray([{ lb: 0, ub: a.length - 1 }]);
    arr.data = a.map((v) => (v === (MISSING as unknown) ? EMPTY : v));
    return arr;
  },
  ubound: (a) => {
    const arr = toArray1(a[0]);
    const d = arr.dims[num(a, 1, 1) - 1];
    if (!d) throw new VbError(9, 'Subscript out of range: the array has no such dimension.');
    return d.ub;
  },
  lbound: (a) => {
    const arr = toArray1(a[0]);
    const d = arr.dims[num(a, 1, 1) - 1];
    if (!d) throw new VbError(9, 'Subscript out of range: the array has no such dimension.');
    return d.lb;
  },
  filter: (a, ctx) => {
    const src = toArray1(a[0]);
    const f = str(a, 1);
    const include = has(a, 2) ? toBool(a[2]) : true;
    const text = has(a, 3) ? num(a, 3) === 1 : ctx.compareText;
    const keep = src.data.map((v) => toStr(v)).filter((s) => (text ? s.toLowerCase().includes(f.toLowerCase()) : s.includes(f)) === include);
    const arr = new VbArray([{ lb: 0, ub: keep.length - 1 }]);
    arr.data = keep;
    return arr;
  },
  // ----- information -----
  isnumeric: (a) => {
    const v = a[0];
    if (typeof v === 'number' || typeof v === 'boolean') return true;
    if (typeof v === 'string') return /^\s*[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?\s*$/i.test(v.replace(/[,$]/g, ''));
    return v === EMPTY;
  },
  isdate: (a) => {
    if (a[0] instanceof VbDate) return true;
    if (typeof a[0] !== 'string' || !/[/:-]|\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i.test(a[0])) return false;
    try {
      toDate(a[0]);
      return true;
    } catch {
      return false;
    }
  },
  isempty: (a) => a[0] === EMPTY || a[0] === undefined,
  isnull: (a) => a[0] === NULL,
  isarray: (a) => a[0] instanceof VbArray,
  isobject: (a) => a[0] === null || isObject(a[0]),
  ismissing: (a) => a[0] === (MISSING as unknown),
  iserror: (a) => isObject(a[0]) && a[0].vbType === 'Error',
  typename: (a) => typeName(a[0]),
  vartype: (a) => {
    const v = a[0];
    if (v === EMPTY || v === undefined) return 0;
    if (v === NULL) return 1;
    if (typeof v === 'number') return Number.isInteger(v) ? (Math.abs(v) <= 32767 ? 2 : 3) : 5;
    if (v instanceof VbDate) return 7;
    if (typeof v === 'string') return 8;
    if (typeof v === 'boolean') return 11;
    if (v instanceof VbArray) return 8204;
    return 9;
  },
  // ----- choice -----
  iif: (a) => (toBool(a[0]) ? a[1] : a[2]),
  choose: (a) => {
    const i = Math.floor(toNum(a[0]));
    return i >= 1 && i < a.length ? a[i] : NULL;
  },
  switch: (a) => {
    for (let i = 0; i + 1 < a.length; i += 2) if (toBool(a[i])) return a[i + 1];
    return NULL;
  },
  nz: (a) => (a[0] === NULL || a[0] === EMPTY ? (has(a, 1) ? a[1] : 0) : a[0]),
  // ----- colors -----
  rgb: (a) => (Math.max(0, Math.min(255, num(a, 0))) | 0) + ((Math.max(0, Math.min(255, num(a, 1))) | 0) << 8) + ((Math.max(0, Math.min(255, num(a, 2))) | 0) << 16),
  qbcolor: (a) => [0, 0x800000, 0x8000, 0x808000, 0x80, 0x800080, 0x8080, 0xc0c0c0, 0x808080, 0xff0000, 0xff00, 0xffff00, 0xff, 0xff00ff, 0xffff, 0xffffff][num(a, 0)] ?? 0,
  // ----- interaction -----
  msgbox: async (a, ctx) => ctx.msgBox(str(a, 0), num(a, 1, 0), str(a, 2, 'Gridline')),
  inputbox: async (a, ctx) => (await ctx.inputBox(str(a, 0), str(a, 1, 'Gridline'), str(a, 2))) ?? '',
  createobject: (a, ctx) => ctx.createObject(str(a, 0)),
  getobject: blocked('GetObject'),
  doevents: () => 0,
  beep: () => EMPTY,
  environ: () => '',
  cverr: (a) => {
    const n = num(a, 0);
    return { vbType: 'Error', get: () => n, set: () => undefined, dflt: () => n } as Value;
  },
  // ----- outside the workbook: not available -----
  shell: blocked('Shell'),
  kill: blocked('Kill'),
  mkdir: blocked('MkDir'),
  rmdir: blocked('RmDir'),
  dir: blocked('Dir'),
  filecopy: blocked('FileCopy'),
  filelen: blocked('FileLen'),
  freefile: blocked('Opening files'),
  curdir: blocked('CurDir'),
  chdir: blocked('ChDir'),
  sendkeys: blocked('SendKeys'),
  appactivate: blocked('AppActivate'),
  savesetting: blocked('SaveSetting'),
  getsetting: blocked('GetSetting'),
  deletesetting: blocked('DeleteSetting'),
};

const XL: Record<string, number> = {
  // VBA
  vbokonly: 0,
  vbokcancel: 1,
  vbabortretryignore: 2,
  vbyesnocancel: 3,
  vbyesno: 4,
  vbretrycancel: 5,
  vbcritical: 16,
  vbquestion: 32,
  vbexclamation: 48,
  vbinformation: 64,
  vbdefaultbutton1: 0,
  vbdefaultbutton2: 256,
  vbdefaultbutton3: 512,
  vbok: 1,
  vbcancel: 2,
  vbabort: 3,
  vbretry: 4,
  vbignore: 5,
  vbyes: 6,
  vbno: 7,
  vbblack: 0,
  vbred: 0xff,
  vbgreen: 0xff00,
  vbyellow: 0xffff,
  vbblue: 0xff0000,
  vbmagenta: 0xff00ff,
  vbcyan: 0xffff00,
  vbwhite: 0xffffff,
  vbbinarycompare: 0,
  vbtextcompare: 1,
  vbuppercase: 1,
  vblowercase: 2,
  vbpropercase: 3,
  vbsunday: 1,
  vbmonday: 2,
  vbtuesday: 3,
  vbwednesday: 4,
  vbthursday: 5,
  vbfriday: 6,
  vbsaturday: 7,
  vbuseSystemDayOfWeek: 0,
  vbempty: 0,
  vbnull: 1,
  vbinteger: 2,
  vblong: 3,
  vbsingle: 4,
  vbdouble: 5,
  vbcurrency: 6,
  vbdate: 7,
  vbstring: 8,
  vbobject: 9,
  vberror: 10,
  vbboolean: 11,
  vbvariant: 12,
  vbarray: 8192,
  vbgeneraldate: 0,
  vblongdate: 1,
  vbshortdate: 2,
  vblongtime: 3,
  vbshorttime: 4,
  vbobjecterror: -2147221504,
  // Excel
  xlup: -4162,
  xldown: -4121,
  xltoleft: -4159,
  xltoright: -4161,
  xlascending: 1,
  xldescending: 2,
  xlyes: 1,
  xlno: 2,
  xlguess: 0,
  xlcenter: -4108,
  xlleft: -4131,
  xlright: -4152,
  xltop: -4160,
  xlbottom: -4107,
  xlgeneral: 1,
  xljustify: -4130,
  xlcenteracrossselection: 7,
  xlvalues: -4163,
  xlformulas: -4123,
  xlcomments: -4144,
  xlpasteall: -4104,
  xlpastevalues: -4163,
  xlpasteformats: -4122,
  xlpasteformulas: -4123,
  xlpastevaluesandnumberformats: 12,
  xlpastecolumnwidths: 8,
  xlcalculationmanual: -4135,
  xlcalculationautomatic: -4105,
  xlcalculationsemiautomatic: 2,
  xlcontinuous: 1,
  xldash: -4115,
  xldot: -4118,
  xldouble: -4119,
  xllinestylenone: -4142,
  xlhairline: 1,
  xlthin: 2,
  xlmedium: -4138,
  xlthick: 4,
  xlnone: -4142,
  xlautomatic: -4105,
  xlcolorindexnone: -4142,
  xlcolorindexautomatic: -4105,
  xledgetop: 8,
  xledgebottom: 9,
  xledgeleft: 7,
  xledgeright: 10,
  xlinsidevertical: 11,
  xlinsidehorizontal: 12,
  xldiagonaldown: 5,
  xldiagonalup: 6,
  xlcelltypelastcell: 11,
  xlcelltypeblanks: 4,
  xlcelltypeconstants: 2,
  xlcelltypeformulas: -4123,
  xlcelltypevisible: 12,
  xllastcell: 11,
  xlwhole: 1,
  xlpart: 2,
  xlbyrows: 1,
  xlbycolumns: 2,
  xlnext: 1,
  xlprevious: 2,
  xlshiftup: -4162,
  xlshiftdown: -4121,
  xlshifttoleft: -4159,
  xlshifttoright: -4161,
  xlsheetvisible: -1,
  xlsheethidden: 0,
  xlsheetveryhidden: 2,
  xlunderlinestylesingle: 2,
  xlunderlinestyledouble: -4119,
  xlunderlinestylenone: -4142,
  xland: 1,
  xlor: 2,
  xlfiltervalues: 7,
  xltop10items: 3,
  xlsortonvalues: 0,
  xlsortnormal: 0,
  xltoptobottom: 1,
  xllefttoright: 2,
  xlnumbers: 1,
  xltextvalues: 2,
  xllogical: 4,
  xlerrors: 16,
  xlwbatworksheet: -4167,
  xlworksheet: -4167,
  xlopenxmlworkbookmacroenabled: 52,
  xlopenxmlworkbook: 51,
  xlcsv: 6,
  xlmaximized: -4137,
  xlnormal: -4143,
  xlminimized: -4140,
  xlr1c1: -4150,
  xla1: 1,
  xlerrdiv0: 2007,
  xlerrna: 2042,
  xlerrname: 2029,
  xlerrnull: 2000,
  xlerrnum: 2036,
  xlerrref: 2023,
  xlerrvalue: 2015,
};

export const CONSTANTS: Record<string, Value> = {
  ...XL,
  vbcrlf: '\r\n',
  vbcr: '\r',
  vblf: '\n',
  vbnewline: '\r\n',
  vbtab: '\t',
  vbnullstring: '',
  vbnullchar: '\0',
  vbback: '\b',
  vbformfeed: '\f',
  vbverticaltab: '\v',
};

export { EMPTY, MISSING as MISSING_ARG };
