// Runtime values for the VBA interpreter, and VBA's conversion and comparison rules.
import { serialToParts } from '../../dates';

export const EMPTY = Symbol('Empty');
export const NULL = Symbol('Null');

export class VbDate {
  constructor(readonly serial: number) {}
}

/** Error raised by running code; `number` follows VBA's runtime error numbers. */
export class VbError extends Error {
  line = 0;
  proc = '';
  module = '';
  constructor(
    readonly number: number,
    message: string,
  ) {
    super(message);
  }
}

/** Stops the whole run (the End statement, or the user pressing Stop). */
export class StopSignal {
  constructor(readonly reason: 'end' | 'cancel') {}
}

/** An object exposed to VBA code (Range, Worksheet, Collection…). */
export interface VbObject {
  readonly vbType: string;
  /** Property read or method call. `args` are already evaluated. */
  get(name: string, args: Value[], named?: Record<string, Value>): Value | Promise<Value>;
  /** Property write. */
  set(name: string, args: Value[], v: Value): void | Promise<void>;
  /** Default member (`rng(1, 1)`, `x = rng`). */
  dflt?(args: Value[]): Value | Promise<Value>;
  setDflt?(args: Value[], v: Value): void | Promise<void>;
  /** Items for For Each. */
  items?(): Value[];
}

export class VbArray {
  data: Value[];
  constructor(public dims: { lb: number; ub: number }[], fill?: () => Value) {
    const n = dims.reduce((s, d) => s * Math.max(0, d.ub - d.lb + 1), 1);
    this.data = Array.from({ length: n }, fill ?? (() => EMPTY));
  }

  index(idx: number[]): number {
    if (idx.length !== this.dims.length) throw new VbError(9, 'Subscript out of range (wrong number of dimensions).');
    let pos = 0;
    for (let i = 0; i < idx.length; i++) {
      const d = this.dims[i];
      const k = Math.round(idx[i]);
      if (k < d.lb || k > d.ub) throw new VbError(9, `Subscript out of range: index ${k} is outside ${d.lb} to ${d.ub}.`);
      pos = pos * (d.ub - d.lb + 1) + (k - d.lb);
    }
    return pos;
  }

  get(idx: number[]): Value {
    return this.data[this.index(idx)];
  }

  set(idx: number[], v: Value): void {
    this.data[this.index(idx)] = v;
  }

  /** ReDim Preserve: only the last dimension may change. */
  resize(dims: { lb: number; ub: number }[], preserve: boolean): void {
    if (!preserve || !this.dims.length) {
      this.dims = dims;
      this.data = new VbArray(dims).data;
      return;
    }
    const old = this;
    const next = new VbArray(dims);
    if (dims.length === 1 && old.dims.length === 1) {
      for (let k = Math.max(dims[0].lb, old.dims[0].lb); k <= Math.min(dims[0].ub, old.dims[0].ub); k++) next.set([k], old.get([k]));
    } else if (dims.length === 2 && old.dims.length === 2) {
      for (let r = old.dims[0].lb; r <= old.dims[0].ub; r++)
        for (let c = Math.max(dims[1].lb, old.dims[1].lb); c <= Math.min(dims[1].ub, old.dims[1].ub); c++) next.set([r, c], old.get([r, c]));
    }
    this.dims = dims;
    this.data = next.data;
  }
}

export type Value = number | string | boolean | VbDate | VbArray | VbObject | null | typeof EMPTY | typeof NULL | undefined;

// ---------- type checks ----------

export const isObject = (v: Value): v is VbObject => typeof v === 'object' && v !== null && !(v instanceof VbDate) && !(v instanceof VbArray);

export function typeName(v: Value): string {
  if (v === EMPTY || v === undefined) return 'Empty';
  if (v === NULL) return 'Null';
  if (v === null) return 'Nothing';
  if (typeof v === 'number') return Number.isInteger(v) && Math.abs(v) <= 2147483647 ? (Math.abs(v) <= 32767 ? 'Integer' : 'Long') : 'Double';
  if (typeof v === 'string') return 'String';
  if (typeof v === 'boolean') return 'Boolean';
  if (v instanceof VbDate) return 'Date';
  if (v instanceof VbArray) return 'Variant()';
  return v.vbType;
}

// ---------- conversions ----------

/** Number shown the way VBA prints it (15 significant digits, no trailing zeros). */
export function numText(n: number): string {
  if (!Number.isFinite(n)) return n > 0 ? 'Infinity' : n < 0 ? '-Infinity' : 'NaN';
  if (n === 0) return '0';
  const p = Number(n.toPrecision(15));
  const abs = Math.abs(p);
  if (abs >= 1e15 || abs < 1e-4) {
    const [m, e] = p.toExponential().split('e');
    const exp = parseInt(e, 10);
    return `${m.replace(/\.?0+$/, '')}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  return String(p);
}

export function dateText(d: VbDate): string {
  const whole = Math.floor(d.serial);
  const frac = d.serial - whole;
  const p = serialToParts(whole);
  const date = whole === 0 && frac ? '' : `${p.m}/${p.d}/${p.y}`;
  if (!frac) return date;
  let secs = Math.round(frac * 86400);
  const h = Math.floor(secs / 3600);
  secs -= h * 3600;
  const m = Math.floor(secs / 60);
  const s = secs - m * 60;
  const time = `${((h + 11) % 12) + 1}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')} ${h < 12 ? 'AM' : 'PM'}`;
  return date ? `${date} ${time}` : time;
}

export function toStr(v: Value): string {
  if (v === EMPTY || v === undefined) return '';
  if (v === NULL) throw new VbError(94, 'Invalid use of Null.');
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numText(v);
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (v instanceof VbDate) return dateText(v);
  if (v === null) throw new VbError(91, 'Object variable not set (Nothing).');
  if (v instanceof VbArray) throw new VbError(13, 'Type mismatch: an array can’t be used as text.');
  throw new VbError(13, `Type mismatch: a ${v.vbType} can’t be used as text.`);
}

/** Parses text as a number the way VBA's implicit conversion does (or returns null). */
export function parseNum(s: string): number | null {
  const t = s.trim();
  if (!t) return null;
  if (/^&H[0-9a-f]+$/i.test(t)) return parseInt(t.slice(2), 16);
  const c = t.replace(/[,$]/g, '');
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?%?$/i.test(c)) return c.endsWith('%') ? parseFloat(c) / 100 : parseFloat(c);
  if (/^\((\d+\.?\d*)\)$/.test(c)) return -parseFloat(c.slice(1, -1));
  return null;
}

export function toNum(v: Value): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? -1 : 0;
  if (v === EMPTY || v === undefined) return 0;
  if (v instanceof VbDate) return v.serial;
  if (typeof v === 'string') {
    const n = parseNum(v);
    if (n === null) {
      if (/^(true|false)$/i.test(v.trim())) return /^true$/i.test(v.trim()) ? -1 : 0;
      throw new VbError(13, `Type mismatch: "${v.length > 40 ? v.slice(0, 40) + '…' : v}" isn't a number.`);
    }
    return n;
  }
  if (v === NULL) throw new VbError(94, 'Invalid use of Null.');
  if (v === null) throw new VbError(91, 'Object variable not set (Nothing).');
  throw new VbError(13, `Type mismatch: a ${typeName(v)} isn't a number.`);
}

export function toBool(v: Value): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'string') {
    const t = v.trim().toLowerCase();
    if (t === 'true') return true;
    if (t === 'false') return false;
  }
  return toNum(v) !== 0;
}

/** VBA rounds to even ("banker's rounding") when converting to whole numbers. */
export function bankersRound(n: number, digits = 0): number {
  const f = 10 ** digits;
  const x = n * f;
  const r = Math.round(x);
  const diff = Math.abs(x - Math.trunc(x));
  if (Math.abs(diff - 0.5) < 1e-9) return (2 * Math.round(x / 2)) / f;
  return r / f;
}

export function toInt(v: Value, max = 2147483647): number {
  const n = bankersRound(toNum(v));
  if (n > max || n < -max - 1) throw new VbError(6, 'Overflow: the number is too large.');
  return n;
}

export function toDate(v: Value): VbDate {
  if (v instanceof VbDate) return v;
  if (typeof v === 'number') return new VbDate(v);
  if (typeof v === 'string') {
    const t = v.trim();
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
    const us = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?)?$/i.exec(t);
    const timeOnly = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(t);
    const mk = (y: number, m: number, d: number, h = 0, mi = 0, s = 0) => new VbDate(Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000) + (h * 3600 + mi * 60 + s) / 86400);
    if (iso) return mk(+iso[1], +iso[2], +iso[3], +(iso[4] ?? 0), +(iso[5] ?? 0), +(iso[6] ?? 0));
    if (us) {
      let y = +us[3];
      if (y < 100) y += y < 30 ? 2000 : 1900;
      let h = +(us[4] ?? 0);
      if (us[7]) h = (h % 12) + (/pm/i.test(us[7]) ? 12 : 0);
      return mk(y, +us[1], +us[2], h, +(us[5] ?? 0), +(us[6] ?? 0));
    }
    if (timeOnly) {
      let h = +timeOnly[1];
      if (timeOnly[4]) h = (h % 12) + (/pm/i.test(timeOnly[4]) ? 12 : 0);
      return new VbDate((h * 3600 + +timeOnly[2] * 60 + +(timeOnly[3] ?? 0)) / 86400);
    }
    const parsed = Date.parse(t);
    if (!Number.isNaN(parsed)) {
      const d = new Date(parsed);
      return mk(d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds());
    }
  }
  throw new VbError(13, `Type mismatch: "${toStrSafe(v)}" isn't a date.`);
}

export function toStrSafe(v: Value): string {
  try {
    return toStr(v);
  } catch {
    return typeName(v);
  }
}

// ---------- comparison ----------

function isNumericLike(v: Value): boolean {
  return typeof v === 'number' || typeof v === 'boolean' || v instanceof VbDate;
}

/** Compares two values: negative, zero or positive. Text compares case-sensitively unless `text` is set. */
export function compare(a: Value, b: Value, text = false): number {
  if (a === NULL || b === NULL) throw new VbError(94, 'Invalid use of Null.');
  if (a === undefined) a = EMPTY;
  if (b === undefined) b = EMPTY;
  if (a === EMPTY && b === EMPTY) return 0;
  if (a === EMPTY) a = typeof b === 'string' ? '' : 0;
  if (b === EMPTY) b = typeof a === 'string' ? '' : 0;
  if (isNumericLike(a) && isNumericLike(b)) return Math.sign(toNum(a) - toNum(b));
  if (typeof a === 'string' && typeof b === 'string') {
    const x = text ? a.toLowerCase() : a;
    const y = text ? b.toLowerCase() : b;
    return x < y ? -1 : x > y ? 1 : 0;
  }
  // Number vs text: the text must look like a number
  if (isNumericLike(a) && typeof b === 'string') {
    const n = parseNum(b);
    if (n === null) return -1; // numbers sort before text, as in Excel
    return Math.sign(toNum(a) - n);
  }
  if (typeof a === 'string' && isNumericLike(b)) return -compare(b, a, text);
  throw new VbError(13, `Type mismatch: can't compare a ${typeName(a)} with a ${typeName(b)}.`);
}

/** VBA's Like operator: * ? # [abc] [!abc] [a-z]. */
export function like(s: string, pattern: string, text = false): boolean {
  let re = '^';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === '*') re += '[\\s\\S]*';
    else if (c === '?') re += '[\\s\\S]';
    else if (c === '#') re += '\\d';
    else if (c === '[') {
      const end = pattern.indexOf(']', i + 1);
      if (end < 0) throw new VbError(93, 'Invalid pattern string.');
      let body = pattern.slice(i + 1, end);
      const neg = body.startsWith('!');
      if (neg) body = body.slice(1);
      re += `[${neg ? '^' : ''}${body.replace(/[\\\]^]/g, '\\$&')}]`;
      i = end;
    } else re += c.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(re + '$', text ? 'i' : '').test(s);
}
