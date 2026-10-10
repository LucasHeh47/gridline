// Type coercion rules, matching Excel's behaviour.
import { MAX_COLS, MAX_ROWS, formatCell } from '../address';
import { parseNumberLike } from '../input';
import { Errors, ErrorValue, isError, type Scalar } from '../values';
import { Matrix, RangeRef, type Arg, type Ctx } from './types';

export function isArrayLike(v: Arg): v is RangeRef | Matrix {
  return (v instanceof RangeRef && !v.isCell) || v instanceof Matrix;
}

/** Reads a value in a scalar position (implicit intersection for ranges). */
export function scalarOf(ctx: Ctx, v: Arg): Scalar {
  if (v === undefined) return null;
  if (v instanceof RangeRef) {
    if (v.isCell) return ctx.get(v.sheet, v.r1, v.c1);
    // Implicit intersection with the formula's row or column
    if (v.cols === 1 && ctx.row >= v.r1 && ctx.row <= v.r2 && v.sheet === ctx.sheet) return ctx.get(v.sheet, ctx.row, v.c1);
    if (v.rows === 1 && ctx.col >= v.c1 && ctx.col <= v.c2 && v.sheet === ctx.sheet) return ctx.get(v.sheet, v.r1, ctx.col);
    return Errors.value('A range was used where a single value is expected.');
  }
  if (v instanceof Matrix) return v.get(0, 0);
  return v;
}

/** Clamps whole-column/row references to the used area of the sheet. */
export function clampRef(ctx: Ctx, ref: RangeRef): { r2: number; c2: number } {
  let r2 = ref.r2;
  let c2 = ref.c2;
  if (r2 === MAX_ROWS - 1 || c2 === MAX_COLS - 1) {
    const b = ctx.bounds(ref.sheet);
    if (r2 === MAX_ROWS - 1) r2 = Math.max(ref.r1, Math.min(r2, b.maxRow));
    if (c2 === MAX_COLS - 1) c2 = Math.max(ref.c1, Math.min(c2, b.maxCol));
  }
  return { r2, c2 };
}

export function toMatrix(ctx: Ctx, v: Arg): Matrix {
  if (v instanceof Matrix) return v;
  if (v instanceof RangeRef) {
    const { r2, c2 } = clampRef(ctx, v);
    const rows = r2 - v.r1 + 1;
    const cols = c2 - v.c1 + 1;
    return Matrix.fill(rows, cols, (r, c) => ctx.get(v.sheet, v.r1 + r, v.c1 + c));
  }
  return new Matrix([[v === undefined ? null : v]]);
}

const NUM_TEXT_MSG = (s: string) =>
  `"${s.length > 30 ? s.slice(0, 30) + '…' : s}" is text, not a number. Remove extra characters or convert it with VALUE().`;

export function toNumber(v: Scalar | undefined): number | ErrorValue {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof ErrorValue) return v;
  const n = parseNumberLike(v);
  if (n === null) return Errors.value(v === '' ? 'An empty text value ("") was used where a number is needed.' : NUM_TEXT_MSG(v));
  return n;
}

/** Converts a number to text the way Excel does in formulas (up to 15 significant digits). */
export function numberToText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e21) return String(n);
  let s = n.toPrecision(15);
  if (s.includes('e')) {
    const [m, e] = s.split('e');
    const mant = m.includes('.') ? m.replace(/\.?0+$/, '') : m;
    const ex = parseInt(e, 10);
    return `${mant}E${ex < 0 ? '-' : '+'}${String(Math.abs(ex)).padStart(2, '0')}`;
  }
  if (s.includes('.')) s = s.replace(/\.?0+$/, '');
  return s;
}

export function toText(v: Scalar | undefined): string | ErrorValue {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numberToText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v;
}

export function toBool(v: Scalar | undefined): boolean | ErrorValue {
  if (v === null || v === undefined) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v instanceof ErrorValue) return v;
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return Errors.value(`"${v}" can't be used as TRUE/FALSE.`);
}

/** Argument helpers that read a scalar and coerce it. */
export function argNum(ctx: Ctx, v: Arg, dflt?: number): number | ErrorValue {
  if (v === undefined && dflt !== undefined) return dflt;
  return toNumber(scalarOf(ctx, v));
}

export function argInt(ctx: Ctx, v: Arg, dflt?: number): number | ErrorValue {
  const n = argNum(ctx, v, dflt);
  return isError(n) ? n : Math.trunc(n);
}

export function argText(ctx: Ctx, v: Arg, dflt?: string): string | ErrorValue {
  if (v === undefined && dflt !== undefined) return dflt;
  return toText(scalarOf(ctx, v));
}

export function argBool(ctx: Ctx, v: Arg, dflt?: boolean): boolean | ErrorValue {
  if (v === undefined && dflt !== undefined) return dflt;
  return toBool(scalarOf(ctx, v));
}

export interface CollectOpts {
  /** Count text and booleans inside ranges (the "A" functions: COUNTA, AVERAGEA...). */
  includeAll?: boolean;
  /** Skip errors instead of propagating them. */
  ignoreErrors?: boolean;
  /** Skip rows hidden by a filter or manually (SUBTOTAL 101-111). */
  skipHidden?: 'filtered' | 'all';
  /** Skip cells containing SUBTOTAL/AGGREGATE formulas. */
  skipSubtotals?: boolean;
}

/**
 * Collects the numbers from function arguments (SUM, AVERAGE...).
 * Direct arguments are coerced ("3" -> 3, TRUE -> 1); text in ranges is ignored.
 */
export function collectNumbers(ctx: Ctx, args: Arg[], opts: CollectOpts = {}): number[] | ErrorValue {
  const out: number[] = [];
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof RangeRef) {
      let error: ErrorValue | null = null;
      const sheet = a.sheet;
      ctx.forEachInRange(a, (v, r, c) => {
        if (opts.skipHidden) {
          if (opts.skipHidden === 'all' ? sheet.isRowHidden(r) : sheet.filterHidden.has(r)) return;
        }
        if (opts.skipSubtotals) {
          const f = ctx.formulaText(sheet, r, c);
          if (f && /\b(SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) return;
        }
        if (typeof v === 'number') out.push(v);
        else if (v instanceof ErrorValue) {
          if (!opts.ignoreErrors) {
            error = v;
            return false;
          }
        } else if (opts.includeAll) {
          if (typeof v === 'boolean') out.push(v ? 1 : 0);
          else if (typeof v === 'string') out.push(0);
        }
      });
      if (error) return error;
    } else if (a instanceof Matrix) {
      for (const row of a.data)
        for (const v of row) {
          if (typeof v === 'number') out.push(v);
          else if (v instanceof ErrorValue) {
            if (!opts.ignoreErrors) return v;
          } else if (opts.includeAll && v !== null) {
            out.push(typeof v === 'boolean' ? (v ? 1 : 0) : 0);
          }
        }
    } else {
      if (a === null) {
        if (opts.includeAll) out.push(0);
        continue;
      }
      const n = toNumber(a);
      if (isError(n)) {
        if (opts.ignoreErrors) continue;
        return n;
      }
      out.push(n);
    }
  }
  return out;
}

/** Iterates every value (including blanks) of an argument in row-major order. */
export function eachValue(ctx: Ctx, a: Arg, cb: (v: Scalar) => void): void {
  if (a === undefined) return;
  if (a instanceof RangeRef || a instanceof Matrix) {
    const m = toMatrix(ctx, a);
    for (const row of m.data) for (const v of row) cb(v);
  } else cb(a);
}

/** Flattens an argument to a list of values (row-major). */
export function flatValues(ctx: Ctx, a: Arg): Scalar[] {
  const out: Scalar[] = [];
  eachValue(ctx, a, (v) => out.push(v));
  return out;
}

/** Excel comparison: numbers < text < booleans; text is case-insensitive. Blank equals 0 / "" / FALSE. */
export function compareValues(a: Scalar, b: Scalar): number {
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ta = typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : typeof a === 'boolean' ? 2 : 3;
  const tb = typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : typeof b === 'boolean' ? 2 : 3;
  if (ta !== tb) return ta - tb;
  if (ta === 0) return (a as number) < (b as number) ? -1 : (a as number) > (b as number) ? 1 : 0;
  if (ta === 1) {
    const x = (a as string).toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (ta === 2) return (a ? 1 : 0) - (b ? 1 : 0);
  return 0;
}

export function describeRef(ref: RangeRef, ctx: Ctx): string {
  const prefix = ref.sheet !== ctx.sheet ? `${ref.sheet.name}!` : '';
  if (ref.isCell) return prefix + formatCell(ref.r1, ref.c1);
  return `${prefix}${formatCell(ref.r1, ref.c1)}:${formatCell(ref.r2, ref.c2)}`;
}

export function numResult(n: number): number | ErrorValue {
  if (!isFinite(n) || isNaN(n)) return Errors.num('The result is too large, too small, or not a real number.');
  return n;
}
