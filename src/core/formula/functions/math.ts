import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import {
  argInt,
  argNum,
  argText,
  collectNumbers,
  numResult,
  scalarOf,
  toMatrix,
  type CollectOpts,
} from '../coerce';
import { makeMatcher } from '../criteria';
import { Matrix, RangeRef, type Arg, type Ctx, type Value } from '../types';
import { define } from './registry';
import * as stats from './stats-core';

/** Excel-style rounding: half away from zero, robust to binary floating point (1.005 -> 1.01). */
export function roundHalfAway(x: number, digits: number): number {
  if (!isFinite(x)) return x;
  const sign = x < 0 ? -1 : 1;
  const abs = Math.abs(x);
  const shifted = Number(abs + 'e' + digits);
  if (!isFinite(shifted)) return x;
  const r = Number(Math.round(shifted) + 'e' + -digits);
  return sign * (isNaN(r) ? abs : r);
}

function roundFn(mode: 'round' | 'up' | 'down') {
  return (ctx: Ctx, [a, b]: Arg[]): Value => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const d = argInt(ctx, b, 0);
    if (isError(d)) return d;
    if (mode === 'round') return roundHalfAway(x, d);
    const f = Math.pow(10, d);
    const y = Number((Math.abs(x) * f).toPrecision(15));
    const r = (mode === 'up' ? Math.ceil(y) : Math.floor(y)) / f;
    return x < 0 ? -r : r;
  };
}

function unary(f: (x: number) => number | ErrorValue) {
  return (ctx: Ctx, [a]: Arg[]): Value => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const r = f(x);
    return isError(r) ? r : numResult(r);
  };
}

function sumOf(nums: number[]): number {
  // Kahan summation reduces floating point drift (0.1+0.2+... style errors)
  let s = 0;
  let c = 0;
  for (const n of nums) {
    const y = n - c;
    const t = s + y;
    c = t - s - y;
    s = t;
  }
  return s;
}

export function sumArgs(ctx: Ctx, args: Arg[], opts?: CollectOpts): Value {
  const nums = collectNumbers(ctx, args, opts);
  if (isError(nums)) return nums;
  return numResult(sumOf(nums));
}

define('SUM', 'Math', 'number1, [number2], ...', 'Adds all the numbers in a range of cells.', (ctx, args) => sumArgs(ctx, args));

define('SUMSQ', 'Math', 'number1, [number2], ...', 'Returns the sum of the squares of the arguments.', (ctx, args) => {
  const nums = collectNumbers(ctx, args);
  if (isError(nums)) return nums;
  return numResult(sumOf(nums.map((n) => n * n)));
});

define('PRODUCT', 'Math', 'number1, [number2], ...', 'Multiplies all the numbers given as arguments.', (ctx, args) => {
  const nums = collectNumbers(ctx, args);
  if (isError(nums)) return nums;
  if (nums.length === 0) return 0;
  return numResult(nums.reduce((a, b) => a * b, 1));
});

define(
  'SUMPRODUCT',
  'Math',
  'array1, [array2], ...',
  'Multiplies corresponding items in the given arrays and returns the sum of those products.',
  (ctx, args) => {
    const mats = args.map((a) => toMatrix(ctx, a === undefined ? null : a));
    const rows = mats[0].rows;
    const cols = mats[0].cols;
    if (mats.some((m) => m.rows !== rows || m.cols !== cols))
      return Errors.value('All arrays in SUMPRODUCT must be the same size.');
    let total = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let p = 1;
        for (const m of mats) {
          const v = m.get(r, c);
          if (v instanceof ErrorValue) return v;
          p *= typeof v === 'number' ? v : typeof v === 'boolean' && mats.length === 1 ? 0 : typeof v === 'boolean' ? (v ? 1 : 0) : 0;
        }
        total += p;
      }
    }
    return numResult(total);
  },
);

/** Shared implementation for SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS. */
export function conditionalValues(
  ctx: Ctx,
  target: Arg | null,
  pairs: [Arg, Arg][],
): { values: Scalar[]; count: number } | ErrorValue {
  const mats: Matrix[] = [];
  const matchers = [];
  for (const [range, crit] of pairs) {
    if (!(range instanceof RangeRef) && !(range instanceof Matrix))
      return Errors.value('Criteria ranges must be cell ranges, like A2:A100.');
    mats.push(toMatrix(ctx, range));
    matchers.push(makeMatcher(scalarOf(ctx, crit)));
  }
  const rows = mats[0].rows;
  const cols = mats[0].cols;
  for (const m of mats) {
    if (m.rows !== rows || m.cols !== cols) return Errors.value('All criteria ranges must be the same size.');
  }
  let tm: Matrix | null = null;
  if (target !== null && target !== undefined) {
    if (target instanceof RangeRef) {
      // Like Excel, the sum range is resized to match the criteria range.
      const first = pairs[0][0];
      const tr = first instanceof RangeRef ? Math.max(rows, 1) : rows;
      const ref = new RangeRef(target.sheet, target.r1, target.c1, target.r1 + tr - 1, target.c1 + cols - 1);
      tm = Matrix.fill(rows, cols, (r, c) => ctx.get(ref.sheet, ref.r1 + r, ref.c1 + c));
    } else tm = toMatrix(ctx, target);
  }
  const values: Scalar[] = [];
  let count = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let ok = true;
      for (let k = 0; k < mats.length; k++) {
        if (!matchers[k](mats[k].get(r, c))) {
          ok = false;
          break;
        }
      }
      if (ok) {
        count++;
        if (tm) values.push(tm.get(r, c));
      }
    }
  }
  return { values, count };
}

function pairsFrom(args: Arg[], start: number): [Arg, Arg][] | ErrorValue {
  const rest = args.slice(start);
  if (rest.length % 2 !== 0) return Errors.value('Criteria must come in pairs: a range, then its condition.');
  const out: [Arg, Arg][] = [];
  for (let i = 0; i < rest.length; i += 2) out.push([rest[i], rest[i + 1]]);
  return out;
}

export function numbersOnly(vals: Scalar[]): number[] | ErrorValue {
  const out: number[] = [];
  for (const v of vals) {
    if (typeof v === 'number') out.push(v);
    else if (v instanceof ErrorValue) return v;
  }
  return out;
}

define('SUMIF', 'Math', 'range, criteria, [sum_range]', 'Adds the cells that meet a condition, e.g. SUMIF(A:A, ">100").', (ctx, [range, crit, sumRange]) => {
  const r = conditionalValues(ctx, sumRange ?? range ?? null, [[range, crit]]);
  if (isError(r)) return r;
  const nums = numbersOnly(r.values);
  return isError(nums) ? nums : numResult(sumOf(nums));
});

define(
  'SUMIFS',
  'Math',
  'sum_range, criteria_range1, criteria1, [criteria_range2], [criteria2], ...',
  'Adds the cells that meet several conditions.',
  (ctx, args) => {
    const pairs = pairsFrom(args, 1);
    if (isError(pairs)) return pairs;
    const r = conditionalValues(ctx, args[0] ?? null, pairs);
    if (isError(r)) return r;
    const nums = numbersOnly(r.values);
    return isError(nums) ? nums : numResult(sumOf(nums));
  },
  { min: 3 },
);

define('ABS', 'Math', 'number', 'Returns the absolute value of a number.', unary(Math.abs), { lift: true });
define('ROUND', 'Math', 'number, num_digits', 'Rounds a number to a given number of digits.', roundFn('round'), { lift: true });
define('ROUNDUP', 'Math', 'number, num_digits', 'Rounds a number up, away from zero.', roundFn('up'), { lift: true });
define('ROUNDDOWN', 'Math', 'number, num_digits', 'Rounds a number down, toward zero.', roundFn('down'), { lift: true });
define('INT', 'Math', 'number', 'Rounds a number down to the nearest integer.', unary(Math.floor), { lift: true });
define(
  'TRUNC',
  'Math',
  'number, [num_digits]',
  'Truncates a number by removing its fractional part.',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const d = argInt(ctx, b, 0);
    if (isError(d)) return d;
    const f = Math.pow(10, d);
    return Math.trunc(Number((x * f).toPrecision(15))) / f;
  },
  { lift: true },
);

define(
  'MROUND',
  'Math',
  'number, multiple',
  'Rounds a number to the nearest multiple.',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const m = argNum(ctx, b);
    if (isError(m)) return m;
    if (m === 0) return 0;
    if (x * m < 0) return Errors.num('The number and multiple must have the same sign.');
    return roundHalfAway(Math.round(Number((x / m).toPrecision(15)) + (x / m > 0 ? 1e-12 : -1e-12)) * m, 10);
  },
  { lift: true },
);

function ceilFloor(up: boolean, math: boolean) {
  return (ctx: Ctx, [a, b, c]: Arg[]): Value => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    let sig = argNum(ctx, b, math ? 1 : x < 0 ? -1 : 1);
    if (isError(sig)) return sig;
    if (sig === 0) return 0;
    if (math) {
      const mode = argNum(ctx, c, 0);
      if (isError(mode)) return mode;
      sig = Math.abs(sig);
      const q = x / sig;
      let r: number;
      if (x < 0 && mode !== 0) r = up ? Math.floor(q) : Math.ceil(q);
      else r = up ? Math.ceil(q - 1e-12) : Math.floor(q + 1e-12);
      return roundHalfAway(r * sig, 12);
    }
    if (x > 0 && sig < 0) return Errors.num('For a positive number, the significance must be positive.');
    const q = x / sig;
    const r = up ? Math.ceil(q - 1e-12) : Math.floor(q + 1e-12);
    return roundHalfAway(r * sig, 12);
  };
}

define('CEILING', 'Math', 'number, significance', 'Rounds a number up to the nearest multiple of significance.', ceilFloor(true, false), { lift: true });
define('CEILING.MATH', 'Math', 'number, [significance], [mode]', 'Rounds a number up to the nearest integer or multiple.', ceilFloor(true, true), { lift: true });
define('FLOOR', 'Math', 'number, significance', 'Rounds a number down to the nearest multiple of significance.', ceilFloor(false, false), { lift: true });
define('FLOOR.MATH', 'Math', 'number, [significance], [mode]', 'Rounds a number down to the nearest integer or multiple.', ceilFloor(false, true), { lift: true });

define(
  'MOD',
  'Math',
  'number, divisor',
  'Returns the remainder after division. The result has the same sign as the divisor.',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const d = argNum(ctx, b);
    if (isError(d)) return d;
    if (d === 0) return Errors.div0('MOD\'s divisor is 0.');
    const r = x - d * Math.floor(x / d);
    return roundHalfAway(r, 12);
  },
  { lift: true },
);

define(
  'QUOTIENT',
  'Math',
  'numerator, denominator',
  'Returns the integer portion of a division.',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const d = argNum(ctx, b);
    if (isError(d)) return d;
    if (d === 0) return Errors.div0('QUOTIENT\'s denominator is 0.');
    return Math.trunc(x / d);
  },
  { lift: true },
);

define(
  'POWER',
  'Math',
  'number, power',
  'Returns a number raised to a power.',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const p = argNum(ctx, b);
    if (isError(p)) return p;
    if (x === 0 && p === 0) return Errors.num('0 raised to the power 0 is undefined.');
    if (x === 0 && p < 0) return Errors.div0('0 raised to a negative power divides by zero.');
    if (x < 0 && !Number.isInteger(p)) return Errors.num('A negative number can\'t be raised to a fractional power.');
    return numResult(Math.pow(x, p));
  },
  { lift: true },
);

define('SQRT', 'Math', 'number', 'Returns the square root of a number.', unary((x) => (x < 0 ? Errors.num('Can\'t take the square root of a negative number.') : Math.sqrt(x))), { lift: true });
define('SQRTPI', 'Math', 'number', 'Returns the square root of (number × pi).', unary((x) => (x < 0 ? Errors.num() : Math.sqrt(x * Math.PI))), { lift: true });
define('EXP', 'Math', 'number', 'Returns e raised to the power of a number.', unary(Math.exp), { lift: true });
define('LN', 'Math', 'number', 'Returns the natural logarithm of a number.', unary((x) => (x <= 0 ? Errors.num('Logarithms need a number greater than 0.') : Math.log(x))), { lift: true });
define('LOG10', 'Math', 'number', 'Returns the base-10 logarithm of a number.', unary((x) => (x <= 0 ? Errors.num('Logarithms need a number greater than 0.') : Math.log10(x))), { lift: true });
define(
  'LOG',
  'Math',
  'number, [base]',
  'Returns the logarithm of a number to the base you specify (default 10).',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const base = argNum(ctx, b, 10);
    if (isError(base)) return base;
    if (x <= 0 || base <= 0) return Errors.num('Logarithms need numbers greater than 0.');
    if (base === 1) return Errors.div0('A logarithm base of 1 divides by zero.');
    return numResult(Math.log(x) / Math.log(base));
  },
  { lift: true },
);

define('PI', 'Math', '', 'Returns the value of pi (3.14159…).', () => Math.PI);
define('RAND', 'Math', '', 'Returns a random number between 0 and 1. Changes on every recalculation.', () => Math.random(), { volatile: true });
define(
  'RANDBETWEEN',
  'Math',
  'bottom, top',
  'Returns a random whole number between the numbers you specify.',
  (ctx, [a, b]) => {
    const lo = argNum(ctx, a);
    if (isError(lo)) return lo;
    const hi = argNum(ctx, b);
    if (isError(hi)) return hi;
    const l = Math.ceil(lo);
    const h = Math.floor(hi);
    if (l > h) return Errors.num('bottom must be less than or equal to top.');
    return l + Math.floor(Math.random() * (h - l + 1));
  },
  { volatile: true },
);
define(
  'RANDARRAY',
  'Math',
  '[rows], [columns], [min], [max], [whole_number]',
  'Returns an array of random numbers.',
  (ctx, [a, b, c, d, e]) => {
    const rows = argInt(ctx, a, 1);
    const cols = argInt(ctx, b, 1);
    const min = argNum(ctx, c, 0);
    const max = argNum(ctx, d, 1);
    const whole = e === undefined ? false : !!scalarOf(ctx, e);
    for (const x of [rows, cols, min, max]) if (isError(x)) return x;
    const R = rows as number;
    const C = cols as number;
    const lo = min as number;
    const hi = max as number;
    if (R < 1 || C < 1) return Errors.calc('RANDARRAY needs at least 1 row and 1 column.');
    if (lo > hi) return Errors.value('min must be less than max.');
    return Matrix.fill(R, C, () => (whole ? Math.floor(lo + Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo)));
  },
  { volatile: true },
);

define('SIGN', 'Math', 'number', 'Returns 1, 0 or -1 depending on the sign of a number.', unary(Math.sign), { lift: true });
define(
  'EVEN',
  'Math',
  'number',
  'Rounds a number up to the nearest even integer (away from zero).',
  unary((x) => {
    const s = x < 0 ? -1 : 1;
    let n = Math.ceil(Math.abs(x));
    if (n % 2) n++;
    return s * n;
  }),
  { lift: true },
);
define(
  'ODD',
  'Math',
  'number',
  'Rounds a number up to the nearest odd integer (away from zero).',
  unary((x) => {
    const s = x < 0 ? -1 : 1;
    let n = Math.ceil(Math.abs(x));
    if (n % 2 === 0) n++;
    return s * n;
  }),
  { lift: true },
);

function factorial(n: number): number {
  let r = 1;
  for (let i = 2; i <= n; i++) r *= i;
  return r;
}

define('FACT', 'Math', 'number', 'Returns the factorial of a number.', unary((x) => (x < 0 ? Errors.num('FACT needs a non-negative number.') : factorial(Math.floor(x)))), { lift: true });
define(
  'FACTDOUBLE',
  'Math',
  'number',
  'Returns the double factorial of a number.',
  unary((x) => {
    if (x < 0) return Errors.num();
    let r = 1;
    for (let i = Math.floor(x); i > 1; i -= 2) r *= i;
    return r;
  }),
  { lift: true },
);

function combin(n: number, k: number): number {
  if (k > n - k) k = n - k;
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
}

define(
  'COMBIN',
  'Math',
  'number, number_chosen',
  'Returns the number of combinations for a given number of items.',
  (ctx, [a, b]) => {
    const n = argInt(ctx, a);
    if (isError(n)) return n;
    const k = argInt(ctx, b);
    if (isError(k)) return k;
    if (n < 0 || k < 0 || k > n) return Errors.num('Need 0 ≤ number_chosen ≤ number.');
    return combin(n, k);
  },
  { lift: true },
);
define(
  'COMBINA',
  'Math',
  'number, number_chosen',
  'Returns the number of combinations with repetitions.',
  (ctx, [a, b]) => {
    const n = argInt(ctx, a);
    if (isError(n)) return n;
    const k = argInt(ctx, b);
    if (isError(k)) return k;
    if (n < 0 || k < 0) return Errors.num();
    return combin(n + k - 1, k);
  },
  { lift: true },
);
define(
  'PERMUT',
  'Math',
  'number, number_chosen',
  'Returns the number of permutations for a given number of objects.',
  (ctx, [a, b]) => {
    const n = argInt(ctx, a);
    if (isError(n)) return n;
    const k = argInt(ctx, b);
    if (isError(k)) return k;
    if (n < 0 || k < 0 || k > n) return Errors.num();
    let r = 1;
    for (let i = n - k + 1; i <= n; i++) r *= i;
    return r;
  },
  { lift: true },
);

function gcd(a: number, b: number): number {
  while (b) [a, b] = [b, a % b];
  return a;
}

define('GCD', 'Math', 'number1, [number2], ...', 'Returns the greatest common divisor.', (ctx, args) => {
  const nums = collectNumbers(ctx, args);
  if (isError(nums)) return nums;
  if (nums.some((n) => n < 0)) return Errors.num('GCD needs non-negative numbers.');
  return nums.map(Math.floor).reduce((a, b) => gcd(a, b), 0);
});
define('LCM', 'Math', 'number1, [number2], ...', 'Returns the least common multiple.', (ctx, args) => {
  const nums = collectNumbers(ctx, args);
  if (isError(nums)) return nums;
  if (nums.some((n) => n < 0)) return Errors.num('LCM needs non-negative numbers.');
  return nums.map(Math.floor).reduce((a, b) => (a === 0 || b === 0 ? 0 : (a * b) / gcd(a, b)), 1);
});

define('DEGREES', 'Math', 'angle', 'Converts radians to degrees.', unary((x) => (x * 180) / Math.PI), { lift: true });
define('RADIANS', 'Math', 'angle', 'Converts degrees to radians.', unary((x) => (x * Math.PI) / 180), { lift: true });
define('SIN', 'Math', 'number', 'Returns the sine of an angle (in radians).', unary(Math.sin), { lift: true });
define('COS', 'Math', 'number', 'Returns the cosine of an angle (in radians).', unary(Math.cos), { lift: true });
define('TAN', 'Math', 'number', 'Returns the tangent of an angle (in radians).', unary(Math.tan), { lift: true });
define('ASIN', 'Math', 'number', 'Returns the arcsine of a number, in radians.', unary((x) => (Math.abs(x) > 1 ? Errors.num('ASIN needs a number between -1 and 1.') : Math.asin(x))), { lift: true });
define('ACOS', 'Math', 'number', 'Returns the arccosine of a number, in radians.', unary((x) => (Math.abs(x) > 1 ? Errors.num('ACOS needs a number between -1 and 1.') : Math.acos(x))), { lift: true });
define('ATAN', 'Math', 'number', 'Returns the arctangent of a number, in radians.', unary(Math.atan), { lift: true });
define(
  'ATAN2',
  'Math',
  'x_num, y_num',
  'Returns the arctangent from x and y coordinates, in radians.',
  (ctx, [a, b]) => {
    const x = argNum(ctx, a);
    if (isError(x)) return x;
    const y = argNum(ctx, b);
    if (isError(y)) return y;
    if (x === 0 && y === 0) return Errors.div0('ATAN2 of (0, 0) is undefined.');
    return Math.atan2(y, x);
  },
  { lift: true },
);
define('SINH', 'Math', 'number', 'Returns the hyperbolic sine of a number.', unary(Math.sinh), { lift: true });
define('COSH', 'Math', 'number', 'Returns the hyperbolic cosine of a number.', unary(Math.cosh), { lift: true });
define('TANH', 'Math', 'number', 'Returns the hyperbolic tangent of a number.', unary(Math.tanh), { lift: true });
define('ASINH', 'Math', 'number', 'Returns the inverse hyperbolic sine.', unary(Math.asinh), { lift: true });
define('ACOSH', 'Math', 'number', 'Returns the inverse hyperbolic cosine.', unary((x) => (x < 1 ? Errors.num() : Math.acosh(x))), { lift: true });
define('ATANH', 'Math', 'number', 'Returns the inverse hyperbolic tangent.', unary((x) => (Math.abs(x) >= 1 ? Errors.num() : Math.atanh(x))), { lift: true });

const SUBTOTAL_FNS: Record<number, (nums: number[]) => number | ErrorValue> = {
  1: (n) => (n.length ? sumOf(n) / n.length : Errors.div0('There are no numbers to average.')),
  2: (n) => n.length,
  3: (n) => n.length,
  4: (n) => (n.length ? Math.max(...n) : 0),
  5: (n) => (n.length ? Math.min(...n) : 0),
  6: (n) => n.reduce((a, b) => a * b, 1),
  7: (n) => stats.stdev(n, true),
  8: (n) => stats.stdev(n, false),
  9: (n) => sumOf(n),
  10: (n) => stats.variance(n, true),
  11: (n) => stats.variance(n, false),
};

define(
  'SUBTOTAL',
  'Math',
  'function_num, ref1, [ref2], ...',
  'Returns a subtotal that ignores rows hidden by a filter (1=AVERAGE, 2=COUNT, 3=COUNTA, 4=MAX, 5=MIN, 6=PRODUCT, 7=STDEV, 9=SUM, 10=VAR; add 100 to also ignore manually hidden rows).',
  (ctx, args) => {
    const fnum = argInt(ctx, args[0]);
    if (isError(fnum)) return fnum;
    const base = fnum > 100 ? fnum - 100 : fnum;
    const f = SUBTOTAL_FNS[base];
    if (!f) return Errors.value('SUBTOTAL\'s function_num must be 1-11 or 101-111.');
    const refs = args.slice(1);
    if (base === 3) {
      let count = 0;
      for (const r of refs) {
        if (!(r instanceof RangeRef)) continue;
        ctx.forEachInRange(r, (v, row, col) => {
          if (fnum > 100 ? r.sheet.isRowHidden(row) : r.sheet.filterHidden.has(row)) return;
          const ft = ctx.formulaText(r.sheet, row, col);
          if (ft && /\bSUBTOTAL\s*\(/i.test(ft)) return;
          if (v !== null && v !== '') count++;
        });
      }
      return count;
    }
    const nums = collectNumbers(ctx, refs, { skipHidden: fnum > 100 ? 'all' : 'filtered', skipSubtotals: true });
    if (isError(nums)) return nums;
    const r = f(nums);
    return isError(r) ? r : numResult(r);
  },
  { min: 2 },
);

define(
  'AGGREGATE',
  'Math',
  'function_num, options, ref1, [ref2], ...',
  'Like SUBTOTAL, but can also ignore errors (options 2, 3, 6, 7) — e.g. AGGREGATE(9, 6, A:A) sums while skipping errors.',
  (ctx, args) => {
    const fnum = argInt(ctx, args[0]);
    if (isError(fnum)) return fnum;
    const opt = argInt(ctx, args[1], 0);
    if (isError(opt)) return opt;
    const ignoreErrors = [2, 3, 6, 7].includes(opt);
    const ignoreHidden = [1, 3, 5, 7].includes(opt);
    const skipSubtotals = opt <= 3;
    if (fnum >= 14 && fnum <= 19) {
      const nums = collectNumbers(ctx, [args[2]], { ignoreErrors, skipHidden: ignoreHidden ? 'all' : undefined, skipSubtotals });
      if (isError(nums)) return nums;
      const k = argNum(ctx, args[3]);
      if (isError(k)) return k;
      switch (fnum) {
        case 14:
          return stats.kth(nums, k, true);
        case 15:
          return stats.kth(nums, k, false);
        case 16:
          return stats.percentile(nums, k, true);
        case 17:
          return stats.percentile(nums, k / 4, true);
        case 18:
          return stats.percentile(nums, k, false);
        default:
          return stats.percentile(nums, k / 4, false);
      }
    }
    const nums = collectNumbers(ctx, args.slice(2), { ignoreErrors, skipHidden: ignoreHidden ? 'all' : undefined, skipSubtotals });
    if (isError(nums)) return nums;
    switch (fnum) {
      case 12:
        return stats.median(nums);
      case 13:
        return stats.mode(nums);
    }
    const f = SUBTOTAL_FNS[fnum];
    if (!f) return Errors.value('AGGREGATE\'s function_num must be 1-19.');
    const r = f(nums);
    return isError(r) ? r : numResult(r);
  },
  { min: 3 },
);

define(
  'SEQUENCE',
  'Math',
  'rows, [columns], [start], [step]',
  'Generates a list of sequential numbers, e.g. SEQUENCE(10) gives 1 to 10.',
  (ctx, [a, b, c, d]) => {
    const rows = argInt(ctx, a);
    const cols = argInt(ctx, b, 1);
    const start = argNum(ctx, c, 1);
    const step = argNum(ctx, d, 1);
    for (const x of [rows, cols, start, step]) if (isError(x)) return x;
    if ((rows as number) < 1 || (cols as number) < 1) return Errors.calc('SEQUENCE needs at least 1 row and 1 column.');
    if ((rows as number) * (cols as number) > 5_000_000) return Errors.num('That sequence is too large.');
    return Matrix.fill(rows as number, cols as number, (r, cc) => (start as number) + (r * (cols as number) + cc) * (step as number));
  },
);

define('MMULT', 'Math', 'array1, array2', 'Returns the matrix product of two arrays.', (ctx, [a, b]) => {
  const x = toMatrix(ctx, a);
  const y = toMatrix(ctx, b);
  if (x.cols !== y.rows) return Errors.value('MMULT needs array1\'s column count to equal array2\'s row count.');
  return Matrix.fill(x.rows, y.cols, (r, c) => {
    let s = 0;
    for (let k = 0; k < x.cols; k++) {
      const p = x.get(r, k);
      const q = y.get(k, c);
      if (typeof p !== 'number' || typeof q !== 'number') return Errors.value('MMULT arrays must contain only numbers.');
      s += p * q;
    }
    return s;
  });
});

define(
  'ROMAN',
  'Math',
  'number, [form]',
  'Converts a number to Roman numerals.',
  (ctx, [a]) => {
    const n = argInt(ctx, a);
    if (isError(n)) return n;
    if (n < 0 || n > 3999) return Errors.value('ROMAN works for 0 to 3999.');
    const table: [number, string][] = [
      [1000, 'M'],
      [900, 'CM'],
      [500, 'D'],
      [400, 'CD'],
      [100, 'C'],
      [90, 'XC'],
      [50, 'L'],
      [40, 'XL'],
      [10, 'X'],
      [9, 'IX'],
      [5, 'V'],
      [4, 'IV'],
      [1, 'I'],
    ];
    let x = n;
    let out = '';
    for (const [v, s] of table)
      while (x >= v) {
        out += s;
        x -= v;
      }
    return out;
  },
  { lift: [0] },
);

define(
  'ARABIC',
  'Math',
  'text',
  'Converts a Roman numeral to a number.',
  (ctx, [a]) => {
    const t = argText(ctx, a);
    if (isError(t)) return t;
    const map: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
    const s = t.trim().toUpperCase();
    let total = 0;
    for (let i = 0; i < s.length; i++) {
      const v = map[s[i]];
      if (!v) return Errors.value(`"${t}" isn't a Roman numeral.`);
      const next = map[s[i + 1]] ?? 0;
      total += v < next ? -v : v;
    }
    return total;
  },
  { lift: true },
);

define(
  'BASE',
  'Math',
  'number, radix, [min_length]',
  'Converts a number to text in another base (radix 2-36).',
  (ctx, [a, b, c]) => {
    const n = argInt(ctx, a);
    if (isError(n)) return n;
    const radix = argInt(ctx, b);
    if (isError(radix)) return radix;
    const len = argInt(ctx, c, 0);
    if (isError(len)) return len;
    if (n < 0 || radix < 2 || radix > 36) return Errors.num();
    return n.toString(radix).toUpperCase().padStart(len, '0');
  },
  { lift: true },
);

define(
  'DECIMAL',
  'Math',
  'text, radix',
  'Converts text in a given base to a decimal number.',
  (ctx, [a, b]) => {
    const t = argText(ctx, a);
    if (isError(t)) return t;
    const radix = argInt(ctx, b);
    if (isError(radix)) return radix;
    if (radix < 2 || radix > 36) return Errors.num();
    const n = parseInt(t, radix);
    return isNaN(n) ? Errors.num() : n;
  },
  { lift: true },
);

/** Used by statistical functions that only accept number lists. */
export function numsFrom(ctx: Ctx, args: Arg[], opts?: CollectOpts): number[] | ErrorValue {
  return collectNumbers(ctx, args, opts);
}

export { sumOf };
