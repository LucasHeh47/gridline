// Arithmetic that gives the same results as Excel, bit for bit.
//
// Excel calculates on the x87 floating-point unit: each operation is done with a 64-bit mantissa and
// the result is then stored as a 53-bit double. Rounding twice gives a different last bit from plain
// JavaScript arithmetic in roughly 1 case in 1,000 (for example =19*107.85*0.95 is 1946.6925 in Excel
// but 1946.6924999999999 in JavaScript). These helpers compute the correctly rounded result plus its
// exact error, and move to the neighboring double when Excel's double rounding would.
//
// Excel also sets the result of a formula's final + or - to exactly 0 when it is within 8 units in the
// last place of the operands (so =0.1+0.2-0.3 is 0), and SUM adds left to right in plain doubles.

const view = new DataView(new ArrayBuffer(8));

function isOddMantissa(x: number): boolean {
  view.setFloat64(0, x);
  return (view.getUint8(7) & 1) === 1;
}

/** The next double above (dir > 0) or below x. */
function nextAfter(x: number, dir: number): number {
  if (x === 0) return dir > 0 ? Number.MIN_VALUE : -Number.MIN_VALUE;
  view.setFloat64(0, x);
  let bits = view.getBigUint64(0);
  bits += (x > 0) === (dir > 0) ? 1n : -1n;
  view.setBigUint64(0, bits);
  return view.getFloat64(0);
}

/**
 * Given p = fl(exact) and err = exact - p (computed exactly or nearly so), returns the double that
 * rounding exact first to 64 bits and then to 53 bits produces.
 */
function doubleRound(p: number, err: number): number {
  if (err === 0 || !Number.isFinite(p) || p === 0) return p;
  // Fast exit: only results within a hair of the midpoint between two doubles can differ. Half the
  // spacing is at least a quarter of p's binade spacing (smaller just below a power of two).
  const binade = 2 ** (Math.floor(Math.log2(Math.abs(p))) - 52);
  if (Math.abs(err) < (binade / 4) * (1 - 2 ** -10)) return p;
  const dir = err > 0 ? 1 : -1;
  const neighbor = nextAfter(p, dir);
  const u = Math.abs(neighbor - p);
  const half = u / 2;
  // The 64-bit rounding lands exactly on the midpoint when exact is within half a 64-bit unit of it
  if (half - Math.abs(err) <= u * 2 ** -12 && Math.abs(err) < half) return isOddMantissa(p) ? neighbor : p;
  return p;
}

const SPLITTER = 134217729; // 2^27 + 1
const SAFE = 2 ** 995;

/** Exact error of a + b (TwoSum). */
function sumError(a: number, b: number, s: number): number {
  const bb = s - a;
  return a - (s - bb) + (b - bb);
}

/** Exact error of a * b (Dekker's TwoProduct). */
function productError(a: number, b: number, p: number): number {
  let c = SPLITTER * a;
  const ah = c - (c - a);
  const al = a - ah;
  c = SPLITTER * b;
  const bh = c - (c - b);
  const bl = b - bh;
  return ah * bh - p + ah * bl + al * bh + al * bl;
}

export function xAdd(a: number, b: number): number {
  const s = a + b;
  if (!Number.isFinite(s)) return s;
  return doubleRound(s, sumError(a, b, s));
}

export function xSub(a: number, b: number): number {
  return xAdd(a, -b);
}

export function xMul(a: number, b: number): number {
  const p = a * b;
  if (!Number.isFinite(p) || p === 0 || Math.abs(a) > SAFE || Math.abs(b) > SAFE || Math.abs(p) < 2 ** -969) return p;
  return doubleRound(p, productError(a, b, p));
}

export function xDiv(a: number, b: number): number {
  const q = a / b;
  if (!Number.isFinite(q) || q === 0 || Math.abs(q) > SAFE || Math.abs(b) > SAFE || Math.abs(q) < 2 ** -969 || Math.abs(b) < 2 ** -969) return q;
  // a - q*b is exact; (a - q*b) / b is the error of q
  const qb = q * b;
  const rem = a - qb - productError(q, b, qb);
  return doubleRound(q, rem / b);
}

/** Spacing of doubles at x's binade (the "unit in the last place" of a normal double). */
function ulpOf(x: number): number {
  const ax = Math.abs(x);
  if (ax === 0 || !Number.isFinite(ax)) return 0;
  return 2 ** (Math.floor(Math.log2(ax)) - 52);
}

/** Excel's rule for the last + or - of a formula: results this close to zero become 0. */
export function snapNearZero(r: number, a: number, b: number): number {
  if (r === 0) return r;
  const big = Math.max(Math.abs(a), Math.abs(b));
  return Math.abs(r) < 8 * ulpOf(big) ? 0 : r;
}

/** SUM as Excel does it: left to right in doubles, with the near-zero rule on the last addition. */
export function excelSum(nums: number[]): number {
  if (!nums.length) return 0;
  let s = nums[0];
  for (let i = 1; i < nums.length; i++) {
    const next = xAdd(s, nums[i]);
    s = i === nums.length - 1 ? snapNearZero(next, s, nums[i]) : next;
  }
  return s;
}

/**
 * Reads a typed decimal number the way Excel does: only the first 15 significant digits count, and the
 * rest are dropped (not rounded), so 0.66666666666666666 is 0.666666666666666.
 */
export function parseExcelNumber(text: string): number {
  const m = /^([+-]?)(\d*)(?:\.(\d*))?(?:[eE]([+-]?\d+))?$/.exec(text.trim());
  if (!m) return parseFloat(text);
  const [, sign, intPart = '', fracPart = '', exp] = m;
  const all = intPart + fracPart;
  const first = all.search(/[1-9]/);
  if (first < 0 || all.length - first <= 15) return parseFloat(text);
  const cut = first + 15;
  const kept = all.slice(0, cut) + '0'.repeat(all.length - cut);
  return parseFloat(`${sign}${kept.slice(0, intPart.length) || '0'}.${kept.slice(intPart.length) || '0'}${exp !== undefined ? 'e' + exp : ''}`);
}
