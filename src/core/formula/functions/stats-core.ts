// Pure numeric helpers for statistical functions.
import { Errors, type ErrorValue } from '../../values';

export function mean(n: number[]): number | ErrorValue {
  if (!n.length) return Errors.div0('There are no numbers to average.');
  let s = 0;
  for (const x of n) s += x;
  return s / n.length;
}

export function variance(n: number[], sample: boolean): number | ErrorValue {
  const k = n.length;
  if (k < (sample ? 2 : 1)) return Errors.div0(sample ? 'Need at least two numbers for a sample variance.' : 'There are no numbers.');
  const m = n.reduce((a, b) => a + b, 0) / k;
  let ss = 0;
  for (const x of n) ss += (x - m) * (x - m);
  return ss / (sample ? k - 1 : k);
}

export function stdev(n: number[], sample: boolean): number | ErrorValue {
  const v = variance(n, sample);
  return typeof v === 'number' ? Math.sqrt(v) : v;
}

export function median(n: number[]): number | ErrorValue {
  if (!n.length) return Errors.num('There are no numbers.');
  const s = [...n].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

export function mode(n: number[]): number | ErrorValue {
  const counts = new Map<number, number>();
  let best = NaN;
  let bestCount = 1;
  for (const x of n) {
    const c = (counts.get(x) ?? 0) + 1;
    counts.set(x, c);
    if (c > bestCount) {
      bestCount = c;
      best = x;
    }
  }
  if (isNaN(best)) return Errors.na('No value appears more than once.');
  return best;
}

export function kth(n: number[], k: number, largest: boolean): number | ErrorValue {
  const kk = Math.ceil(k);
  if (kk < 1 || kk > n.length) return Errors.num(`k must be between 1 and ${n.length} (the count of numbers).`);
  const s = [...n].sort((a, b) => (largest ? b - a : a - b));
  return s[kk - 1];
}

/** PERCENTILE.INC (inclusive=true) or PERCENTILE.EXC. */
export function percentile(n: number[], p: number, inclusive: boolean): number | ErrorValue {
  if (!n.length) return Errors.num('There are no numbers.');
  const s = [...n].sort((a, b) => a - b);
  const len = s.length;
  if (inclusive) {
    if (p < 0 || p > 1) return Errors.num('The percentile must be between 0 and 1.');
    const pos = p * (len - 1);
    const lo = Math.floor(pos);
    const frac = pos - lo;
    return lo + 1 < len ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
  }
  const pos = p * (len + 1) - 1;
  if (pos < 0 || pos > len - 1) return Errors.num('The percentile is outside the range PERCENTILE.EXC can calculate.');
  const lo = Math.floor(pos);
  const frac = pos - lo;
  return lo + 1 < len ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
}

/** Error function approximation (Abramowitz-Stegun 7.1.26 is too coarse; use a series/continued fraction). */
export function erf(x: number): number {
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x);
  if (x < 3) {
    // Taylor series
    let sum = x;
    let term = x;
    for (let n = 1; n < 100; n++) {
      term *= (-x * x) / n;
      const add = term / (2 * n + 1);
      sum += add;
      if (Math.abs(add) < 1e-16) break;
    }
    return (sign * 2 * sum) / Math.sqrt(Math.PI);
  }
  // Continued fraction for erfc
  let f = 0;
  for (let n = 60; n >= 1; n--) f = n / 2 / (x + f);
  const erfc = Math.exp(-x * x) / Math.sqrt(Math.PI) / (x + f);
  return sign * (1 - erfc);
}

export function normCdf(z: number): number {
  return 0.5 * (1 + erf(z / Math.SQRT2));
}

export function normPdf(z: number): number {
  return Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
}

/** Inverse standard normal CDF (Acklam's algorithm, refined with one Newton step). */
export function normInv(p: number): number {
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const pl = 0.02425;
  let q: number;
  let x: number;
  if (p < pl) {
    q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - pl) {
    q = p - 0.5;
    const r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  const e = normCdf(x) - p;
  const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
  return x - u / (1 + (x * u) / 2);
}

export function lnGamma(x: number): number {
  const g = 7;
  const coef = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ];
  if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - lnGamma(1 - x);
  x -= 1;
  let a = coef[0];
  const t = x + g + 0.5;
  for (let i = 1; i < g + 2; i++) a += coef[i] / (x + i);
  return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
}
