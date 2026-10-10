import { Errors, ErrorValue, isError } from '../../values';
import { argNum, collectNumbers, numResult, toMatrix } from '../coerce';
import type { Arg, Ctx, Value } from '../types';
import { define } from './registry';

function nums(ctx: Ctx, args: Arg[], defaults: (number | undefined)[]): number[] | ErrorValue {
  const out: number[] = [];
  for (let i = 0; i < defaults.length; i++) {
    const v = argNum(ctx, args[i], defaults[i]);
    if (isError(v)) return v;
    out.push(v);
  }
  return out;
}

export function pmt(rate: number, nper: number, pv: number, fv = 0, type = 0): number {
  if (rate === 0) return -(pv + fv) / nper;
  const f = Math.pow(1 + rate, nper);
  return (-rate * (fv + pv * f)) / ((1 + rate * type) * (f - 1));
}

export function fv(rate: number, nper: number, payment: number, pv = 0, type = 0): number {
  if (rate === 0) return -(pv + payment * nper);
  const f = Math.pow(1 + rate, nper);
  return -(pv * f + (payment * (1 + rate * type) * (f - 1)) / rate);
}

function ipmt(rate: number, per: number, nper: number, pv: number, fvv = 0, type = 0): number {
  const payment = pmt(rate, nper, pv, fvv, type);
  let interest: number;
  if (per === 1) interest = type === 1 ? 0 : -pv;
  else interest = type === 1 ? fv(rate, per - 2, payment, pv, 1) - payment : fv(rate, per - 1, payment, pv, 0);
  return interest * rate;
}

define('PMT', 'Financial', 'rate, nper, pv, [fv], [type]', 'Calculates the payment for a loan with constant payments and interest rate. Use a monthly rate for monthly payments (e.g. 6%/12).', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, 0, 0]);
  if (isError(n)) return n;
  if (n[1] === 0) return Errors.num('nper (number of periods) can\'t be 0.');
  return numResult(pmt(n[0], n[1], n[2], n[3], n[4] ? 1 : 0));
}, { lift: true });
define('FV', 'Financial', 'rate, nper, pmt, [pv], [type]', 'Returns the future value of an investment.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, 0, 0]);
  if (isError(n)) return n;
  return numResult(fv(n[0], n[1], n[2], n[3], n[4] ? 1 : 0));
}, { lift: true });
define('PV', 'Financial', 'rate, nper, pmt, [fv], [type]', 'Returns the present value of an investment.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, 0, 0]);
  if (isError(n)) return n;
  const [r, np, p, f, t] = n;
  if (r === 0) return -(f + p * np);
  const g = Math.pow(1 + r, np);
  return numResult(-(f + (p * (1 + r * (t ? 1 : 0)) * (g - 1)) / r) / g);
}, { lift: true });
define('NPER', 'Financial', 'rate, pmt, pv, [fv], [type]', 'Returns the number of periods for an investment.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, 0, 0]);
  if (isError(n)) return n;
  const [r, p, pv, f, t0] = n;
  const t = t0 ? 1 : 0;
  if (r === 0) return p === 0 ? Errors.num() : -(pv + f) / p;
  const num = p * (1 + r * t) - f * r;
  const den = pv * r + p * (1 + r * t);
  if (num / den <= 0) return Errors.num('These values never pay off the loan.');
  return numResult(Math.log(num / den) / Math.log(1 + r));
}, { lift: true });
define('IPMT', 'Financial', 'rate, per, nper, pv, [fv], [type]', 'Returns the interest portion of a payment for a given period.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, undefined, 0, 0]);
  if (isError(n)) return n;
  if (n[1] < 1 || n[1] > n[2]) return Errors.num('per must be between 1 and nper.');
  return numResult(ipmt(n[0], n[1], n[2], n[3], n[4], n[5] ? 1 : 0));
}, { lift: true });
define('PPMT', 'Financial', 'rate, per, nper, pv, [fv], [type]', 'Returns the principal portion of a payment for a given period.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, undefined, 0, 0]);
  if (isError(n)) return n;
  if (n[1] < 1 || n[1] > n[2]) return Errors.num('per must be between 1 and nper.');
  const t = n[5] ? 1 : 0;
  return numResult(pmt(n[0], n[2], n[3], n[4], t) - ipmt(n[0], n[1], n[2], n[3], n[4], t));
}, { lift: true });

function cumulative(principal: boolean) {
  return (ctx: Ctx, args: Arg[]): Value => {
    const n = nums(ctx, args, [undefined, undefined, undefined, undefined, undefined, undefined]);
    if (isError(n)) return n;
    const [r, np, pv, start, end, t] = n;
    if (r <= 0 || np <= 0 || pv <= 0 || start < 1 || end < start || end > np) return Errors.num();
    let total = 0;
    const payment = pmt(r, np, pv, 0, t);
    for (let per = start; per <= end; per++) {
      const i = ipmt(r, per, np, pv, 0, t);
      total += principal ? payment - i : i;
    }
    return total;
  };
}
define('CUMIPMT', 'Financial', 'rate, nper, pv, start_period, end_period, type', 'Returns cumulative interest paid between two periods.', cumulative(false));
define('CUMPRINC', 'Financial', 'rate, nper, pv, start_period, end_period, type', 'Returns cumulative principal paid between two periods.', cumulative(true));

define('RATE', 'Financial', 'nper, pmt, pv, [fv], [type], [guess]', 'Returns the interest rate per period.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, 0, 0, 0.1]);
  if (isError(n)) return n;
  const [np, p, pv, f, t0, guess] = n;
  const t = t0 ? 1 : 0;
  let r = guess;
  for (let i = 0; i < 100; i++) {
    const g = (rate: number) => {
      if (Math.abs(rate) < 1e-12) return pv + p * np + f;
      const x = Math.pow(1 + rate, np);
      return pv * x + (p * (1 + rate * t) * (x - 1)) / rate + f;
    };
    const y = g(r);
    const dy = (g(r + 1e-7) - y) / 1e-7;
    if (dy === 0) break;
    const nr = r - y / dy;
    if (Math.abs(nr - r) < 1e-10) return nr;
    r = nr;
  }
  return Errors.num('RATE couldn\'t find an answer. Try giving a guess.');
}, { lift: true });

define('NPV', 'Financial', 'rate, value1, [value2], ...', 'Returns the net present value of future cash flows at a discount rate.', (ctx, [r, ...rest]) => {
  const rate = argNum(ctx, r);
  if (isError(rate)) return rate;
  const values = collectNumbers(ctx, rest);
  if (isError(values)) return values;
  let total = 0;
  values.forEach((v, i) => (total += v / Math.pow(1 + rate, i + 1)));
  return numResult(total);
});

function irrOf(values: number[], guess: number): number | ErrorValue {
  if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) return Errors.num('IRR needs at least one positive and one negative cash flow.');
  const npv = (r: number) => values.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0);
  let r = guess;
  for (let i = 0; i < 100; i++) {
    const y = npv(r);
    const dy = values.reduce((s, v, k) => s - (k * v) / Math.pow(1 + r, k + 1), 0);
    if (dy === 0) break;
    const nr = r - y / dy;
    if (Math.abs(nr - r) < 1e-10) return nr;
    r = nr;
  }
  // Bisection fallback
  let lo = -0.9999;
  let hi = 10;
  if (npv(lo) * npv(hi) > 0) return Errors.num('IRR couldn\'t find an answer.');
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (npv(lo) * npv(mid) <= 0) hi = mid;
    else lo = mid;
  }
  return (lo + hi) / 2;
}

define('IRR', 'Financial', 'values, [guess]', 'Returns the internal rate of return for a series of cash flows.', (ctx, [v, g]) => {
  const values = collectNumbers(ctx, [v]);
  if (isError(values)) return values;
  const guess = argNum(ctx, g, 0.1);
  if (isError(guess)) return guess;
  return irrOf(values, guess);
});
define('MIRR', 'Financial', 'values, finance_rate, reinvest_rate', 'Returns the modified internal rate of return.', (ctx, [v, fr, rr]) => {
  const values = collectNumbers(ctx, [v]);
  if (isError(values)) return values;
  const f = argNum(ctx, fr);
  if (isError(f)) return f;
  const r = argNum(ctx, rr);
  if (isError(r)) return r;
  const n = values.length;
  let pos = 0;
  let neg = 0;
  values.forEach((x, i) => {
    if (x > 0) pos += x * Math.pow(1 + r, n - 1 - i);
    else neg += x / Math.pow(1 + f, i);
  });
  if (neg === 0 || pos === 0) return Errors.div0();
  return Math.pow(-pos / neg, 1 / (n - 1)) - 1;
});

function datedFlows(ctx: Ctx, v: Arg, d: Arg): [number[], number[]] | ErrorValue {
  const vm = toMatrix(ctx, v).data.flat();
  const dm = toMatrix(ctx, d).data.flat();
  if (vm.length !== dm.length) return Errors.num('values and dates must be the same size.');
  const values: number[] = [];
  const dates: number[] = [];
  for (let i = 0; i < vm.length; i++) {
    if (typeof vm[i] !== 'number' || typeof dm[i] !== 'number') return Errors.value('values and dates must all be numbers.');
    values.push(vm[i] as number);
    dates.push(Math.floor(dm[i] as number));
  }
  return [values, dates];
}

define('XNPV', 'Financial', 'rate, values, dates', 'Returns the net present value for cash flows on specific dates.', (ctx, [r, v, d]) => {
  const rate = argNum(ctx, r);
  if (isError(rate)) return rate;
  const flows = datedFlows(ctx, v, d);
  if (isError(flows)) return flows;
  const [values, dates] = flows;
  return numResult(values.reduce((s, x, i) => s + x / Math.pow(1 + rate, (dates[i] - dates[0]) / 365), 0));
});
define('XIRR', 'Financial', 'values, dates, [guess]', 'Returns the internal rate of return for cash flows on specific dates.', (ctx, [v, d, g]) => {
  const flows = datedFlows(ctx, v, d);
  if (isError(flows)) return flows;
  const [values, dates] = flows;
  let r = 0.1;
  const gv = argNum(ctx, g, 0.1);
  if (!isError(gv)) r = gv;
  const f = (rate: number) => values.reduce((s, x, i) => s + x / Math.pow(1 + rate, (dates[i] - dates[0]) / 365), 0);
  for (let i = 0; i < 100; i++) {
    const y = f(r);
    const dy = (f(r + 1e-7) - y) / 1e-7;
    if (dy === 0) break;
    const nr = r - y / dy;
    if (Math.abs(nr - r) < 1e-10) return nr;
    r = nr;
  }
  return Errors.num('XIRR couldn\'t find an answer.');
});

define('SLN', 'Financial', 'cost, salvage, life', 'Straight-line depreciation for one period.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined]);
  if (isError(n)) return n;
  if (n[2] === 0) return Errors.div0('life can\'t be 0.');
  return (n[0] - n[1]) / n[2];
}, { lift: true });
define('SYD', 'Financial', 'cost, salvage, life, per', 'Sum-of-years\' digits depreciation.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, undefined]);
  if (isError(n)) return n;
  const [c, s, l, p] = n;
  if (l <= 0 || p < 1 || p > l) return Errors.num();
  return ((c - s) * (l - p + 1) * 2) / (l * (l + 1));
}, { lift: true });
define('DDB', 'Financial', 'cost, salvage, life, period, [factor]', 'Double-declining balance depreciation.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, undefined, 2]);
  if (isError(n)) return n;
  const [c, s, l, p, f] = n;
  if (l <= 0 || p < 1 || p > l || f <= 0) return Errors.num();
  let book = c;
  let dep = 0;
  for (let i = 1; i <= p; i++) {
    dep = Math.min(book * (f / l), Math.max(0, book - s));
    book -= dep;
  }
  return dep;
}, { lift: true });
define('DB', 'Financial', 'cost, salvage, life, period, [month]', 'Fixed-declining balance depreciation.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined, undefined, undefined, 12]);
  if (isError(n)) return n;
  const [c, s, l, p, m] = n;
  if (c < 0 || s < 0 || l <= 0 || p < 1 || m < 1 || m > 12) return Errors.num();
  const rate = Math.round((1 - Math.pow(s / c, 1 / l)) * 1000) / 1000;
  let total = 0;
  let dep = 0;
  for (let i = 1; i <= p; i++) {
    if (i === 1) dep = (c * rate * m) / 12;
    else if (i === Math.floor(l) + 1) dep = ((c - total) * rate * (12 - m)) / 12;
    else dep = (c - total) * rate;
    total += dep;
  }
  return dep;
}, { lift: true });
define('EFFECT', 'Financial', 'nominal_rate, npery', 'Returns the effective annual interest rate.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined]);
  if (isError(n)) return n;
  const p = Math.trunc(n[1]);
  if (n[0] <= 0 || p < 1) return Errors.num();
  return Math.pow(1 + n[0] / p, p) - 1;
}, { lift: true });
define('NOMINAL', 'Financial', 'effect_rate, npery', 'Returns the nominal annual interest rate.', (ctx, args) => {
  const n = nums(ctx, args, [undefined, undefined]);
  if (isError(n)) return n;
  const p = Math.trunc(n[1]);
  if (n[0] <= 0 || p < 1) return Errors.num();
  return p * (Math.pow(1 + n[0], 1 / p) - 1);
}, { lift: true });
