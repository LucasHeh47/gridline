import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import { argBool, argInt, argNum, collectNumbers, numResult, toMatrix } from '../coerce';
import { RangeRef, Matrix, type Arg, type Ctx, type Value } from '../types';
import { conditionalValues, numbersOnly, sumOf } from './math';
import { define, alias } from './registry';
import * as S from './stats-core';

function wrap(f: (n: number[]) => number | ErrorValue, opts?: { includeAll?: boolean }) {
  return (ctx: Ctx, args: Arg[]): Value => {
    const nums = collectNumbers(ctx, args, opts);
    if (isError(nums)) return nums;
    const r = f(nums);
    return isError(r) ? r : numResult(r);
  };
}

define('AVERAGE', 'Statistical', 'number1, [number2], ...', 'Returns the average (arithmetic mean) of its arguments.', wrap(S.mean));
define('AVERAGEA', 'Statistical', 'value1, [value2], ...', 'Averages values, counting text as 0 and TRUE as 1.', wrap(S.mean, { includeAll: true }));
define('COUNT', 'Statistical', 'value1, [value2], ...', 'Counts how many cells contain numbers.', (ctx, args) => {
  let n = 0;
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof RangeRef) {
      ctx.forEachInRange(a, (v) => {
        if (typeof v === 'number') n++;
      });
    } else if (a instanceof Matrix) {
      for (const row of a.data) for (const v of row) if (typeof v === 'number') n++;
    } else if (typeof a === 'number' || (typeof a === 'boolean') || (typeof a === 'string' && !isNaN(Number(a)) && a.trim() !== '')) n++;
  }
  return n;
});
define('COUNTA', 'Statistical', 'value1, [value2], ...', 'Counts how many cells are not empty.', (ctx, args) => {
  let n = 0;
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof RangeRef) {
      ctx.forEachInRange(a, (v) => {
        if (v !== null) n++;
      });
    } else if (a instanceof Matrix) {
      for (const row of a.data) for (const v of row) if (v !== null) n++;
    } else n++;
  }
  return n;
});
define('COUNTBLANK', 'Statistical', 'range', 'Counts empty cells in a range (cells with "" count as empty).', (ctx, [a]) => {
  if (a instanceof RangeRef) {
    let filled = 0;
    ctx.forEachInRange(a, (v) => {
      if (v !== null && v !== '') filled++;
    });
    return a.rows * a.cols - filled;
  }
  const m = toMatrix(ctx, a);
  let n = 0;
  for (const row of m.data) for (const v of row) if (v === null || v === '') n++;
  return n;
});

define('COUNTIF', 'Statistical', 'range, criteria', 'Counts cells that meet a condition, e.g. COUNTIF(A:A, "Yes").', (ctx, [range, crit]) => {
  const r = conditionalValues(ctx, null, [[range, crit]]);
  return isError(r) ? r : r.count;
});
define(
  'COUNTIFS',
  'Statistical',
  'criteria_range1, criteria1, [criteria_range2], [criteria2], ...',
  'Counts cells that meet several conditions.',
  (ctx, args) => {
    if (args.length % 2) return Errors.value('Criteria must come in pairs: a range, then its condition.');
    const pairs: [Arg, Arg][] = [];
    for (let i = 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const r = conditionalValues(ctx, null, pairs);
    return isError(r) ? r : r.count;
  },
  { min: 2 },
);

function ifsAgg(f: (n: number[]) => number | ErrorValue, emptyValue?: number) {
  return (ctx: Ctx, args: Arg[]): Value => {
    const rest = args.slice(1);
    if (rest.length % 2) return Errors.value('Criteria must come in pairs: a range, then its condition.');
    const pairs: [Arg, Arg][] = [];
    for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]]);
    const r = conditionalValues(ctx, args[0] ?? null, pairs);
    if (isError(r)) return r;
    const nums = numbersOnly(r.values);
    if (isError(nums)) return nums;
    if (!nums.length && emptyValue !== undefined) return emptyValue;
    const v = f(nums);
    return isError(v) ? v : numResult(v);
  };
}

define('AVERAGEIF', 'Statistical', 'range, criteria, [average_range]', 'Averages the cells that meet a condition.', (ctx, [range, crit, avg]) => {
  const r = conditionalValues(ctx, avg ?? range ?? null, [[range, crit]]);
  if (isError(r)) return r;
  const nums = numbersOnly(r.values);
  if (isError(nums)) return nums;
  return nums.length ? numResult(sumOf(nums) / nums.length) : Errors.div0('No cells matched the condition, so there\'s nothing to average.');
});
define(
  'AVERAGEIFS',
  'Statistical',
  'average_range, criteria_range1, criteria1, [criteria_range2], [criteria2], ...',
  'Averages the cells that meet several conditions.',
  ifsAgg((n) => (n.length ? sumOf(n) / n.length : Errors.div0('No cells matched all the conditions.'))),
  { min: 3 },
);
define('MAXIFS', 'Statistical', 'max_range, criteria_range1, criteria1, ...', 'Returns the largest value among cells that meet conditions.', ifsAgg((n) => Math.max(...n), 0), { min: 3 });
define('MINIFS', 'Statistical', 'min_range, criteria_range1, criteria1, ...', 'Returns the smallest value among cells that meet conditions.', ifsAgg((n) => Math.min(...n), 0), { min: 3 });

function maxOf(n: number[]): number {
  let m = -Infinity;
  for (const x of n) if (x > m) m = x;
  return n.length ? m : 0;
}
function minOf(n: number[]): number {
  let m = Infinity;
  for (const x of n) if (x < m) m = x;
  return n.length ? m : 0;
}

define('MAX', 'Statistical', 'number1, [number2], ...', 'Returns the largest number.', wrap(maxOf));
define('MAXA', 'Statistical', 'value1, [value2], ...', 'Returns the largest value, counting text as 0 and TRUE as 1.', wrap(maxOf, { includeAll: true }));
define('MIN', 'Statistical', 'number1, [number2], ...', 'Returns the smallest number.', wrap(minOf));
define('MINA', 'Statistical', 'value1, [value2], ...', 'Returns the smallest value, counting text as 0 and TRUE as 1.', wrap(minOf, { includeAll: true }));
define('MEDIAN', 'Statistical', 'number1, [number2], ...', 'Returns the middle number of the given numbers.', wrap(S.median));
define('MODE.SNGL', 'Statistical', 'number1, [number2], ...', 'Returns the most frequently occurring value.', wrap(S.mode));
alias('MODE', 'MODE.SNGL');
define('MODE.MULT', 'Statistical', 'number1, [number2], ...', 'Returns a vertical array of the most frequent values.', (ctx, args) => {
  const nums = collectNumbers(ctx, args);
  if (isError(nums)) return nums;
  const counts = new Map<number, number>();
  for (const x of nums) counts.set(x, (counts.get(x) ?? 0) + 1);
  const best = Math.max(0, ...counts.values());
  if (best < 2) return Errors.na('No value appears more than once.');
  return Matrix.column([...counts].filter(([, c]) => c === best).map(([v]) => v));
});
define('STDEV.S', 'Statistical', 'number1, [number2], ...', 'Estimates standard deviation from a sample.', wrap((n) => S.stdev(n, true)));
define('STDEV.P', 'Statistical', 'number1, [number2], ...', 'Standard deviation of an entire population.', wrap((n) => S.stdev(n, false)));
alias('STDEV', 'STDEV.S');
alias('STDEVP', 'STDEV.P');
define('STDEVA', 'Statistical', 'value1, [value2], ...', 'Sample standard deviation counting text as 0 and TRUE as 1.', wrap((n) => S.stdev(n, true), { includeAll: true }));
define('VAR.S', 'Statistical', 'number1, [number2], ...', 'Estimates variance from a sample.', wrap((n) => S.variance(n, true)));
define('VAR.P', 'Statistical', 'number1, [number2], ...', 'Variance of an entire population.', wrap((n) => S.variance(n, false)));
alias('VAR', 'VAR.S');
alias('VARP', 'VAR.P');

define('LARGE', 'Statistical', 'array, k', 'Returns the k-th largest value, e.g. LARGE(A:A, 2) is the second largest.', (ctx, [a, k]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  if (k instanceof RangeRef || k instanceof Matrix) {
    const km = toMatrix(ctx, k);
    return Matrix.fill(km.rows, km.cols, (r, c) => {
      const kv = km.get(r, c);
      return typeof kv === 'number' ? S.kth(nums, kv, true) : Errors.value();
    });
  }
  const kk = argNum(ctx, k);
  return isError(kk) ? kk : S.kth(nums, kk, true);
});
define('SMALL', 'Statistical', 'array, k', 'Returns the k-th smallest value.', (ctx, [a, k]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  if (k instanceof RangeRef || k instanceof Matrix) {
    const km = toMatrix(ctx, k);
    return Matrix.fill(km.rows, km.cols, (r, c) => {
      const kv = km.get(r, c);
      return typeof kv === 'number' ? S.kth(nums, kv, false) : Errors.value();
    });
  }
  const kk = argNum(ctx, k);
  return isError(kk) ? kk : S.kth(nums, kk, false);
});

function rank(avg: boolean) {
  return (ctx: Ctx, [n, ref, order]: Arg[]): Value => {
    const x = argNum(ctx, n);
    if (isError(x)) return x;
    const asc = argBool(ctx, order, false);
    if (isError(asc)) return asc;
    const nums = collectNumbers(ctx, [ref], { ignoreErrors: false });
    if (isError(nums)) return nums;
    let better = 0;
    let equal = 0;
    for (const v of nums) {
      if (v === x) equal++;
      else if (asc ? v < x : v > x) better++;
    }
    if (!equal) return Errors.na('The number isn\'t in the list being ranked.');
    return avg ? better + (equal + 1) / 2 : better + 1;
  };
}
define('RANK.EQ', 'Statistical', 'number, ref, [order]', 'Returns the rank of a number in a list (0 = descending, 1 = ascending).', rank(false), { lift: [0] });
alias('RANK', 'RANK.EQ');
define('RANK.AVG', 'Statistical', 'number, ref, [order]', 'Returns the rank of a number; ties get the average rank.', rank(true), { lift: [0] });

define('PERCENTILE.INC', 'Statistical', 'array, k', 'Returns the k-th percentile (k between 0 and 1).', (ctx, [a, k]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  const p = argNum(ctx, k);
  return isError(p) ? p : S.percentile(nums, p, true);
});
alias('PERCENTILE', 'PERCENTILE.INC');
define('PERCENTILE.EXC', 'Statistical', 'array, k', 'Returns the k-th percentile, exclusive of 0 and 1.', (ctx, [a, k]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  const p = argNum(ctx, k);
  return isError(p) ? p : S.percentile(nums, p, false);
});
define('QUARTILE.INC', 'Statistical', 'array, quart', 'Returns a quartile (0=min, 1=25%, 2=median, 3=75%, 4=max).', (ctx, [a, q]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  const qq = argInt(ctx, q);
  if (isError(qq)) return qq;
  if (qq < 0 || qq > 4) return Errors.num('quart must be 0, 1, 2, 3 or 4.');
  return S.percentile(nums, qq / 4, true);
});
alias('QUARTILE', 'QUARTILE.INC');
define('QUARTILE.EXC', 'Statistical', 'array, quart', 'Returns a quartile, exclusive (quart 1-3).', (ctx, [a, q]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  const qq = argInt(ctx, q);
  if (isError(qq)) return qq;
  if (qq < 1 || qq > 3) return Errors.num('quart must be 1, 2 or 3.');
  return S.percentile(nums, qq / 4, false);
});

define('PERCENTRANK.INC', 'Statistical', 'array, x, [significance]', 'Returns the rank of a value as a percentage of the data set.', (ctx, [a, xv, sig]) => {
  const nums = collectNumbers(ctx, [a]);
  if (isError(nums)) return nums;
  const x = argNum(ctx, xv);
  if (isError(x)) return x;
  const digits = argInt(ctx, sig, 3);
  if (isError(digits)) return digits;
  const s = [...nums].sort((p, q) => p - q);
  if (!s.length || x < s[0] || x > s[s.length - 1]) return Errors.na('x is outside the range of the data.');
  let lo = 0;
  while (lo < s.length - 1 && s[lo + 1] <= x) lo++;
  let rankPos: number;
  if (s[lo] === x) {
    const first = s.indexOf(x);
    rankPos = first / (s.length - 1);
  } else {
    rankPos = (lo + (x - s[lo]) / (s[lo + 1] - s[lo])) / (s.length - 1);
  }
  const f = Math.pow(10, digits);
  return Math.floor(rankPos * f) / f;
});
alias('PERCENTRANK', 'PERCENTRANK.INC');

function pairedNumbers(ctx: Ctx, a: Arg, b: Arg): [number[], number[]] | ErrorValue {
  const ma = toMatrix(ctx, a);
  const mb = toMatrix(ctx, b);
  if (ma.rows * ma.cols !== mb.rows * mb.cols) return Errors.na('Both arrays must have the same number of values.');
  const xs: number[] = [];
  const ys: number[] = [];
  const fa: Scalar[] = ma.data.flat();
  const fb: Scalar[] = mb.data.flat();
  for (let i = 0; i < fa.length; i++) {
    const x = fa[i];
    const y = fb[i];
    if (x instanceof ErrorValue) return x;
    if (y instanceof ErrorValue) return y;
    if (typeof x === 'number' && typeof y === 'number') {
      xs.push(x);
      ys.push(y);
    }
  }
  return [xs, ys];
}

function regression(ys: number[], xs: number[]) {
  const n = xs.length;
  const mx = xs.reduce((a, b) => a + b, 0) / n;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
    syy += (ys[i] - my) ** 2;
  }
  return { n, mx, my, sxy, sxx, syy };
}

define('CORREL', 'Statistical', 'array1, array2', 'Returns the correlation coefficient between two data sets.', (ctx, [a, b]) => {
  const p = pairedNumbers(ctx, a, b);
  if (isError(p)) return p;
  const r = regression(p[1], p[0]);
  if (r.n < 2 || r.sxx === 0 || r.syy === 0) return Errors.div0('Need at least two varying values in each array.');
  return r.sxy / Math.sqrt(r.sxx * r.syy);
});
alias('PEARSON', 'CORREL');
define('RSQ', 'Statistical', 'known_ys, known_xs', 'Returns the R² of a linear regression line.', (ctx, [a, b]) => {
  const p = pairedNumbers(ctx, b, a);
  if (isError(p)) return p;
  const r = regression(p[1], p[0]);
  if (r.n < 2 || r.sxx === 0 || r.syy === 0) return Errors.div0();
  return (r.sxy * r.sxy) / (r.sxx * r.syy);
});
define('SLOPE', 'Statistical', 'known_ys, known_xs', 'Returns the slope of the linear regression line.', (ctx, [ys, xs]) => {
  const p = pairedNumbers(ctx, xs, ys);
  if (isError(p)) return p;
  const r = regression(p[1], p[0]);
  if (r.n < 2 || r.sxx === 0) return Errors.div0('Need at least two different x values.');
  return r.sxy / r.sxx;
});
define('INTERCEPT', 'Statistical', 'known_ys, known_xs', 'Returns the y-intercept of the linear regression line.', (ctx, [ys, xs]) => {
  const p = pairedNumbers(ctx, xs, ys);
  if (isError(p)) return p;
  const r = regression(p[1], p[0]);
  if (r.n < 2 || r.sxx === 0) return Errors.div0('Need at least two different x values.');
  return r.my - (r.sxy / r.sxx) * r.mx;
});
define('STEYX', 'Statistical', 'known_ys, known_xs', 'Returns the standard error of the predicted y-values.', (ctx, [ys, xs]) => {
  const p = pairedNumbers(ctx, xs, ys);
  if (isError(p)) return p;
  const r = regression(p[1], p[0]);
  if (r.n < 3 || r.sxx === 0) return Errors.div0();
  return Math.sqrt((r.syy - (r.sxy * r.sxy) / r.sxx) / (r.n - 2));
});
define(
  'FORECAST.LINEAR',
  'Statistical',
  'x, known_ys, known_xs',
  'Predicts a value along a linear trend.',
  (ctx, [xv, ys, xs]) => {
    const x = argNum(ctx, xv);
    if (isError(x)) return x;
    const p = pairedNumbers(ctx, xs, ys);
    if (isError(p)) return p;
    const r = regression(p[1], p[0]);
    if (r.n < 2 || r.sxx === 0) return Errors.div0('Need at least two different x values.');
    const slope = r.sxy / r.sxx;
    return r.my - slope * r.mx + slope * x;
  },
  { lift: [0] },
);
alias('FORECAST', 'FORECAST.LINEAR');
define('TREND', 'Statistical', "known_ys, [known_xs], [new_xs]", 'Returns values along a linear trend.', (ctx, [ys, xs, nx]) => {
  const ym = toMatrix(ctx, ys);
  const yvals = ym.data.flat();
  const xvals = xs !== undefined ? toMatrix(ctx, xs).data.flat() : yvals.map((_, i) => i + 1);
  const px: number[] = [];
  const py: number[] = [];
  for (let i = 0; i < yvals.length; i++) {
    if (typeof yvals[i] === 'number' && typeof xvals[i] === 'number') {
      px.push(xvals[i] as number);
      py.push(yvals[i] as number);
    }
  }
  const r = regression(py, px);
  if (r.n < 2 || r.sxx === 0) return Errors.div0();
  const slope = r.sxy / r.sxx;
  const icpt = r.my - slope * r.mx;
  const target = nx !== undefined ? toMatrix(ctx, nx) : xs !== undefined ? toMatrix(ctx, xs) : Matrix.fill(ym.rows, ym.cols, (rr, cc) => rr * ym.cols + cc + 1);
  return Matrix.fill(target.rows, target.cols, (rr, cc) => {
    const v = target.get(rr, cc);
    return typeof v === 'number' ? icpt + slope * v : Errors.value();
  });
});

function covar(sample: boolean) {
  return (ctx: Ctx, [a, b]: Arg[]): Value => {
    const p = pairedNumbers(ctx, a, b);
    if (isError(p)) return p;
    const r = regression(p[1], p[0]);
    if (r.n < (sample ? 2 : 1)) return Errors.div0();
    return r.sxy / (sample ? r.n - 1 : r.n);
  };
}
define('COVARIANCE.P', 'Statistical', 'array1, array2', 'Returns population covariance.', covar(false));
alias('COVAR', 'COVARIANCE.P');
define('COVARIANCE.S', 'Statistical', 'array1, array2', 'Returns sample covariance.', covar(true));

define('GEOMEAN', 'Statistical', 'number1, [number2], ...', 'Returns the geometric mean.', wrap((n) => {
  if (!n.length || n.some((x) => x <= 0)) return Errors.num('GEOMEAN needs positive numbers.');
  return Math.exp(n.reduce((a, b) => a + Math.log(b), 0) / n.length);
}));
define('HARMEAN', 'Statistical', 'number1, [number2], ...', 'Returns the harmonic mean.', wrap((n) => {
  if (!n.length || n.some((x) => x <= 0)) return Errors.num('HARMEAN needs positive numbers.');
  return n.length / n.reduce((a, b) => a + 1 / b, 0);
}));
define('AVEDEV', 'Statistical', 'number1, [number2], ...', 'Returns the average of absolute deviations from the mean.', wrap((n) => {
  if (!n.length) return Errors.num();
  const m = n.reduce((a, b) => a + b, 0) / n.length;
  return n.reduce((a, b) => a + Math.abs(b - m), 0) / n.length;
}));
define('DEVSQ', 'Statistical', 'number1, [number2], ...', 'Returns the sum of squared deviations from the mean.', wrap((n) => {
  if (!n.length) return Errors.num();
  const m = n.reduce((a, b) => a + b, 0) / n.length;
  return n.reduce((a, b) => a + (b - m) ** 2, 0);
}));
define('SKEW', 'Statistical', 'number1, [number2], ...', 'Returns the skewness of a distribution.', wrap((n) => {
  const k = n.length;
  const sd = S.stdev(n, true);
  if (k < 3 || isError(sd) || sd === 0) return Errors.div0();
  const m = n.reduce((a, b) => a + b, 0) / k;
  return (k / ((k - 1) * (k - 2))) * n.reduce((a, b) => a + ((b - m) / sd) ** 3, 0);
}));
define('KURT', 'Statistical', 'number1, [number2], ...', 'Returns the kurtosis of a data set.', wrap((n) => {
  const k = n.length;
  const sd = S.stdev(n, true);
  if (k < 4 || isError(sd) || sd === 0) return Errors.div0();
  const m = n.reduce((a, b) => a + b, 0) / k;
  const s4 = n.reduce((a, b) => a + ((b - m) / sd) ** 4, 0);
  return ((k * (k + 1)) / ((k - 1) * (k - 2) * (k - 3))) * s4 - (3 * (k - 1) ** 2) / ((k - 2) * (k - 3));
}));

define('STANDARDIZE', 'Statistical', 'x, mean, standard_dev', 'Returns a normalized value (z-score).', (ctx, [x, m, s]) => {
  const xv = argNum(ctx, x);
  const mv = argNum(ctx, m);
  const sv = argNum(ctx, s);
  for (const v of [xv, mv, sv]) if (isError(v)) return v;
  if ((sv as number) <= 0) return Errors.num('standard_dev must be greater than 0.');
  return ((xv as number) - (mv as number)) / (sv as number);
}, { lift: true });

define('NORM.DIST', 'Statistical', 'x, mean, standard_dev, cumulative', 'Returns the normal distribution.', (ctx, [x, m, s, c]) => {
  const xv = argNum(ctx, x);
  const mv = argNum(ctx, m);
  const sv = argNum(ctx, s);
  const cv = argBool(ctx, c);
  for (const v of [xv, mv, sv, cv]) if (isError(v)) return v as ErrorValue;
  if ((sv as number) <= 0) return Errors.num('standard_dev must be greater than 0.');
  const z = ((xv as number) - (mv as number)) / (sv as number);
  return cv ? S.normCdf(z) : S.normPdf(z) / (sv as number);
}, { lift: true });
alias('NORMDIST', 'NORM.DIST');
define('NORM.S.DIST', 'Statistical', 'z, cumulative', 'Returns the standard normal distribution.', (ctx, [z, c]) => {
  const zv = argNum(ctx, z);
  if (isError(zv)) return zv;
  const cv = argBool(ctx, c, true);
  if (isError(cv)) return cv;
  return cv ? S.normCdf(zv) : S.normPdf(zv);
}, { lift: true, min: 1 });
define('NORMSDIST', 'Statistical', 'z', 'Returns the cumulative standard normal distribution.', (ctx, [z]) => {
  const zv = argNum(ctx, z);
  return isError(zv) ? zv : S.normCdf(zv);
}, { lift: true });
define('NORM.INV', 'Statistical', 'probability, mean, standard_dev', 'Returns the inverse of the normal cumulative distribution.', (ctx, [p, m, s]) => {
  const pv = argNum(ctx, p);
  const mv = argNum(ctx, m);
  const sv = argNum(ctx, s);
  for (const v of [pv, mv, sv]) if (isError(v)) return v;
  if ((pv as number) <= 0 || (pv as number) >= 1) return Errors.num('probability must be between 0 and 1.');
  if ((sv as number) <= 0) return Errors.num();
  return (mv as number) + (sv as number) * S.normInv(pv as number);
}, { lift: true });
alias('NORMINV', 'NORM.INV');
define('NORM.S.INV', 'Statistical', 'probability', 'Returns the inverse of the standard normal distribution.', (ctx, [p]) => {
  const pv = argNum(ctx, p);
  if (isError(pv)) return pv;
  if (pv <= 0 || pv >= 1) return Errors.num('probability must be between 0 and 1.');
  return S.normInv(pv);
}, { lift: true });
alias('NORMSINV', 'NORM.S.INV');

define('BINOM.DIST', 'Statistical', 'number_s, trials, probability_s, cumulative', 'Returns the binomial distribution probability.', (ctx, [k, n, p, c]) => {
  const kv = argInt(ctx, k);
  const nv = argInt(ctx, n);
  const pv = argNum(ctx, p);
  const cv = argBool(ctx, c);
  for (const v of [kv, nv, pv, cv]) if (isError(v)) return v as ErrorValue;
  const K = kv as number;
  const N = nv as number;
  const P = pv as number;
  if (K < 0 || K > N || P < 0 || P > 1) return Errors.num();
  const pmf = (i: number) => Math.exp(S.lnGamma(N + 1) - S.lnGamma(i + 1) - S.lnGamma(N - i + 1) + i * Math.log(P || 1e-300) + (N - i) * Math.log(1 - P || 1e-300));
  if (!cv) return pmf(K);
  let s = 0;
  for (let i = 0; i <= K; i++) s += pmf(i);
  return Math.min(1, s);
}, { lift: true });
alias('BINOMDIST', 'BINOM.DIST');
define('POISSON.DIST', 'Statistical', 'x, mean, cumulative', 'Returns the Poisson distribution.', (ctx, [x, m, c]) => {
  const xv = argInt(ctx, x);
  const mv = argNum(ctx, m);
  const cv = argBool(ctx, c);
  for (const v of [xv, mv, cv]) if (isError(v)) return v as ErrorValue;
  const X = xv as number;
  const M = mv as number;
  if (X < 0 || M < 0) return Errors.num();
  const pmf = (i: number) => Math.exp(-M + i * Math.log(M || 1e-300) - S.lnGamma(i + 1));
  if (!cv) return pmf(X);
  let s = 0;
  for (let i = 0; i <= X; i++) s += pmf(i);
  return Math.min(1, s);
}, { lift: true });
alias('POISSON', 'POISSON.DIST');
define('EXPON.DIST', 'Statistical', 'x, lambda, cumulative', 'Returns the exponential distribution.', (ctx, [x, l, c]) => {
  const xv = argNum(ctx, x);
  const lv = argNum(ctx, l);
  const cv = argBool(ctx, c);
  for (const v of [xv, lv, cv]) if (isError(v)) return v as ErrorValue;
  if ((xv as number) < 0 || (lv as number) <= 0) return Errors.num();
  return cv ? 1 - Math.exp(-(lv as number) * (xv as number)) : (lv as number) * Math.exp(-(lv as number) * (xv as number));
}, { lift: true });
alias('EXPONDIST', 'EXPON.DIST');

define('FREQUENCY', 'Statistical', 'data_array, bins_array', 'Counts how many values fall within each bin; returns a vertical array.', (ctx, [d, b]) => {
  const data = collectNumbers(ctx, [d]);
  if (isError(data)) return data;
  const bins = collectNumbers(ctx, [b]);
  if (isError(bins)) return bins;
  const sorted = bins.map((v, i) => ({ v, i })).sort((x, y) => x.v - y.v);
  const counts = new Array(bins.length + 1).fill(0);
  for (const x of data) {
    let placed = false;
    for (const s of sorted) {
      if (x <= s.v) {
        counts[s.i]++;
        placed = true;
        break;
      }
    }
    if (!placed) counts[bins.length]++;
  }
  return Matrix.column(counts);
});

