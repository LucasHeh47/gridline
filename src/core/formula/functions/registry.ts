import type { Arg, Ctx, Value } from '../types';

export type Category =
  | 'Math'
  | 'Statistical'
  | 'Logical'
  | 'Lookup'
  | 'Text'
  | 'Date & time'
  | 'Information'
  | 'Financial'
  | 'Personal finance'
  | 'NFL';

export interface FnDef {
  name: string;
  category: Category;
  /** Argument list as shown to users, e.g. "number1, [number2], ..." */
  args: string;
  desc: string;
  min: number;
  max: number;
  /** Receives unevaluated AST nodes (for IF, IFERROR, LET...). */
  lazy?: boolean;
  volatile?: boolean;
  /**
   * Argument indexes that take single values; if an array is passed, the function is applied
   * to each element and returns an array (like Excel's dynamic arrays). true = all arguments.
   */
  lift?: true | number[];
  /** Uses downloaded data (recalculated when the download finishes). */
  external?: boolean;
  /** Gridline-only function (not available in Excel). */
  custom?: boolean;
  /** Descriptions of individual arguments, shown in the formula hint. */
  argHelp?: Record<string, string>;
  /** Arguments that auto fill advances when dragged (e.g. season +1 per row). */
  series?: { season?: number; week?: number };
  /** Argument positions that get suggestions while typing (stat names, teams...). */
  suggest?: Record<number, SuggestKind>;
  fn: (ctx: Ctx, args: Arg[]) => Value;
}

export type SuggestKind = 'nflPlayerStat' | 'nflTeamStat' | 'nflTeam' | 'nflPlayer' | 'nflGameField' | 'nflInfoField' | 'nflSeasonType' | 'nflGroup' | 'stockField' | 'currency' | 'interval';

export const FUNCTIONS = new Map<string, FnDef>();

export interface DefineOpts {
  lazy?: boolean;
  volatile?: boolean;
  lift?: true | number[];
  min?: number;
  max?: number;
  external?: boolean;
  custom?: boolean;
  argHelp?: Record<string, string>;
  series?: { season?: number; week?: number };
  suggest?: Record<number, SuggestKind>;
}

/** Splits "a, [b], ..." into argument names. */
export function argNames(args: string): string[] {
  if (!args.trim()) return [];
  return args.split(',').map((s) => s.trim());
}

export function define(
  name: string,
  category: Category,
  args: string,
  desc: string,
  fn: (ctx: Ctx, args: Arg[]) => Value,
  opts: DefineOpts = {},
): void {
  const names = argNames(args);
  const variadic = names.includes('...');
  const required = names.filter((n) => n !== '...' && !n.startsWith('[')).length;
  const total = names.filter((n) => n !== '...').length;
  FUNCTIONS.set(name, {
    name,
    category,
    args,
    desc,
    min: opts.min ?? required,
    max: opts.max ?? (variadic ? 255 : total),
    lazy: opts.lazy,
    volatile: opts.volatile,
    lift: opts.lift,
    external: opts.external,
    custom: opts.custom || opts.external,
    argHelp: opts.argHelp,
    series: opts.series,
    suggest: opts.suggest,
    fn,
  });
}

/** Defines an alias that shares another function's implementation. */
export function alias(name: string, target: string, desc?: string): void {
  const t = FUNCTIONS.get(target);
  if (!t) throw new Error(`alias target ${target} missing`);
  FUNCTIONS.set(name, { ...t, name, desc: desc ?? t.desc });
}

function levenshtein(a: string, b: string): number {
  const m = a.length;
  const n = b.length;
  const d: number[] = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    let prev = d[0];
    d[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = d[j];
      d[j] = Math.min(d[j] + 1, d[j - 1] + 1, prev + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = tmp;
    }
  }
  return d[n];
}

export function suggestFunction(name: string): string | null {
  const up = name.toUpperCase();
  let best: string | null = null;
  let bestD = Infinity;
  for (const k of FUNCTIONS.keys()) {
    const dist = levenshtein(up, k);
    if (dist < bestD) {
      bestD = dist;
      best = k;
    }
  }
  const limit = up.length <= 4 ? 1 : 2;
  return bestD <= limit ? best : null;
}

export const VOLATILE_FUNCS = new Set(['RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY', 'OFFSET', 'INDIRECT', 'CELL', 'INFO']);

/** Functions that only exist in newer Excel versions and need the _xlfn. prefix in .xlsx files. */
export const XLFN_FUNCS = new Set([
  'XLOOKUP',
  'XMATCH',
  'IFS',
  'SWITCH',
  'MAXIFS',
  'MINIFS',
  'CONCAT',
  'TEXTJOIN',
  'LET',
  'SEQUENCE',
  'RANDARRAY',
  'UNIQUE',
  'SORT',
  'SORTBY',
  'FILTER',
  'XOR',
  'IFNA',
  'DAYS',
  'ISOWEEKNUM',
  'CEILING.MATH',
  'FLOOR.MATH',
  'STDEV.S',
  'STDEV.P',
  'VAR.S',
  'VAR.P',
  'MODE.SNGL',
  'MODE.MULT',
  'PERCENTILE.INC',
  'PERCENTILE.EXC',
  'QUARTILE.INC',
  'QUARTILE.EXC',
  'RANK.EQ',
  'RANK.AVG',
  'NORM.DIST',
  'NORM.INV',
  'NORM.S.DIST',
  'NORM.S.INV',
  'BINOM.DIST',
  'POISSON.DIST',
  'EXPON.DIST',
  'COVARIANCE.P',
  'COVARIANCE.S',
  'FORECAST.LINEAR',
  'PERCENTRANK.INC',
  'AGGREGATE',
  'NUMBERVALUE',
  'UNICHAR',
  'UNICODE',
  'TEXTBEFORE',
  'TEXTAFTER',
  'TEXTSPLIT',
  'VSTACK',
  'HSTACK',
  'TAKE',
  'DROP',
  'TOCOL',
  'TOROW',
  'CHOOSECOLS',
  'CHOOSEROWS',
  'FORMULATEXT',
  'ISFORMULA',
  'SHEET',
  'SHEETS',
  'COMBINA',
  'ARABIC',
  'BASE',
  'DECIMAL',
  'VALUETOTEXT',
]);

/** Dynamic-array functions additionally need _xlws. in some cases (FILTER, SORT). */
export const XLWS_FUNCS = new Set(['FILTER', 'SORT']);
