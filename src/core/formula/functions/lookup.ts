import { colToName, MAX_COLS, MAX_ROWS, quoteSheetName } from '../../address';
import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import {
  argBool,
  argInt,
  argText,
  clampRef,
  compareValues,
  describeRef,
  numberToText,
  scalarOf,
  toBool,
  toMatrix,
} from '../coerce';
import { hasWildcard, wildcardToRegExp } from '../criteria';
import { tokenize } from '../tokenizer';
import { Matrix, RangeRef, type Arg, type Ctx, type Value } from '../types';
import { define } from './registry';

function show(v: Scalar): string {
  if (typeof v === 'string') return `"${v.length > 40 ? v.slice(0, 40) + '…' : v}"`;
  if (typeof v === 'number') return numberToText(v);
  if (v === null) return 'an empty value';
  return String(v);
}

/** Values of a 1-D slice of a range or matrix. */
function vectorOf(ctx: Ctx, v: Arg, which: 'firstCol' | 'firstRow' | 'flat'): Scalar[] {
  if (v instanceof RangeRef) {
    const { r2, c2 } = clampRef(ctx, v);
    const out: Scalar[] = [];
    if (which === 'firstCol') for (let r = v.r1; r <= r2; r++) out.push(ctx.get(v.sheet, r, v.c1));
    else if (which === 'firstRow') for (let c = v.c1; c <= c2; c++) out.push(ctx.get(v.sheet, v.r1, c));
    else for (let r = v.r1; r <= r2; r++) for (let c = v.c1; c <= c2; c++) out.push(ctx.get(v.sheet, r, c));
    return out;
  }
  const m = toMatrix(ctx, v);
  if (which === 'firstCol') return m.data.map((row) => row[0]);
  if (which === 'firstRow') return [...(m.data[0] ?? [])];
  return m.data.flat();
}

function sameKind(a: Scalar, b: Scalar): boolean {
  if (a === null || b === null) return false;
  return typeof a === typeof b;
}

function exactMatcher(target: Scalar, wildcards: boolean): (v: Scalar) => boolean {
  if (typeof target === 'string' && wildcards && hasWildcard(target)) {
    const re = wildcardToRegExp(target);
    return (v) => typeof v === 'string' && re.test(v);
  }
  if (typeof target === 'string') {
    const t = target.toLowerCase();
    return (v) => typeof v === 'string' && v.toLowerCase() === t;
  }
  if (target === null) return (v) => v === null || v === '';
  return (v) => sameKind(v, target) && v === target;
}

function findExact(values: Scalar[], target: Scalar, wildcards: boolean, reverse = false): number {
  const m = exactMatcher(target, wildcards);
  if (reverse) {
    for (let i = values.length - 1; i >= 0; i--) if (m(values[i])) return i;
  } else {
    for (let i = 0; i < values.length; i++) if (m(values[i])) return i;
  }
  return -1;
}

/** Binary search: index of the largest value <= target (ascending data). */
function approxAsc(values: Scalar[], target: Scalar): number {
  let lo = 0;
  let hi = values.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let v = values[mid];
    // Skip blanks and mismatched types by scanning to the nearest comparable value
    let probe = mid;
    while (probe <= hi && (v === null || (typeof v !== typeof target && typeof target !== 'object'))) {
      probe++;
      v = values[probe];
    }
    if (probe > hi) {
      hi = mid - 1;
      continue;
    }
    const c = compareValues(v, target);
    if (c <= 0) {
      if (sameKind(v, target)) best = probe;
      lo = probe + 1;
    } else hi = mid - 1;
  }
  return best;
}

/** Binary search: index of the smallest value >= target (descending data). */
function approxDesc(values: Scalar[], target: Scalar): number {
  let lo = 0;
  let hi = values.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const c = compareValues(values[mid], target);
    if (c >= 0) {
      if (sameKind(values[mid], target)) best = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return best;
}

/** Explains common reasons a lookup fails (extra spaces, numbers stored as text). */
function nearMiss(values: Scalar[], target: Scalar): string {
  if (typeof target === 'string') {
    const t = target.trim().toLowerCase();
    const hit = values.find((v) => typeof v === 'string' && v.trim().toLowerCase() === t);
    if (hit !== undefined) return ` There's a close match ${show(hit)} — check for extra spaces.`;
    const n = Number(target);
    if (target.trim() !== '' && !isNaN(n) && values.some((v) => v === n))
      return ` The lookup value is text "${target}" but the list contains the number ${target}. Convert one of them (e.g. VALUE()).`;
  } else if (typeof target === 'number') {
    const s = numberToText(target);
    if (values.some((v) => typeof v === 'string' && v.trim() === s))
      return ` The list contains "${s}" stored as text, but the lookup value is a number.`;
  }
  return '';
}

function notFound(fn: string, target: Scalar, where: string, values: Scalar[]): ErrorValue {
  return Errors.na(`${fn} couldn't find ${show(target)} in ${where}.${nearMiss(values, target)}`);
}

function vhlookup(vertical: boolean) {
  const fname = vertical ? 'VLOOKUP' : 'HLOOKUP';
  return (ctx: Ctx, [lv, table, idx, approx]: Arg[]): Value => {
    const target = scalarOf(ctx, lv);
    if (target instanceof ErrorValue) return target;
    if (!(table instanceof RangeRef) && !(table instanceof Matrix)) return Errors.value(`${fname}'s table_array must be a range.`);
    const index = argInt(ctx, idx);
    if (isError(index)) return index;
    const isApprox = argBool(ctx, approx, true);
    if (isError(isApprox)) return isApprox;
    const span = table instanceof RangeRef ? (vertical ? table.cols : table.rows) : vertical ? table.cols : table.rows;
    if (index < 1) return Errors.value(`${vertical ? 'col' : 'row'}_index_num must be at least 1.`);
    if (index > span)
      return Errors.ref(`${vertical ? 'col' : 'row'}_index_num is ${index}, but the table only has ${span} ${vertical ? 'columns' : 'rows'}.`);
    const keys = vectorOf(ctx, table, vertical ? 'firstCol' : 'firstRow');
    const i = isApprox ? approxAsc(keys, target) : findExact(keys, target, true);
    const where = table instanceof RangeRef ? `the first ${vertical ? 'column' : 'row'} of ${describeRef(table, ctx)}` : 'the table';
    if (i < 0) {
      if (isApprox && keys.length) return Errors.na(`${fname} found no value less than or equal to ${show(target)}. With approximate match the first ${vertical ? 'column' : 'row'} must be sorted ascending — or pass FALSE for an exact match.`);
      return notFound(fname, target, where, keys);
    }
    if (table instanceof RangeRef) {
      return vertical ? ctx.get(table.sheet, table.r1 + i, table.c1 + index - 1) : ctx.get(table.sheet, table.r1 + index - 1, table.c1 + i);
    }
    return vertical ? table.get(i, index - 1) : table.get(index - 1, i);
  };
}

define(
  'VLOOKUP',
  'Lookup',
  'lookup_value, table_array, col_index_num, [range_lookup]',
  'Looks for a value in the first column of a table and returns a value from another column in the same row. Use FALSE as the last argument for an exact match.',
  vhlookup(true),
  { lift: [0] },
);
define(
  'HLOOKUP',
  'Lookup',
  'lookup_value, table_array, row_index_num, [range_lookup]',
  'Looks for a value in the first row of a table and returns a value from another row in the same column.',
  vhlookup(false),
  { lift: [0] },
);

/** Returns a single row/column slice of a range or matrix (keeps references when possible). */
function slice(ctx: Ctx, arr: Arg, i: number, alongRows: boolean): Value {
  if (arr instanceof RangeRef) {
    const { r2, c2 } = clampRef(ctx, arr);
    if (alongRows) {
      if (arr.cols === 1 && arr.rows > 1) return ctx.get(arr.sheet, arr.r1 + i, arr.c1);
      return arr.rows === 1 ? ctx.get(arr.sheet, arr.r1, arr.c1 + i) : new RangeRef(arr.sheet, arr.r1 + i, arr.c1, arr.r1 + i, c2);
    }
    return new RangeRef(arr.sheet, arr.r1, arr.c1 + i, r2, arr.c1 + i);
  }
  const m = toMatrix(ctx, arr);
  if (alongRows) {
    if (m.cols === 1) return m.get(i, 0);
    if (m.rows === 1) return m.get(0, i);
    return new Matrix([m.data[i] ?? []]);
  }
  return new Matrix(m.data.map((row) => [row[i]]));
}

function xFind(values: Scalar[], target: Scalar, matchMode: number, searchMode: number): number | ErrorValue {
  if (searchMode === 2) return matchMode === 0 ? exactBinary(values, target, true) : matchMode === -1 ? approxAsc(values, target) : nextLargerAsc(values, target);
  if (searchMode === -2) return matchMode === 0 ? exactBinary(values, target, false) : approxDesc(values, target);
  const reverse = searchMode === -1;
  const exact = findExact(values, target, matchMode === 2, reverse);
  if (exact >= 0 || matchMode === 0 || matchMode === 2) return exact;
  let best = -1;
  for (let k = 0; k < values.length; k++) {
    const i = reverse ? values.length - 1 - k : k;
    const v = values[i];
    if (!sameKind(v, target)) continue;
    const c = compareValues(v, target);
    if (matchMode === -1 && c < 0 && (best < 0 || compareValues(v, values[best]) > 0)) best = i;
    if (matchMode === 1 && c > 0 && (best < 0 || compareValues(v, values[best]) < 0)) best = i;
  }
  return best;
}

function exactBinary(values: Scalar[], target: Scalar, asc: boolean): number {
  let lo = 0;
  let hi = values.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const c = compareValues(values[mid], target) * (asc ? 1 : -1);
    if (c === 0) return mid;
    if (c < 0) lo = mid + 1;
    else hi = mid - 1;
  }
  return -1;
}

function nextLargerAsc(values: Scalar[], target: Scalar): number {
  let lo = 0;
  let hi = values.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (compareValues(values[mid], target) >= 0) {
      best = mid;
      hi = mid - 1;
    } else lo = mid + 1;
  }
  return best;
}

define(
  'XLOOKUP',
  'Lookup',
  'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]',
  'Searches a range for a match and returns the corresponding item from another range. Exact match by default. match_mode: 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcard.',
  (ctx, [lv, la, ra, nf, mm, sm]) => {
    const target = scalarOf(ctx, lv);
    if (target instanceof ErrorValue) return target;
    const matchMode = argInt(ctx, mm, 0);
    if (isError(matchMode)) return matchMode;
    const searchMode = argInt(ctx, sm, 1);
    if (isError(searchMode)) return searchMode;
    if (![0, -1, 1, 2].includes(matchMode)) return Errors.value('match_mode must be 0, -1, 1 or 2.');
    if (![1, -1, 2, -2].includes(searchMode)) return Errors.value('search_mode must be 1, -1, 2 or -2.');
    const lrows = la instanceof RangeRef ? la.rows : toMatrix(ctx, la).rows;
    const vertical = lrows > 1 || !(la instanceof RangeRef ? la.cols > 1 : toMatrix(ctx, la).cols > 1);
    const values = vectorOf(ctx, la, vertical ? 'firstCol' : 'firstRow');
    const retSize = ra instanceof RangeRef ? (vertical ? clampRef(ctx, ra).r2 - ra.r1 + 1 : clampRef(ctx, ra).c2 - ra.c1 + 1) : vertical ? toMatrix(ctx, ra).rows : toMatrix(ctx, ra).cols;
    if (retSize < values.length && !(ra instanceof RangeRef))
      return Errors.value('lookup_array and return_array must be the same size.');
    const i = xFind(values, target, matchMode, searchMode);
    if (isError(i)) return i;
    if (i < 0) {
      if (nf !== undefined) return nf;
      const where = la instanceof RangeRef ? describeRef(la, ctx) : 'the lookup array';
      return notFound('XLOOKUP', target, where, values);
    }
    return slice(ctx, ra, i, vertical);
  },
  { lift: [0] },
);

define(
  'LOOKUP',
  'Lookup',
  'lookup_value, lookup_vector, [result_vector]',
  'Looks up a value in a sorted one-row or one-column range (approximate match).',
  (ctx, [lv, lvec, rvec]) => {
    const target = scalarOf(ctx, lv);
    if (target instanceof ErrorValue) return target;
    let keys: Scalar[];
    let results: Scalar[];
    if (rvec === undefined) {
      const m = toMatrix(ctx, lvec);
      if (m.cols > m.rows) {
        keys = m.data[0];
        results = m.data[m.rows - 1];
      } else {
        keys = m.data.map((r) => r[0]);
        results = m.data.map((r) => r[m.cols - 1]);
      }
    } else {
      keys = vectorOf(ctx, lvec, 'flat');
      results = vectorOf(ctx, rvec, 'flat');
    }
    const i = approxAsc(keys, target);
    if (i < 0) return Errors.na(`LOOKUP found no value less than or equal to ${show(target)}.`);
    return results[i] ?? Errors.na();
  },
  { lift: [0] },
);

define(
  'MATCH',
  'Lookup',
  'lookup_value, lookup_array, [match_type]',
  'Returns the position of a value in a range. Use 0 as match_type for an exact match.',
  (ctx, [lv, la, mt]) => {
    const target = scalarOf(ctx, lv);
    if (target instanceof ErrorValue) return target;
    const type = argInt(ctx, mt, 1);
    if (isError(type)) return type;
    const values = vectorOf(ctx, la, 'flat');
    const i = type === 0 ? findExact(values, target, true) : type > 0 ? approxAsc(values, target) : approxDesc(values, target);
    if (i < 0) {
      const where = la instanceof RangeRef ? describeRef(la, ctx) : 'the lookup array';
      return type === 0 ? notFound('MATCH', target, where, values) : Errors.na(`MATCH found no suitable value for ${show(target)} — is the list sorted? Use 0 as match_type for an exact match.`);
    }
    return i + 1;
  },
  { lift: [0] },
);

define(
  'XMATCH',
  'Lookup',
  'lookup_value, lookup_array, [match_mode], [search_mode]',
  'Returns the position of an item in a range (exact match by default).',
  (ctx, [lv, la, mm, sm]) => {
    const target = scalarOf(ctx, lv);
    if (target instanceof ErrorValue) return target;
    const matchMode = argInt(ctx, mm, 0);
    if (isError(matchMode)) return matchMode;
    const searchMode = argInt(ctx, sm, 1);
    if (isError(searchMode)) return searchMode;
    const values = vectorOf(ctx, la, 'flat');
    const i = xFind(values, target, matchMode, searchMode);
    if (isError(i)) return i;
    if (i < 0) return notFound('XMATCH', target, la instanceof RangeRef ? describeRef(la, ctx) : 'the lookup array', values);
    return i + 1;
  },
  { lift: [0] },
);

define(
  'INDEX',
  'Lookup',
  'array, row_num, [column_num]',
  'Returns the value at a given row and column of a range. Use 0 for a whole row or column.',
  (ctx, [arr, rn, cn]) => {
    let row = argInt(ctx, rn, 0);
    if (isError(row)) return row;
    let col = argInt(ctx, cn, 0);
    if (isError(col)) return col;
    const rows = arr instanceof RangeRef ? arr.rows : toMatrix(ctx, arr).rows;
    const cols = arr instanceof RangeRef ? arr.cols : toMatrix(ctx, arr).cols;
    if (cn === undefined && rows === 1 && cols > 1) {
      col = row;
      row = 1;
    }
    if (cn === undefined && cols === 1) col = 1;
    if (row < 0 || col < 0) return Errors.value('row_num and column_num can\'t be negative.');
    if (row > rows) return Errors.ref(`row_num ${row} is outside the range, which has ${rows} row${rows === 1 ? '' : 's'}.`);
    if (col > cols) return Errors.ref(`column_num ${col} is outside the range, which has ${cols} column${cols === 1 ? '' : 's'}.`);
    if (arr instanceof RangeRef) {
      const r1 = row === 0 ? arr.r1 : arr.r1 + row - 1;
      const r2 = row === 0 ? arr.r2 : r1;
      const c1 = col === 0 ? arr.c1 : arr.c1 + col - 1;
      const c2 = col === 0 ? arr.c2 : c1;
      return new RangeRef(arr.sheet, r1, c1, r2, c2);
    }
    const m = toMatrix(ctx, arr);
    if (row === 0 && col === 0) return m;
    if (row === 0) return new Matrix(m.data.map((r) => [r[col - 1]]));
    if (col === 0) return new Matrix([m.data[row - 1]]);
    return m.get(row - 1, col - 1);
  },
);

define(
  'OFFSET',
  'Lookup',
  'reference, rows, cols, [height], [width]',
  'Returns a reference shifted by a number of rows and columns from a starting cell.',
  (ctx, [ref, rs, cs, h, w]) => {
    if (!(ref instanceof RangeRef)) return Errors.value('OFFSET needs a cell reference as its first argument.');
    const dr = argInt(ctx, rs);
    if (isError(dr)) return dr;
    const dc = argInt(ctx, cs);
    if (isError(dc)) return dc;
    const height = argInt(ctx, h, ref.rows);
    if (isError(height)) return height;
    const width = argInt(ctx, w, ref.cols);
    if (isError(width)) return width;
    if (height < 1 || width < 1) return Errors.ref('height and width must be at least 1.');
    const r1 = ref.r1 + dr;
    const c1 = ref.c1 + dc;
    const r2 = r1 + height - 1;
    const c2 = c1 + width - 1;
    if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) return Errors.ref('OFFSET moved the reference off the edge of the sheet.');
    return new RangeRef(ref.sheet, r1, c1, r2, c2);
  },
  { volatile: true },
);

define(
  'INDIRECT',
  'Lookup',
  'ref_text, [a1]',
  'Returns the reference described by a text string, e.g. INDIRECT("B" & A1).',
  (ctx, [t, a1]) => {
    const text = argText(ctx, t);
    if (isError(text)) return text;
    const isA1 = argBool(ctx, a1, true);
    if (isError(isA1)) return isA1;
    const s = text.trim();
    if (!isA1) {
      const m = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?R(\d+)C(\d+)(?::R(\d+)C(\d+))?$/i.exec(s);
      if (!m) return Errors.ref(`"${text}" isn't a valid R1C1 reference.`);
      const sheetName = m[1]?.replace(/''/g, "'") ?? m[2];
      const sheet = sheetName ? ctx.wb.sheetByName(sheetName) : ctx.sheet;
      if (!sheet) return Errors.ref(`There's no sheet named "${sheetName}".`);
      const r1 = +m[3] - 1;
      const c1 = +m[4] - 1;
      const r2 = m[5] ? +m[5] - 1 : r1;
      const c2 = m[6] ? +m[6] - 1 : c1;
      return new RangeRef(sheet, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
    }
    const toks = tokenize(s).filter((x) => x.type !== 'ws');
    if (toks.length === 1 && toks[0].type === 'ref') {
      const ref = toks[0].ref!;
      const sheet = ref.sheet !== undefined ? ctx.wb.sheetByName(ref.sheet) : ctx.sheet;
      if (!sheet) return Errors.ref(`There's no sheet named "${ref.sheet}".`);
      return new RangeRef(sheet, ref.r1, ref.c1, ref.r2, ref.c2);
    }
    if (toks.length === 1 && toks[0].type === 'name') {
      const dn = ctx.wb.getName(toks[0].text);
      if (dn) {
        const inner = tokenize(dn.ref).filter((x) => x.type !== 'ws');
        if (inner.length === 1 && inner[0].type === 'ref') {
          const ref = inner[0].ref!;
          const sheet = ref.sheet !== undefined ? ctx.wb.sheetByName(ref.sheet) : ctx.sheet;
          if (sheet) return new RangeRef(sheet, ref.r1, ref.c1, ref.r2, ref.c2);
        }
      }
    }
    return Errors.ref(`"${text}" isn't a valid cell reference.`);
  },
  { volatile: true },
);

define('ROW', 'Lookup', '[reference]', 'Returns the row number of a reference (or of the current cell).', (ctx, [ref]) => {
  if (ref === undefined) return ctx.row + 1 + ctx.dr;
  if (!(ref instanceof RangeRef)) return Errors.value('ROW needs a cell reference.');
  if (ref.rows === 1) return ref.r1 + 1;
  const { r2 } = clampRef(ctx, ref);
  return Matrix.fill(r2 - ref.r1 + 1, 1, (r) => ref.r1 + r + 1);
});
define('COLUMN', 'Lookup', '[reference]', 'Returns the column number of a reference (or of the current cell).', (ctx, [ref]) => {
  if (ref === undefined) return ctx.col + 1 + ctx.dc;
  if (!(ref instanceof RangeRef)) return Errors.value('COLUMN needs a cell reference.');
  if (ref.cols === 1) return ref.c1 + 1;
  const { c2 } = clampRef(ctx, ref);
  return Matrix.fill(1, c2 - ref.c1 + 1, (_r, c) => ref.c1 + c + 1);
});
define('ROWS', 'Lookup', 'array', 'Returns the number of rows in a range or array.', (ctx, [a]) => {
  if (a instanceof RangeRef) return a.rows;
  if (a instanceof Matrix) return a.rows;
  return 1;
});
define('COLUMNS', 'Lookup', 'array', 'Returns the number of columns in a range or array.', (ctx, [a]) => {
  if (a instanceof RangeRef) return a.cols;
  if (a instanceof Matrix) return a.cols;
  return 1;
});

define(
  'ADDRESS',
  'Lookup',
  'row_num, column_num, [abs_num], [a1], [sheet_text]',
  'Creates a cell address as text, e.g. ADDRESS(2, 3) returns "$C$2".',
  (ctx, [r, c, abs, a1, sh]) => {
    const row = argInt(ctx, r);
    if (isError(row)) return row;
    const col = argInt(ctx, c);
    if (isError(col)) return col;
    const absNum = argInt(ctx, abs, 1);
    if (isError(absNum)) return absNum;
    const isA1 = argBool(ctx, a1, true);
    if (isError(isA1)) return isA1;
    if (row < 1 || col < 1 || row > MAX_ROWS || col > MAX_COLS) return Errors.value('Row and column must be within the sheet.');
    let out: string;
    if (isA1) {
      const absRow = absNum === 1 || absNum === 2;
      const absCol = absNum === 1 || absNum === 3;
      out = `${absCol ? '$' : ''}${colToName(col - 1)}${absRow ? '$' : ''}${row}`;
    } else {
      const rr = absNum === 1 || absNum === 2 ? `R${row}` : `R[${row}]`;
      const cc = absNum === 1 || absNum === 3 ? `C${col}` : `C[${col}]`;
      out = rr + cc;
    }
    if (sh !== undefined) {
      const s = argText(ctx, sh);
      if (isError(s)) return s;
      if (s) out = quoteSheetName(s) + '!' + out;
    }
    return out;
  },
  { lift: [0, 1] },
);

define('TRANSPOSE', 'Lookup', 'array', 'Flips rows and columns of a range.', (ctx, [a]) => {
  const m = toMatrix(ctx, a);
  return Matrix.fill(m.cols, m.rows, (r, c) => m.get(c, r) ?? 0);
});

define(
  'FILTER',
  'Lookup',
  'array, include, [if_empty]',
  'Returns only the rows (or columns) of a range that meet a condition, e.g. FILTER(A2:C100, B2:B100>50).',
  (ctx, [a, inc, ifEmpty]) => {
    const m = toMatrix(ctx, a);
    const im = toMatrix(ctx, inc);
    const byRows = im.cols === 1 && im.rows === m.rows;
    const byCols = im.rows === 1 && im.cols === m.cols;
    if (!byRows && !byCols) return Errors.value(`The include argument must have the same number of rows (${m.rows}) or columns (${m.cols}) as the array.`);
    const keep: number[] = [];
    const n = byRows ? m.rows : m.cols;
    for (let i = 0; i < n; i++) {
      const v = byRows ? im.get(i, 0) : im.get(0, i);
      if (v instanceof ErrorValue) return v;
      const b = toBool(v);
      if (isError(b)) return b;
      if (b) keep.push(i);
    }
    if (!keep.length) return ifEmpty !== undefined ? ifEmpty : Errors.calc('FILTER found no matching rows. Add a third argument to show something else, like FILTER(…, …, "None").');
    if (byRows) return new Matrix(keep.map((i) => m.data[i]));
    return new Matrix(m.data.map((row) => keep.map((i) => row[i])));
  },
);

function sortRows(rows: Scalar[][], keys: { idx: number; desc: boolean }[]): Scalar[][] {
  return rows
    .map((r, i) => ({ r, i }))
    .sort((x, y) => {
      for (const k of keys) {
        const a = x.r[k.idx];
        const b = y.r[k.idx];
        // Blanks always sort last
        const ab = a === null || a === '';
        const bb = b === null || b === '';
        if (ab !== bb) return ab ? 1 : -1;
        const c = compareValues(a, b);
        if (c) return k.desc ? -c : c;
      }
      return x.i - y.i;
    })
    .map((x) => x.r);
}

function transpose(d: Scalar[][]): Scalar[][] {
  const rows = d.length;
  const cols = d[0]?.length ?? 0;
  const out: Scalar[][] = [];
  for (let c = 0; c < cols; c++) {
    const row: Scalar[] = [];
    for (let r = 0; r < rows; r++) row.push(d[r][c]);
    out.push(row);
  }
  return out;
}

define(
  'SORT',
  'Lookup',
  'array, [sort_index], [sort_order], [by_col]',
  'Sorts the contents of a range. sort_order: 1 ascending, -1 descending.',
  (ctx, [a, si, so, bc]) => {
    const m = toMatrix(ctx, a);
    const idx = argInt(ctx, si, 1);
    if (isError(idx)) return idx;
    const order = argInt(ctx, so, 1);
    if (isError(order)) return order;
    const byCol = argBool(ctx, bc, false);
    if (isError(byCol)) return byCol;
    const data = byCol ? transpose(m.data) : m.data;
    const width = data[0]?.length ?? 0;
    if (idx < 1 || idx > width) return Errors.value(`sort_index must be between 1 and ${width}.`);
    if (order !== 1 && order !== -1) return Errors.value('sort_order must be 1 (ascending) or -1 (descending).');
    const sorted = sortRows(data, [{ idx: idx - 1, desc: order === -1 }]);
    return new Matrix(byCol ? transpose(sorted) : sorted);
  },
);

define(
  'SORTBY',
  'Lookup',
  'array, by_array1, [sort_order1], [by_array2], [sort_order2], ...',
  'Sorts a range based on the values in other ranges.',
  (ctx, [a, ...rest]) => {
    const m = toMatrix(ctx, a);
    const keys: { vals: Scalar[]; desc: boolean }[] = [];
    for (let i = 0; i < rest.length; i += 2) {
      const bm = toMatrix(ctx, rest[i]);
      const order = argInt(ctx, rest[i + 1], 1);
      if (isError(order)) return order;
      const vals = bm.cols === 1 ? bm.data.map((r) => r[0]) : bm.data[0];
      if (vals.length !== m.rows) return Errors.value('Each by_array must have the same number of rows as the array.');
      keys.push({ vals, desc: order === -1 });
    }
    const idx = m.data.map((_, i) => i);
    idx.sort((x, y) => {
      for (const k of keys) {
        const c = compareValues(k.vals[x], k.vals[y]);
        if (c) return k.desc ? -c : c;
      }
      return x - y;
    });
    return new Matrix(idx.map((i) => m.data[i]));
  },
  { min: 2 },
);

define(
  'UNIQUE',
  'Lookup',
  'array, [by_col], [exactly_once]',
  'Returns the unique rows (or values) of a range, removing duplicates.',
  (ctx, [a, bc, eo]) => {
    const m = toMatrix(ctx, a);
    const byCol = argBool(ctx, bc, false);
    if (isError(byCol)) return byCol;
    const once = argBool(ctx, eo, false);
    if (isError(once)) return once;
    const data = byCol ? transpose(m.data) : m.data;
    const keyOf = (row: Scalar[]) =>
      row.map((v) => (typeof v === 'string' ? 's' + v.toLowerCase() : v instanceof ErrorValue ? 'e' + v.code : typeof v + String(v))).join('\u0001');
    const counts = new Map<string, { row: Scalar[]; n: number }>();
    for (const row of data) {
      const k = keyOf(row);
      const e = counts.get(k);
      if (e) e.n++;
      else counts.set(k, { row, n: 1 });
    }
    const out = [...counts.values()].filter((e) => !once || e.n === 1).map((e) => e.row.map((v) => (v === null ? 0 : v)));
    if (!out.length) return Errors.calc('UNIQUE found no values that appear exactly once.');
    return new Matrix(byCol ? transpose(out) : out);
  },
);

define('TAKE', 'Lookup', 'array, rows, [columns]', 'Returns rows or columns from the start (positive) or end (negative) of an array.', (ctx, [a, r, c]) => {
  const m = toMatrix(ctx, a);
  const rows = argInt(ctx, r, m.rows);
  if (isError(rows)) return rows;
  const cols = argInt(ctx, c, m.cols);
  if (isError(cols)) return cols;
  if (rows === 0 || cols === 0) return Errors.calc('TAKE returned an empty array.');
  const rs = rows > 0 ? m.data.slice(0, rows) : m.data.slice(rows);
  return new Matrix(rs.map((row) => (cols > 0 ? row.slice(0, cols) : row.slice(cols))));
});
define('DROP', 'Lookup', 'array, rows, [columns]', 'Removes rows or columns from the start (positive) or end (negative) of an array.', (ctx, [a, r, c]) => {
  const m = toMatrix(ctx, a);
  const rows = argInt(ctx, r, 0);
  if (isError(rows)) return rows;
  const cols = argInt(ctx, c, 0);
  if (isError(cols)) return cols;
  const rs = rows >= 0 ? m.data.slice(rows) : m.data.slice(0, rows);
  const out = rs.map((row) => (cols >= 0 ? row.slice(cols) : row.slice(0, cols)));
  if (!out.length || !out[0].length) return Errors.calc('DROP removed everything.');
  return new Matrix(out);
});

function stack(vertical: boolean) {
  return (ctx: Ctx, args: Arg[]): Value => {
    const mats = args.filter((a) => a !== undefined).map((a) => toMatrix(ctx, a));
    if (vertical) {
      const width = Math.max(...mats.map((m) => m.cols));
      const rows: Scalar[][] = [];
      for (const m of mats) for (const row of m.data) rows.push(Array.from({ length: width }, (_, i) => (i < row.length ? row[i] : Errors.na())));
      return new Matrix(rows);
    }
    const height = Math.max(...mats.map((m) => m.rows));
    return Matrix.fill(height, mats.reduce((s, m) => s + m.cols, 0), (r, c) => {
      let cc = c;
      for (const m of mats) {
        if (cc < m.cols) return r < m.rows ? m.get(r, cc) : Errors.na();
        cc -= m.cols;
      }
      return Errors.na();
    });
  };
}
define('VSTACK', 'Lookup', 'array1, [array2], ...', 'Stacks arrays on top of each other.', stack(true));
define('HSTACK', 'Lookup', 'array1, [array2], ...', 'Places arrays side by side.', stack(false));

function toVector(column: boolean) {
  return (ctx: Ctx, [a, ig, sc]: Arg[]): Value => {
    const m = toMatrix(ctx, a);
    const ignore = argInt(ctx, ig, 0);
    if (isError(ignore)) return ignore;
    const byCol = argBool(ctx, sc, false);
    if (isError(byCol)) return byCol;
    const data = byCol ? transpose(m.data) : m.data;
    const vals = data.flat().filter((v) => {
      if ((ignore === 1 || ignore === 3) && v === null) return false;
      if ((ignore === 2 || ignore === 3) && v instanceof ErrorValue) return false;
      return true;
    });
    if (!vals.length) return Errors.calc();
    return column ? Matrix.column(vals) : Matrix.row(vals);
  };
}
define('TOCOL', 'Lookup', 'array, [ignore], [scan_by_column]', 'Turns an array into a single column.', toVector(true));
define('TOROW', 'Lookup', 'array, [ignore], [scan_by_column]', 'Turns an array into a single row.', toVector(false));

define('CHOOSECOLS', 'Lookup', 'array, col_num1, [col_num2], ...', 'Returns specific columns from an array.', (ctx, [a, ...nums]) => {
  const m = toMatrix(ctx, a);
  const idx: number[] = [];
  for (const n of nums) {
    const i = argInt(ctx, n);
    if (isError(i)) return i;
    const k = i < 0 ? m.cols + i : i - 1;
    if (k < 0 || k >= m.cols) return Errors.value(`Column ${i} is outside the array.`);
    idx.push(k);
  }
  return new Matrix(m.data.map((row) => idx.map((k) => row[k])));
});
define('CHOOSEROWS', 'Lookup', 'array, row_num1, [row_num2], ...', 'Returns specific rows from an array.', (ctx, [a, ...nums]) => {
  const m = toMatrix(ctx, a);
  const out: Scalar[][] = [];
  for (const n of nums) {
    const i = argInt(ctx, n);
    if (isError(i)) return i;
    const k = i < 0 ? m.rows + i : i - 1;
    if (k < 0 || k >= m.rows) return Errors.value(`Row ${i} is outside the array.`);
    out.push(m.data[k]);
  }
  return new Matrix(out);
});

