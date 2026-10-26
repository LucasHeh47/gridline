// Sorting, filtering and finding.
import { cellKey, type Range } from '../address';
import type { Doc } from '../doc';
import { compareValues } from '../formula/coerce';
import { translateFormula } from '../formula/refs';
import type { Transaction } from '../history';
import { formatValue } from '../numfmt';
import { ErrorValue, type Scalar } from '../values';
import type { Cell, FilterCond, FilterState, Sheet } from '../workbook';
import { normalizeFormula } from './cells';

export interface SortKey {
  col: number; // absolute column index
  desc: boolean;
}

function isBlank(v: Scalar): boolean {
  return v === null || v === '';
}

/** Excel order: numbers < text < booleans < errors; blanks always last. */
export function compareForSort(a: Scalar, b: Scalar): number {
  const ab = isBlank(a);
  const bb = isBlank(b);
  if (ab || bb) return ab === bb ? 0 : ab ? 1 : -1;
  const ae = a instanceof ErrorValue;
  const be = b instanceof ErrorValue;
  if (ae || be) return ae === be ? 0 : ae ? 1 : -1;
  return compareValues(a, b);
}

/** Guesses whether the first row of a range is a header (text above numbers / different types). */
export function detectHeader(doc: Doc, sheet: Sheet, range: Range): boolean {
  if (range.r2 <= range.r1) return false;
  let textHeaders = 0;
  let differs = 0;
  for (let c = range.c1; c <= range.c2; c++) {
    const h = doc.engine.getValue(sheet, range.r1, c);
    const v = doc.engine.getValue(sheet, range.r1 + 1, c);
    if (typeof h === 'string' && h !== '') textHeaders++;
    if (typeof h === 'string' && typeof v !== 'string' && v !== null) differs++;
    const hs = doc.wb.styles.get(sheet.styleIdAt(range.r1, c));
    const vs = doc.wb.styles.get(sheet.styleIdAt(range.r1 + 1, c));
    if (hs.bold && !vs.bold) differs++;
  }
  return textHeaders === range.c2 - range.c1 + 1 && differs > 0;
}

export function sortRange(tx: Transaction, doc: Doc, sheet: Sheet, range: Range, keys: SortKey[], hasHeader: boolean): void {
  const start = hasHeader ? range.r1 + 1 : range.r1;
  if (start >= range.r2 + (hasHeader ? 1 : 0) && start > range.r2) return;
  const rows: { r: number; cells: (Cell | undefined)[]; keys: Scalar[] }[] = [];
  for (let r = start; r <= range.r2; r++) {
    const cells: (Cell | undefined)[] = [];
    for (let c = range.c1; c <= range.c2; c++) cells.push(sheet.cells.get(cellKey(r, c)));
    rows.push({ r, cells, keys: keys.map((k) => doc.engine.getValue(sheet, r, k.col)) });
  }
  rows.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const x = a.keys[i];
      const y = b.keys[i];
      // Blanks last regardless of direction
      if (isBlank(x) !== isBlank(y)) return isBlank(x) ? 1 : -1;
      const c = compareForSort(x, y);
      if (c) return keys[i].desc ? -c : c;
    }
    return a.r - b.r;
  });
  rows.forEach((row, i) => {
    const target = start + i;
    if (row.r === target) return;
    row.cells.forEach((cell, j) => {
      const c = range.c1 + j;
      if (!cell) {
        tx.setCell(sheet, target, c, undefined);
        return;
      }
      const next = cell.f !== undefined ? { ...cell, f: translateFormula(cell.f, target - row.r, 0) } : cell;
      tx.setCell(sheet, target, c, next);
    });
  });
}

/** The text a filter list shows for a cell. */
export function displayText(doc: Doc, sheet: Sheet, r: number, c: number): string {
  const v = doc.engine.getValue(sheet, r, c);
  if (v === null) return '';
  return formatValue(v, typeof v === 'number' ? doc.numFmtAt(sheet, r, c) : doc.wb.styles.get(sheet.styleIdAt(r, c)).numFmt, 255).text;
}

export function columnValues(doc: Doc, sheet: Sheet, range: Range, col: number): { text: string; count: number }[] {
  const counts = new Map<string, number>();
  for (let r = range.r1 + 1; r <= range.r2; r++) {
    const t = displayText(doc, sheet, r, col);
    counts.set(t, (counts.get(t) ?? 0) + 1);
  }
  const entries = [...counts].map(([text, count]) => ({ text, count }));
  // Sort numerically when possible
  entries.sort((a, b) => {
    if (a.text === '') return 1;
    if (b.text === '') return -1;
    const na = Number(a.text.replace(/[$,%]/g, ''));
    const nb = Number(b.text.replace(/[$,%]/g, ''));
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.text.localeCompare(b.text, undefined, { numeric: true, sensitivity: 'base' });
  });
  return entries;
}

function matchesCond(doc: Doc, sheet: Sheet, r: number, col: number, cond: FilterCond, stats: { avg: number; topN?: number }): boolean {
  if (cond.values) {
    const t = displayText(doc, sheet, r, col);
    if (!cond.values.includes(t)) return false;
  }
  if (cond.op) {
    const v = doc.engine.getValue(sheet, r, col);
    const text = displayText(doc, sheet, r, col).toLowerCase();
    const q = cond.op.value.toLowerCase();
    const num = Number(cond.op.value);
    const num2 = Number(cond.op.value2);
    switch (cond.op.type) {
      case 'contains':
        return text.includes(q);
      case 'notContains':
        return !text.includes(q);
      case 'beginsWith':
        return text.startsWith(q);
      case 'endsWith':
        return text.endsWith(q);
      case 'equals':
        return typeof v === 'number' && !isNaN(num) ? v === num : text === q;
      case 'notEquals':
        return typeof v === 'number' && !isNaN(num) ? v !== num : text !== q;
      case 'gt':
        return typeof v === 'number' && v > num;
      case 'gte':
        return typeof v === 'number' && v >= num;
      case 'lt':
        return typeof v === 'number' && v < num;
      case 'lte':
        return typeof v === 'number' && v <= num;
      case 'between':
        return typeof v === 'number' && v >= Math.min(num, num2) && v <= Math.max(num, num2);
      case 'aboveAvg':
        return typeof v === 'number' && v > stats.avg;
      case 'belowAvg':
        return typeof v === 'number' && v < stats.avg;
      case 'top10':
        return typeof v === 'number' && stats.topN !== undefined && v >= stats.topN;
    }
  }
  return true;
}

/** Sets the sheet's filter and recomputes which rows filters hide. */
export function applyFilter(tx: Transaction, doc: Doc, sheet: Sheet, filter: FilterState | null): void {
  tx.setProp(sheet, 'filter', filter);
  refreshFilters(tx, doc, sheet);
}

/** Recomputes which rows are hidden by the sheet's filter and its tables' filters. */
export function refreshFilters(tx: Transaction, doc: Doc, sheet: Sheet): void {
  const hidden = new Set<number>();
  if (sheet.filter) hideFiltered(doc, sheet, sheet.filter, hidden);
  for (const t of sheet.tables) if (Object.keys(t.filter).length) hideFiltered(doc, sheet, { range: t.range, conds: t.filter }, hidden);
  tx.setProp(sheet, 'filterHidden', hidden, true);
}

function hideFiltered(doc: Doc, sheet: Sheet, filter: FilterState, hidden: Set<number>): void {
  {
    const statsByCol = new Map<number, { avg: number; topN?: number }>();
    for (const [k, cond] of Object.entries(filter.conds)) {
      const col = +k;
      const nums: number[] = [];
      for (let r = filter.range.r1 + 1; r <= filter.range.r2; r++) {
        const v = doc.engine.getValue(sheet, r, col);
        if (typeof v === 'number') nums.push(v);
      }
      const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
      let topN: number | undefined;
      if (cond.op?.type === 'top10') {
        const n = Math.max(1, parseInt(cond.op.value, 10) || 10);
        const sorted = [...nums].sort((a, b) => b - a);
        topN = sorted[Math.min(n, sorted.length) - 1];
      }
      statsByCol.set(col, { avg, topN });
    }
    for (let r = filter.range.r1 + 1; r <= filter.range.r2; r++) {
      for (const [k, cond] of Object.entries(filter.conds)) {
        if (!matchesCond(doc, sheet, r, +k, cond, statsByCol.get(+k)!)) {
          hidden.add(r);
          break;
        }
      }
    }
  }
}

/** Expands a single cell to the surrounding block of data (like Excel's Ctrl+A / current region). */
export function currentRegion(doc: Doc, sheet: Sheet, row: number, col: number): Range {
  const filled = (r: number, c: number) => r >= 0 && c >= 0 && doc.engine.getValue(sheet, r, c) !== null;
  let r1 = row;
  let r2 = row;
  let c1 = col;
  let c2 = col;
  let grew = true;
  let guard = 0;
  while (grew && guard++ < 10000) {
    grew = false;
    // Expand while any cell on the border (including diagonals) has data
    const rowHas = (r: number) => {
      for (let c = Math.max(0, c1 - 1); c <= c2 + 1; c++) if (filled(r, c)) return true;
      return false;
    };
    const colHas = (c: number) => {
      for (let r = Math.max(0, r1 - 1); r <= r2 + 1; r++) if (filled(r, c)) return true;
      return false;
    };
    if (r1 > 0 && rowHas(r1 - 1)) {
      r1--;
      grew = true;
    }
    if (rowHas(r2 + 1)) {
      r2++;
      grew = true;
    }
    if (c1 > 0 && colHas(c1 - 1)) {
      c1--;
      grew = true;
    }
    if (colHas(c2 + 1)) {
      c2++;
      grew = true;
    }
  }
  return { r1, c1, r2, c2 };
}

export interface FindOptions {
  matchCase: boolean;
  wholeCell: boolean;
  inFormulas: boolean;
  allSheets: boolean;
}

export interface FindHit {
  sheetId: number;
  row: number;
  col: number;
}

export function findAll(doc: Doc, activeSheet: Sheet, query: string, opts: FindOptions): FindHit[] {
  if (!query) return [];
  const q = opts.matchCase ? query : query.toLowerCase();
  const hits: FindHit[] = [];
  const sheets = opts.allSheets ? doc.wb.sheets : [activeSheet];
  for (const sheet of sheets) {
    const keys = [...sheet.cells.keys()].sort((a, b) => a - b);
    for (const key of keys) {
      const cell = sheet.cells.get(key)!;
      const r = Math.floor(key / 16384);
      const c = key % 16384;
      let text: string;
      if (opts.inFormulas && cell.f !== undefined) text = '=' + cell.f;
      else {
        if (cell.f === undefined && (cell.v === undefined || cell.v === null)) continue;
        text = displayText(doc, sheet, r, c);
      }
      const t = opts.matchCase ? text : text.toLowerCase();
      if (opts.wholeCell ? t === q : t.includes(q)) hits.push({ sheetId: sheet.id, row: r, col: c });
    }
  }
  return hits;
}

/** Replaces text in matching cells. Returns the number of cells changed. */
export function replaceIn(tx: Transaction, doc: Doc, hits: FindHit[], query: string, replacement: string, opts: FindOptions): number {
  let n = 0;
  const re = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), opts.matchCase ? 'g' : 'gi');
  for (const h of hits) {
    const sheet = doc.wb.sheetById(h.sheetId);
    if (!sheet) continue;
    const cell = sheet.getCell(h.row, h.col);
    if (!cell) continue;
    if (cell.f !== undefined) {
      if (!opts.inFormulas) continue;
      const f = opts.wholeCell ? replacement.replace(/^=/, '') : cell.f.replace(re, replacement);
      if (f !== cell.f) {
        tx.setCell(sheet, h.row, h.col, { ...cell, f: normalizeFormula(f) });
        n++;
      }
    } else if (typeof cell.v === 'string') {
      const v = opts.wholeCell ? replacement : cell.v.replace(re, replacement);
      if (v !== cell.v) {
        tx.setCell(sheet, h.row, h.col, { ...cell, v });
        n++;
      }
    } else if (cell.v !== undefined && cell.v !== null) {
      const text = displayText(doc, sheet, h.row, h.col);
      const v = opts.wholeCell ? replacement : text.replace(re, replacement);
      if (v !== text) {
        const num = Number(v);
        tx.setCell(sheet, h.row, h.col, { ...cell, v: v.trim() !== '' && !isNaN(num) ? num : v });
        n++;
      }
    }
  }
  return n;
}

/** Removes duplicate rows within a range (by the given columns). Returns how many were removed. */
export function removeDuplicates(tx: Transaction, doc: Doc, sheet: Sheet, range: Range, cols: number[], hasHeader: boolean): number {
  const start = hasHeader ? range.r1 + 1 : range.r1;
  const seen = new Set<string>();
  const keep: number[] = [];
  for (let r = start; r <= range.r2; r++) {
    const key = cols
      .map((c) => {
        const v = doc.engine.getValue(sheet, r, c);
        return typeof v === 'string' ? 's' + v.toLowerCase() : String(v);
      })
      .join('\u0001');
    if (seen.has(key)) continue;
    seen.add(key);
    keep.push(r);
  }
  const removed = range.r2 - start + 1 - keep.length;
  if (!removed) return 0;
  const snapshot = new Map<number, (Cell | undefined)[]>();
  for (let r = start; r <= range.r2; r++) {
    const row: (Cell | undefined)[] = [];
    for (let c = range.c1; c <= range.c2; c++) row.push(sheet.getCell(r, c));
    snapshot.set(r, row);
  }
  let target = start;
  for (const r of keep) {
    const row = snapshot.get(r)!;
    row.forEach((cell, j) => {
      const next = cell?.f !== undefined ? { ...cell, f: translateFormula(cell.f, target - r, 0) } : cell;
      tx.setCell(sheet, target, range.c1 + j, next);
    });
    target++;
  }
  for (let r = target; r <= range.r2; r++)
    for (let c = range.c1; c <= range.c2; c++) {
      const cell = sheet.getCell(r, c);
      if (cell) tx.setCell(sheet, r, c, cell.s !== undefined ? { s: cell.s } : undefined);
    }
  return removed;
}
