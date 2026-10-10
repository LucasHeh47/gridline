// Cell editing operations: input, clearing, styling, borders, merging.
import { cellKey, isFullColumn, isFullRow, keyCol, keyRow, rangesIntersect, type Range } from '../address';
import type { Doc } from '../doc';
import { normalizeFuncName } from '../formula/parser';
import { translateFormula } from '../formula/refs';
import { refToString, tokenize } from '../formula/tokenizer';
import type { Transaction } from '../history';
import { parseInput } from '../input';
import { adjustDecimals } from '../numfmt';
import type { BorderSide, Style, StylePatch } from '../styles';
import type { Cell, Sheet, Workbook } from '../workbook';

/** Iterates existing cells inside a range without visiting every empty address. */
export function forEachCellIn(sheet: Sheet, range: Range, fn: (row: number, col: number, cell: Cell) => void): void {
  const area = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  if (area <= sheet.cells.size * 2) {
    for (let r = range.r1; r <= range.r2; r++)
      for (let c = range.c1; c <= range.c2; c++) {
        const cell = sheet.cells.get(cellKey(r, c));
        if (cell) fn(r, c, cell);
      }
    return;
  }
  const hits: [number, Cell][] = [];
  for (const [key, cell] of sheet.cells) {
    const r = keyRow(key);
    const c = keyCol(key);
    if (r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2) hits.push([key, cell]);
  }
  for (const [key, cell] of hits) fn(keyRow(key), keyCol(key), cell);
}

/** Upper-cases function names, references and booleans the way Excel tidies formulas on entry. */
export function normalizeFormula(text: string): string {
  return tokenize(text)
    .map((t) => {
      if (t.type === 'func') {
        const n = normalizeFuncName(t.text);
        return t.text.toUpperCase().startsWith('_XL') ? t.text : n;
      }
      if (t.type === 'ref' && t.ref) return refToString(t.ref);
      if (t.type === 'bool') return t.text.toUpperCase();
      return t.text;
    })
    .join('');
}

const DATE_FUNCS: Record<string, string> = {
  TODAY: 'm/d/yyyy',
  NOW: 'm/d/yyyy h:mm',
  DATE: 'm/d/yyyy',
  EDATE: 'm/d/yyyy',
  EOMONTH: 'm/d/yyyy',
  WORKDAY: 'm/d/yyyy',
  'WORKDAY.INTL': 'm/d/yyyy',
  DATEVALUE: 'm/d/yyyy',
  TIME: 'h:mm AM/PM',
  TIMEVALUE: 'h:mm AM/PM',
};

/** Picks a sensible number format for a new formula (dates for TODAY(), or the format of the first referenced cell). */
function inferFormulaFormat(wb: Workbook, sheet: Sheet, formula: string): string | undefined {
  const toks = tokenize(formula).filter((t) => t.type !== 'ws');
  const first = toks[0];
  if (first?.type === 'func') {
    const f = DATE_FUNCS[normalizeFuncName(first.text)];
    if (f) return f;
    const name = normalizeFuncName(first.text);
    if (!['SUM', 'AVERAGE', 'MIN', 'MAX', 'MEDIAN', 'ROUND', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS', 'SUBTOTAL', 'ABS'].includes(name))
      return undefined;
  }
  // Arithmetic-only formulas inherit the format of the first referenced cell
  const allowed = new Set(['ref', 'op', 'num', 'lparen', 'rparen', 'func', 'comma']);
  if (!toks.every((t) => allowed.has(t.type))) return undefined;
  if (toks.some((t) => t.type === 'op' && (t.text === '/' || t.text === '^' || t.text === '*'))) {
    // Multiplying a currency by a number is still currency; dividing usually isn't.
    if (toks.some((t) => t.type === 'op' && (t.text === '/' || t.text === '^'))) return undefined;
  }
  const ref = toks.find((t) => t.type === 'ref')?.ref;
  if (!ref) return undefined;
  const target = ref.sheet ? wb.sheetByName(ref.sheet) : sheet;
  if (!target) return undefined;
  const fmt = wb.styles.get(target.styleIdAt(ref.r1, ref.c1)).numFmt;
  if (!fmt || fmt === 'General' || fmt === '@') return undefined;
  return fmt;
}

export function setCellInput(tx: Transaction, doc: Doc, sheet: Sheet, row: number, col: number, raw: string): void {
  const wb = doc.wb;
  const cell = sheet.getCell(row, col);
  const styleId = sheet.styleIdAt(row, col);
  const style = wb.styles.get(styleId);
  const parsed = parseInput(raw, style.numFmt);
  const next: Cell = {};
  if (cell?.s !== undefined) next.s = cell.s;
  else if (styleId !== undefined) next.s = styleId;
  if (cell?.note) next.note = cell.note;
  if (cell?.link) next.link = cell.link;
  if (parsed.formula !== undefined) {
    next.f = normalizeFormula(parsed.formula);
    if (!style.numFmt || style.numFmt === 'General') {
      const fmt = inferFormulaFormat(wb, sheet, next.f);
      if (fmt) next.s = wb.styles.apply(next.s, { numFmt: fmt });
    }
  } else {
    if (parsed.value !== null && parsed.value !== undefined) next.v = parsed.value;
    if (parsed.numFmt) next.s = wb.styles.apply(next.s, { numFmt: parsed.numFmt });
    if (typeof parsed.value === 'string' && /^https?:\/\/\S+$/i.test(parsed.value.trim())) next.link = parsed.value.trim();
    else if (cell?.link && cell.v !== parsed.value) delete next.link;
  }
  if (typeof parsed.value === 'string' && parsed.value.includes('\n')) {
    next.s = wb.styles.apply(next.s, { wrap: true });
  }
  tx.setCell(sheet, row, col, next);
}

/** Ctrl+Enter: enters the same input in every selected cell (formulas adjust relative to the active cell). */
export function fillSelectionWithInput(
  tx: Transaction,
  doc: Doc,
  sheet: Sheet,
  ranges: Range[],
  activeRow: number,
  activeCol: number,
  raw: string,
): void {
  const isFormula = raw.startsWith('=') && raw.length > 1;
  for (const r of ranges) {
    for (let row = r.r1; row <= r.r2; row++)
      for (let col = r.c1; col <= r.c2; col++) {
        if (isFormula) {
          const f = translateFormula(normalizeFormula(raw.slice(1)), row - activeRow, col - activeCol);
          setCellInput(tx, doc, sheet, row, col, '=' + f);
        } else setCellInput(tx, doc, sheet, row, col, raw);
      }
  }
}

export type ClearWhat = 'contents' | 'formats' | 'all' | 'notes' | 'links';

export function clearRanges(tx: Transaction, sheet: Sheet, ranges: Range[], what: ClearWhat): void {
  for (const range of ranges) {
    forEachCellIn(sheet, range, (r, c, cell) => {
      const next: Cell = { ...cell };
      if (what === 'contents' || what === 'all') {
        delete next.v;
        delete next.f;
        delete next.cached;
        delete next.link;
      }
      if (what === 'formats' || what === 'all') delete next.s;
      if (what === 'notes' || what === 'all') delete next.note;
      if (what === 'links') delete next.link;
      tx.setCell(sheet, r, c, next);
    });
    if (what === 'formats' || what === 'all') {
      if (isFullColumn(range)) for (let c = range.c1; c <= range.c2; c++) tx.setMap(sheet.colStyles, c, undefined, sheet);
      if (isFullRow(range)) for (let r = range.r1; r <= range.r2; r++) tx.setMap(sheet.rowStyles, r, undefined, sheet);
      if (what === 'all' || what === 'formats') {
        // Conditional formats fully inside the cleared area are removed too
        if (what === 'all') {
          const remaining = sheet.cf.filter((rule) => !rule.ranges.every((rr) => contains(range, rr)));
          if (remaining.length !== sheet.cf.length) tx.setProp(sheet, 'cf', remaining);
        }
      }
    }
  }
}

function contains(outer: Range, inner: Range): boolean {
  return inner.r1 >= outer.r1 && inner.r2 <= outer.r2 && inner.c1 >= outer.c1 && inner.c2 <= outer.c2;
}

export type PatchFn = StylePatch | ((current: Style, row: number, col: number) => StylePatch | null);

/** Applies a style change to ranges. Whole rows/columns use row/column styles so empty cells are formatted too. */
export function applyStyle(tx: Transaction, wb: Workbook, sheet: Sheet, ranges: Range[], patch: PatchFn): void {
  const resolve = (id: number | undefined, r: number, c: number): number | undefined => {
    const p = typeof patch === 'function' ? patch(wb.styles.get(id), r, c) : patch;
    if (!p) return id;
    return wb.styles.apply(id, p);
  };
  for (const range of ranges) {
    const fullCols = isFullColumn(range);
    const fullRows = isFullRow(range);
    if (fullCols) {
      for (let c = range.c1; c <= range.c2; c++) {
        const cur = sheet.colStyles.get(c);
        const next = resolve(cur, 0, c);
        tx.setMap(sheet.colStyles, c, next || undefined, sheet);
      }
      if (fullRows) {
        for (const [r, id] of [...sheet.rowStyles]) tx.setMap(sheet.rowStyles, r, resolve(id, r, 0) || undefined, sheet);
      }
    } else if (fullRows) {
      for (let r = range.r1; r <= range.r2; r++) {
        const cur = sheet.rowStyles.get(r) ?? (sheet.colStyles.size ? undefined : undefined);
        tx.setMap(sheet.rowStyles, r, resolve(cur, r, 0) || undefined, sheet);
      }
    }
    if (fullCols || fullRows) {
      forEachCellIn(sheet, range, (r, c, cell) => {
        const base = cell.s;
        if (base === undefined) return; // inherits the new row/column style
        const next = resolve(base, r, c);
        if (next !== base) tx.patchCell(sheet, r, c, { s: ownStyle(sheet, r, c, next) });
      });
      continue;
    }
    for (let r = range.r1; r <= range.r2; r++)
      for (let c = range.c1; c <= range.c2; c++) {
        const base = sheet.styleIdAt(r, c);
        const next = resolve(base, r, c);
        if (next === base) continue;
        const s = ownStyle(sheet, r, c, next);
        if (s !== sheet.getCell(r, c)?.s) tx.patchCell(sheet, r, c, { s });
      }
  }
}

/** The style id a cell must store to display `id` (undefined when it would inherit the same thing). */
function ownStyle(sheet: Sheet, r: number, c: number, id: number | undefined): number | undefined {
  const inherited = sheet.rowStyles.get(r) ?? sheet.colStyles.get(c);
  if ((id === undefined || id === 0) && (inherited === undefined || inherited === 0)) return undefined;
  return id ?? 0;
}

export type BorderPreset =
  | 'all'
  | 'outer'
  | 'inner'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'none'
  | 'thickOuter'
  | 'bottomDouble'
  | 'horizontal'
  | 'vertical';

export function applyBorders(
  tx: Transaction,
  wb: Workbook,
  sheet: Sheet,
  range: Range,
  preset: BorderPreset,
  side: BorderSide = { style: 'thin' },
): void {
  const thick: BorderSide = { ...side, style: 'medium' };
  const dbl: BorderSide = { ...side, style: 'double' };
  applyStyle(tx, wb, sheet, [range], (_s, r, c) => {
    const top = r === range.r1;
    const bottom = r === range.r2;
    const left = c === range.c1;
    const right = c === range.c2;
    const p: StylePatch = {};
    switch (preset) {
      case 'none':
        return { borderTop: null, borderBottom: null, borderLeft: null, borderRight: null };
      case 'all':
        return { borderTop: side, borderBottom: side, borderLeft: side, borderRight: side };
      case 'outer':
      case 'thickOuter': {
        const s = preset === 'thickOuter' ? thick : side;
        if (top) p.borderTop = s;
        if (bottom) p.borderBottom = s;
        if (left) p.borderLeft = s;
        if (right) p.borderRight = s;
        return Object.keys(p).length ? p : null;
      }
      case 'inner':
        if (!top) p.borderTop = side;
        if (!bottom) p.borderBottom = side;
        if (!left) p.borderLeft = side;
        if (!right) p.borderRight = side;
        return Object.keys(p).length ? p : null;
      case 'horizontal':
        if (!top) p.borderTop = side;
        if (!bottom) p.borderBottom = side;
        return Object.keys(p).length ? p : null;
      case 'vertical':
        if (!left) p.borderLeft = side;
        if (!right) p.borderRight = side;
        return Object.keys(p).length ? p : null;
      case 'top':
        return top ? { borderTop: side } : null;
      case 'bottom':
        return bottom ? { borderBottom: side } : null;
      case 'bottomDouble':
        return bottom ? { borderBottom: dbl } : null;
      case 'left':
        return left ? { borderLeft: side } : null;
      case 'right':
        return right ? { borderRight: side } : null;
    }
  });
  // Clearing borders also clears the matching edges of neighbouring cells
  if (preset === 'none') {
    const neighbors: [Range, StylePatch][] = [];
    if (range.r1 > 0) neighbors.push([{ r1: range.r1 - 1, r2: range.r1 - 1, c1: range.c1, c2: range.c2 }, { borderBottom: null }]);
    neighbors.push([{ r1: range.r2 + 1, r2: range.r2 + 1, c1: range.c1, c2: range.c2 }, { borderTop: null }]);
    if (range.c1 > 0) neighbors.push([{ r1: range.r1, r2: range.r2, c1: range.c1 - 1, c2: range.c1 - 1 }, { borderRight: null }]);
    neighbors.push([{ r1: range.r1, r2: range.r2, c1: range.c2 + 1, c2: range.c2 + 1 }, { borderLeft: null }]);
    for (const [rg, p] of neighbors) {
      applyStyle(tx, wb, sheet, [rg], (s) => {
        const k = Object.keys(p)[0] as keyof Style;
        return s[k] ? p : null;
      });
    }
  }
}

/** Changes the number of decimals shown (the .0 / .00 buttons). */
export function changeDecimals(tx: Transaction, doc: Doc, sheet: Sheet, ranges: Range[], delta: number, sampleRow: number, sampleCol: number): void {
  const sample = doc.engine.getValue(sheet, sampleRow, sampleCol);
  const baseFmt = doc.wb.styles.get(sheet.styleIdAt(sampleRow, sampleCol)).numFmt;
  const nextFmt = adjustDecimals(baseFmt, delta, sample);
  applyStyle(tx, doc.wb, sheet, ranges, { numFmt: nextFmt });
}

export type MergeMode = 'merge' | 'center' | 'across' | 'unmerge';

/** Merges cells. Returns true if values other than the top-left were discarded. */
export function mergeCells(tx: Transaction, doc: Doc, sheet: Sheet, range: Range, mode: MergeMode): boolean {
  let discarded = false;
  const others = sheet.merges.filter((m) => !rangesIntersect(m, range));
  if (mode === 'unmerge') {
    tx.setProp(sheet, 'merges', others);
    return false;
  }
  const targets: Range[] = mode === 'across' ? Array.from({ length: range.r2 - range.r1 + 1 }, (_, i) => ({ r1: range.r1 + i, r2: range.r1 + i, c1: range.c1, c2: range.c2 })) : [range];
  const merges = [...others];
  for (const t of targets) {
    if (t.r1 === t.r2 && t.c1 === t.c2) continue;
    forEachCellIn(sheet, t, (r, c, cell) => {
      if (r === t.r1 && c === t.c1) return;
      if (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== '')) {
        discarded = true;
        tx.patchCell(sheet, r, c, { v: undefined, f: undefined, cached: undefined });
      }
    });
    merges.push(t);
  }
  tx.setProp(sheet, 'merges', merges);
  if (mode === 'center') applyStyle(tx, doc.wb, sheet, [{ r1: range.r1, c1: range.c1, r2: range.r1, c2: range.c1 }], { hAlign: 'center', vAlign: 'middle' });
  return discarded;
}

export function setNote(tx: Transaction, sheet: Sheet, row: number, col: number, note: string | undefined): void {
  tx.patchCell(sheet, row, col, { note: note && note.trim() ? note : undefined });
}

export function setLink(tx: Transaction, sheet: Sheet, row: number, col: number, link: string | undefined): void {
  tx.patchCell(sheet, row, col, { link: link && link.trim() ? link.trim() : undefined });
}
