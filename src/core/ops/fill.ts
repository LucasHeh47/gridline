// Auto fill (dragging the fill handle).
import { cellKey, type Range } from '../address';
import { dateToSerial, DAYS, MONTHS, serialToParts } from '../dates';
import type { Doc } from '../doc';
import { FUNCTIONS } from '../formula/functions';
import { normalizeFuncName } from '../formula/parser';
import { translateFormula } from '../formula/refs';
import { tokenize } from '../formula/tokenizer';
import type { Transaction } from '../history';
import { isDateFormat } from '../numfmt';
import type { Cell, Sheet } from '../workbook';

export type FillMode = 'auto' | 'copy' | 'series' | 'formats' | 'values';

const LISTS: string[][] = [
  MONTHS,
  MONTHS.map((m) => m.slice(0, 3)),
  DAYS,
  DAYS.map((d) => d.slice(0, 3)),
];

function matchCase(template: string, word: string): string {
  if (template === template.toUpperCase()) return word.toUpperCase();
  if (template === template.toLowerCase()) return word.toLowerCase();
  return word;
}

type Gen = (k: number) => Cell | undefined;

/**
 * For functions with a "series" argument (NFL season/week), dragging the fill handle steps the
 * typed-in value: =NFL.PLAYER("Tom Brady","Passing TDs",2015) filled down gives 2016, 2017…
 * If a week is given, the week steps instead.
 */
export function bumpSeriesArgs(formula: string, delta: number): string {
  if (!delta || !/[A-Za-z_.]+\s*\(/.test(formula)) return formula;
  const toks = tokenize(formula);
  interface Frame {
    def: { season?: number; week?: number } | undefined;
    arg: number;
    argToks: Map<number, number[]>;
  }
  const stack: Frame[] = [];
  let pendingFunc: string | null = null;
  const edits = new Map<number, string>();
  const finish = (f: Frame) => {
    if (!f.def) return;
    const single = (argIdx: number | undefined) => {
      if (argIdx === undefined) return null;
      const list = (f.argToks.get(argIdx) ?? []).filter((i) => toks[i].type !== 'ws');
      return list.length === 1 && toks[list[0]].type === 'num' ? list[0] : null;
    };
    const week = single(f.def.week);
    const season = single(f.def.season);
    const target = week ?? season;
    if (target === null) return;
    const n = Number(toks[target].value) + delta;
    edits.set(target, String(n));
  };
  toks.forEach((t, i) => {
    const top = stack[stack.length - 1];
    if (t.type === 'func') {
      pendingFunc = normalizeFuncName(t.text);
      return;
    }
    if (t.type === 'lparen') {
      stack.push({ def: pendingFunc ? FUNCTIONS.get(pendingFunc)?.series : undefined, arg: 0, argToks: new Map() });
      pendingFunc = null;
      return;
    }
    pendingFunc = null;
    if (t.type === 'rparen') {
      const f = stack.pop();
      if (f) finish(f);
      if (stack.length) {
        const parent = stack[stack.length - 1];
        const list = parent.argToks.get(parent.arg) ?? [];
        list.push(-1); // nested call: the argument isn't a plain literal
        parent.argToks.set(parent.arg, list);
      }
      return;
    }
    if (!top) return;
    if (t.type === 'comma') {
      top.arg++;
      return;
    }
    const list = top.argToks.get(top.arg) ?? [];
    list.push(i);
    top.argToks.set(top.arg, list);
  });
  if (!edits.size) return formula;
  return toks.map((t, i) => edits.get(i) ?? t.text).join('');
}

/** Builds a generator for one row/column of source cells. k is the offset from the first source cell. */
function makeGenerator(doc: Doc, sheet: Sheet, cells: (Cell | undefined)[], mode: FillMode, vertical: boolean): Gen {
  const n = cells.length;
  const mod = (k: number) => ((k % n) + n) % n;
  const copyGen: Gen = (k) => {
    const i = mod(k);
    const cell = cells[i];
    if (!cell) return undefined;
    if (cell.f !== undefined) {
      const d = k - i;
      let f = translateFormula(cell.f, vertical ? d : 0, vertical ? 0 : d);
      if (mode !== 'copy' && mode !== 'formats') f = bumpSeriesArgs(f, d);
      return { ...cell, f };
    }
    return cell;
  };
  if (mode === 'copy' || mode === 'formats') return copyGen;
  if (cells.some((c) => c?.f !== undefined)) return copyGen;

  const values = cells.map((c) => c?.v ?? null);
  const fmt = (c: Cell | undefined) => doc.wb.styles.get(c?.s).numFmt;
  const withValue = (k: number, v: number | string): Cell => {
    const base = cells[mod(k)];
    return { ...(base ?? {}), v };
  };

  // Numbers (and dates)
  if (values.every((v) => typeof v === 'number')) {
    const nums = values as number[];
    const isDate = isDateFormat(fmt(cells[0]));
    if (n === 1) {
      if (mode === 'series' || isDate) return (k) => withValue(k, nums[0] + k);
      return copyGen;
    }
    if (isDate) {
      // Monthly / yearly patterns keep the day of month
      const parts = nums.map(serialToParts);
      const sameDay = parts.every((p) => p.d === parts[0].d);
      if (sameDay) {
        const months = parts.map((p) => p.y * 12 + p.m - 1);
        const step = months[1] - months[0];
        if (step !== 0 && months.every((m, i) => m - months[0] === step * i)) {
          return (k) => {
            const total = months[0] + step * k;
            const y = Math.floor(total / 12);
            const m = (total % 12) + 1;
            const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
            return withValue(k, dateToSerial(y, m, Math.min(parts[0].d, dim)));
          };
        }
      }
    }
    // Least-squares linear trend (like Excel)
    const xs = nums.map((_, i) => i);
    const mx = (n - 1) / 2;
    const my = nums.reduce((a, b) => a + b, 0) / n;
    let sxy = 0;
    let sxx = 0;
    for (let i = 0; i < n; i++) {
      sxy += (xs[i] - mx) * (nums[i] - my);
      sxx += (xs[i] - mx) ** 2;
    }
    const slope = sxx ? sxy / sxx : 0;
    const icpt = my - slope * mx;
    return (k) => withValue(k, Number((icpt + slope * k).toPrecision(15)));
  }

  // Named lists: months and weekdays
  if (values.every((v) => typeof v === 'string')) {
    const strs = values as string[];
    for (const list of LISTS) {
      const idx = strs.map((s) => list.findIndex((x) => x.toLowerCase() === s.toLowerCase()));
      if (idx.every((i) => i >= 0)) {
        const step = n === 1 ? 1 : idx[1] - idx[0];
        const len = list.length;
        return (k) => withValue(k, matchCase(strs[0], list[(((idx[0] + step * k) % len) + len) % len]));
      }
    }
    // Text with a trailing number: "Item 1", "Q3", "Week 07"
    const re = /^(.*?)(\d+)(\D*)$/;
    const ms = strs.map((s) => re.exec(s));
    if (ms.every((m) => m && m[1] === ms[0]![1] && m[3] === ms[0]![3])) {
      const nums = ms.map((m) => parseInt(m![2], 10));
      const width = ms[0]![2].length;
      const step = n === 1 ? 1 : nums[1] - nums[0];
      if (n === 1 || nums.every((x, i) => x - nums[0] === step * i)) {
        return (k) => {
          const v = Math.max(0, nums[0] + step * k);
          const digits = String(v).padStart(ms[0]![2].startsWith('0') ? width : 1, '0');
          return withValue(k, ms[0]![1] + digits + ms[0]![3]);
        };
      }
    }
  }
  return copyGen;
}

/**
 * Fills `target` (which contains `src`) by extending the pattern in `src`.
 */
export function autofill(tx: Transaction, doc: Doc, sheet: Sheet, src: Range, target: Range, mode: FillMode): void {
  const vertical = target.r1 < src.r1 || target.r2 > src.r2;
  if (vertical) {
    for (let c = src.c1; c <= src.c2; c++) {
      const cells: (Cell | undefined)[] = [];
      for (let r = src.r1; r <= src.r2; r++) cells.push(sheet.cells.get(cellKey(r, c)));
      const gen = makeGenerator(doc, sheet, cells, mode, true);
      for (let r = target.r1; r <= target.r2; r++) {
        if (r >= src.r1 && r <= src.r2) continue;
        write(tx, sheet, r, c, gen(r - src.r1), mode);
      }
    }
  } else {
    for (let r = src.r1; r <= src.r2; r++) {
      const cells: (Cell | undefined)[] = [];
      for (let c = src.c1; c <= src.c2; c++) cells.push(sheet.cells.get(cellKey(r, c)));
      const gen = makeGenerator(doc, sheet, cells, mode, false);
      for (let c = target.c1; c <= target.c2; c++) {
        if (c >= src.c1 && c <= src.c2) continue;
        write(tx, sheet, r, c, gen(c - src.c1), mode);
      }
    }
  }
}

function write(tx: Transaction, sheet: Sheet, r: number, c: number, cell: Cell | undefined, mode: FillMode): void {
  const existing = sheet.getCell(r, c);
  if (mode === 'formats') {
    tx.patchCell(sheet, r, c, { s: cell?.s });
    return;
  }
  if (mode === 'values') {
    const next: Cell = { ...(cell ?? {}) };
    if (existing?.s !== undefined) next.s = existing.s;
    else delete next.s;
    tx.setCell(sheet, r, c, next);
    return;
  }
  const next: Cell | undefined = cell ? { ...cell } : existing?.note ? { note: existing.note } : undefined;
  if (next) delete next.note;
  if (existing?.note && next) next.note = existing.note;
  tx.setCell(sheet, r, c, next);
}

/** Double-clicking the fill handle: how far down to fill, based on the neighbouring column's data. */
export function fillDownExtent(doc: Doc, sheet: Sheet, src: Range): number | null {
  const has = (r: number, c: number) => c >= 0 && doc.engine.getValue(sheet, r, c) !== null;
  const probeCols = [src.c1 - 1, src.c2 + 1];
  for (const c of probeCols) {
    if (c < 0 || !has(src.r2, c) && !has(src.r2 + 1, c)) continue;
    let r = src.r2;
    while (has(r + 1, c)) r++;
    if (r > src.r2) return r;
  }
  return null;
}
