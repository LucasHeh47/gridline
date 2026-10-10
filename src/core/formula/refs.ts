// Rewrites references inside formula text (copy/paste, insert/delete rows, rename sheet...).
import { MAX_COLS, MAX_ROWS, type Range } from '../address';
import { refToString, tokenize, type RefInfo } from './tokenizer';

/**
 * Maps every reference in a formula. The callback returns:
 *  - undefined to keep the reference unchanged,
 *  - null to replace it with #REF!,
 *  - a RefInfo to replace it.
 */
export function mapRefs(text: string, fn: (ref: RefInfo) => RefInfo | null | undefined): string {
  const toks = tokenize(text);
  let changed = false;
  const out = toks.map((t) => {
    if (t.type !== 'ref' || !t.ref) return t.text;
    const r = fn(t.ref);
    if (r === undefined) return t.text;
    changed = true;
    if (r === null) return '#REF!';
    return refToString(r);
  });
  return changed ? out.join('') : text;
}

function inBounds(r: RefInfo): boolean {
  return r.r1 >= 0 && r.c1 >= 0 && r.r2 < MAX_ROWS && r.c2 < MAX_COLS && r.r1 <= r.r2 && r.c1 <= r.c2;
}

function normalized(r: RefInfo): RefInfo {
  const out = { ...r };
  if (out.r1 > out.r2) {
    [out.r1, out.r2] = [out.r2, out.r1];
    [out.ar1, out.ar2] = [out.ar2, out.ar1];
  }
  if (out.c1 > out.c2) {
    [out.c1, out.c2] = [out.c2, out.c1];
    [out.ac1, out.ac2] = [out.ac2, out.ac1];
  }
  return out;
}

/** Shifts relative references, as when copying a formula dRow rows down and dCol columns right. */
export function translateFormula(text: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return text;
  return mapRefs(text, (ref) => {
    const r = { ...ref };
    if (ref.kind !== 'cols') {
      if (!ref.ar1) r.r1 += dRow;
      if (!ref.ar2) r.r2 += dRow;
    }
    if (ref.kind !== 'rows') {
      if (!ref.ac1) r.c1 += dCol;
      if (!ref.ac2) r.c2 += dCol;
    }
    if (r.r1 === ref.r1 && r.r2 === ref.r2 && r.c1 === ref.c1 && r.c2 === ref.c2) return undefined;
    if (r.r1 < 0 || r.c1 < 0 || r.r2 >= MAX_ROWS || r.c2 >= MAX_COLS || r.r2 < 0 || r.c2 < 0 || r.r1 >= MAX_ROWS || r.c1 >= MAX_COLS)
      return null;
    return normalized(r);
  });
}

export interface StructuralChange {
  sheetName: string;
  axis: 'row' | 'col';
  at: number;
  /** Positive = insert, negative = delete. */
  count: number;
}

function sameSheet(ref: RefInfo, formulaSheet: string, target: string): boolean {
  const s = ref.sheet ?? formulaSheet;
  return s.toLowerCase() === target.toLowerCase();
}

function shiftAxis(lo: number, hi: number, at: number, count: number, max: number): [number, number] | null {
  if (count > 0) {
    const nlo = lo >= at ? lo + count : lo;
    const nhi = hi >= at ? hi + count : hi;
    return [Math.min(nlo, max - 1), Math.min(nhi, max - 1)];
  }
  const n = -count;
  const end = at + n; // exclusive
  let nlo: number;
  let nhi: number;
  if (lo < at) nlo = lo;
  else if (lo >= end) nlo = lo - n;
  else nlo = at;
  if (hi < at) nhi = hi;
  else if (hi >= end) nhi = hi - n;
  else nhi = at - 1;
  if (nhi < nlo) return null;
  return [nlo, nhi];
}

/** Adjusts references after inserting or deleting rows/columns. */
export function adjustForStructure(text: string, formulaSheet: string, ch: StructuralChange): string {
  return mapRefs(text, (ref) => {
    if (!sameSheet(ref, formulaSheet, ch.sheetName)) return undefined;
    if (ch.axis === 'row') {
      if (ref.kind === 'cols') return undefined;
      const res = shiftAxis(ref.r1, ref.r2, ch.at, ch.count, MAX_ROWS);
      if (!res) return null;
      if (ref.kind === 'cell' && res[0] !== res[1]) return null;
      if (res[0] === ref.r1 && res[1] === ref.r2) return undefined;
      return { ...ref, r1: res[0], r2: res[1] };
    }
    if (ref.kind === 'rows') return undefined;
    const res = shiftAxis(ref.c1, ref.c2, ch.at, ch.count, MAX_COLS);
    if (!res) return null;
    if (res[0] === ref.c1 && res[1] === ref.c2) return undefined;
    return { ...ref, c1: res[0], c2: res[1] };
  });
}

/** Adjusts a range (merges, CF ranges, charts) after a structural change. Returns null if deleted. */
export function adjustRange(r: Range, ch: StructuralChange): Range | null {
  if (ch.axis === 'row') {
    const res = shiftAxis(r.r1, r.r2, ch.at, ch.count, MAX_ROWS);
    return res ? { ...r, r1: res[0], r2: res[1] } : null;
  }
  const res = shiftAxis(r.c1, r.c2, ch.at, ch.count, MAX_COLS);
  return res ? { ...r, c1: res[0], c2: res[1] } : null;
}

/**
 * Updates references pointing into a moved block (cut & paste). References fully inside
 * `src` on sheet `srcSheet` are shifted by (dRow, dCol) and re-pointed at `dstSheet`.
 */
export function adjustForMove(
  text: string,
  formulaSheet: string,
  srcSheet: string,
  src: Range,
  dRow: number,
  dCol: number,
  dstSheet: string,
): string {
  return mapRefs(text, (ref) => {
    if (!sameSheet(ref, formulaSheet, srcSheet)) return undefined;
    if (ref.r1 < src.r1 || ref.r2 > src.r2 || ref.c1 < src.c1 || ref.c2 > src.c2) return undefined;
    const r: RefInfo = { ...ref, r1: ref.r1 + dRow, r2: ref.r2 + dRow, c1: ref.c1 + dCol, c2: ref.c2 + dCol };
    if (!inBounds(r)) return null;
    const formulaOnDst = formulaSheet.toLowerCase() === dstSheet.toLowerCase();
    if (dstSheet.toLowerCase() !== srcSheet.toLowerCase() || ref.sheet !== undefined) {
      r.sheet = formulaOnDst && ref.sheet === undefined ? undefined : dstSheet;
    }
    return r;
  });
}

export function renameSheetRefs(text: string, oldName: string, newName: string): string {
  const lower = oldName.toLowerCase();
  return mapRefs(text, (ref) => {
    if (ref.sheet === undefined || ref.sheet.toLowerCase() !== lower) return undefined;
    return { ...ref, sheet: newName };
  });
}

export function deleteSheetRefs(text: string, name: string): string {
  const lower = name.toLowerCase();
  return mapRefs(text, (ref) => (ref.sheet !== undefined && ref.sheet.toLowerCase() === lower ? null : undefined));
}

/** Toggles $ signs on the reference at the caret (F4 key): A1 -> $A$1 -> A$1 -> $A1 -> A1. */
export function cycleAbsolute(text: string, caret: number): { text: string; caret: number } | null {
  const toks = tokenize(text);
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref) continue;
    if (caret < t.start || caret > t.end) continue;
    const r = { ...t.ref };
    const cycle = (ar: boolean, ac: boolean): [boolean, boolean] => {
      if (!ar && !ac) return [true, true];
      if (ar && ac) return [true, false];
      if (ar && !ac) return [false, true];
      return [false, false];
    };
    [r.ar1, r.ac1] = cycle(r.ar1, r.ac1);
    [r.ar2, r.ac2] = r.kind === 'cell' ? [r.ar1, r.ac1] : cycle(t.ref.ar2, t.ref.ac2);
    const s = refToString(r);
    return { text: text.slice(0, t.start) + s + text.slice(t.end), caret: t.start + s.length };
  }
  return null;
}
