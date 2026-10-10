// The current selection: one or more ranges plus the active cell.
import { MAX_COLS, MAX_ROWS, normalize, rangeContains, type CellAddr, type Range } from '../core/address';
import type { SelectionSnapshot } from '../core/history';
import type { Sheet } from '../core/workbook';

export class Selection {
  ranges: Range[] = [{ r1: 0, c1: 0, r2: 0, c2: 0 }];
  active: CellAddr = { row: 0, col: 0 };
  /** The fixed corner when extending with Shift. */
  anchor: CellAddr = { row: 0, col: 0 };
  /** The moving corner when extending with Shift+arrows. */
  cursor: CellAddr | null = null;

  get range(): Range {
    return this.ranges[this.ranges.length - 1];
  }

  get isSingleCell(): boolean {
    const r = this.range;
    return this.ranges.length === 1 && r.r1 === r.r2 && r.c1 === r.c2;
  }

  setCell(sheet: Sheet, row: number, col: number): void {
    this.cursor = null;
    const m = sheet.mergeAt(row, col);
    this.active = m ? { row: m.r1, col: m.c1 } : { row, col };
    this.anchor = { ...this.active };
    this.ranges = [m ? { ...m } : { r1: row, c1: col, r2: row, c2: col }];
  }

  /** Extends the current range from the anchor to (row, col). */
  extendTo(sheet: Sheet, row: number, col: number): void {
    this.cursor = { row, col };
    const r = expandMerges(sheet, normalize({ r1: this.anchor.row, c1: this.anchor.col, r2: row, c2: col }));
    this.ranges[this.ranges.length - 1] = r;
  }

  addCell(sheet: Sheet, row: number, col: number): void {
    this.cursor = null;
    const m = sheet.mergeAt(row, col);
    this.active = m ? { row: m.r1, col: m.c1 } : { row, col };
    this.anchor = { ...this.active };
    this.ranges.push(m ? { ...m } : { r1: row, c1: col, r2: row, c2: col });
  }

  setRange(r: Range, active?: CellAddr): void {
    this.cursor = null;
    this.ranges = [normalize(r)];
    this.active = active ?? { row: r.r1, col: r.c1 };
    this.anchor = { ...this.active };
  }

  selectColumns(c1: number, c2: number, add = false): void {
    this.cursor = null;
    const r = { r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) };
    if (add) this.ranges.push(r);
    else this.ranges = [r];
    this.active = { row: 0, col: c1 };
    this.anchor = { row: 0, col: c1 };
  }

  selectRows(r1: number, r2: number, add = false): void {
    this.cursor = null;
    const r = { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 };
    if (add) this.ranges.push(r);
    else this.ranges = [r];
    this.active = { row: r1, col: 0 };
    this.anchor = { row: r1, col: 0 };
  }

  selectAll(): void {
    this.cursor = null;
    this.ranges = [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }];
    this.active = { row: 0, col: 0 };
    this.anchor = { row: 0, col: 0 };
  }

  contains(row: number, col: number): boolean {
    return this.ranges.some((r) => rangeContains(r, row, col));
  }

  /** Columns fully selected (whole-column selections). */
  fullColumns(): number[] {
    const out = new Set<number>();
    for (const r of this.ranges) if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) for (let c = r.c1; c <= r.c2; c++) out.add(c);
    return [...out].sort((a, b) => a - b);
  }

  fullRows(): number[] {
    const out = new Set<number>();
    for (const r of this.ranges) if (r.c1 === 0 && r.c2 === MAX_COLS - 1) for (let i = r.r1; i <= r.r2; i++) out.add(i);
    return [...out].sort((a, b) => a - b);
  }

  snapshot(sheetId: number): SelectionSnapshot {
    return { sheetId, ranges: this.ranges.map((r) => ({ ...r })), active: { ...this.active } };
  }

  restore(s: SelectionSnapshot): void {
    this.cursor = null;
    this.ranges = s.ranges.map((r) => ({ ...r }));
    this.active = { ...s.active };
    this.anchor = { ...s.active };
  }
}

/** Grows a range until it fully contains every merged block it touches. */
export function expandMerges(sheet: Sheet, r: Range): Range {
  if (!sheet.merges.length) return r;
  let out = { ...r };
  let changed = true;
  let guard = 0;
  while (changed && guard++ < 50) {
    changed = false;
    for (const m of sheet.merges) {
      if (m.r1 <= out.r2 && out.r1 <= m.r2 && m.c1 <= out.c2 && out.c1 <= m.c2) {
        const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
        if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) {
          out = n;
          changed = true;
        }
      }
    }
  }
  return out;
}
