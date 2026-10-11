// Inserting/deleting rows and columns, resizing, hiding.
import { keyCol, keyRow, MAX_COLS, MAX_ROWS, type Range } from '../address';
import { adjustForStructure, adjustRange, type StructuralChange } from '../formula/refs';
import type { Transaction } from '../history';
import type { Cell, ChartDef, Sheet, Workbook } from '../workbook';

function shiftIndex(i: number, ch: StructuralChange): number | null {
  if (ch.count > 0) return i >= ch.at ? i + ch.count : i;
  const n = -ch.count;
  if (i < ch.at) return i;
  if (i < ch.at + n) return null;
  return i - n;
}

function shiftMap<V>(map: Map<number, V>, ch: StructuralChange, limit: number): Map<number, V> {
  const out = new Map<number, V>();
  for (const [k, v] of map) {
    const nk = shiftIndex(k, ch);
    if (nk !== null && nk < limit) out.set(nk, v);
  }
  return out;
}

function shiftSet(set: Set<number>, ch: StructuralChange, limit: number): Set<number> {
  const out = new Set<number>();
  for (const k of set) {
    const nk = shiftIndex(k, ch);
    if (nk !== null && nk < limit) out.add(nk);
  }
  return out;
}

function adjustChartRef(text: string | undefined, sheet: Sheet, ch: StructuralChange): string | undefined {
  if (!text) return text;
  return adjustForStructure(text, sheet.name, ch);
}

/** Inserts (count > 0) or deletes (count < 0) rows or columns on a sheet, updating every formula in the workbook. */
export function restructure(tx: Transaction, wb: Workbook, sheet: Sheet, axis: 'row' | 'col', at: number, count: number): void {
  const ch: StructuralChange = { sheetName: sheet.name, axis, at, count };
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;

  // 1. Move cells on this sheet and fix their formulas
  const cells = new Map<number, Cell>();
  for (const [key, cell] of sheet.cells) {
    let r = keyRow(key);
    let c = keyCol(key);
    const idx = shiftIndex(axis === 'row' ? r : c, ch);
    if (idx === null || idx >= limit) continue;
    if (axis === 'row') r = idx;
    else c = idx;
    let next = cell;
    if (cell.f !== undefined) {
      const f = adjustForStructure(cell.f, sheet.name, ch);
      if (f !== cell.f) next = { ...cell, f };
    }
    cells.set(r * MAX_COLS + c, next);
  }
  tx.setProp(sheet, 'cells', cells, true);

  // 2. Fix formulas on other sheets that point at this sheet
  for (const other of wb.sheets) {
    if (other === sheet) continue;
    for (const [key, cell] of other.cells) {
      if (cell.f === undefined) continue;
      const f = adjustForStructure(cell.f, other.name, ch);
      if (f !== cell.f) tx.setCell(other, keyRow(key), keyCol(key), { ...cell, f });
    }
  }

  // 3. Sizes, visibility and row/column styles
  if (axis === 'row') {
    tx.setProp(sheet, 'rowHeights', shiftMap(sheet.rowHeights, ch, limit));
    tx.setProp(sheet, 'hiddenRows', shiftSet(sheet.hiddenRows, ch, limit));
    tx.setProp(sheet, 'filterHidden', shiftSet(sheet.filterHidden, ch, limit));
    tx.setProp(sheet, 'rowStyles', shiftMap(sheet.rowStyles, ch, limit));
  } else {
    tx.setProp(sheet, 'colWidths', shiftMap(sheet.colWidths, ch, limit));
    tx.setProp(sheet, 'hiddenCols', shiftSet(sheet.hiddenCols, ch, limit));
    tx.setProp(sheet, 'colStyles', shiftMap(sheet.colStyles, ch, limit));
  }

  // 4. Merges, conditional formats, filter, charts
  const fix = (r: Range) => adjustRange(r, ch);
  tx.setProp(sheet, 'merges', sheet.merges.map(fix).filter((r): r is Range => !!r && (r.r1 !== r.r2 || r.c1 !== r.c2)));
  tx.setProp(
    sheet,
    'cf',
    sheet.cf
      .map((rule) => {
        const ranges = rule.ranges.map(fix).filter((r): r is Range => !!r);
        if (!ranges.length) return null;
        const next = { ...rule, ranges };
        if (next.type === 'expression') next.formula = adjustForStructure(next.formula, sheet.name, ch);
        if (next.type === 'cellIs') next.values = next.values.map((v) => (v.startsWith('=') ? '=' + adjustForStructure(v.slice(1), sheet.name, ch) : v));
        return next;
      })
      .filter((r): r is NonNullable<typeof r> => !!r),
  );
  if (sheet.validations.length) {
    const adj = (f: string | undefined) => (f === undefined || f.trim().startsWith('"') ? f : adjustForStructure(f, sheet.name, ch));
    tx.setProp(
      sheet,
      'validations',
      sheet.validations
        .map((dv) => {
          const ranges = dv.ranges.map(fix).filter((r): r is Range => !!r);
          return ranges.length ? { ...dv, ranges, formula1: adj(dv.formula1) ?? '', formula2: adj(dv.formula2) } : null;
        })
        .filter((d): d is NonNullable<typeof d> => !!d),
    );
  }
  if (sheet.filter) {
    const range = fix(sheet.filter.range);
    if (!range) tx.setProp(sheet, 'filter', null);
    else {
      const conds: typeof sheet.filter.conds = {};
      for (const [k, v] of Object.entries(sheet.filter.conds)) {
        const nk = axis === 'col' ? shiftIndex(+k, ch) : +k;
        if (nk !== null) conds[nk] = v;
      }
      tx.setProp(sheet, 'filter', { range, conds });
    }
  }
  for (const s of wb.sheets) {
    if (!s.charts.length) continue;
    const charts = s.charts.map((c): ChartDef => {
      const anchor = { ...c.anchor };
      if (s === sheet) {
        const idx = shiftIndex(axis === 'row' ? anchor.row : anchor.col, ch);
        if (idx !== null) {
          if (axis === 'row') anchor.row = idx;
          else anchor.col = idx;
        }
      }
      return {
        ...c,
        anchor,
        series: c.series.map((se) => ({
          ...se,
          name: se.name?.startsWith('=') ? '=' + adjustChartRef(se.name.slice(1), s, ch) : se.name,
          categories: adjustChartRef(se.categories, s, ch),
          values: adjustChartRef(se.values, s, ch) ?? se.values,
        })),
      };
    });
    tx.setProp(s, 'charts', charts);
  }

  if (sheet.buttons.length) {
    const buttons = sheet.buttons.map((b) => {
      const anchor = { ...b.anchor };
      const idx = shiftIndex(axis === 'row' ? anchor.row : anchor.col, ch);
      if (idx !== null) {
        if (axis === 'row') anchor.row = idx;
        else anchor.col = idx;
      }
      return { ...b, anchor };
    });
    tx.setProp(sheet, 'buttons', buttons);
  }

  // 5. Defined names
  if (wb.names.size) {
    const names = new Map(wb.names);
    for (const [k, dn] of names) {
      const ref = adjustForStructure(dn.ref, sheet.name, ch);
      if (ref !== dn.ref) names.set(k, { ...dn, ref });
    }
    tx.setProp(wb, 'names', names, true);
  }

  // 6. Frozen panes shrink if their rows/cols are deleted
  if (count < 0) {
    const f = { ...sheet.freeze };
    if (axis === 'row' && at < f.rows) f.rows = Math.max(at, f.rows + count);
    if (axis === 'col' && at < f.cols) f.cols = Math.max(at, f.cols + count);
    if (f.rows !== sheet.freeze.rows || f.cols !== sheet.freeze.cols) tx.setProp(sheet, 'freeze', f);
  }
}

export function setColumnWidth(tx: Transaction, sheet: Sheet, cols: number[], width: number | undefined): void {
  for (const c of cols) tx.setMap(sheet.colWidths, c, width === undefined ? undefined : Math.max(2, Math.round(width)), sheet);
}

export function setRowHeight(tx: Transaction, sheet: Sheet, rows: number[], height: number | undefined): void {
  for (const r of rows) tx.setMap(sheet.rowHeights, r, height === undefined ? undefined : Math.max(2, Math.round(height)), sheet);
}

export function setHidden(tx: Transaction, sheet: Sheet, axis: 'row' | 'col', indexes: number[], hidden: boolean): void {
  const set = axis === 'row' ? sheet.hiddenRows : sheet.hiddenCols;
  for (const i of indexes) tx.setMember(set, i, hidden, sheet);
  tx.structural = tx.structural || axis === 'row'; // SUBTOTAL(10x) depends on hidden rows
}
