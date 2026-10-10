// Copy, cut and paste.
import { cellKey, keyCol, keyRow, MAX_COLS, MAX_ROWS, rangesIntersect, type Range } from '../address';
import type { Doc } from '../doc';
import { adjustForMove, translateFormula } from '../formula/refs';
import type { Transaction } from '../history';
import { formatValue } from '../numfmt';
import type { Scalar } from '../values';
import type { Cell, Sheet } from '../workbook';
import { setCellInput } from './cells';

export interface ClipData {
  sheetId: number;
  range: Range;
  cells: (Cell | undefined)[][];
  values: Scalar[][];
  text: string;
  html: string;
  cut: boolean;
  merges: Range[];
  /** Source row index for each copied row (filtered rows are skipped). */
  srcRows: number[];
}

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas' | 'transpose' | 'noBorders';

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function tsvField(s: string): string {
  return /[\t\n\r"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

export function clampToUsed(doc: Doc, sheet: Sheet, range: Range): Range {
  const b = doc.engine.usedBounds(sheet);
  return {
    r1: range.r1,
    c1: range.c1,
    r2: range.r2 === MAX_ROWS - 1 ? Math.max(range.r1, b.maxRow) : range.r2,
    c2: range.c2 === MAX_COLS - 1 ? Math.max(range.c1, b.maxCol) : range.c2,
  };
}

export function copyRange(doc: Doc, sheet: Sheet, input: Range, cut: boolean): ClipData {
  const range = clampToUsed(doc, sheet, input);
  const cells: (Cell | undefined)[][] = [];
  const values: Scalar[][] = [];
  const textRows: string[] = [];
  const htmlRows: string[] = [];
  const srcRows: number[] = [];
  for (let r = range.r1; r <= range.r2; r++) {
    if (sheet.filterHidden.has(r) && !cut && range.r2 > range.r1) continue; // Like Excel, skip filtered rows
    srcRows.push(r);
    const rowCells: (Cell | undefined)[] = [];
    const rowVals: Scalar[] = [];
    const texts: string[] = [];
    const tds: string[] = [];
    for (let c = range.c1; c <= range.c2; c++) {
      const cell = sheet.cells.get(cellKey(r, c));
      const v = doc.engine.getValue(sheet, r, c);
      const style = doc.wb.styles.get(sheet.styleIdAt(r, c));
      rowCells.push(cell);
      rowVals.push(v);
      const shown = v === null ? '' : formatValue(v, typeof v === 'number' ? doc.numFmtAt(sheet, r, c) : style.numFmt, 255).text;
      texts.push(tsvField(shown));
      const css: string[] = [];
      if (style.bold) css.push('font-weight:bold');
      if (style.italic) css.push('font-style:italic');
      if (style.underline) css.push('text-decoration:underline');
      if (style.color) css.push(`color:${style.color}`);
      if (style.fill) css.push(`background:${style.fill}`);
      if (style.hAlign && style.hAlign !== 'general') css.push(`text-align:${style.hAlign}`);
      tds.push(`<td${css.length ? ` style="${css.join(';')}"` : ''}>${escapeHtml(shown)}</td>`);
    }
    cells.push(rowCells);
    values.push(rowVals);
    textRows.push(texts.join('\t'));
    htmlRows.push(`<tr>${tds.join('')}</tr>`);
  }
  const merges = sheet.merges.filter(
    (m) => m.r1 >= range.r1 && m.r2 <= range.r2 && m.c1 >= range.c1 && m.c2 <= range.c2,
  );
  return {
    sheetId: sheet.id,
    range,
    cells,
    values,
    text: textRows.join('\r\n') + (textRows.length > 1 ? '\r\n' : ''),
    html: `<table>${htmlRows.join('')}</table>`,
    cut,
    merges,
    srcRows,
  };
}

/** Parses tab-separated text (with quoted fields) into rows. */
export function parseTSV(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  let quoted = false;
  const t = text.replace(/\r\n?/g, '\n');
  while (i < t.length) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && field === '') {
      quoted = true;
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(field);
      field = '';
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += ch;
    i++;
  }
  if (field !== '' || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** Pastes data copied inside Gridline. Returns the destination range. */
export function pasteClip(tx: Transaction, doc: Doc, clip: ClipData, dst: Sheet, sel: Range, mode: PasteMode): Range {
  const src = doc.wb.sheetById(clip.sheetId);
  const srcRows = clip.cells.length;
  const srcCols = clip.cells[0]?.length ?? 0;
  if (!srcRows || !srcCols) return sel;
  const transpose = mode === 'transpose';
  const h = transpose ? srcCols : srcRows;
  const w = transpose ? srcRows : srcCols;
  // Tile when the selection is an exact multiple of the copied size
  const selH = sel.r2 - sel.r1 + 1;
  const selW = sel.c2 - sel.c1 + 1;
  const tilesY = !clip.cut && selH % h === 0 ? selH / h : 1;
  const tilesX = !clip.cut && selW % w === 0 ? selW / w : 1;
  const out: Range = { r1: sel.r1, c1: sel.c1, r2: sel.r1 + h * tilesY - 1, c2: sel.c1 + w * tilesX - 1 };
  if (out.r2 >= MAX_ROWS || out.c2 >= MAX_COLS) throw new Error('The paste area extends past the edge of the sheet.');

  for (let ty = 0; ty < tilesY; ty++)
    for (let tx2 = 0; tx2 < tilesX; tx2++) {
      for (let i = 0; i < srcRows; i++)
        for (let j = 0; j < srcCols; j++) {
          const di = transpose ? j : i;
          const dj = transpose ? i : j;
          const r = sel.r1 + ty * h + di;
          const c = sel.c1 + tx2 * w + dj;
          const cell = clip.cells[i][j];
          const srcR = clip.srcRows[i] ?? clip.range.r1 + i;
          const srcC = clip.range.c1 + j;
          const existing = dst.getCell(r, c);
          let next: Cell | undefined;
          switch (mode) {
            case 'values': {
              const v = clip.values[i][j];
              next = { ...(existing?.s !== undefined ? { s: existing.s } : {}) };
              if (v !== null) next.v = v;
              break;
            }
            case 'formats':
              next = { ...(existing ?? {}), s: cell?.s };
              if (next.s === undefined) delete next.s;
              break;
            case 'formulas':
              next = { ...(existing?.s !== undefined ? { s: existing.s } : {}) };
              if (cell?.f !== undefined) next.f = clip.cut ? cell.f : translateFormula(cell.f, r - srcR, c - srcC);
              else if (cell?.v !== undefined) next.v = cell.v;
              break;
            case 'noBorders':
            case 'all':
            case 'transpose':
            default: {
              if (!cell) {
                next = undefined;
                break;
              }
              next = { ...cell };
              if (cell.f !== undefined && !clip.cut) next.f = translateFormula(cell.f, r - srcR, c - srcC);
              if (mode === 'noBorders' && next.s !== undefined) {
                next.s = doc.wb.styles.apply(next.s, { borderTop: null, borderBottom: null, borderLeft: null, borderRight: null });
              }
            }
          }
          tx.setCell(dst, r, c, next);
        }
    }

  if (mode === 'all' && clip.merges.length) {
    const dr = sel.r1 - clip.range.r1;
    const dc = sel.c1 - clip.range.c1;
    const moved = clip.merges.map((m) => ({ r1: m.r1 + dr, r2: m.r2 + dr, c1: m.c1 + dc, c2: m.c2 + dc }));
    const merges = dst.merges.filter((m) => !rangesIntersect(m, out)).concat(moved);
    tx.setProp(dst, 'merges', merges);
  }

  if (clip.cut && src) {
    // Clear the source cells that weren't overwritten
    for (let i = 0; i < srcRows; i++)
      for (let j = 0; j < srcCols; j++) {
        const r = clip.srcRows[i] ?? clip.range.r1 + i;
        const c = clip.range.c1 + j;
        if (src === dst && r >= out.r1 && r <= out.r2 && c >= out.c1 && c <= out.c2) continue;
        if (src.getCell(r, c)) tx.setCell(src, r, c, undefined);
      }
    if (src !== dst || sel.r1 !== clip.range.r1 || sel.c1 !== clip.range.c1) {
      const merges = src.merges.filter((m) => !clip.merges.some((x) => x.r1 === m.r1 && x.c1 === m.c1 && x.r2 === m.r2 && x.c2 === m.c2));
      if (merges.length !== src.merges.length) tx.setProp(src, 'merges', merges);
    }
    // Re-point every reference to the moved cells
    const dr = sel.r1 - clip.range.r1;
    const dc = sel.c1 - clip.range.c1;
    for (const s of doc.wb.sheets) {
      for (const [key, cell] of [...s.cells]) {
        if (cell.f === undefined) continue;
        const f = adjustForMove(cell.f, s.name, src.name, clip.range, dr, dc, dst.name);
        if (f !== cell.f) tx.setCell(s, keyRow(key), keyCol(key), { ...cell, f });
      }
    }
  }
  return out;
}

/** Pastes plain text (from another app). Returns the destination range. */
export function pasteText(tx: Transaction, doc: Doc, text: string, dst: Sheet, row: number, col: number): Range {
  let rows = parseTSV(text);
  if (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows = rows.slice(0, -1);
  if (!rows.length) return { r1: row, c1: col, r2: row, c2: col };
  const width = Math.max(...rows.map((r) => r.length));
  if (row + rows.length > MAX_ROWS || col + width > MAX_COLS) throw new Error('The pasted data extends past the edge of the sheet.');
  rows.forEach((fields, i) => {
    fields.forEach((f, j) => setCellInput(tx, doc, dst, row + i, col + j, f));
  });
  return { r1: row, c1: col, r2: row + rows.length - 1, c2: col + width - 1 };
}
