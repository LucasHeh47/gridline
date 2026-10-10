// Cell and range addressing. All indices are 0-based internally; A1 strings are 1-based.

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export interface CellAddr {
  row: number;
  col: number;
}

/** A normalized rectangular range (r1 <= r2, c1 <= c2), inclusive. */
export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function cellKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

export function keyRow(key: number): number {
  return Math.floor(key / MAX_COLS);
}

export function keyCol(key: number): number {
  return key % MAX_COLS;
}

export function colToName(col: number): string {
  let s = '';
  let n = col + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Converts column letters to a 0-based index, or -1 if invalid. */
export function nameToCol(name: string): number {
  let n = 0;
  for (let i = 0; i < name.length; i++) {
    const ch = name.charCodeAt(i) & ~32; // upper-case
    if (ch < 65 || ch > 90) return -1;
    n = n * 26 + (ch - 64);
  }
  return n - 1;
}

export function formatCell(row: number, col: number, absRow = false, absCol = false): string {
  return `${absCol ? '$' : ''}${colToName(col)}${absRow ? '$' : ''}${row + 1}`;
}

export function formatRange(r: Range): string {
  if (r.r1 === r.r2 && r.c1 === r.c2) return formatCell(r.r1, r.c1);
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return `${colToName(r.c1)}:${colToName(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return `${r.r1 + 1}:${r.r2 + 1}`;
  return `${formatCell(r.r1, r.c1)}:${formatCell(r.r2, r.c2)}`;
}

/** Formats a range always as A1:B2 (never collapsing to a single cell). */
export function formatRangeFull(r: Range, abs = false): string {
  return `${formatCell(r.r1, r.c1, abs, abs)}:${formatCell(r.r2, r.c2, abs, abs)}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

export function parseCell(s: string): CellAddr | null {
  const m = CELL_RE.exec(s.trim());
  if (!m) return null;
  const col = nameToCol(m[1]);
  const row = parseInt(m[2], 10) - 1;
  if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) return null;
  return { row, col };
}

/** Parses "A1", "A1:B2", "A:C" or "3:5" (no sheet prefix). */
export function parseRange(s: string): Range | null {
  const t = s.trim();
  const parts = t.split(':');
  if (parts.length === 1) {
    const c = parseCell(parts[0]);
    return c ? { r1: c.row, c1: c.col, r2: c.row, c2: c.col } : null;
  }
  if (parts.length !== 2) return null;
  const [a, b] = parts;
  const ca = parseCell(a);
  const cb = parseCell(b);
  if (ca && cb) return normalize({ r1: ca.row, c1: ca.col, r2: cb.row, c2: cb.col });
  const colA = /^\$?([A-Za-z]{1,3})$/.exec(a);
  const colB = /^\$?([A-Za-z]{1,3})$/.exec(b);
  if (colA && colB) {
    const c1 = nameToCol(colA[1]);
    const c2 = nameToCol(colB[1]);
    if (c1 < 0 || c2 < 0 || c1 >= MAX_COLS || c2 >= MAX_COLS) return null;
    return normalize({ r1: 0, c1, r2: MAX_ROWS - 1, c2 });
  }
  const rowA = /^\$?(\d+)$/.exec(a);
  const rowB = /^\$?(\d+)$/.exec(b);
  if (rowA && rowB) {
    const r1 = parseInt(rowA[1], 10) - 1;
    const r2 = parseInt(rowB[1], 10) - 1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return null;
    return normalize({ r1, c1: 0, r2, c2: MAX_COLS - 1 });
  }
  return null;
}

export function normalize(r: Range): Range {
  return {
    r1: Math.min(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2),
    c2: Math.max(r.c1, r.c2),
  };
}

export function cellRange(row: number, col: number): Range {
  return { r1: row, c1: col, r2: row, c2: col };
}

export function rangeContains(r: Range, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

export function intersect(a: Range, b: Range): Range | null {
  if (!rangesIntersect(a, b)) return null;
  return {
    r1: Math.max(a.r1, b.r1),
    c1: Math.max(a.c1, b.c1),
    r2: Math.min(a.r2, b.r2),
    c2: Math.min(a.c2, b.c2),
  };
}

export function union(a: Range, b: Range): Range {
  return {
    r1: Math.min(a.r1, b.r1),
    c1: Math.min(a.c1, b.c1),
    r2: Math.max(a.r2, b.r2),
    c2: Math.max(a.c2, b.c2),
  };
}

export function rangeEquals(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

export function rangeRows(r: Range): number {
  return r.r2 - r.r1 + 1;
}

export function rangeCols(r: Range): number {
  return r.c2 - r.c1 + 1;
}

export function rangeSize(r: Range): number {
  return rangeRows(r) * rangeCols(r);
}

export function isFullColumn(r: Range): boolean {
  return r.r1 === 0 && r.r2 === MAX_ROWS - 1;
}

export function isFullRow(r: Range): boolean {
  return r.c1 === 0 && r.c2 === MAX_COLS - 1;
}

/** Quotes a sheet name for use in a formula if needed. */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseCell(name) && !/^(TRUE|FALSE)$/i.test(name)) {
    return name;
  }
  return `'${name.replace(/'/g, "''")}'`;
}
