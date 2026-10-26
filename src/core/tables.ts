// Tables (Ctrl+T): a header row plus rows of data, made for sorting. Compared with Excel's tables:
//  • Making one is instant: no dialog, and nothing in the cells changes: values, formats, column
//    widths and row positions stay as they are. The table adds sort buttons to its first row and
//    an outline; colors are optional. A sheet filter over the same cells becomes the table's own.
//  • Sorting is one click on a column's arrow: smallest first, then largest first, then back to
//    the original order, which Excel can't do once you've sorted. Shift+click adds a second column.
//  • Sorting understands the data: numbers typed as text ("1,200", "$5") sort as numbers, "Item 2"
//    comes before "Item 10", capitals don't matter, and empty cells always go last.
//  • The arrow shows when rows changed after sorting (a new row, an edited value); clicking it
//    sorts again.
//  • Typing just below or right of a table makes it bigger, and formulas that use a whole column
//    (=SUM(C2:C20)) grow with it, so you don't need structured references (they work too). A total
//    you type under a table stays out of it, so sorting never mixes it in.
import { formatCell, MAX_COLS, MAX_ROWS, normalize, rangesIntersect, type Range } from './address';
import type { Doc } from './doc';
import { adjustRange, mapRefs, translateFormula } from './formula/refs';
import type { Transaction } from './history';
import { formatValue } from './numfmt';
import { refreshFilters } from './ops/data';
import { restructure } from './ops/structure';
import { expandStructured, renameColumnInFormula, renameTableInFormula, tableAt } from './tableRefs';
import { ErrorValue, type Scalar } from './values';
import { hasContent, type Cell, type FilterCond, type Sheet, type TableColor, type TableDef, type TableSortKey, type Workbook } from './workbook';

export { tableAt };

/** Colors for the light theme; the grid tints the light ones for dark mode. */
export const TABLE_COLORS: Record<TableColor, { label: string; header: string; band: string; line: string }> = {
  none: { label: 'No colors (keep the cells’ look)', header: '', band: '', line: '#8f97a6' },
  blue: { label: 'Blue', header: '#3567c8', band: '#eaf0fb', line: '#a9bfe9' },
  green: { label: 'Green', header: '#2b8a57', band: '#e6f4ec', line: '#9fd0b3' },
  orange: { label: 'Orange', header: '#d9682a', band: '#fcede3', line: '#efb894' },
  gold: { label: 'Gold', header: '#b8860b', band: '#fbf3dc', line: '#e3cc8a' },
  gray: { label: 'Gray', header: '#566273', band: '#eef0f3', line: '#bcc3cc' },
};
export const TABLE_COLOR_IDS = Object.keys(TABLE_COLORS) as TableColor[];

// ---------- names ----------

/** Why a table can't have this name (Excel's rules), or null. */
export function tableNameProblem(name: string, wb: Workbook, except?: TableDef): string | null {
  const n = name.trim();
  if (!n) return 'Give the table a name.';
  if (n.length > 255) return 'Keep the name under 255 characters.';
  if (/\s/.test(n)) return 'Table names can’t have spaces. Use _ instead, like Sales_2026.';
  if (!/^[A-Za-z_\\À-￿][\w.\\À-￿]*$/.test(n)) return 'Start with a letter or _, and use only letters, numbers, _ and periods.';
  if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^R\d*C\d*$/i.test(n) || /^[RC]$/i.test(n)) return `${n} looks like a cell address. Pick another name.`;
  const other = wb.table(n);
  if (other && other.table.id !== except?.id) return `There’s already a table named ${other.table.name}.`;
  if (wb.getName(n)) return `${n} is already the name of a range.`;
  return null;
}

/** Table1, Table2… (or Sales, Sales2… for another base). */
export function uniqueTableName(wb: Workbook, base = 'Table'): string {
  const b = base.replace(/\d+$/, '') || 'Table';
  for (let i = 1; ; i++) {
    const n = `${b}${i}`;
    if (!wb.table(n) && !wb.getName(n)) return n;
  }
}

// ---------- header row ----------

type ValueAt = (r: number, c: number) => Scalar;

function headerText(v: Scalar, numFmt?: string): string {
  if (v === null) return '';
  if (v instanceof ErrorValue) return v.code;
  if (typeof v === 'number') return formatValue(v, numFmt, 30).text;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v.trim();
}

/** Column names from header texts: empty ones become Column3, repeats get a number (Amount2), like Excel. */
export function columnNames(texts: string[]): string[] {
  const used = new Set<string>();
  const out: string[] = [];
  texts.forEach((t, i) => {
    let name = t || `Column${i + 1}`;
    if (used.has(name.toLowerCase())) {
      const base = name;
      for (let k = 2; used.has(name.toLowerCase()); k++) name = `${base}${k}`;
    }
    used.add(name.toLowerCase());
    out.push(name);
  });
  return out;
}

/** The value of a cell for the header row: its own value, or a formula's result. */
function valueAtFn(doc: Doc, sheet: Sheet): ValueAt {
  return (r, c) => {
    const cell = sheet.getCell(r, c);
    if (!cell) return null;
    if (cell.f !== undefined) {
      const v = doc.engine.evalText(cell.f, sheet, r, c);
      return v instanceof Object && 'rows' in v ? null : (v as Scalar);
    }
    return cell.v ?? null;
  };
}

/** Puts the column names in a header row that was added for them. */
function writeHeaders(tx: Transaction, sheet: Sheet, range: Range, names: string[]): void {
  names.forEach((name, i) => {
    const c = range.c1 + i;
    const cell = sheet.getCell(range.r1, c);
    if (cell?.f === undefined && cell?.v === name) return;
    tx.patchCell(sheet, range.r1, c, { v: name, f: undefined, cached: undefined });
  });
}

/**
 * Guesses whether the first row of a block is a header: every cell has text, no two are alike, and
 * either the rows below look different (numbers, not bold) or no header word appears in its column.
 */
export function looksLikeHeaderRow(doc: Doc, sheet: Sheet, range: Range): boolean {
  if (range.r2 <= range.r1) return true;
  const seen = new Set<string>();
  const heads: string[] = [];
  for (let c = range.c1; c <= range.c2; c++) {
    const h = doc.engine.getValue(sheet, range.r1, c);
    if (typeof h !== 'string' || !h.trim()) return false;
    const k = h.trim().toLowerCase();
    if (seen.has(k)) return false;
    seen.add(k);
    heads.push(k);
  }
  let differs = 0;
  for (let c = range.c1; c <= range.c2; c++) {
    const v = doc.engine.getValue(sheet, range.r1 + 1, c);
    if (v !== null && typeof v !== 'string') differs++;
    if (doc.wb.styles.get(sheet.styleIdAt(range.r1, c)).bold && !doc.wb.styles.get(sheet.styleIdAt(range.r1 + 1, c)).bold) differs++;
  }
  if (differs) return true;
  // Text over text: a header word doesn't turn up again in its own column
  const last = Math.min(range.r2, range.r1 + 500);
  for (let c = range.c1; c <= range.c2; c++)
    for (let r = range.r1 + 1; r <= last; r++) {
      const v = doc.engine.getValue(sheet, r, c);
      if (typeof v === 'string' && v.trim().toLowerCase() === heads[c - range.c1]) return false;
    }
  return true;
}

// ---------- making and changing tables ----------

/** Replaces a table (tables are never changed in place, so undo can bring the old one back). */
export function putTable(tx: Transaction, sheet: Sheet, table: TableDef, structural = true): void {
  tx.setProp(
    sheet,
    'tables',
    sheet.tables.map((t) => (t.id === table.id ? table : t)),
    structural,
  );
}

export function tableById(sheet: Sheet, id: string): TableDef | undefined {
  return sheet.tables.find((t) => t.id === id);
}

export interface MadeTable {
  table: TableDef;
  /** A header row was added above the data (only when asked for with `header: false`). */
  addedHeader: boolean;
  /** The first row gives the column names, but it looks like data (numbers, repeats, blanks). */
  firstRowLooksLikeData: boolean;
}

/** The cells include merged cells, which can't be sorted. Nothing was changed. */
export class MergedCellsError extends Error {
  constructor(readonly merges: Range[]) {
    super(`Tables can't contain merged cells (${merges.map((m) => formatCell(m.r1, m.c1) + ':' + formatCell(m.r2, m.c2)).join(', ')}), because sorting moves rows apart.`);
  }
}

/** Blank rows and columns at the edges of a block are left out (a whole-column selection, an empty row above). */
function trimBlankEdges(sheet: Sheet, r: Range): Range {
  const filled = (r1: number, c1: number, r2: number, c2: number) => {
    for (let row = r1; row <= r2; row++) for (let c = c1; c <= c2; c++) if (hasContent(sheet.getCell(row, c))) return true;
    return false;
  };
  let { r1, c1, r2, c2 } = r;
  while (r1 < r2 && !filled(r1, c1, r1, c2)) r1++;
  while (r2 > r1 && !filled(r2, c1, r2, c2)) r2--;
  while (c1 < c2 && !filled(r1, c1, r2, c1)) c1++;
  while (c2 > c1 && !filled(r1, c2, r2, c2)) c2--;
  return { r1, c1, r2, c2 };
}

/** A title above the header row (a cell or two, with a full row of names under it) stays out of the table. */
function skipTitleRow(doc: Doc, sheet: Sheet, r: Range): Range {
  if (r.c2 === r.c1 || r.r2 - r.r1 < 2) return r;
  const count = (row: number) => {
    let n = 0;
    for (let c = r.c1; c <= r.c2; c++) if (hasContent(sheet.getCell(row, c))) n++;
    return n;
  };
  const top = count(r.r1);
  const next = count(r.r1 + 1);
  if (top < next && top <= Math.ceil((r.c2 - r.c1 + 1) / 2) && looksLikeHeaderRow(doc, sheet, { ...r, r1: r.r1 + 1 })) return { ...r, r1: r.r1 + 1 };
  return r;
}

/**
 * Makes a table from a block of cells. The cells don't change: the first row gives the column
 * names (Column2 for an empty one, Amount2 for a repeat; the names live in the table, not the
 * cells). Only with `header: false` is a header row added above the data, named Column1, Column2…
 * Merged cells throw a MergedCellsError unless `unmerge` is set.
 */
export function createTable(tx: Transaction, doc: Doc, sheet: Sheet, range: Range, opts: { header?: boolean; name?: string; color?: TableColor; unmerge?: boolean } = {}): MadeTable {
  let r = normalize(range);
  // Whole columns or rows: just the part with data
  if (r.r2 === MAX_ROWS - 1 || r.c2 === MAX_COLS - 1) {
    const b = sheet.contentBounds();
    if (r.r2 === MAX_ROWS - 1) r = { ...r, r2: Math.max(r.r1, b.maxRow) };
    if (r.c2 === MAX_COLS - 1) r = { ...r, c2: Math.max(r.c1, b.maxCol) };
  }
  r = trimBlankEdges(sheet, r);
  if (r.r1 === r.r2 && r.c1 === r.c2 && !hasContent(sheet.getCell(r.r1, r.c1))) throw new Error('There’s no data here. Click inside your data (or select it), then press Ctrl+T.');
  if (opts.header === undefined) r = skipTitleRow(doc, sheet, r);
  if (r.c2 - r.c1 >= 1000) throw new Error('A table can have at most 1,000 columns.');
  const clash = sheet.tables.find((t) => rangesIntersect(t.range, r));
  if (clash) throw new Error(`That overlaps the table ${clash.name}. A cell can only be in one table.`);

  // Tables can't have merged cells: they're only split when asked to
  const merged = sheet.merges.filter((m) => rangesIntersect(m, r));
  if (merged.length) {
    if (!opts.unmerge) throw new MergedCellsError(merged);
    tx.setProp(sheet, 'merges', sheet.merges.filter((m) => !merged.includes(m)), true);
  }
  // A sheet filter over the same cells becomes the table's filter
  let filter: Record<number, FilterCond> = {};
  if (sheet.filter && rangesIntersect(sheet.filter.range, r)) {
    filter = { ...sheet.filter.conds };
    tx.setProp(sheet, 'filter', null);
  }
  const addedHeader = opts.header === false;
  if (addedHeader) {
    restructure(tx, doc.wb, sheet, 'row', r.r1, 1);
    r = { ...r, r2: r.r2 + 1 };
    for (const k of Object.keys(filter)) if (+k < r.c1 || +k > r.c2) delete filter[+k];
  }
  // A header alone gets the empty row under it to type into
  if (r.r2 === r.r1) r = { ...r, r2: r.r1 + 1 };
  const names = columnNames(addedHeader ? Array.from({ length: r.c2 - r.c1 + 1 }, () => '') : headerTexts(doc, sheet, r));
  if (addedHeader) writeHeaders(tx, sheet, r, names);
  const name = opts.name && !tableNameProblem(opts.name, doc.wb) ? opts.name.trim() : uniqueTableName(doc.wb);
  const color = opts.color ?? 'none';
  const table: TableDef = { id: 't' + Math.random().toString(36).slice(2, 10), name, range: r, columns: names, color, banded: color !== 'none', sort: [], filter };
  tx.setProp(sheet, 'tables', [...sheet.tables, table], true);
  if (Object.keys(filter).length) refreshFilters(tx, doc, sheet);
  return { table, addedHeader, firstRowLooksLikeData: !addedHeader && r.r2 > r.r1 && !looksLikeHeaderRow(doc, sheet, r) };
}

/** What the header cells show (their formatted values), for the column names. */
function headerTexts(doc: Doc, sheet: Sheet, r: Range): string[] {
  const valueAt = valueAtFn(doc, sheet);
  const texts: string[] = [];
  for (let c = r.c1; c <= r.c2; c++) texts.push(headerText(valueAt(r.r1, c), doc.wb.styles.get(sheet.styleIdAt(r.r1, c)).numFmt));
  return texts;
}

/** Turns a table back into ordinary cells (the data stays; the colors and buttons go). */
export function removeTable(tx: Transaction, doc: Doc, sheet: Sheet, table: TableDef): void {
  // Formulas that name the table get the cells it covered, so they keep working
  rewriteFormulas(tx, doc.wb, (f, fs, r, c) => expandStructured(f, doc.wb, fs, r, c, table));
  tx.setProp(
    sheet,
    'tables',
    sheet.tables.filter((t) => t.id !== table.id),
    true,
  );
  if (Object.keys(table.filter).length) refreshFilters(tx, doc, sheet);
}

export function renameTable(tx: Transaction, wb: Workbook, sheet: Sheet, table: TableDef, name: string): void {
  const problem = tableNameProblem(name, wb, table);
  if (problem) throw new Error(problem);
  const next = name.trim();
  if (next === table.name) return;
  putTable(tx, sheet, { ...table, name: next });
  rewriteFormulas(tx, wb, (f) => renameTableInFormula(f, table.name, next));
}

export function setTableLook(tx: Transaction, sheet: Sheet, table: TableDef, look: { color?: TableColor; banded?: boolean }): void {
  const next = { ...table, ...look };
  if (look.color && look.color !== table.color) delete next.fileStyle;
  putTable(tx, sheet, next, false);
}

export function setTableFilter(tx: Transaction, doc: Doc, sheet: Sheet, table: TableDef, conds: Record<number, FilterCond>): void {
  putTable(tx, sheet, { ...table, filter: conds }, false);
  refreshFilters(tx, doc, sheet);
}

/** Calls `fn` for every formula in the workbook (and conditional formats); changed text is saved. */
function rewriteFormulas(tx: Transaction, wb: Workbook, fn: (f: string, sheet: Sheet, row: number, col: number) => string): void {
  for (const s of wb.sheets) {
    for (const [key, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const r = Math.floor(key / MAX_COLS);
      const c = key % MAX_COLS;
      const f = fn(cell.f, s, r, c);
      if (f !== cell.f) tx.setCell(s, r, c, { ...cell, f });
    }
    if (s.cf.some((rule) => rule.type === 'expression')) {
      const cf = s.cf.map((rule) => {
        if (rule.type !== 'expression') return rule;
        const at = rule.ranges[0] ?? { r1: 0, c1: 0 };
        const formula = fn(rule.formula, s, at.r1, at.c1);
        return formula === rule.formula ? rule : { ...rule, formula };
      });
      if (cf.some((rule, i) => rule !== s.cf[i])) tx.setProp(s, 'cf', cf);
    }
  }
}

// ---------- sorting ----------

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
const NUMERIC_TEXT = /^\s*([-+])?\s*[$€£¥]?\s*([-+])?(\d{1,3}(?:,\d{3})+|\d+)?(\.\d+)?\s*(%)?\s*$/;

/** A number written as text ("1,200", "$5.50", "-3", "12%"), or null. */
function numberInText(s: string): number | null {
  const m = NUMERIC_TEXT.exec(s);
  if (!m || (!m[3] && !m[4])) return null;
  let n = parseFloat(`${m[3] ?? '0'}`.replace(/,/g, '') + (m[4] ?? ''));
  if (m[1] === '-' || m[2] === '-') n = -n;
  return m[5] ? n / 100 : n;
}

const isBlank = (v: Scalar) => v === null || (typeof v === 'string' && v.trim() === '');

/** Order for sorting: numbers (also typed as text), then text, then TRUE/FALSE, then errors. */
export function smartCompare(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar): [number, number | string] => {
    if (typeof v === 'number') return [0, v];
    if (typeof v === 'string') {
      const n = numberInText(v);
      return n !== null ? [0, n] : [1, v];
    }
    if (typeof v === 'boolean') return [2, v ? 1 : 0];
    return [3, 0];
  };
  const [ra, va] = rank(a);
  const [rb, vb] = rank(b);
  if (ra !== rb) return ra - rb;
  if (ra === 1) return collator.compare(va as string, vb as string);
  return (va as number) - (vb as number);
}

/** Compares two rows' sort values; empty cells go last whichever way it sorts. */
function compareKeys(a: Scalar[], b: Scalar[], keys: TableSortKey[]): number {
  for (let i = 0; i < keys.length; i++) {
    const x = a[i];
    const y = b[i];
    if (isBlank(x) || isBlank(y)) {
      if (isBlank(x) !== isBlank(y)) return isBlank(x) ? 1 : -1;
      continue;
    }
    const c = smartCompare(x, y);
    if (c) return keys[i].desc ? -c : c;
  }
  return 0;
}

/** What one click on a column's arrow does next. */
export function nextSort(table: TableDef, col: number, opts: { additive?: boolean; stale?: boolean } = {}): TableSortKey[] | 'original' {
  const keys = table.sort;
  const idx = keys.findIndex((k) => k.col === col);
  if (!opts.additive) {
    if (idx === 0) {
      // Changed since it was sorted: sort the same way again
      if (opts.stale) return keys;
      if (!keys[0].desc) return [{ col, desc: true }];
      return table.order ? 'original' : [{ col, desc: false }];
    }
    return [{ col, desc: false }];
  }
  if (idx < 0) return [...keys, { col, desc: false }];
  if (!keys[idx].desc) return keys.map((k, i) => (i === idx ? { col, desc: true } : k));
  const rest = keys.filter((_, i) => i !== idx);
  return rest.length ? rest : table.order ? 'original' : [];
}

/**
 * Sorts a table's rows (the header stays on top), or puts them back in the order they were in
 * before the first sort. Formulas move with their rows.
 */
export function sortTable(tx: Transaction, doc: Doc, sheet: Sheet, tableId: string, keys: TableSortKey[] | 'original'): void {
  const t = tableById(sheet, tableId);
  if (!t) return;
  const top = t.range.r1 + 1;
  const n = t.range.r2 - top + 1;
  if (n < 1) return;
  const order = t.order && t.order.length === n ? t.order : Array.from({ length: n }, (_, i) => i);
  const sortKeys = keys === 'original' ? [] : keys.filter((k) => k.col >= t.range.c1 && k.col <= t.range.c2);
  const rows = Array.from({ length: n }, (_, i) => {
    const r = top + i;
    const cells: (Cell | undefined)[] = [];
    for (let c = t.range.c1; c <= t.range.c2; c++) cells.push(sheet.getCell(r, c));
    return { i, r, cells, rank: order[i], keys: sortKeys.map((k) => doc.engine.getValue(sheet, r, k.col)) };
  });
  if (keys === 'original') rows.sort((a, b) => a.rank - b.rank || a.i - b.i);
  else rows.sort((a, b) => compareKeys(a.keys, b.keys, sortKeys) || a.i - b.i);
  rows.forEach((row, i) => {
    const target = top + i;
    if (row.r === target) return;
    row.cells.forEach((cell, j) => {
      const c = t.range.c1 + j;
      tx.setCell(sheet, target, c, cell?.f !== undefined ? { ...cell, f: translateFormula(cell.f, target - row.r, 0) } : cell);
    });
  });
  putTable(tx, sheet, { ...t, sort: sortKeys, order: rows.map((r) => r.rank) }, false);
  if (Object.keys(t.filter).length || sheet.filter) refreshFilters(tx, doc, sheet);
}

/** Whether rows changed so they're no longer in the table's sort order (a new row, an edited value). */
export function sortIsStale(doc: Doc, sheet: Sheet, table: TableDef): boolean {
  if (!table.sort.length) return false;
  let prev: Scalar[] | null = null;
  for (let r = table.range.r1 + 1; r <= table.range.r2; r++) {
    const cur = table.sort.map((k) => doc.engine.getValue(sheet, r, k.col));
    if (prev && compareKeys(prev, cur, table.sort) > 0) return true;
    prev = cur;
  }
  return false;
}

/** Whether rows are in a different order than before the first sort. */
export function canRestoreOrder(table: TableDef): boolean {
  return !!table.order && table.order.some((v, i) => i > 0 && v < table.order![i - 1]);
}

/** What a column mostly holds, for menu labels ("Smallest to largest", "Oldest to newest", "A to Z"). */
export function columnKind(doc: Doc, sheet: Sheet, table: TableDef, col: number): 'number' | 'date' | 'text' {
  let nums = 0;
  let dates = 0;
  let texts = 0;
  const last = Math.min(table.range.r2, table.range.r1 + 300);
  for (let r = table.range.r1 + 1; r <= last; r++) {
    const v = doc.engine.getValue(sheet, r, col);
    if (isBlank(v)) continue;
    if (typeof v === 'number') {
      if (/[dmy]/i.test((doc.numFmtAt(sheet, r, col) ?? '').replace(/"[^"]*"|\[[^\]]*\]/g, ''))) dates++;
      else nums++;
    } else if (typeof v === 'string' && numberInText(v) !== null) nums++;
    else texts++;
  }
  if (dates > nums && dates >= texts) return 'date';
  return nums > texts ? 'number' : 'text';
}

export function sortLabel(kind: 'number' | 'date' | 'text', desc: boolean): string {
  if (kind === 'number') return desc ? 'Largest to smallest' : 'Smallest to largest';
  if (kind === 'date') return desc ? 'Newest to oldest' : 'Oldest to newest';
  return desc ? 'Z to A' : 'A to Z';
}

// ---------- growing as you type, and the header row ----------

/**
 * After an edit: tables grow over rows typed just below them and columns typed just right of them,
 * and their column names follow the header cells (renaming one updates the formulas that use it).
 */
export function afterEdit(tx: Transaction, doc: Doc): void {
  if (!doc.wb.hasTables) return;
  const bySheet = new Map<Sheet, { r: number; c: number; added: boolean }[]>();
  for (const ch of tx.changes) {
    if (ch.k !== 'cell' || !ch.sheet.tables.length) continue;
    let list = bySheet.get(ch.sheet);
    if (!list) bySheet.set(ch.sheet, (list = []));
    const key = ch.key;
    list.push({ r: Math.floor(key / MAX_COLS), c: key % MAX_COLS, added: hasContent(ch.after) && !hasContent(ch.before) });
  }
  for (const [sheet, changes] of bySheet) {
    for (const id of sheet.tables.map((t) => t.id)) {
      const t = tableById(sheet, id);
      if (t) updateTable(tx, doc, sheet, t, changes);
    }
  }
}

function updateTable(tx: Transaction, doc: Doc, sheet: Sheet, t: TableDef, changes: { r: number; c: number; added: boolean }[]): void {
  const { r1, c1, c2 } = t.range;
  let { r2 } = t.range;
  const blocked = (r: Range) => sheet.tables.some((o) => o.id !== t.id && rangesIntersect(o.range, r)) || sheet.merges.some((m) => rangesIntersect(m, r));

  // Rows typed just below (a pasted block can add several)
  let bottom = r2;
  for (;;) {
    const row = bottom + 1;
    const typed = changes.filter((ch) => ch.r === row && ch.c >= c1 && ch.c <= c2 && ch.added);
    if (!typed.length || blocked({ r1: row, c1, r2: row, c2 }) || typed.some((ch) => isTotal(sheet, t, ch.r, ch.c))) break;
    bottom = row;
  }
  // Columns typed just right
  let right = c2;
  for (;;) {
    const col = right + 1;
    if (col >= MAX_COLS) break;
    const typed = changes.filter((ch) => ch.c === col && ch.r >= r1 && ch.r <= bottom && ch.added);
    if (!typed.length || blocked({ r1, c1: col, r2: bottom, c2: col })) break;
    right = col;
  }
  let next: TableDef = t;
  if (bottom > r2 || right > c2) {
    const range = { r1, c1, r2: bottom, c2: right };
    if (bottom > r2) {
      fillNewRows(tx, sheet, t, bottom);
      growColumnRefs(tx, doc.wb, sheet, t.range, bottom);
      if (t.order) {
        let max = Math.max(-1, ...t.order);
        next = { ...next, order: [...t.order, ...Array.from({ length: bottom - r2 }, () => ++max)] };
      }
    }
    next = { ...next, range };
    r2 = bottom;
  }
  // Column names follow the header cells (an empty one is Column4, a repeat gets a number). The
  // names live in the table: the cells keep exactly what was typed.
  const headerEdited = changes.some((ch) => ch.r === r1 && ch.c >= c1 && ch.c <= next.range.c2);
  if (headerEdited || right > c2) {
    const names = columnNames(headerTexts(doc, sheet, next.range));
    names.forEach((name, i) => {
      const old = t.columns[i];
      if (old !== undefined && old !== name) renameColumnRefs(tx, doc.wb, next, old, name);
    });
    next = { ...next, columns: names };
  }
  if (next !== t) putTable(tx, sheet, next);
}

/** A total under a table (=SUM(C2:C20) of its rows) isn't a new row of it. */
function isTotal(sheet: Sheet, t: TableDef, r: number, c: number): boolean {
  const f = sheet.getCell(r, c)?.f;
  if (f === undefined) return false;
  let sums = false;
  mapRefs(f, (ref) => {
    if (ref.sheet && ref.sheet.toLowerCase() !== sheet.name.toLowerCase()) return undefined;
    if (ref.r2 > ref.r1 && ref.r1 <= t.range.r2 && ref.r2 >= t.range.r1 + 1 && ref.c1 <= t.range.c2 && ref.c2 >= t.range.c1) sums = true;
    return undefined;
  });
  return sums || /\[#?(all|data)?\]|\b\w+\[[^@]/i.test(f);
}

/** New rows get the formulas of formula columns ("calculated columns") and the formats of the row above. */
function fillNewRows(tx: Transaction, sheet: Sheet, t: TableDef, bottom: number): void {
  const last = t.range.r2;
  for (let c = t.range.c1; c <= t.range.c2; c++) {
    const above = sheet.getCell(last, c);
    if (!above) continue;
    const calculated = above.f !== undefined && isCalculatedColumn(sheet, t, c);
    for (let r = last + 1; r <= bottom; r++) {
      const cell = sheet.getCell(r, c);
      if (calculated && !hasContent(cell)) tx.setCell(sheet, r, c, { ...(cell ?? {}), f: translateFormula(above.f!, r - last, 0), s: cell?.s ?? above.s });
      else if (cell?.s === undefined && above.s !== undefined) tx.patchCell(sheet, r, c, { s: above.s });
    }
  }
}

/** A column whose rows all have the same formula (relative to their row). */
function isCalculatedColumn(sheet: Sheet, t: TableDef, c: number): boolean {
  const from = Math.max(t.range.r1 + 1, t.range.r2 - 50);
  let prev: string | undefined;
  for (let r = from; r <= t.range.r2; r++) {
    const f = sheet.getCell(r, c)?.f;
    if (f === undefined) return false;
    if (prev !== undefined && translateFormula(prev, 1, 0) !== f) return false;
    prev = f;
  }
  return true;
}

/** References to a whole table column (or the whole table) grow when the table gets more rows. */
function growColumnRefs(tx: Transaction, wb: Workbook, sheet: Sheet, old: Range, bottom: number): void {
  const fits = (r1: number, r2: number, a: number, b: number) => (r1 === old.r1 || r1 === old.r1 + 1) && r2 === old.r2 && a >= old.c1 && b <= old.c2;
  for (const s of wb.sheets) {
    for (const [key, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const r = Math.floor(key / MAX_COLS);
      const c = key % MAX_COLS;
      // A formula in the new rows keeps its references (it would include itself)
      if (s === sheet && r > old.r2 && r <= bottom && c >= old.c1 && c <= old.c2) continue;
      const f = mapRefs(cell.f, (ref) => {
        const target = ref.sheet ?? s.name;
        if (target.toLowerCase() !== sheet.name.toLowerCase() || ref.kind !== 'range' || !fits(ref.r1, ref.r2, ref.c1, ref.c2)) return undefined;
        return { ...ref, r2: bottom };
      });
      if (f !== cell.f) tx.setCell(s, r, c, { ...cell, f });
    }
  }
  const grow = (x: Range) => (fits(x.r1, x.r2, x.c1, x.c2) ? { ...x, r2: bottom } : x);
  if (sheet.cf.some((rule) => rule.ranges.some((x) => fits(x.r1, x.r2, x.c1, x.c2))))
    tx.setProp(
      sheet,
      'cf',
      sheet.cf.map((rule) => ({ ...rule, ranges: rule.ranges.map(grow) })),
    );
  if (sheet.validations.some((dv) => dv.ranges.some((x) => fits(x.r1, x.r2, x.c1, x.c2))))
    tx.setProp(
      sheet,
      'validations',
      sheet.validations.map((dv) => ({ ...dv, ranges: dv.ranges.map(grow) })),
    );
  for (const s of wb.sheets) {
    if (!s.charts.length) continue;
    const growText = (text: string | undefined) =>
      text === undefined
        ? undefined
        : mapRefs(text, (ref) => ((ref.sheet ?? s.name).toLowerCase() === sheet.name.toLowerCase() && ref.kind === 'range' && fits(ref.r1, ref.r2, ref.c1, ref.c2) ? { ...ref, r2: bottom } : undefined));
    const charts = s.charts.map((ch) => {
      const series = ch.series.map((se) => ({ ...se, categories: growText(se.categories), values: growText(se.values) ?? se.values }));
      return series.some((se, i) => se.categories !== ch.series[i].categories || se.values !== ch.series[i].values) ? { ...ch, series } : ch;
    });
    if (charts.some((ch, i) => ch !== s.charts[i])) tx.setProp(s, 'charts', charts);
  }
}

function renameColumnRefs(tx: Transaction, wb: Workbook, t: TableDef, from: string, to: string): void {
  rewriteFormulas(tx, wb, (f, s, r, c) => {
    if (!f.includes('[')) return f;
    const inside = tableAt(s, r, c)?.id === t.id;
    return renameColumnInFormula(f, t.name, inside, from, to);
  });
}

// ---------- rows and columns inserted or deleted ----------

/** Tables after inserting or deleting rows or columns (count < 0 deletes). */
export function tablesAfterStructure(tables: TableDef[], axis: 'row' | 'col', at: number, count: number): TableDef[] {
  const ch = { sheetName: '', axis, at, count };
  // Where an index goes (null: deleted)
  const shift = (i: number): number | null => (count > 0 ? (i >= at ? i + count : i) : i < at ? i : i < at - count ? null : i + count);
  /** Applies the change to a list with one entry per row or column, starting at `first`. */
  const splice = <T>(list: T[], first: number, last: number, filler: (k: number) => T): T[] => {
    const out = [...list];
    if (count > 0 && at > first && at <= last) out.splice(at - first, 0, ...Array.from({ length: count }, (_, k) => filler(k)));
    else if (count < 0) {
      const from = Math.max(at, first);
      const to = Math.min(at - count, last + 1);
      if (to > from) out.splice(from - first, to - from);
    }
    return out;
  };
  const out: TableDef[] = [];
  for (const t of tables) {
    const range = adjustRange(t.range, ch);
    if (!range) continue;
    if (axis === 'row') {
      // Deleting the header row ends the table (its data stays)
      if (count < 0 && t.range.r1 >= at && t.range.r1 < at - count) continue;
      if (range.r2 <= range.r1) range.r2 = range.r1 + 1;
      let order = t.order;
      if (order) {
        const top = t.range.r1 + 1;
        // New rows inside go between their neighbours in the original order
        order = splice(order, top, t.range.r2, (k) => {
          const before = order![at - top - 1] ?? -1;
          const after = order![at - top] ?? before + 1;
          return before + ((after - before) * (k + 1)) / (count + 1);
        });
      }
      out.push({ ...t, range, order });
    } else {
      const columns = splice(t.columns, t.range.c1, t.range.c2, () => '');
      const sort = t.sort.flatMap((k) => {
        const col = shift(k.col);
        return col === null ? [] : [{ ...k, col }];
      });
      const filter: Record<number, FilterCond> = {};
      for (const [k, v] of Object.entries(t.filter)) {
        const col = shift(+k);
        if (col !== null) filter[col] = v;
      }
      out.push({ ...t, range, columns, sort, filter });
    }
  }
  return out;
}

/** Header cells for columns inserted inside tables (they get Column3…). */
export function nameNewColumns(tx: Transaction, sheet: Sheet): void {
  const tables = sheet.tables.map((t) => {
    if (!t.columns.some((c) => !c)) return t;
    const texts = t.columns.map((name, i) => name || headerText(sheet.getCell(t.range.r1, t.range.c1 + i)?.v ?? null));
    const names = columnNames(texts);
    writeHeaders(tx, sheet, t.range, names);
    return { ...t, columns: names };
  });
  if (tables.some((t, i) => t !== sheet.tables[i])) tx.setProp(sheet, 'tables', tables, true);
}

/** Copies a sheet's tables for a duplicated sheet, with new names; formulas on the copy use them. */
export function copyTables(wb: Workbook, from: Sheet, to: Sheet): void {
  const renames: [string, string][] = [];
  const taken = (n: string) => !!wb.table(n) || !!wb.getName(n) || to.tables.some((x) => x.name.toLowerCase() === n.toLowerCase());
  to.tables = [];
  for (const t of from.tables) {
    const base = t.name.replace(/\d+$/, '') || 'Table';
    let i = 1;
    while (taken(`${base}${i}`)) i++;
    const name = `${base}${i}`;
    renames.push([t.name, name]);
    to.tables.push({ ...t, id: 't' + Math.random().toString(36).slice(2, 10), name, range: { ...t.range }, columns: [...t.columns], sort: t.sort.map((k) => ({ ...k })), order: t.order && [...t.order], filter: { ...t.filter } });
  }
  if (!renames.length) return;
  for (const [key, cell] of to.cells) {
    if (cell.f === undefined) continue;
    let f = cell.f;
    for (const [a, b] of renames) f = renameTableInFormula(f, a, b);
    if (f !== cell.f) to.cells.set(key, { ...cell, f });
  }
}

/** A short description for the status bar: "Sales, sorted by Amount (largest first), then Date". */
export function describeSort(table: TableDef): string {
  if (!table.sort.length) return '';
  return table.sort
    .map((k, i) => {
      const name = table.columns[k.col - table.range.c1] ?? formatCell(table.range.r1, k.col);
      return `${i ? 'then ' : 'sorted by '}${name}${k.desc ? ' ↓' : ' ↑'}`;
    })
    .join(', ');
}
