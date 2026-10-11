// How the AI assistant sees a workbook: a compact overview (sheets, headers, formula patterns that
// link sheets, charts and rules) plus read tools for exact values, formulas and dependencies.
import { cellKey, colToName, formatCell, formatRange, keyCol, keyRow, MAX_COLS, MAX_ROWS, parseRange, type Range } from '../address';
import { describeRule as describeCf } from '../cf';
import type { Doc } from '../doc';
import { translateFormula } from '../formula/refs';
import { Matrix } from '../formula/types';
import { displayText } from '../ops/data';
import { describeRule as describeValidation, rangesText } from '../validation';
import { ErrorValue, type Scalar } from '../values';
import type { Sheet } from '../workbook';

export class ReadError extends Error {}

const MAX_TEXT = 160;

function clip(s: string, n = MAX_TEXT): string {
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}

export function findSheet(doc: Doc, name: string | undefined): Sheet {
  if (!name) return doc.wb.activeSheet;
  const s = doc.wb.sheetByName(name.replace(/^'|'$/g, ''));
  if (!s) throw new ReadError(`There's no sheet called "${name}". Sheets: ${doc.wb.sheets.map((x) => x.name).join(', ')}.`);
  return s;
}

/** "Data!B2:D9" or "B2:D9" (+ default sheet). Whole rows/columns are limited to the used part. */
export function resolve(doc: Doc, sheetName: string | undefined, addr: string): { sheet: Sheet; range: Range } {
  let a = (addr ?? '').trim().replace(/\$/g, '');
  let name = sheetName;
  const bang = a.lastIndexOf('!');
  if (bang > 0) {
    name = a.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    a = a.slice(bang + 1);
  }
  const sheet = findSheet(doc, name);
  const r = parseRange(a);
  if (!r) throw new ReadError(`"${addr}" isn't a cell or range address.`);
  const b = doc.engine.usedBounds(sheet);
  return {
    sheet,
    range: { r1: r.r1, c1: r.c1, r2: r.r2 === MAX_ROWS - 1 ? Math.max(r.r1, b.maxRow) : r.r2, c2: r.c2 === MAX_COLS - 1 ? Math.max(r.c1, b.maxCol) : r.c2 },
  };
}

function shown(doc: Doc, sheet: Sheet, r: number, c: number): string {
  return clip(displayText(doc, sheet, r, c).replace(/\t/g, ' ').replace(/\r?\n/g, ' ⏎ '));
}

/** The row that holds column headings (text over data), or -1. */
export function headerRow(doc: Doc, sheet: Sheet): number {
  const b = doc.engine.usedBounds(sheet);
  for (let r = 0; r <= Math.min(b.maxRow - 1, 20); r++) {
    let text = 0;
    let filled = 0;
    for (let c = 0; c <= Math.min(b.maxCol, 60); c++) {
      const v = doc.engine.getValue(sheet, r, c);
      if (v === null || v === '') continue;
      filled++;
      // Headings are short labels, not sentences
      if (typeof v === 'string' && v.length <= 40) text++;
    }
    if (filled >= 2 && text / filled >= 0.6) {
      // The next non-empty row should hold data
      for (let c = 0; c <= Math.min(b.maxCol, 60); c++) {
        const v = doc.engine.getValue(sheet, r + 1, c);
        if (typeof v === 'number' || (typeof v === 'string' && v !== '')) return r;
      }
    }
  }
  return -1;
}

/** Runs of the same formula down a column: "L5:L1204 =E5*K5*(1-F5)". */
export function formulaRuns(sheet: Sheet, limit = 60): string[] {
  const byCol = new Map<number, number[]>();
  for (const [key, cell] of sheet.cells) {
    if (cell.f === undefined) continue;
    const c = keyCol(key);
    const list = byCol.get(c) ?? [];
    list.push(keyRow(key));
    byCol.set(c, list);
  }
  const out: string[] = [];
  for (const c of [...byCol.keys()].sort((a, b) => a - b)) {
    const rows = byCol.get(c)!.sort((a, b) => a - b);
    let i = 0;
    let shownForCol = 0;
    while (i < rows.length) {
      const r0 = rows[i];
      const f0 = sheet.getCell(r0, c)!.f!;
      let j = i + 1;
      while (j < rows.length && rows[j] === rows[j - 1] + 1 && translateFormula(f0, rows[j] - r0, 0) === sheet.getCell(rows[j], c)!.f) j++;
      const range = j - i > 1 ? `${formatCell(r0, c)}:${formatCell(rows[j - 1], c)}` : formatCell(r0, c);
      if (shownForCol < 4) out.push(`${range} =${clip(f0, 200)}${j - i > 1 ? ` (same formula in ${j - i} rows)` : ''}`);
      else if (shownForCol === 4) out.push(`${colToName(c)}: more formulas…`);
      shownForCol++;
      i = j;
      if (out.length >= limit) return [...out, '…'];
    }
  }
  return out;
}

/**
 * Separate blocks of data on a sheet (cells that touch, including diagonally), largest first.
 * Lets the assistant see layouts like a ledger beside a summary table.
 */
export function dataBlocks(sheet: Sheet, limit = 8): Range[] {
  if (sheet.cells.size > 200_000) return [];
  const seen = new Set<number>();
  const blocks: (Range & { n: number })[] = [];
  for (const start of sheet.cells.keys()) {
    if (seen.has(start)) continue;
    const c0 = sheet.cells.get(start)!;
    if (c0.v === undefined && c0.f === undefined) continue;
    const b = { r1: keyRow(start), c1: keyCol(start), r2: keyRow(start), c2: keyCol(start), n: 0 };
    const stack = [start];
    seen.add(start);
    while (stack.length) {
      const k = stack.pop()!;
      const r = keyRow(k);
      const c = keyCol(k);
      b.n++;
      if (r < b.r1) b.r1 = r;
      if (r > b.r2) b.r2 = r;
      if (c < b.c1) b.c1 = c;
      if (c > b.c2) b.c2 = c;
      for (let dr = -1; dr <= 1; dr++) {
        for (let dc = -1; dc <= 1; dc++) {
          const rr = r + dr;
          const cc = c + dc;
          if (rr < 0 || cc < 0 || cc >= MAX_COLS) continue;
          const nk = cellKey(rr, cc);
          if (seen.has(nk)) continue;
          const cell = sheet.cells.get(nk);
          if (!cell || (cell.v === undefined && cell.f === undefined)) continue;
          seen.add(nk);
          stack.push(nk);
        }
      }
    }
    if (b.n >= 4) blocks.push(b);
  }
  blocks.sort((a, b) => b.n - a.n);
  return blocks.slice(0, limit).map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2 }));
}

function blockLine(doc: Doc, sheet: Sheet, r: Range): string {
  const heads: string[] = [];
  for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 11); c++) {
    const t = shown(doc, sheet, r.r1, c);
    if (t) heads.push(clip(t, 24));
  }
  return `${formatRange(r)} (${r.r2 - r.r1 + 1}×${r.c2 - r.c1 + 1}${heads.length ? `, first row: ${heads.join(' | ')}${r.c2 - r.c1 > 11 ? ' | …' : ''}` : ''})`;
}

/** The overview included with every request. */
export function workbookOverview(doc: Doc, selection?: { sheet: Sheet; ranges: Range[]; active: { row: number; col: number } }): string {
  const wb = doc.wb;
  const lines: string[] = [`Workbook "${doc.fileName}" with ${wb.sheets.length} sheet${wb.sheets.length === 1 ? '' : 's'}.`];
  for (const sheet of wb.sheets) {
    const b = doc.engine.usedBounds(sheet);
    lines.push('');
    if (b.maxRow < 0) {
      lines.push(`## Sheet "${sheet.name}" (empty)${sheet.hidden ? ' — hidden' : ''}`);
      continue;
    }
    lines.push(`## Sheet "${sheet.name}": used range A1:${formatCell(b.maxRow, b.maxCol)} (${b.maxRow + 1} rows × ${b.maxCol + 1} columns)${sheet.hidden ? ' — hidden' : ''}`);
    const h = headerRow(doc, sheet);
    if (h >= 0) {
      const heads: string[] = [];
      for (let c = 0; c <= Math.min(b.maxCol, 40); c++) {
        const t = shown(doc, sheet, h, c);
        if (t) heads.push(`${colToName(c)}=${clip(t, 40)}`);
      }
      lines.push(`Headers in row ${h + 1}: ${heads.join(', ')}${b.maxCol > 40 ? ', …' : ''}. Data rows ${h + 2}–${b.maxRow + 1}.`);
    }
    // Titles above the header (first few rows)
    for (let r = 0; r < Math.min(h >= 0 ? h : 3, 3); r++) {
      const t = shown(doc, sheet, r, 0);
      if (t) lines.push(`Row ${r + 1}: ${clip(t, 120)}`);
    }
    const blocks = dataBlocks(sheet);
    if (blocks.length > 1) lines.push(`Separate blocks of data: ${blocks.map((r) => blockLine(doc, sheet, r)).join('; ')}.`);
    const runs = formulaRuns(sheet, 25);
    if (runs.length) lines.push('Formulas: ' + runs.join('; '));
    if (sheet.freeze.rows || sheet.freeze.cols) lines.push(`Frozen: ${sheet.freeze.rows} rows, ${sheet.freeze.cols} columns.`);
    if (sheet.filter) lines.push(`Filter on ${formatRange(sheet.filter.range)}.`);
    for (const ch of sheet.charts) lines.push(`Chart: ${ch.type}${ch.title ? ` "${ch.title}"` : ''} of ${ch.series.map((s) => s.values).join(', ')}.`);
    for (const r of sheet.cf.slice(0, 8)) lines.push(`Conditional format on ${r.ranges.map(formatRange).join(',')}: ${describeCf(r)}.`);
    for (const v of sheet.validations.slice(0, 8)) lines.push(`Validation on ${rangesText(v)}: ${describeValidation(v)}.`);
  }
  if (wb.names.size) lines.push('', 'Named ranges: ' + [...wb.names.values()].slice(0, 30).map((n) => `${n.name}=${n.ref}`).join(', ') + '.');
  if (selection) {
    const a = selection.active;
    lines.push('', `The user is looking at sheet "${selection.sheet.name}"; selection ${selection.ranges.map(formatRange).join(',')}; active cell ${formatCell(a.row, a.col)} = ${JSON.stringify(shown(doc, selection.sheet, a.row, a.col))}.`);
  }
  return lines.join('\n');
}

// ---------- read tools ----------

export type ReadMode = 'values' | 'formulas' | 'both';

/** Cells as a tab-separated table with row numbers and column letters. */
export function readRange(doc: Doc, sheetName: string | undefined, addr: string, mode: ReadMode, maxCells: number): { text: string; cells: number; label: string } {
  const { sheet, range } = resolve(doc, sheetName, addr);
  const width = range.c2 - range.c1 + 1;
  const height = range.r2 - range.r1 + 1;
  const rowsAllowed = Math.max(1, Math.floor(maxCells / Math.max(1, width)));
  const showRows = Math.min(height, rowsAllowed);
  if (width > maxCells) throw new ReadError(`That range is ${width} columns wide; read fewer columns at a time.`);
  const out: string[] = [`${sheet.name}!${formatRange(range)}`, ['row', ...Array.from({ length: width }, (_, i) => colToName(range.c1 + i))].join('\t')];
  for (let r = range.r1; r < range.r1 + showRows; r++) {
    const cells: string[] = [String(r + 1)];
    for (let c = range.c1; c <= range.c2; c++) {
      const f = sheet.getCell(r, c)?.f;
      const text = shown(doc, sheet, r, c);
      if (mode === 'formulas') cells.push(f !== undefined ? '=' + clip(f, 200) : text);
      else if (mode === 'both' && f !== undefined) cells.push(`${text} [=${clip(f, 200)}]`);
      else cells.push(text);
    }
    out.push(cells.join('\t'));
  }
  if (showRows < height) out.push(`… ${height - showRows} more rows not shown (read up to ${maxCells} cells at a time; ask for rows ${range.r1 + showRows + 1}–${range.r2 + 1} next, or use evaluate for totals and counts).`);
  return { text: out.join('\n'), cells: showRows * width, label: `${sheet.name}!${formatRange(range)}` };
}

export function findCells(doc: Doc, query: string, opts: { sheet?: string; inFormulas?: boolean; exact?: boolean }, limit = 100): { text: string; hits: number } {
  const q = query.toLowerCase();
  const sheets = opts.sheet ? [findSheet(doc, opts.sheet)] : doc.wb.sheets;
  const hits: string[] = [];
  let total = 0;
  for (const sheet of sheets) {
    const keys = [...sheet.cells.keys()].sort((a, b) => a - b);
    for (const key of keys) {
      const cell = sheet.cells.get(key)!;
      const r = keyRow(key);
      const c = keyCol(key);
      const hay = opts.inFormulas ? (cell.f !== undefined ? '=' + cell.f : '') : displayText(doc, sheet, r, c);
      if (!hay) continue;
      const h = hay.toLowerCase();
      if (opts.exact ? h !== q : !h.includes(q)) continue;
      total++;
      if (hits.length < limit) hits.push(`${sheet.name}!${formatCell(r, c)}: ${clip(hay, 120)}`);
    }
  }
  return { text: hits.length ? hits.join('\n') + (total > hits.length ? `\n… ${total - hits.length} more matches` : '') : `No cells match "${query}".`, hits: total };
}

function valueText(v: Scalar): string {
  if (v === null) return '(empty)';
  if (v instanceof ErrorValue) return `${v.code}${v.message ? ` (${v.message})` : ''}`;
  if (typeof v === 'string') return JSON.stringify(clip(v, 300));
  return String(v);
}

export function describeCell(doc: Doc, sheetName: string | undefined, addr: string): string {
  const { sheet, range } = resolve(doc, sheetName, addr);
  const r = range.r1;
  const c = range.c1;
  const cell = sheet.getCell(r, c);
  const v = doc.engine.getValue(sheet, r, c);
  const lines = [`${sheet.name}!${formatCell(r, c)}`, `Value: ${valueText(v)}`, `Shown as: ${JSON.stringify(displayText(doc, sheet, r, c))}`];
  if (cell?.f !== undefined) lines.push(`Formula: =${cell.f}`);
  const fmt = doc.numFmtAt(sheet, r, c);
  if (fmt && fmt !== 'General') lines.push(`Number format: ${fmt}`);
  if (cell?.note) lines.push(`Note: ${clip(cell.note, 300)}`);
  if (cell?.f !== undefined) {
    const prec = doc.engine.precedentsOf(sheet, r, c).slice(0, 20);
    if (prec.length) {
      lines.push('Uses:');
      for (const p of prec) {
        const ps = doc.wb.sheetById(p.sheetId);
        if (!ps) continue;
        const n = (p.range.r2 - p.range.r1 + 1) * (p.range.c2 - p.range.c1 + 1);
        if (n === 1) lines.push(`  ${ps.name}!${formatCell(p.range.r1, p.range.c1)} = ${valueText(doc.engine.getValue(ps, p.range.r1, p.range.c1))}`);
        else lines.push(`  ${ps.name}!${formatRange(p.range)} (${n} cells)`);
      }
    }
  }
  const deps: string[] = [];
  outer: for (const s of doc.wb.sheets) {
    for (const [key, other] of s.cells) {
      if (other.f === undefined) continue;
      const rr = keyRow(key);
      const cc = keyCol(key);
      for (const p of doc.engine.precedentsOf(s, rr, cc)) {
        if (p.sheetId === sheet.id && r >= p.range.r1 && r <= p.range.r2 && c >= p.range.c1 && c <= p.range.c2) {
          deps.push(`${s.name}!${formatCell(rr, cc)}`);
          if (deps.length >= 30) break outer;
          break;
        }
      }
    }
  }
  if (deps.length) lines.push(`Used by: ${deps.join(', ')}${deps.length >= 30 ? ', …' : ''}`);
  return lines.join('\n');
}

export function evaluateFormula(doc: Doc, formula: string, sheetName?: string): string {
  const sheet = findSheet(doc, sheetName);
  const f = formula.trim().replace(/^=/, '');
  const v = doc.engine.evalText(f, sheet, 0, 0);
  if (v instanceof Matrix) {
    const rows = Math.min(v.rows, 50);
    const cols = Math.min(v.cols, 20);
    const out: string[] = [`Array ${v.rows}×${v.cols}${v.rows > rows || v.cols > cols ? ` (first ${rows}×${cols} shown)` : ''}:`];
    for (let r = 0; r < rows; r++) out.push(Array.from({ length: cols }, (_, c) => valueText(v.get(r, c))).join('\t'));
    return out.join('\n');
  }
  return valueText(v);
}

/** Details of one sheet: column summaries, formulas, rules and objects. */
export function describeSheet(doc: Doc, sheetName: string | undefined): string {
  const sheet = findSheet(doc, sheetName);
  const b = doc.engine.usedBounds(sheet);
  if (b.maxRow < 0) return `Sheet "${sheet.name}" is empty.`;
  const h = headerRow(doc, sheet);
  const first = h + 1;
  const lines = [`Sheet "${sheet.name}": A1:${formatCell(b.maxRow, b.maxCol)}. Header row: ${h >= 0 ? h + 1 : 'none found'}.`, 'Columns:'];
  for (let c = 0; c <= Math.min(b.maxCol, 60); c++) {
    let nums = 0;
    let texts = 0;
    let blanks = 0;
    let errors = 0;
    let formulas = 0;
    let min = Infinity;
    let max = -Infinity;
    let sum = 0;
    const samples: string[] = [];
    for (let r = Math.max(0, first); r <= b.maxRow; r++) {
      const v = doc.engine.getValue(sheet, r, c);
      if (sheet.getCell(r, c)?.f !== undefined) formulas++;
      if (v === null || v === '') blanks++;
      else if (typeof v === 'number') {
        nums++;
        min = Math.min(min, v);
        max = Math.max(max, v);
        sum += v;
      } else if (v instanceof ErrorValue) errors++;
      else texts++;
      if (samples.length < 3 && v !== null && v !== '') {
        const t = shown(doc, sheet, r, c);
        if (!samples.includes(t)) samples.push(t);
      }
    }
    const head = h >= 0 ? shown(doc, sheet, h, c) : '';
    if (!nums && !texts && !errors && !head) continue;
    const parts = [`${colToName(c)}${head ? ` "${clip(head, 40)}"` : ''}:`];
    if (nums) parts.push(`${nums} numbers (min ${fmtNum(min)}, max ${fmtNum(max)}, sum ${fmtNum(sum)})`);
    if (texts) parts.push(`${texts} text`);
    if (errors) parts.push(`${errors} errors`);
    if (blanks) parts.push(`${blanks} blank`);
    if (formulas) parts.push(`${formulas} formulas`);
    const fmt = doc.numFmtAt(sheet, Math.max(0, first), c);
    if (fmt && fmt !== 'General') parts.push(`format ${fmt}`);
    if (samples.length) parts.push(`e.g. ${samples.map((s) => JSON.stringify(clip(s, 30))).join(', ')}`);
    lines.push('  ' + parts.join(' '));
  }
  const runs = formulaRuns(sheet);
  if (runs.length) lines.push('Formulas:', ...runs.map((r) => '  ' + r));
  if (sheet.merges.length) lines.push(`Merged: ${sheet.merges.slice(0, 20).map(formatRange).join(', ')}`);
  for (const ch of sheet.charts) lines.push(`Chart: ${ch.type}${ch.title ? ` "${ch.title}"` : ''}; series ${ch.series.map((s) => `${s.name ?? ''} ${s.values}`.trim()).join('; ')}`);
  for (const r of sheet.cf) lines.push(`Conditional format ${r.ranges.map(formatRange).join(',')}: ${describeCf(r)}`);
  for (const v of sheet.validations) lines.push(`Validation ${rangesText(v)}: ${describeValidation(v)}`);
  return lines.join('\n');
}

function fmtNum(n: number): string {
  return Number.isInteger(n) ? String(n) : String(Number(n.toPrecision(10)));
}
