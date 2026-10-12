// Changes the AI assistant can propose, how they're previewed, and how they're applied.
// A proposal is tried on the real workbook inside Doc.dryRun (applied, measured, then reverted in
// the same instant), which gives an exact preview: the cells it writes, the formulas whose results
// change because of it, and any sheets, charts or rules it adds. Approved proposals are applied as
// a single undoable transaction.
import { colToName, formatCell, formatRange, keyCol, keyRow, MAX_COLS, MAX_ROWS, nameToCol, parseRange, type Range } from '../address';
import { createChart } from '../charts';
import type { Doc } from '../doc';
import type { Transaction } from '../history';
import { applyBorders, applyStyle, clearRanges, fillSelectionWithInput, setCellInput, setNote, type BorderPreset } from '../ops/cells';
import { applyFilter, displayText, sortRange } from '../ops/data';
import { addSheet, renameSheet } from '../ops/sheets';
import { restructure, setColumnWidth } from '../ops/structure';
import type { StylePatch } from '../styles';
import { newValidationId } from '../validation';
import { ErrorValue } from '../values';
import { Matrix } from '../formula/types';
import { validateSheetName, type Cell, type CFRuleSpec, type CFStyle, type ChartType, type Sheet } from '../workbook';

export type CellInput = string | number | boolean | null;

export type Change =
  | { op: 'set_values'; sheet?: string; start: string; values: CellInput[][] }
  | { op: 'fill_formula'; sheet?: string; range: string; formula: string }
  | { op: 'clear'; sheet?: string; range: string; what?: 'contents' | 'formats' | 'all' }
  | { op: 'insert_rows'; sheet?: string; before_row: number; count?: number }
  | { op: 'delete_rows'; sheet?: string; row: number; count?: number }
  | { op: 'insert_columns'; sheet?: string; before_column: string; count?: number }
  | { op: 'delete_columns'; sheet?: string; column: string; count?: number }
  | { op: 'add_sheet'; name: string; after?: string }
  | { op: 'rename_sheet'; sheet: string; name: string }
  | {
      op: 'format';
      sheet?: string;
      range: string;
      bold?: boolean;
      italic?: boolean;
      fill?: string | null;
      color?: string | null;
      number_format?: string;
      align?: 'left' | 'center' | 'right';
      wrap?: boolean;
      font_size?: number;
      border?: 'all' | 'outline' | 'bottom' | 'none';
    }
  | { op: 'column_width'; sheet?: string; columns: string; width: number | 'auto' }
  | { op: 'add_chart'; sheet?: string; type: ChartType; data: string; title?: string; at?: string }
  | { op: 'highlight'; sheet?: string; range: string; rule: HighlightRule; value?: string | number; value2?: string | number; color?: 'red' | 'yellow' | 'green' | 'blue' }
  | { op: 'sort'; sheet?: string; range: string; by: { column: string; descending?: boolean }[]; header?: boolean }
  | { op: 'filter'; sheet?: string; range: string }
  | { op: 'freeze'; sheet?: string; rows: number; columns?: number }
  | { op: 'add_note'; sheet?: string; cell: string; text: string }
  | { op: 'dropdown'; sheet?: string; range: string; choices: string[] };

export type HighlightRule = 'duplicates' | 'unique' | 'greater_than' | 'less_than' | 'between' | 'equal' | 'text_contains' | 'formula' | 'blanks' | 'errors' | 'top' | 'bottom';

export class ChangeError extends Error {}

const MAX_WRITE = 100_000;

/** Excel's standard highlight styles (light fill, dark text). */
const COLORS: Record<string, CFStyle> = {
  red: { fill: '#ffc7ce', color: '#9c0006' },
  yellow: { fill: '#ffeb9c', color: '#9c5700' },
  green: { fill: '#c6efce', color: '#006100' },
  blue: { fill: '#ddebf7', color: '#1f4e78' },
};

// ---------- addresses ----------

/** `home` is the sheet that changes without a sheet name apply to (where the user was when it started). */
function sheetOf(doc: Doc, home: Sheet, name: string | undefined): Sheet {
  if (!name) return home;
  const s = doc.wb.sheetByName(name.replace(/^'|'$/g, ''));
  if (!s) throw new ChangeError(`There's no sheet called "${name}".`);
  return s;
}

/** A range like "B2:D9", "C:C", "Data!A1:B5" (the sheet in the address wins). */
function rangeOf(doc: Doc, home: Sheet, sheetName: string | undefined, addr: string): { sheet: Sheet; range: Range } {
  let a = addr.trim().replace(/\$/g, '');
  let name = sheetName;
  const bang = a.lastIndexOf('!');
  if (bang > 0) {
    name = a.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
    a = a.slice(bang + 1);
  }
  const sheet = sheetOf(doc, home, name);
  const range = parseRange(a);
  if (!range) throw new ChangeError(`"${addr}" isn't a cell or range address.`);
  return { sheet, range };
}

/** Full rows/columns limited to the used part of the sheet (for writes and fills). */
function usedPart(doc: Doc, sheet: Sheet, r: Range): Range {
  const b = doc.engine.usedBounds(sheet);
  return {
    r1: r.r1,
    c1: r.c1,
    r2: r.r2 === MAX_ROWS - 1 ? Math.max(r.r1, b.maxRow) : r.r2,
    c2: r.c2 === MAX_COLS - 1 ? Math.max(r.c1, b.maxCol) : r.c2,
  };
}

function size(r: Range): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

function col(letters: string): number {
  const c = nameToCol(letters.replace(/\$/g, '').trim().split(':')[0].replace(/\d+/g, ''));
  if (c < 0) throw new ChangeError(`"${letters}" isn't a column letter.`);
  return c;
}

function where(sheet: Sheet, r: Range | string): string {
  return `${sheet.name}!${typeof r === 'string' ? r : formatRange(r)}`;
}

// ---------- applying ----------

/** Applies one change inside a transaction. Returns a plain-English note for the preview. */
/** Fields each kind of change needs. */
const REQUIRED: Record<string, string[]> = {
  set_values: ['start', 'values'],
  fill_formula: ['range', 'formula'],
  clear: ['range'],
  insert_rows: ['before_row'],
  delete_rows: ['row'],
  insert_columns: ['before_column'],
  delete_columns: ['column'],
  add_sheet: ['name'],
  rename_sheet: ['sheet', 'name'],
  format: ['range'],
  column_width: ['columns', 'width'],
  add_chart: ['type', 'data'],
  highlight: ['range', 'rule'],
  sort: ['range', 'by'],
  filter: ['range'],
  freeze: ['rows'],
  add_note: ['cell', 'text'],
  dropdown: ['range', 'choices'],
};

/**
 * Accepts the field names models commonly use instead of ours (value for values, cell for start…),
 * and explains clearly what's missing instead of failing on it.
 */
export function normalizeChange(input: Change): Change {
  const c = { ...(input as unknown as Record<string, unknown>) };
  const first = (...keys: string[]) => keys.map((k) => c[k]).find((v) => v !== undefined && v !== null && v !== '');
  const firstCell = (v: unknown) => (typeof v === 'string' ? v.split(':')[0] : v);
  switch (c.op) {
    case 'set_values': {
      c.start ??= firstCell(first('cell', 'range', 'at'));
      if (c.values === undefined) {
        const v = first('value', 'formula', 'data');
        if (v !== undefined) c.values = Array.isArray(v) ? v : [[v]];
      }
      // One row written as a flat list
      if (Array.isArray(c.values) && c.values.length && !c.values.some(Array.isArray)) c.values = [c.values];
      break;
    }
    case 'fill_formula':
      c.range ??= first('cells', 'start', 'cell');
      c.formula ??= first('value', 'formulas');
      if (Array.isArray(c.formula)) c.formula = c.formula.flat()[0];
      break;
    case 'add_sheet':
      c.name ??= first('sheet', 'title');
      break;
    case 'column_width':
      c.columns ??= first('column', 'range');
      c.width ??= 'auto';
      break;
    case 'add_note':
      c.cell ??= first('start', 'range');
      c.text ??= first('note', 'value');
      break;
    case 'freeze':
      c.rows ??= 0;
      break;
    case 'dropdown':
      c.choices ??= first('values', 'items', 'options');
      break;
    case 'add_chart':
      c.data ??= first('range', 'source');
      c.type ??= 'column';
      break;
    default:
      c.range ??= first('cell', 'cells', 'start');
  }
  const op = String(c.op);
  if (!(op in REQUIRED)) throw new ChangeError(`Unknown change "${op}". Use one of: ${Object.keys(REQUIRED).join(', ')}.`);
  const missing = REQUIRED[op].filter((k) => c[k] === undefined || c[k] === null || c[k] === '');
  if (missing.length) throw new ChangeError(`A ${op} change needs ${missing.map((k) => `"${k}"`).join(' and ')}.`);
  for (const k of ['start', 'range', 'cell', 'formula', 'name', 'sheet', 'columns', 'data', 'before_column', 'column']) {
    if (c[k] !== undefined && typeof c[k] !== 'string') c[k] = String(c[k]);
  }
  return c as unknown as Change;
}

export function applyChange(tx: Transaction, doc: Doc, input: Change, home: Sheet = doc.wb.activeSheet): string {
  const wb = doc.wb;
  const c = normalizeChange(input);
  switch (c.op) {
    case 'set_values': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.start);
      if (!Array.isArray(c.values) || !c.values.length) throw new ChangeError('set_values needs a 2-D array of values.');
      const rows = c.values.length;
      const cols = Math.max(...c.values.map((r) => (Array.isArray(r) ? r.length : 1)));
      if (rows * cols > MAX_WRITE) throw new ChangeError(`That writes ${rows * cols} cells; split it into smaller steps.`);
      if (range.r1 + rows > MAX_ROWS || range.c1 + cols > MAX_COLS) throw new ChangeError('Those values go past the edge of the sheet.');
      c.values.forEach((row, i) => {
        (Array.isArray(row) ? row : [row]).forEach((v, j) => {
          const r = range.r1 + i;
          const k = range.c1 + j;
          if (v === null || v === undefined || v === '') {
            const cur = sheet.getCell(r, k);
            if (cur && (cur.v !== undefined || cur.f !== undefined)) tx.patchCell(sheet, r, k, { v: undefined, f: undefined, cached: undefined });
          } else if (typeof v === 'object') {
            throw new ChangeError('Each value in set_values must be text, a number, true/false or null (one value per cell).');
          } else if (typeof v === 'number' || typeof v === 'boolean') {
            if (typeof v === 'number' && !Number.isFinite(v)) throw new ChangeError('A value is not a finite number.');
            tx.patchCell(sheet, r, k, { v, f: undefined, cached: undefined });
          } else setCellInput(tx, doc, sheet, r, k, String(v));
        });
      });
      return `Writes ${rows * cols === 1 ? '1 cell' : `${rows * cols} cells`} at ${where(sheet, { r1: range.r1, c1: range.c1, r2: range.r1 + rows - 1, c2: range.c1 + cols - 1 })}.`;
    }
    case 'fill_formula': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      const r = usedPart(doc, sheet, range);
      if (size(r) > MAX_WRITE) throw new ChangeError('That range is too large to fill in one step.');
      const f = c.formula.trim().startsWith('=') ? c.formula.trim() : '=' + c.formula.trim();
      // A formula that returns a whole table (NFL.LEADERS, FILTER…) belongs in one cell: filled
      // into every cell, each copy would block the others (#SPILL!)
      if (size(r) > 1) {
        const v = doc.engine.evalText(f.slice(1), sheet, r.r1, r.c1);
        if (v instanceof Matrix && v.rows * v.cols > 1) {
          setCellInput(tx, doc, sheet, r.r1, r.c1, f);
          return `Puts ${f} in ${where(sheet, formatCell(r.r1, r.c1))} only: it returns a table (${v.rows} rows × ${v.cols} columns) that fills the cells below and to the right by itself.`;
        }
      }
      fillSelectionWithInput(tx, doc, sheet, [r], r.r1, r.c1, f);
      return `Fills ${where(sheet, r)} with ${f}${size(r) > 1 ? ' (adjusted for each cell)' : ''}.`;
    }
    case 'clear': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      clearRanges(tx, sheet, [range], c.what ?? 'contents');
      return `Clears ${c.what === 'all' ? 'everything' : c.what === 'formats' ? 'the formatting' : 'the contents'} in ${where(sheet, range)}.`;
    }
    case 'insert_rows':
    case 'delete_rows': {
      const sheet = sheetOf(doc, home, c.sheet);
      const row = Math.round(c.op === 'insert_rows' ? c.before_row : c.row) - 1;
      const n = Math.max(1, Math.round(c.count ?? 1));
      if (!(row >= 0 && row < MAX_ROWS)) throw new ChangeError('That row number is outside the sheet.');
      restructure(tx, wb, sheet, 'row', row, c.op === 'insert_rows' ? n : -n);
      return c.op === 'insert_rows' ? `Inserts ${n} row${n === 1 ? '' : 's'} above row ${row + 1} in ${sheet.name}.` : `Deletes row${n === 1 ? '' : 's'} ${row + 1}${n > 1 ? `–${row + n}` : ''} in ${sheet.name}.`;
    }
    case 'insert_columns':
    case 'delete_columns': {
      const sheet = sheetOf(doc, home, c.sheet);
      const at = col(c.op === 'insert_columns' ? c.before_column : c.column);
      const n = Math.max(1, Math.round(c.count ?? 1));
      restructure(tx, wb, sheet, 'col', at, c.op === 'insert_columns' ? n : -n);
      return c.op === 'insert_columns'
        ? `Inserts ${n} column${n === 1 ? '' : 's'} before column ${colToName(at)} in ${sheet.name}.`
        : `Deletes column${n === 1 ? '' : 's'} ${colToName(at)}${n > 1 ? `–${colToName(at + n - 1)}` : ''} in ${sheet.name}.`;
    }
    case 'add_sheet': {
      const problem = validateSheetName(c.name, wb);
      if (problem) throw new ChangeError(problem);
      const after = c.after ? wb.sheets.indexOf(sheetOf(doc, home, c.after)) + 1 : wb.sheets.length;
      addSheet(tx, wb, c.name.trim(), after);
      return `Adds a sheet named "${c.name.trim()}".`;
    }
    case 'rename_sheet': {
      const sheet = sheetOf(doc, home, c.sheet);
      const problem = validateSheetName(c.name, wb, sheet.id);
      if (problem) throw new ChangeError(problem);
      const old = sheet.name;
      renameSheet(tx, wb, sheet, c.name.trim());
      return `Renames sheet "${old}" to "${c.name.trim()}".`;
    }
    case 'format': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      const p: StylePatch = {};
      if (c.bold !== undefined) p.bold = c.bold || null;
      if (c.italic !== undefined) p.italic = c.italic || null;
      if (c.fill !== undefined) p.fill = c.fill;
      if (c.color !== undefined) p.color = c.color;
      if (c.number_format !== undefined) p.numFmt = c.number_format;
      if (c.align !== undefined) p.hAlign = c.align;
      if (c.wrap !== undefined) p.wrap = c.wrap || null;
      if (c.font_size !== undefined) p.fontSize = c.font_size;
      if (Object.keys(p).length) applyStyle(tx, wb, sheet, [range], p);
      if (c.border) applyBorders(tx, wb, sheet, range, (c.border === 'outline' ? 'outer' : c.border) as BorderPreset);
      return `Formats ${where(sheet, range)}${describeFormat(c)}.`;
    }
    case 'column_width': {
      const sheet = sheetOf(doc, home, c.sheet);
      const parts = c.columns.replace(/\$/g, '').split(':');
      const c1 = col(parts[0]);
      const c2 = col(parts[1] ?? parts[0]);
      for (let k = Math.min(c1, c2); k <= Math.max(c1, c2); k++) {
        let px: number;
        if (c.width === 'auto') {
          let longest = 0;
          const b = doc.engine.usedBounds(sheet);
          for (let r = 0; r <= Math.min(b.maxRow, 2000); r++) longest = Math.max(longest, displayText(doc, sheet, r, k).length);
          px = Math.max(48, Math.min(480, longest * 7 + 16));
        } else px = Math.round(Number(c.width) * 7 + 5);
        setColumnWidth(tx, sheet, [k], px);
      }
      return `Sets the width of column${c1 === c2 ? '' : 's'} ${c.columns} in ${sheet.name}${c.width === 'auto' ? ' to fit' : ''}.`;
    }
    case 'add_chart': {
      const data = rangeOf(doc, home, c.sheet, c.data);
      const target = c.sheet ? sheetOf(doc, home, c.sheet) : data.sheet;
      const at = c.at ? rangeOf(doc, home, target.name, c.at).range : { r1: data.range.r1, c1: Math.min(MAX_COLS - 1, data.range.c2 + 2) };
      const chart = createChart(doc, data.sheet, usedPart(doc, data.sheet, data.range), c.type, at.r1, at.c1);
      if (c.title) chart.title = c.title;
      tx.setProp(target, 'charts', [...target.charts, chart]);
      return `Adds a ${c.type} chart${c.title ? ` "${c.title}"` : ''} of ${where(data.sheet, data.range)} on ${target.name} at ${formatCell(at.r1, at.c1)}.`;
    }
    case 'highlight': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      const style = COLORS[c.color ?? 'red'] ?? COLORS.red;
      const v = (x: string | number | undefined) => (x === undefined ? '' : String(x));
      let spec: CFRuleSpec;
      switch (c.rule) {
        case 'duplicates':
          spec = { type: 'duplicate', style };
          break;
        case 'unique':
          spec = { type: 'unique', style };
          break;
        case 'greater_than':
          spec = { type: 'cellIs', operator: 'greaterThan', values: [v(c.value)], style };
          break;
        case 'less_than':
          spec = { type: 'cellIs', operator: 'lessThan', values: [v(c.value)], style };
          break;
        case 'between':
          spec = { type: 'cellIs', operator: 'between', values: [v(c.value), v(c.value2)], style };
          break;
        case 'equal':
          spec = { type: 'cellIs', operator: 'equal', values: [typeof c.value === 'string' && !/^[-\d.]+$/.test(c.value) && !c.value.startsWith('=') ? `"${c.value}"` : v(c.value)], style };
          break;
        case 'text_contains':
          spec = { type: 'text', operator: 'contains', text: v(c.value), style };
          break;
        case 'formula':
          spec = { type: 'expression', formula: v(c.value).replace(/^=/, ''), style };
          break;
        case 'blanks':
          spec = { type: 'blanks', style };
          break;
        case 'errors':
          spec = { type: 'errors', style };
          break;
        case 'top':
        case 'bottom':
          spec = { type: 'top', rank: Math.max(1, Math.round(Number(c.value) || 10)), percent: false, bottom: c.rule === 'bottom', style };
          break;
        default:
          throw new ChangeError(`Unknown highlight rule "${c.rule}".`);
      }
      tx.setProp(sheet, 'cf', [...sheet.cf, { ...spec, id: 'cf-ai-' + Math.random().toString(36).slice(2, 9), ranges: [usedPart(doc, sheet, range)] }]);
      return `Highlights ${describeHighlight(c)} in ${where(sheet, usedPart(doc, sheet, range))} (${c.color ?? 'red'}).`;
    }
    case 'sort': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      const r = usedPart(doc, sheet, range);
      const keys = c.by.map((b) => {
        const k = col(b.column);
        if (k < r.c1 || k > r.c2) throw new ChangeError(`Column ${b.column} isn't inside ${formatRange(r)}.`);
        return { col: k, desc: !!b.descending };
      });
      sortRange(tx, doc, sheet, r, keys, c.header !== false);
      return `Sorts ${where(sheet, r)} by ${c.by.map((b) => `column ${b.column.toUpperCase()}${b.descending ? ' (largest first)' : ''}`).join(', then ')}.`;
    }
    case 'filter': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      applyFilter(tx, doc, sheet, { range: usedPart(doc, sheet, range), conds: {} });
      return `Adds filter buttons to ${where(sheet, usedPart(doc, sheet, range))}.`;
    }
    case 'freeze': {
      const sheet = sheetOf(doc, home, c.sheet);
      const f = { rows: Math.max(0, Math.round(c.rows || 0)), cols: Math.max(0, Math.round(c.columns || 0)) };
      tx.setProp(sheet, 'freeze', f);
      return `Freezes ${f.rows ? `the top ${f.rows} row${f.rows === 1 ? '' : 's'}` : ''}${f.rows && f.cols ? ' and ' : ''}${f.cols ? `the first ${f.cols} column${f.cols === 1 ? '' : 's'}` : ''}${!f.rows && !f.cols ? 'nothing (unfreezes)' : ''} in ${sheet.name}.`;
    }
    case 'add_note': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.cell);
      setNote(tx, sheet, range.r1, range.c1, c.text);
      return `Adds a note to ${where(sheet, formatCell(range.r1, range.c1))}.`;
    }
    case 'dropdown': {
      const { sheet, range } = rangeOf(doc, home, c.sheet, c.range);
      const r = usedPart(doc, sheet, range);
      const items = c.choices.map((x) => String(x).trim()).filter(Boolean);
      if (!items.length) throw new ChangeError('A dropdown needs at least one choice.');
      tx.setProp(sheet, 'validations', [
        ...sheet.validations,
        { id: newValidationId(), ranges: [r], type: 'list', operator: 'between', formula1: `"${items.join(',').replace(/"/g, '""')}"`, allowBlank: true, dropdown: true, showError: true, errorStyle: 'stop', showPrompt: false },
      ]);
      return `Adds a dropdown (${items.join(', ')}) to ${where(sheet, r)}.`;
    }
  }
  throw new ChangeError(`Unknown change "${(c as { op: string }).op}".`);
}

function describeFormat(c: Extract<Change, { op: 'format' }>): string {
  const parts: string[] = [];
  if (c.bold) parts.push('bold');
  if (c.italic) parts.push('italic');
  if (c.fill) parts.push(`fill ${c.fill}`);
  if (c.fill === null) parts.push('no fill');
  if (c.color) parts.push(`text ${c.color}`);
  if (c.number_format) parts.push(`number format ${c.number_format}`);
  if (c.align) parts.push(`aligned ${c.align}`);
  if (c.wrap) parts.push('wrapped');
  if (c.font_size) parts.push(`${c.font_size} pt`);
  if (c.border) parts.push(c.border === 'none' ? 'no borders' : `${c.border} borders`);
  return parts.length ? `: ${parts.join(', ')}` : '';
}

function describeHighlight(c: Extract<Change, { op: 'highlight' }>): string {
  switch (c.rule) {
    case 'duplicates':
      return 'duplicate values';
    case 'unique':
      return 'unique values';
    case 'greater_than':
      return `values greater than ${c.value}`;
    case 'less_than':
      return `values less than ${c.value}`;
    case 'between':
      return `values between ${c.value} and ${c.value2}`;
    case 'equal':
      return `values equal to ${c.value}`;
    case 'text_contains':
      return `text containing "${c.value}"`;
    case 'formula':
      return `cells where =${String(c.value ?? '').replace(/^=/, '')} is TRUE`;
    case 'blanks':
      return 'blank cells';
    case 'errors':
      return 'errors';
    case 'top':
      return `the top ${c.value ?? 10} values`;
    case 'bottom':
      return `the bottom ${c.value ?? 10} values`;
  }
}

// ---------- preview ----------

export interface CellDiff {
  sheet: string;
  cell: string;
  before: string;
  after: string;
  beforeFormula?: string;
  afterFormula?: string;
  /** A formula that returns a table: the cells it fills, like "A2:E12". */
  spill?: string;
}

export interface Preview {
  /** Plain-English list of what the proposal does. */
  notes: string[];
  /** Cells the proposal writes (values or formulas), with their shown results. */
  edits: CellDiff[];
  editCount: number;
  /** Cells whose formatting changes only. */
  formatCount: number;
  /** Formula results elsewhere that change because of the edits. */
  recalcs: CellDiff[];
  recalcCount: number;
  /** Formulas the change writes that would show an error (#SPILL!, #NAME?, #REF!…). */
  errors: { sheet: string; cell: string; code: string; message?: string }[];
  errorCount: number;
  /** Rows or columns move, so formula results elsewhere aren't compared. */
  structural: boolean;
  /** Set when the proposal couldn't be applied (nothing changes). */
  error?: string;
}

const SHOW = 400;
const FORMULA_MISTAKES = new Set(['#SPILL!', '#NAME?', '#REF!', '#VALUE!', '#CALC!', '#CIRC!', '#NULL!']);

function formulaCells(doc: Doc): Map<string, { sheet: Sheet; key: number }> {
  const out = new Map<string, { sheet: Sheet; key: number }>();
  for (const s of doc.wb.sheets) for (const [key, c] of s.cells) if (c.f !== undefined) out.set(`${s.id}:${key}`, { sheet: s, key });
  return out;
}

function shown(doc: Doc, sheet: Sheet, key: number): string {
  return displayText(doc, sheet, keyRow(key), keyCol(key));
}

/** True while some of the preview's results are still downloading (live data). */
export function previewLoading(p: Preview): boolean {
  return [...p.edits, ...p.recalcs].some((d) => d.after.startsWith('#BUSY!'));
}

/** Tries the changes and reports exactly what they would do, without changing anything. */
export function previewChanges(doc: Doc, changes: Change[]): Preview {
  // Results of every formula now, to spot the ones that change as a side effect
  const before = new Map<string, string>();
  for (const [id, { sheet, key }] of formulaCells(doc)) before.set(id, shown(doc, sheet, key));
  const notes: string[] = [];
  const home = doc.wb.activeSheet;
  let result: Preview;
  try {
    result = doc.dryRun(
      (tx) => {
        for (const c of changes) notes.push(applyChange(tx, doc, c, home));
      },
      (tx) => {
        const structural = tx.structural;
        const edits: CellDiff[] = [];
        let editCount = 0;
        let formatCount = 0;
        const errors: Preview['errors'] = [];
        let errorCount = 0;
        const edited = new Set<string>();
        for (const ch of tx.changes) {
          if (ch.k !== 'cell') continue;
          const id = `${ch.sheet.id}:${ch.key}`;
          if (edited.has(id)) continue;
          edited.add(id);
          const a = ch.sheet.cells.get(ch.key);
          const b = ch.before;
          if (sameCell(a, b)) continue;
          const contentChanged = b?.f !== a?.f || b?.v !== a?.v;
          if (!contentChanged) {
            formatCount++;
            continue;
          }
          editCount++;
          if (a?.f !== undefined) {
            const v = doc.engine.getValue(ch.sheet, keyRow(ch.key), keyCol(ch.key));
            // Only errors that mean the formula itself is wrong (not #N/A from a lookup, #DIV/0! from
            // missing data, or live data still loading or offline)
            if (v instanceof ErrorValue && FORMULA_MISTAKES.has(v.code)) {
              errorCount++;
              if (errors.length < 12) errors.push({ sheet: ch.sheet.name, cell: formatCell(keyRow(ch.key), keyCol(ch.key)), code: v.code, message: v.message });
            }
          }
          const sp = a?.f !== undefined ? doc.engine.spillAt(ch.sheet, keyRow(ch.key), keyCol(ch.key)) : null;
          const spill = sp && (sp.r2 > sp.r1 || sp.c2 > sp.c1) ? formatRange(sp) : undefined;
          if (spill) notes.push(`${ch.sheet.name}!${formatCell(keyRow(ch.key), keyCol(ch.key))} fills ${ch.sheet.name}!${spill} with a table (${sp!.r2 - sp!.r1 + 1} rows × ${sp!.c2 - sp!.c1 + 1} columns).`);
          if (edits.length < SHOW) {
            edits.push({
              sheet: ch.sheet.name,
              cell: formatCell(keyRow(ch.key), keyCol(ch.key)),
              before: b?.f !== undefined ? '' : rawText(b?.v),
              beforeFormula: b?.f !== undefined ? '=' + b.f : undefined,
              after: shown(doc, ch.sheet, ch.key),
              afterFormula: a?.f !== undefined ? '=' + a.f : undefined,
              spill,
            });
          }
        }
        const recalcs: CellDiff[] = [];
        let recalcCount = 0;
        if (!structural) {
          for (const [id, { sheet, key }] of formulaCells(doc)) {
            if (edited.has(id) || !before.has(id)) continue;
            const now = shown(doc, sheet, key);
            if (now === before.get(id)) continue;
            recalcCount++;
            if (recalcs.length < SHOW) recalcs.push({ sheet: sheet.name, cell: formatCell(keyRow(key), keyCol(key)), before: before.get(id)!, after: now });
          }
        }
        return { notes: [...notes], edits, editCount, formatCount, recalcs, recalcCount, errors, errorCount, structural };
      },
    );
  } catch (e) {
    return { notes: [], edits: [], editCount: 0, formatCount: 0, recalcs: [], recalcCount: 0, errors: [], errorCount: 0, structural: false, error: (e as Error).message || String(e) };
  }
  // How edited cells looked before (now that the trial has been reverted). When rows or columns
  // moved, addresses no longer line up with the original layout, so keep the raw values instead.
  if (result.structural) return result;
  for (const d of result.edits) {
    const s = doc.wb.sheetByName(d.sheet);
    const p = parseRange(d.cell);
    if (s && p) d.before = displayText(doc, s, p.r1, p.c1);
  }
  return result;
}

/** Same content, style and note (a rewrite that changes nothing). */
function sameCell(a: Cell | undefined, b: Cell | undefined): boolean {
  if (a === b) return true;
  const keys = new Set([...Object.keys(a ?? {}), ...Object.keys(b ?? {})]);
  keys.delete('cached');
  for (const k of keys) {
    const x = (a as Record<string, unknown> | undefined)?.[k];
    const y = (b as Record<string, unknown> | undefined)?.[k];
    if (x !== y && JSON.stringify(x) !== JSON.stringify(y)) return false;
  }
  return true;
}

function rawText(v: unknown): string {
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'object') return String(v);
  return String(v);
}

/** Applies approved changes as one undoable step. Returns the transaction (for one-click undo). */
export function applyChanges(doc: Doc, title: string, changes: Change[]): Transaction {
  const home = doc.wb.activeSheet;
  return doc.transact(`AI: ${title}`.slice(0, 80), (tx) => {
    for (const c of changes) applyChange(tx, doc, c, home);
  });
}

/**
 * Undoes an applied AI change: a normal undo if it's the latest edit, otherwise reverts the parts
 * nobody has changed since (cells edited afterwards are kept and counted in `skipped`).
 */
export function revertChanges(doc: Doc, tx: Transaction): { reverted: number; skipped: number } {
  if (doc.history.peekUndo() === tx) {
    doc.undo();
    return { reverted: tx.changes.length, skipped: 0 };
  }
  let result = { reverted: 0, skipped: 0 };
  doc.transact(`Undo ${tx.label}`.slice(0, 80), (t) => {
    result = tx.revertInto(t);
  });
  return result;
}
