// Structured references: formulas that name a table and its columns instead of cell addresses, as
// in Excel. =SUM(Sales[Amount]) adds up the Amount column, however many rows Sales has; inside
// the table, [@Price]*[@Qty] uses this row's Price and Qty. The engine turns them into ordinary
// ranges before calculating (expandStructured), so everything else works with A1 references.
import { formatCell, quoteSheetName, type Range } from './address';
import type { Sheet, TableDef, Workbook } from './workbook';

type Special = '#all' | '#data' | '#headers' | '#totals' | '#this row';
const SPECIALS: Record<string, Special> = { '#all': '#all', '#data': '#data', '#headers': '#headers', '#totals': '#totals', '#this row': '#this row' };
const SPECIAL_TEXT: Record<Special, string> = { '#all': '#All', '#data': '#Data', '#headers': '#Headers', '#totals': '#Totals', '#this row': '#This Row' };

/** What a reference's brackets ask for: which rows (specials, or this row) and which columns. */
export interface Spec {
  items: Special[];
  /** [@Col] or [#This Row]. */
  thisRow: boolean;
  /** No columns = every column; two = a span [A]:[C]. */
  cols: string[];
}

/** A structured reference found in formula text, or a bare table name (=SUM(Sales)). */
export interface Found {
  start: number;
  end: number;
  /** The table named before the brackets, or null for [Col] inside a table. */
  table: string | null;
  /** The bracket part, or null for a bare name. */
  spec: Spec | null;
  text: string;
}

const IDENT_START = /[A-Za-z_\\À-￿]/;
const IDENT_CHAR = /[\w.\\À-￿]/;

function skipString(s: string, i: number): number {
  for (let k = i + 1; k < s.length; k++) {
    if (s[k] !== '"') continue;
    if (s[k + 1] === '"') k++;
    else return k + 1;
  }
  return s.length;
}

function skipQuoted(s: string, i: number): number {
  for (let k = i + 1; k < s.length; k++) {
    if (s[k] !== "'") continue;
    if (s[k + 1] === "'") k++;
    else return k + 1;
  }
  return s.length;
}

/** The index after the bracket that closes the one at `i`, or -1. A ' escapes the next character. */
function closeBracket(s: string, i: number): number {
  let depth = 0;
  for (let k = i; k < s.length; k++) {
    const c = s[k];
    if (c === "'") k++;
    else if (c === '[') depth++;
    else if (c === ']' && --depth === 0) return k + 1;
  }
  return -1;
}

const unescapeName = (s: string) => s.replace(/'(.)/g, '$1');
const escapeName = (s: string) => s.replace(/(['#[\]])/g, "'$1");

/** Reads "[Amount]", "[@Amount]", "[@[Unit Price]]", "[[#Headers],[A]:[C]]", "[#All]"… */
export function parseSpec(body: string): Spec | null {
  if (!body.startsWith('[') || !body.endsWith(']')) return null;
  const inner = body.slice(1, -1).trim();
  const spec: Spec = { items: [], thisRow: false, cols: [] };
  const addPart = (part: string): boolean => {
    const t = part.trim();
    if (t.startsWith('#')) {
      const sp = SPECIALS[t.toLowerCase().replace(/\s+/g, ' ')];
      if (!sp) return false;
      if (sp === '#this row') spec.thisRow = true;
      else spec.items.push(sp);
      return true;
    }
    spec.cols.push(unescapeName(t));
    return true;
  };
  // A list of bracketed parts: [#This Row],[Amount] or [A]:[C]
  const list = (s: string): boolean => {
    let k = 0;
    let span = false;
    while (k < s.length) {
      while (s[k] === ' ') k++;
      if (k >= s.length) break;
      if (s[k] !== '[') return false;
      const end = closeBracket(s, k);
      if (end < 0) return false;
      const part = s.slice(k + 1, end - 1);
      const before = spec.cols.length;
      if (!addPart(part)) return false;
      if (span && spec.cols.length !== before + 1) return false;
      span = false;
      k = end;
      while (s[k] === ' ') k++;
      if (s[k] === ',') k++;
      else if (s[k] === ':') {
        span = true;
        k++;
      } else if (k < s.length) return false;
    }
    return !span && spec.cols.length <= 2;
  };
  if (inner.startsWith('@')) {
    spec.thisRow = true;
    const rest = inner.slice(1).trim();
    if (!rest) return spec;
    if (rest.startsWith('[')) return list(rest) ? spec : null;
    spec.cols.push(unescapeName(rest));
    return spec;
  }
  if (inner.startsWith('[')) return list(inner) ? spec : null;
  if (!inner) return spec;
  return addPart(inner) ? spec : null;
}

/** The way Excel shows a reference's brackets: [Amount], [@Amount], [@[Unit Price]], [[#Headers],[Amount]]. */
export function formatSpec(spec: Spec, forFile = false): string {
  const col = (n: string) => `[${escapeName(n)}]`;
  const cols = spec.cols.length === 2 ? `${col(spec.cols[0])}:${col(spec.cols[1])}` : spec.cols.length === 1 ? col(spec.cols[0]) : '';
  if (spec.thisRow && !spec.items.length && !forFile) {
    if (!cols) return '[@]';
    if (spec.cols.length === 1 && /^[A-Za-z0-9_À-￿]+$/.test(spec.cols[0])) return `[@${spec.cols[0]}]`;
    return `[@${cols}]`;
  }
  const parts = spec.items.map((x) => `[${SPECIAL_TEXT[x]}]`);
  if (spec.thisRow) parts.unshift('[#This Row]');
  if (!parts.length) return cols || '[]';
  if (cols) parts.push(cols);
  if (parts.length === 1) return parts[0];
  return `[${parts.join(',')}]`;
}

/** Every structured reference (and bare word that might be a table name) in formula text. */
export function findStructured(text: string): Found[] {
  const out: Found[] = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (ch === '"') {
      i = skipString(text, i);
      continue;
    }
    if (ch === "'") {
      i = skipQuoted(text, i);
      continue;
    }
    if (IDENT_START.test(ch) && !(i > 0 && IDENT_CHAR.test(text[i - 1]))) {
      let j = i + 1;
      while (j < text.length && IDENT_CHAR.test(text[j])) j++;
      const name = text.slice(i, j);
      if (text[j] === '[') {
        const end = closeBracket(text, j);
        const spec = end > 0 ? parseSpec(text.slice(j, end)) : null;
        if (spec) {
          out.push({ start: i, end, table: name, spec, text: text.slice(i, end) });
          i = end;
          continue;
        }
      } else if (text[j] !== '(' && text[j] !== '!' && text[i - 1] !== '!') out.push({ start: i, end: j, table: name, spec: null, text: name });
      i = j;
      continue;
    }
    if (ch === '[') {
      const end = closeBracket(text, i);
      const spec = end > 0 ? parseSpec(text.slice(i, end)) : null;
      if (spec) {
        out.push({ start: i, end, table: null, spec, text: text.slice(i, end) });
        i = end;
        continue;
      }
    }
    i++;
  }
  return out;
}

/** Rewrites structured references; `fn` returns replacement text, or undefined to keep one. */
export function mapStructured(text: string, fn: (f: Found) => string | undefined): string {
  if (!text.includes('[') && !/[A-Za-z_]/.test(text)) return text;
  const found = findStructured(text);
  if (!found.length) return text;
  let out = '';
  let last = 0;
  for (const f of found) {
    const next = fn(f);
    if (next === undefined) continue;
    out += text.slice(last, f.start) + next;
    last = f.end;
  }
  return last ? out + text.slice(last) : text;
}

/** The table a cell is in (header or data rows). */
export function tableAt(sheet: Sheet, row: number, col: number): TableDef | undefined {
  for (const t of sheet.tables) if (row >= t.range.r1 && row <= t.range.r2 && col >= t.range.c1 && col <= t.range.c2) return t;
  return undefined;
}

function columnIndex(t: TableDef, name: string): number {
  const n = name.trim().toLowerCase();
  return t.columns.findIndex((c) => c.toLowerCase() === n);
}

/** The cells a reference stands for, or the error it gives. */
export function resolveStructured(
  wb: Workbook,
  sheet: Sheet,
  row: number,
  col: number,
  tableName: string | null,
  spec: Spec,
): { sheet: Sheet; range: Range } | '#REF!' | '#VALUE!' {
  const found = tableName !== null ? wb.table(tableName) : (() => {
    const t = tableAt(sheet, row, col);
    return t ? { sheet, table: t } : undefined;
  })();
  if (!found) return '#REF!';
  const { table: t } = found;
  let c1 = t.range.c1;
  let c2 = t.range.c2;
  if (spec.cols.length) {
    const a = columnIndex(t, spec.cols[0]);
    const b = spec.cols.length === 2 ? columnIndex(t, spec.cols[1]) : a;
    if (a < 0 || b < 0) return '#REF!';
    c1 = t.range.c1 + Math.min(a, b);
    c2 = t.range.c1 + Math.max(a, b);
  }
  const top = t.range.r1 + 1;
  let r1 = top;
  let r2 = t.range.r2;
  if (spec.thisRow) {
    // This row only works on the table's own rows
    if (found.sheet !== sheet || row < top || row > t.range.r2) return '#VALUE!';
    r1 = r2 = row;
  } else if (spec.items.includes('#totals')) return '#REF!';
  else if (spec.items.includes('#all')) r1 = t.range.r1;
  else if (spec.items.includes('#headers')) {
    r1 = t.range.r1;
    if (!spec.items.includes('#data')) r2 = t.range.r1;
  }
  return { sheet: found.sheet, range: { r1, c1, r2, c2 } };
}

/**
 * Formula text with structured references replaced by A1 references (for calculating). With
 * `only`, just the references to that table are replaced (when it stops being a table).
 */
export function expandStructured(text: string, wb: Workbook, sheet: Sheet, row: number, col: number, only?: TableDef): string {
  return mapStructured(text, (f) => {
    if (!f.spec) {
      // A bare table name stands for its data rows, unless it's also a defined name
      if (!f.table || wb.getName(f.table) || !wb.table(f.table)) return undefined;
    }
    if (only && (f.table ? f.table.toLowerCase() !== only.name.toLowerCase() : tableAt(sheet, row, col)?.id !== only.id)) return undefined;
    const r = resolveStructured(wb, sheet, row, col, f.table, f.spec ?? { items: [], thisRow: false, cols: [] });
    if (typeof r === 'string') return only ? undefined : r;
    const { range } = r;
    const prefix = `${quoteSheetName(r.sheet.name)}!`;
    const a = formatCell(range.r1, range.c1, true, true);
    if (range.r1 === range.r2 && range.c1 === range.c2) return prefix + a;
    return `${prefix}${a}:${formatCell(range.r2, range.c2, true, true)}`;
  });
}

/** Formula text after a table is renamed. */
export function renameTableInFormula(text: string, from: string, to: string): string {
  const lower = from.toLowerCase();
  return mapStructured(text, (f) => (f.table && f.table.toLowerCase() === lower ? to + (f.spec ? f.text.slice(f.table.length) : '') : undefined));
}

/**
 * Formula text after a column of `table` is renamed. `inside` says whether the formula is in that
 * table (where [Col] without a table name means this table).
 */
export function renameColumnInFormula(text: string, table: string, inside: boolean, from: string, to: string): string {
  const lowerTable = table.toLowerCase();
  const lower = from.toLowerCase();
  return mapStructured(text, (f) => {
    if (!f.spec || !f.spec.cols.some((c) => c.toLowerCase() === lower)) return undefined;
    if (f.table ? f.table.toLowerCase() !== lowerTable : !inside) return undefined;
    const spec = { ...f.spec, cols: f.spec.cols.map((c) => (c.toLowerCase() === lower ? to : c)) };
    return (f.table ?? '') + formatSpec(spec);
  });
}

/** For saving: Excel files always name the table and write [@Col] as [[#This Row],[Col]]. */
export function structuredForFile(text: string, wb: Workbook, sheet: Sheet, row: number, col: number): string {
  return mapStructured(text, (f) => {
    // A bare table name is written Sales[] (Excel can't read just Sales in a file)
    if (!f.spec) return f.table && !wb.getName(f.table) && wb.table(f.table) ? `${wb.table(f.table)!.table.name}[]` : undefined;
    const name = f.table ?? tableAt(sheet, row, col)?.name;
    if (!name) return undefined;
    if (f.table && !f.spec.thisRow) return undefined;
    return name + formatSpec(f.spec, true);
  });
}

/**
 * After opening a file, references read the way Excel shows them: Sales[] as Sales, and
 * Sales[[#This Row],[Qty]] inside Sales as [@Qty].
 */
export function structuredFromFile(text: string, sheet: Sheet, row: number, col: number): string {
  const here = tableAt(sheet, row, col);
  return mapStructured(text, (f) => {
    if (!f.spec || !f.table) return undefined;
    if (!f.spec.cols.length && !f.spec.items.length && !f.spec.thisRow) return f.table;
    if (!here || f.table.toLowerCase() !== here.name.toLowerCase() || !f.spec.thisRow || f.spec.items.length) return undefined;
    return formatSpec(f.spec);
  });
}
