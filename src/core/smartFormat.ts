// Advanced format: looks at a block of data, works out what each part is (title, header, totals,
// dates, money, percentages, years, IDs, text…) and formats it in one undoable step.
// analyze() produces a plain FormatPlan, so a different "brain" (e.g. an AI assistant) can supply
// a plan later and reuse applyPlan().
import { cellKey, type Range } from './address';
import type { Doc } from './doc';
import type { Transaction } from './history';
import { parseInput } from './input';
import { formatValue, isDateFormat } from './numfmt';
import { applyBorders, applyStyle, setCellInput } from './ops/cells';
import { applyFilter } from './ops/data';
import { setColumnWidth } from './ops/structure';
import type { HAlign, StylePatch } from './styles';
import { ErrorValue, type Scalar } from './values';
import type { CFRule, Sheet } from './workbook';

export type SmartStyle = 'banded' | 'clean' | 'bold' | 'grid';
export const SMART_STYLES: SmartStyle[] = ['banded', 'clean', 'bold', 'grid'];
export const SMART_STYLE_NAMES: Record<SmartStyle, string> = { banded: 'Banded rows', clean: 'Clean lines', bold: 'Bold header', grid: 'Boxed grid' };

export interface SmartOptions {
  style: SmartStyle;
  /** Header color name: blue, green, gray, orange, purple, red (or a #hex). */
  color: string;
  header: 'auto' | 'always' | 'never';
  convertText: boolean;
  currency: string;
  dateFormat: string;
  negatives: 'minus' | 'red' | 'parens' | 'redParens';
  thousands: boolean;
  decimals: 'auto' | '0' | '1' | '2' | '3';
  boldTotals: boolean;
  autofit: boolean;
  freeze: boolean;
  filter: boolean;
}

export const DEFAULT_SMART_OPTIONS: SmartOptions = {
  style: 'banded',
  color: 'blue',
  header: 'auto',
  convertText: true,
  currency: '$',
  dateFormat: 'mmm d, yyyy',
  negatives: 'minus',
  thousands: true,
  decimals: 'auto',
  boldTotals: true,
  autofit: true,
  freeze: true,
  filter: false,
};

export type ColumnKind = 'empty' | 'text' | 'longText' | 'code' | 'boolean' | 'number' | 'integer' | 'currency' | 'percent' | 'percentNumber' | 'date' | 'datetime' | 'time' | 'year' | 'id' | 'zip' | 'phone';

export const KIND_NAMES: Record<ColumnKind, string> = {
  empty: 'empty',
  text: 'text',
  longText: 'long text',
  code: 'codes',
  boolean: 'yes/no',
  number: 'numbers',
  integer: 'whole numbers',
  currency: 'money',
  percent: 'percentages',
  percentNumber: 'percentages',
  date: 'dates',
  datetime: 'dates & times',
  time: 'times',
  year: 'years',
  id: 'IDs',
  zip: 'ZIP codes',
  phone: 'phone numbers',
};

export interface ColumnPlan {
  col: number;
  kind: ColumnKind;
  numFmt?: string;
  hAlign?: HAlign;
  wrap?: boolean;
  /** Cells holding numbers or dates as text that will be converted. */
  convert: { row: number; raw: string }[];
  width?: number;
}

export interface FormatPlan {
  range: Range;
  titleRow: number | null;
  headerRow: number | null;
  dataStart: number;
  dataEnd: number;
  totalRow: number | null;
  columns: ColumnPlan[];
}

const PALETTE: Record<string, { base: string; tint: string; line: string; band: string }> = {
  blue: { base: '#1d4ed8', tint: '#dbeafe', line: '#93c5fd', band: '#eff6ff' },
  green: { base: '#15803d', tint: '#dcfce7', line: '#86efac', band: '#f0fdf4' },
  gray: { base: '#374151', tint: '#e5e7eb', line: '#9ca3af', band: '#f3f4f6' },
  orange: { base: '#c2410c', tint: '#ffedd5', line: '#fdba74', band: '#fff7ed' },
  purple: { base: '#6d28d9', tint: '#ede9fe', line: '#c4b5fd', band: '#f5f3ff' },
  red: { base: '#b91c1c', tint: '#fee2e2', line: '#fca5a5', band: '#fef2f2' },
};

function paletteFor(color: string): (typeof PALETTE)[string] {
  if (PALETTE[color]) return PALETTE[color];
  if (/^#[0-9a-f]{6}$/i.test(color)) {
    const mix = (to: number, amt: number) => {
      const n = parseInt(color.slice(1), 16);
      const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((c) => Math.round(c + (to - c) * amt));
      return '#' + ch.map((c) => c.toString(16).padStart(2, '0')).join('');
    };
    return { base: mix(0, 0.25), tint: mix(255, 0.82), line: mix(255, 0.55), band: mix(255, 0.92) };
  }
  return PALETTE.blue;
}

// ---------- analysis ----------

const RX = {
  money: /\b(price|cost|amount|total|balance|revenue|sales|income|expenses?|salary|wages?|pay|payment|fee|fees|budget|budgeted|spent|paid|owed|owe|value|profit|loss|tax|rent|bill|debt|loan|deposit|withdrawal|net|gross|subtotal|charge|cash|fund|funds|spend|savings|actual|planned|cost basis|market value|dividend|bonus|tip|refund)\b|[$€£¥₹]|\b(usd|eur|gbp)\b/i,
  percent: /%|\b(percent|percentage|pct|rate|growth|change|margin|share|yield|apr|apy|ratio|interest|return|roi|weight|completion)\b/i,
  date: /\b(date|day|due|dob|birth|birthday|created|updated|modified|joined|start|end|deadline|when|paid on|posted|gameday|period)\b/i,
  time: /\b(time|clock|kickoff|hour)\b/i,
  year: /\b(year|yr|season|fy|vintage)\b/i,
  zip: /\b(zip|postal|postcode)\b/i,
  phone: /\b(phone|tel|telephone|mobile|cell|fax)\b/i,
  id: /(^|[\s_])(id|#|no\.?|code|sku|account|acct|order|invoice|ticket|ref|reference|employee|serial|upc|isbn)($|[\s_#])/i,
  total: /^(grand\s+)?(sub)?totals?\b|^sum\b|^overall\b/i,
};

interface ColStats {
  values: Scalar[];
  /** Every numeric value, including numbers stored as text. */
  numbers: number[];
  raw: (Scalar | undefined)[];
  fmts: (string | undefined)[];
  nums: number;
  textNums: number;
  textPct: number;
  textMoney: number;
  textDates: number;
  bools: number;
  texts: number;
  blanks: number;
  ints: number;
  fracs: number;
  maxDecimals: number;
  min: number;
  max: number;
  dateFmts: number;
  pctFmts: number;
  moneyFmts: number;
  avgTextLen: number;
  maxTextLen: number;
}

function decimalsOf(n: number): number {
  if (Number.isInteger(n)) return 0;
  const s = String(Number(n.toPrecision(12)));
  const i = s.indexOf('.');
  return i < 0 ? 0 : Math.min(6, s.length - i - 1);
}

function statsFor(doc: Doc, sheet: Sheet, col: number, r1: number, r2: number, skip: number | null): ColStats {
  const st: ColStats = { values: [], numbers: [], raw: [], fmts: [], nums: 0, textNums: 0, textPct: 0, textMoney: 0, textDates: 0, bools: 0, texts: 0, blanks: 0, ints: 0, fracs: 0, maxDecimals: 0, min: Infinity, max: -Infinity, dateFmts: 0, pctFmts: 0, moneyFmts: 0, avgTextLen: 0, maxTextLen: 0 };
  let textLen = 0;
  for (let r = r1; r <= r2; r++) {
    if (r === skip || sheet.isRowHidden(r)) continue;
    const v = doc.engine.getValue(sheet, r, col);
    const cell = sheet.cells.get(cellKey(r, col));
    const fmt = doc.numFmtAt(sheet, r, col);
    st.values.push(v);
    st.raw.push(cell?.f === undefined ? cell?.v : undefined);
    st.fmts.push(fmt);
    if (v === null || v === '') {
      st.blanks++;
      continue;
    }
    if (typeof v === 'number') {
      st.numbers.push(v);
      st.nums++;
      if (Number.isInteger(v)) st.ints++;
      else st.fracs++;
      st.maxDecimals = Math.max(st.maxDecimals, decimalsOf(v));
      st.min = Math.min(st.min, v);
      st.max = Math.max(st.max, v);
      if (fmt && isDateFormat(fmt)) st.dateFmts++;
      else if (fmt && /%/.test(fmt.replace(/"[^"]*"/g, ''))) st.pctFmts++;
      else if (fmt && /[$€£¥₹]|\[\$/.test(fmt)) st.moneyFmts++;
    } else if (typeof v === 'boolean') st.bools++;
    else if (typeof v === 'string') {
      const t = v.trim();
      if (/^(yes|no|y|n|true|false)$/i.test(t)) st.bools++;
      const p = parseInput(t);
      if (cell?.f === undefined && typeof p.value === 'number') {
        if (p.numFmt && isDateFormat(p.numFmt)) st.textDates++;
        else {
          st.textNums++;
          if (/%/.test(t)) st.textPct++;
          if (/[$€£¥₹]/.test(t)) st.textMoney++;
        }
        const n = p.value;
        st.numbers.push(n);
        if (Number.isInteger(n)) st.ints++;
        else st.fracs++;
        st.maxDecimals = Math.max(st.maxDecimals, decimalsOf(n));
        st.min = Math.min(st.min, n);
        st.max = Math.max(st.max, n);
      } else {
        st.texts++;
        textLen += t.length;
        st.maxTextLen = Math.max(st.maxTextLen, t.length);
      }
    } else if (v instanceof ErrorValue) st.texts++;
  }
  st.avgTextLen = st.texts ? textLen / st.texts : 0;
  return st;
}

function nonEmpty(doc: Doc, sheet: Sheet, r: number, c1: number, c2: number): Scalar[] {
  const out: Scalar[] = [];
  for (let c = c1; c <= c2; c++) {
    const v = doc.engine.getValue(sheet, r, c);
    if (v !== null && v !== '') out.push(v);
  }
  return out;
}

function isHeaderRow(doc: Doc, sheet: Sheet, r: number, range: Range): boolean {
  const vals = nonEmpty(doc, sheet, r, range.c1, range.c2);
  if (!vals.length || r >= range.r2) return false;
  if (!vals.every((v) => typeof v === 'string')) return false;
  const width = range.c2 - range.c1 + 1;
  // Header cells are short labels
  if (vals.some((v) => (v as string).length > 60)) return false;
  // A text label above numbers or dates is the clearest signal
  for (let c = range.c1; c <= range.c2; c++) {
    const top = doc.engine.getValue(sheet, r, c);
    if (typeof top !== 'string' || !top) continue;
    const st = statsFor(doc, sheet, c, r + 1, Math.min(range.r2, r + 200), null);
    const filled = st.values.length - st.blanks;
    if (filled && (st.nums + st.textNums + st.textDates + st.bools) / filled >= 0.6) return true;
  }
  // All-text tables: unique, non-repeating labels in most columns
  const lower = vals.map((v) => (v as string).trim().toLowerCase());
  if (new Set(lower).size !== lower.length || vals.length < Math.ceil(width / 2)) return false;
  const style = doc.wb.styles.get(sheet.styleIdAt(r, range.c1));
  if (style.bold) return true;
  for (let c = range.c1; c <= range.c2; c++) {
    const top = doc.engine.getValue(sheet, r, c);
    if (typeof top !== 'string') continue;
    for (let rr = r + 1; rr <= Math.min(range.r2, r + 200); rr++) if (doc.engine.getValue(sheet, rr, c) === top) return false;
  }
  return range.r2 - r >= 2;
}

function classify(st: ColStats, header: string): ColumnKind {
  const filled = st.values.length - st.blanks;
  if (!filled) return 'empty';
  const numeric = st.nums + st.textNums + st.textDates;
  const h = header.toLowerCase();
  if (st.bools / filled >= 0.7 && numeric / filled < 0.5) return 'boolean';
  if (numeric / filled < 0.7) {
    if (st.avgTextLen > 40) return 'longText';
    if (st.maxTextLen <= 4 && st.texts >= 2) return 'code';
    return 'text';
  }
  const allInts = st.fracs === 0;
  if (st.dateFmts / filled >= 0.5 || st.textDates / filled >= 0.5) {
    if (st.max < 1 && st.min >= 0) return 'time';
    // Dates with a time part
    const fractional = st.numbers.some((v) => !Number.isInteger(v));
    return fractional ? 'datetime' : 'date';
  }
  if (RX.date.test(h) && !RX.year.test(h) && st.min >= 18000 && st.max <= 80000) return st.fracs ? 'datetime' : 'date';
  if (RX.time.test(h) && st.min >= 0 && st.max < 1) return 'time';
  if (allInts && (RX.year.test(h) || (!RX.money.test(h) && st.min >= 1900 && st.max <= 2100 && filled >= 2 && !RX.id.test(h)))) return 'year';
  if (allInts && RX.zip.test(h) && st.max < 100000) return 'zip';
  if (allInts && RX.phone.test(h) && st.min >= 1e9 && st.max < 1e10) return 'phone';
  if (allInts && RX.id.test(h) && !RX.money.test(h)) return 'id';
  if (st.textPct / filled >= 0.5 || st.pctFmts / filled >= 0.5) return 'percent';
  if (st.textMoney / filled >= 0.5 || st.moneyFmts / filled >= 0.5) return 'currency';
  if (RX.percent.test(h) && !RX.money.test(h.replace(/%/g, ''))) return Math.abs(st.min) <= 1.5 && Math.abs(st.max) <= 1.5 ? 'percent' : 'percentNumber';
  if (RX.money.test(h)) return 'currency';
  return allInts ? 'integer' : 'number';
}

function moneyFormat(opts: SmartOptions, decimals: number): string {
  const sym = opts.currency ? `"${opts.currency}"` : '';
  const body = `${sym}${opts.thousands ? '#,##0' : '0'}${decimals ? '.' + '0'.repeat(decimals) : ''}`;
  switch (opts.negatives) {
    case 'red':
      return `${body};[Red]-${body}`;
    case 'parens':
      return `${body};(${body})`;
    case 'redParens':
      return `${body};[Red](${body})`;
    default:
      return body;
  }
}

function numberFormat(opts: SmartOptions, decimals: number): string {
  return `${opts.thousands ? '#,##0' : '0'}${decimals ? '.' + '0'.repeat(decimals) : ''}`;
}

function formatFor(kind: ColumnKind, st: ColStats, opts: SmartOptions): { numFmt?: string; hAlign?: HAlign; wrap?: boolean } {
  const fixed = opts.decimals === 'auto' ? null : Number(opts.decimals);
  switch (kind) {
    case 'currency': {
      const d = fixed ?? (st.fracs === 0 && Math.abs(st.max) >= 1000 ? 0 : 2);
      return { numFmt: moneyFormat(opts, d), hAlign: 'right' };
    }
    case 'percent': {
      // Show enough decimals that 12.5% doesn't round to 13%
      const scaled = st.numbers.map((v) => v * 100);
      const d = fixed ?? Math.min(2, Math.max(0, ...scaled.map(decimalsOf)));
      return { numFmt: `0${d ? '.' + '0'.repeat(d) : ''}%`, hAlign: 'right' };
    }
    case 'percentNumber': {
      const d = fixed ?? Math.min(2, st.maxDecimals);
      return { numFmt: `0${d ? '.' + '0'.repeat(d) : ''}"%"`, hAlign: 'right' };
    }
    case 'number':
      return { numFmt: numberFormat(opts, fixed ?? Math.min(3, Math.max(1, st.maxDecimals))), hAlign: 'right' };
    case 'integer':
      return { numFmt: numberFormat(opts, fixed ?? 0), hAlign: 'right' };
    case 'date':
      return { numFmt: opts.dateFormat, hAlign: 'right' };
    case 'datetime':
      return { numFmt: `${opts.dateFormat} h:mm AM/PM`, hAlign: 'right' };
    case 'time':
      return { numFmt: 'h:mm AM/PM', hAlign: 'right' };
    case 'year':
    case 'id':
      return { numFmt: '0', hAlign: kind === 'year' ? 'center' : 'left' };
    case 'zip':
      return { numFmt: '00000', hAlign: 'left' };
    case 'phone':
      return { numFmt: '(###) ###-####', hAlign: 'left' };
    case 'boolean':
    case 'code':
      return { hAlign: 'center' };
    case 'longText':
      return { hAlign: 'left', wrap: true };
    case 'text':
      return { hAlign: 'left' };
    default:
      return {};
  }
}

/** Expands the selection to the data it's about: a single cell becomes its whole table. */
export function targetRange(doc: Doc, sheet: Sheet, sel: Range, region: (r: number, c: number) => Range): Range {
  const b = doc.engine.usedBounds(sheet);
  const r: Range = { r1: sel.r1, c1: sel.c1, r2: Math.min(sel.r2, b.maxRow), c2: Math.min(sel.c2, b.maxCol) };
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
    const reg = region(sel.r1, sel.c1);
    // Include a title that sits one blank row above the table
    if (reg.r1 >= 2 && nonEmpty(doc, sheet, reg.r1 - 1, reg.c1, reg.c2).length === 0) {
      const t = nonEmpty(doc, sheet, reg.r1 - 2, reg.c1, reg.c2);
      if (t.length === 1 && typeof t[0] === 'string' && nonEmpty(doc, sheet, reg.r1 - 2, Math.max(0, reg.c1 - 1), reg.c2 + 1).length === 1) return { ...reg, r1: reg.r1 - 2 };
    }
    return reg;
  }
  // Trim empty edges
  const rowEmpty = (row: number) => nonEmpty(doc, sheet, row, r.c1, r.c2).length === 0;
  const colEmpty = (col: number) => {
    for (let row = r.r1; row <= r.r2; row++) {
      const v = doc.engine.getValue(sheet, row, col);
      if (v !== null && v !== '') return false;
    }
    return true;
  };
  while (r.r2 > r.r1 && rowEmpty(r.r2)) r.r2--;
  while (r.r1 < r.r2 && rowEmpty(r.r1)) r.r1++;
  while (r.c2 > r.c1 && colEmpty(r.c2)) r.c2--;
  while (r.c1 < r.c2 && colEmpty(r.c1)) r.c1++;
  return r;
}

export function analyze(doc: Doc, sheet: Sheet, range: Range, opts: SmartOptions): FormatPlan {
  let first = range.r1;
  let titleRow: number | null = null;
  // A lone label above a wider table is a title
  const top = nonEmpty(doc, sheet, first, range.c1, range.c2);
  let next = first + 1;
  while (next < range.r2 && nonEmpty(doc, sheet, next, range.c1, range.c2).length === 0) next++;
  if (range.r2 - range.r1 >= 2 && range.c2 > range.c1 && top.length === 1 && typeof top[0] === 'string' && nonEmpty(doc, sheet, next, range.c1, range.c2).length >= 2) {
    titleRow = first;
    first++;
    // Skip a blank spacer row under the title
    if (nonEmpty(doc, sheet, first, range.c1, range.c2).length === 0 && first < range.r2) first++;
  }
  let headerRow: number | null = null;
  if (opts.header === 'always') headerRow = first;
  else if (opts.header === 'auto' && isHeaderRow(doc, sheet, first, range)) headerRow = first;
  const dataStart = headerRow === null ? first : first + 1;
  const dataEnd = range.r2;

  let totalRow: number | null = null;
  if (dataEnd > dataStart) {
    const lastVals = nonEmpty(doc, sheet, dataEnd, range.c1, range.c2);
    const label = lastVals.find((v) => typeof v === 'string') as string | undefined;
    const sums = [];
    for (let c = range.c1; c <= range.c2; c++) {
      const f = sheet.cells.get(cellKey(dataEnd, c))?.f;
      if (f && /^\s*(SUM|SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) sums.push(c);
    }
    if ((label && RX.total.test(label.trim())) || sums.length) totalRow = dataEnd;
  }

  const columns: ColumnPlan[] = [];
  for (let c = range.c1; c <= range.c2; c++) {
    const header = headerRow !== null ? String(doc.engine.getValue(sheet, headerRow, c) ?? '') : '';
    const st = statsFor(doc, sheet, c, dataStart, dataEnd, totalRow);
    const kind = classify(st, header);
    const f = formatFor(kind, st, opts);
    const convert: ColumnPlan['convert'] = [];
    const numericKind = !['empty', 'text', 'longText', 'code', 'boolean'].includes(kind);
    if (opts.convertText && numericKind) {
      for (let r = dataStart; r <= dataEnd; r++) {
        const cell = sheet.cells.get(cellKey(r, c));
        if (!cell || cell.f !== undefined || typeof cell.v !== 'string') continue;
        const p = parseInput(cell.v.trim());
        if (typeof p.value === 'number') convert.push({ row: r, raw: cell.v.trim() });
      }
    }
    columns.push({ col: c, kind, ...f, convert });
  }
  return { range, titleRow, headerRow, dataStart, dataEnd, totalRow, columns };
}

// ---------- applying ----------

function estimateWidth(doc: Doc, sheet: Sheet, plan: FormatPlan, cp: ColumnPlan): number {
  let max = 0;
  const end = Math.min(plan.dataEnd, plan.dataStart + 500);
  for (let r = plan.dataStart; r <= end; r++) {
    const v = doc.engine.getValue(sheet, r, cp.col);
    if (v === null) continue;
    const text = typeof v === 'number' && cp.numFmt ? formatValue(v, cp.numFmt).text : String(v instanceof ErrorValue ? v.code : v);
    max = Math.max(max, text.length);
  }
  if (plan.totalRow !== null) {
    const v = doc.engine.getValue(sheet, plan.totalRow, cp.col);
    if (typeof v === 'number' && cp.numFmt) max = Math.max(max, formatValue(v, cp.numFmt).text.length + 1);
  }
  let px = max * 7.4 + 22;
  if (plan.headerRow !== null) {
    const hdr = String(doc.engine.getValue(sheet, plan.headerRow, cp.col) ?? '');
    // Bold header text, capped so long headers wrap instead of making huge columns
    px = Math.max(px, Math.min(hdr.length, 22) * 8.2 + 26);
  }
  if (cp.kind === 'longText') px = Math.min(px, 320);
  return Math.round(Math.max(48, Math.min(px, 380)));
}

/** Applies a plan as part of a transaction. */
export function applyPlan(tx: Transaction, doc: Doc, sheet: Sheet, plan: FormatPlan, opts: SmartOptions): void {
  const wb = doc.wb;
  const { range } = plan;
  const pal = paletteFor(opts.color);
  const firstTable = plan.headerRow ?? plan.dataStart;
  const table: Range = { r1: firstTable, c1: range.c1, r2: plan.dataEnd, c2: range.c2 };
  const data: Range = { r1: plan.dataStart, c1: range.c1, r2: plan.dataEnd, c2: range.c2 };

  // Convert numbers stored as text first, so formats apply to real numbers
  for (const cp of plan.columns) for (const cv of cp.convert) setCellInput(tx, doc, sheet, cv.row, cp.col, cv.raw);

  // Start from a clean slate for the things Advanced format controls
  const reset: StylePatch = { fill: null, color: null, bold: null, italic: null, borderTop: null, borderBottom: null, borderLeft: null, borderRight: null };
  applyStyle(tx, wb, sheet, [table], reset);
  // Remove earlier banding rules on this block
  const keepCf = sheet.cf.filter((r: CFRule) => !(r.type === 'expression' && r.formula.startsWith('MOD(ROW()') && r.ranges.some((x) => x.r1 === data.r1 && x.c1 === data.c1)));
  if (keepCf.length !== sheet.cf.length) tx.setProp(sheet, 'cf', keepCf);

  // Columns
  for (const cp of plan.columns) {
    const col: Range = { r1: plan.dataStart, c1: cp.col, r2: plan.dataEnd, c2: cp.col };
    const patch: StylePatch = { hAlign: cp.hAlign ?? null, wrap: cp.wrap ?? null, vAlign: cp.wrap ? 'top' : null };
    if (cp.numFmt) patch.numFmt = cp.numFmt;
    applyStyle(tx, wb, sheet, [col], patch);
  }

  // Title
  if (plan.titleRow !== null) {
    applyStyle(tx, wb, sheet, [{ r1: plan.titleRow, c1: range.c1, r2: plan.titleRow, c2: range.c1 }], (s) => ({ bold: true, fontSize: Math.max(16, s.fontSize ?? 11) }));
  }

  // Header
  if (plan.headerRow !== null) {
    for (const cp of plan.columns) {
      const hdr: StylePatch = { bold: true, vAlign: 'bottom', wrap: true, hAlign: cp.hAlign === 'right' ? 'right' : cp.hAlign === 'center' ? 'center' : 'left' };
      switch (opts.style) {
        case 'banded':
          Object.assign(hdr, { fill: pal.tint, borderBottom: { style: 'thin', color: pal.line } });
          break;
        case 'clean':
          Object.assign(hdr, { borderBottom: { style: 'medium', color: pal.base } });
          break;
        case 'bold':
          Object.assign(hdr, { fill: pal.base, color: '#ffffff' });
          break;
        case 'grid':
          Object.assign(hdr, { fill: '#e5e7eb' });
          break;
      }
      applyStyle(tx, wb, sheet, [{ r1: plan.headerRow, c1: cp.col, r2: plan.headerRow, c2: cp.col }], hdr);
    }
  }

  // Body
  const bodyEnd = plan.totalRow !== null ? plan.totalRow - 1 : plan.dataEnd;
  const body: Range = { r1: plan.dataStart, c1: range.c1, r2: bodyEnd, c2: range.c2 };
  if (bodyEnd >= plan.dataStart) {
    if (opts.style === 'banded' || opts.style === 'bold') {
      const rule = { type: 'expression', formula: `MOD(ROW()-${plan.dataStart + 1},2)=1`, style: { fill: opts.style === 'bold' ? pal.band : '#f1f5f9' }, id: 'cf' + Math.random().toString(36).slice(2, 9), ranges: [body] } as CFRule;
      tx.setProp(sheet, 'cf', [...sheet.cf, rule]);
    }
    if (opts.style === 'clean') applyBorders(tx, wb, sheet, body, 'horizontal', { style: 'thin', color: '#e5e7eb' });
    if (opts.style === 'grid') applyBorders(tx, wb, sheet, table, 'all', { style: 'thin', color: '#9ca3af' });
    if (opts.style === 'banded') applyBorders(tx, wb, sheet, table, 'outer', { style: 'thin', color: '#cbd5e1' });
  }

  // Totals
  if (plan.totalRow !== null && opts.boldTotals) {
    applyStyle(tx, wb, sheet, [{ r1: plan.totalRow, c1: range.c1, r2: plan.totalRow, c2: range.c2 }], { bold: true, fill: null, borderTop: { style: 'thin', color: opts.style === 'bold' ? pal.base : '#6b7280' }, borderBottom: { style: 'double', color: opts.style === 'bold' ? pal.base : '#6b7280' } });
  }

  // Column widths
  if (opts.autofit) {
    for (const cp of plan.columns) {
      if (cp.kind === 'empty') continue;
      const w = estimateWidth(doc, sheet, plan, cp);
      if (Math.abs(w - sheet.colWidth(cp.col)) > 2) setColumnWidth(tx, sheet, [cp.col], w);
    }
  }

  // Freeze the header and add filters
  if (plan.headerRow !== null) {
    if (opts.freeze && !sheet.freeze.rows && !sheet.freeze.cols && plan.headerRow < 30) tx.setProp(sheet, 'freeze', { rows: plan.headerRow + 1, cols: 0 });
    if (opts.filter && !sheet.filter) applyFilter(tx, doc, sheet, { range: { r1: plan.headerRow, c1: range.c1, r2: plan.dataEnd, c2: range.c2 }, conds: {} });
  }
}

/** A one-line, plain-English summary of what was detected. */
export function describePlan(plan: FormatPlan): string {
  const parts: string[] = [];
  if (plan.titleRow !== null) parts.push('title');
  parts.push(plan.headerRow !== null ? 'header row' : 'no header');
  if (plan.totalRow !== null) parts.push('totals row');
  const counts = new Map<string, number>();
  for (const c of plan.columns) if (c.kind !== 'empty') counts.set(KIND_NAMES[c.kind], (counts.get(KIND_NAMES[c.kind]) ?? 0) + 1);
  const cols = [...counts].map(([k, n]) => (n > 1 ? `${n} ${k}` : k));
  let s = `Found ${parts.join(', ')}; columns: ${cols.join(', ')}.`;
  const converted = plan.columns.reduce((n, c) => n + c.convert.length, 0);
  if (converted) s += ` Converted ${converted} number${converted === 1 ? '' : 's'} stored as text.`;
  return s;
}
