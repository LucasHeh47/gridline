// .xlsx import/export using ExcelJS (plus our own chart writer/reader).
import type * as ExcelJSTypes from 'exceljs';
import { cellKey, formatCell, formatRange, keyCol, keyRow, MAX_COLS, MAX_ROWS, parseRange, type Range } from '../address';
import type { Doc } from '../doc';
import { normalizeFuncName, parseFormula, type Node } from '../formula/parser';
import { FUNCTIONS, XLFN_FUNCS, XLWS_FUNCS } from '../formula/functions';
import { Matrix, RangeRef } from '../formula/types';
import { tokenize, type Token } from '../formula/tokenizer';
import { applyTint, type BorderSide, type BorderStyle, type HAlign, type Style, type VAlign } from '../styles';
import { ErrorValue, parseErrorCode, type ErrorCode, type Scalar } from '../values';
import { Sheet, Workbook, type CFRule, type CFStyle, type CFValue, type Cell } from '../workbook';
import { prepareForExcelJS, readCharts, readNotes, writeCharts } from './xlsxCharts';
import { readMacros, writeMacros } from './xlsxMacros';
import { readValidations, writeValidations } from './xlsxValidation';

type ExcelJS = typeof ExcelJSTypes;

let excelPromise: Promise<ExcelJS> | null = null;
async function loadExcelJS(): Promise<ExcelJS> {
  if (!excelPromise) {
    excelPromise = import('exceljs').then((m) => ((m as unknown as { default?: ExcelJS }).default ?? m) as ExcelJS);
  }
  return excelPromise;
}

// ---------- colors ----------

const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF',
  'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366',
  'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080',
  '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00',
  'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

type XColor = Partial<ExcelJSTypes.Color> & { indexed?: number; tint?: number };

/** Converts an Excel color. `auto` marks colors that mean "automatic" (e.g. default black text). */
function fromColor(c: XColor | undefined, role: 'font' | 'fill' | 'border'): string | undefined {
  if (!c) return undefined;
  let hex: string | undefined;
  if (c.argb && /^[0-9A-Fa-f]{8}$/.test(c.argb)) hex = c.argb.slice(2);
  else if (c.argb && /^[0-9A-Fa-f]{6}$/.test(c.argb)) hex = c.argb;
  else if (c.theme !== undefined) {
    if (role === 'font' && c.theme === 1 && !c.tint) return undefined; // automatic text color
    hex = THEME[c.theme];
  } else if (c.indexed !== undefined) {
    if (c.indexed === 64 || c.indexed === 65) return undefined; // system colors
    if (role === 'font' && c.indexed === 8) return undefined;
    hex = INDEXED[c.indexed];
  }
  if (!hex) return undefined;
  let out = '#' + hex.toLowerCase();
  if (c.tint) out = applyTint(out, c.tint);
  if (role === 'font' && out === '#000000') return undefined;
  return out;
}

function toArgb(hex: string): string {
  return 'FF' + hex.replace('#', '').toUpperCase();
}

// ---------- styles ----------

const BORDER_STYLES = new Set<BorderStyle>([
  'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
]);

function fromBorder(b: Partial<ExcelJSTypes.Border> | undefined): BorderSide | undefined {
  if (!b || !b.style || !BORDER_STYLES.has(b.style as BorderStyle)) return undefined;
  return { style: b.style as BorderStyle, color: fromColor(b.color as XColor, 'border') };
}

function fromExcelStyle(st: Partial<ExcelJSTypes.Style> | undefined): Style {
  const s: Style = {};
  if (!st) return s;
  const f = st.font;
  if (f) {
    // The default font is left unset so it follows the app default
    if (f.name && f.name !== 'Calibri') s.fontName = f.name;
    if (f.size && f.size !== 11) s.fontSize = f.size;
    if (f.bold) s.bold = true;
    if (f.italic) s.italic = true;
    if (f.underline && f.underline !== 'none') s.underline = true;
    if (f.strike) s.strike = true;
    const color = fromColor(f.color as XColor, 'font');
    if (color) s.color = color;
  }
  const fill = st.fill;
  if (fill) {
    if (fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
      const c = fromColor(fill.fgColor as XColor, 'fill') ?? fromColor(fill.bgColor as XColor, 'fill');
      if (c) s.fill = c;
    } else if (fill.type === 'gradient' && fill.stops?.length) {
      const c = fromColor(fill.stops[0].color as XColor, 'fill');
      if (c) s.fill = c;
    }
  }
  if (st.numFmt && st.numFmt !== 'General') s.numFmt = st.numFmt;
  const a = st.alignment;
  if (a) {
    if (a.horizontal) s.hAlign = a.horizontal as HAlign;
    if (a.vertical) s.vAlign = (a.vertical === 'top' ? 'top' : a.vertical === 'bottom' ? 'bottom' : 'middle') as VAlign;
    if (a.wrapText) s.wrap = true;
    if (a.indent) s.indent = a.indent;
    if (typeof a.textRotation === 'number' && a.textRotation) s.rotation = a.textRotation;
  }
  const b = st.border;
  if (b) {
    const t = fromBorder(b.top);
    const r = fromBorder(b.right);
    const bt = fromBorder(b.bottom);
    const l = fromBorder(b.left);
    if (t) s.borderTop = t;
    if (r) s.borderRight = r;
    if (bt) s.borderBottom = bt;
    if (l) s.borderLeft = l;
  }
  return s;
}

function toExcelBorder(b: BorderSide | undefined): Partial<ExcelJSTypes.Border> | undefined {
  if (!b) return undefined;
  return { style: b.style, color: { argb: toArgb(b.color ?? '#000000') } };
}

function toExcelStyle(s: Style): Partial<ExcelJSTypes.Style> {
  const out: Partial<ExcelJSTypes.Style> = {};
  const font: Partial<ExcelJSTypes.Font> = {};
  if (s.fontName) font.name = s.fontName;
  if (s.fontSize) font.size = s.fontSize;
  if (s.bold) font.bold = true;
  if (s.italic) font.italic = true;
  if (s.underline) font.underline = true;
  if (s.strike) font.strike = true;
  if (s.color) font.color = { argb: toArgb(s.color) };
  if (Object.keys(font).length) out.font = { name: 'Calibri', size: 11, ...font };
  if (s.fill) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: toArgb(s.fill) } };
  if (s.numFmt) out.numFmt = s.numFmt;
  const al: Partial<ExcelJSTypes.Alignment> = {};
  if (s.hAlign && s.hAlign !== 'general') al.horizontal = s.hAlign as ExcelJSTypes.Alignment['horizontal'];
  if (s.vAlign) al.vertical = s.vAlign;
  if (s.wrap) al.wrapText = true;
  if (s.indent) al.indent = s.indent;
  if (s.rotation) al.textRotation = s.rotation;
  if (Object.keys(al).length) out.alignment = al;
  const border: Partial<ExcelJSTypes.Borders> = {};
  if (s.borderTop) border.top = toExcelBorder(s.borderTop);
  if (s.borderRight) border.right = toExcelBorder(s.borderRight);
  if (s.borderBottom) border.bottom = toExcelBorder(s.borderBottom);
  if (s.borderLeft) border.left = toExcelBorder(s.borderLeft);
  if (Object.keys(border).length) out.border = border;
  return out;
}

// ---------- formulas ----------

/** Removes _xlfn./_xlws. prefixes Excel uses in files. */
function fromFileFormula(f: string): string {
  return (
    f
      .replace(/_xlfn\.ANCHORARRAY\(([^()]+)\)/gi, '$1#')
      .replace(/_xlfn\.SINGLE\(([^()]+)\)/gi, '@$1')
      .replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xlpm\./gi, '')
      // _xludf. marks a function Excel didn't know when saving: some tools write built-ins (CONCAT, LET,
      // XLOOKUP…) that way, and functions from VBA code can end up so too. Either way the plain name is
      // what calculates (a built-in, or the workbook's VBA Function); saving writes it correctly.
      .replace(/_xludf\.([A-Za-z_][A-Za-z0-9_.]*)(?=\()/gi, (_m, name: string) => name)
      .replace(/^=/, '')
  );
}

/** Variable names declared by LET and LAMBDA (Excel writes them as _xlpm.name inside files). */
function lambdaParams(toks: Token[]): Set<string> {
  const names = new Set<string>();
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'func') continue;
    const fn = normalizeFuncName(t.text);
    if (fn !== 'LET' && fn !== 'LAMBDA') continue;
    let j = i + 1;
    while (j < toks.length && toks[j].type === 'ws') j++;
    if (toks[j]?.type !== 'lparen') continue;
    const args: Token[][] = [[]];
    let depth = 0;
    for (j = j + 1; j < toks.length; j++) {
      const k = toks[j];
      if (k.type === 'lparen' || k.type === 'lbrace') depth++;
      else if (k.type === 'rparen' || k.type === 'rbrace') {
        if (depth === 0) break;
        depth--;
      } else if (k.type === 'comma' && depth === 0) {
        args.push([]);
        continue;
      }
      args[args.length - 1].push(k);
    }
    args.forEach((a, idx) => {
      // LET(name1, value1, name2, value2, …, result); LAMBDA(param1, param2, …, body)
      if (idx === args.length - 1 || (fn === 'LET' && idx % 2 === 1)) return;
      const real = a.filter((x) => x.type !== 'ws');
      if (real.length === 1 && real[0].type === 'name') names.add(real[0].text.toLowerCase());
    });
  }
  return names;
}

/** Adds the _xlfn. prefixes newer functions need inside .xlsx files (and _xlpm. for LET/LAMBDA names). */
export function toFileFormula(f: string): string {
  const toks = tokenize(f);
  const params = lambdaParams(toks);
  const out = toks.map((t) => {
    // Without the prefix Excel can't read the formula and refuses to open the whole file
    if ((t.type === 'name' || t.type === 'func') && params.has(t.text.toLowerCase())) return '_xlpm.' + t.text;
    if (t.type !== 'func') return t.text;
    const name = normalizeFuncName(t.text);
    if (XLWS_FUNCS.has(name)) return '_xlfn._xlws.' + name;
    if (XLFN_FUNCS.has(name)) return '_xlfn.' + name;
    return t.text;
  });
  // In files, A1# is written _xlfn.ANCHORARRAY(A1) and @x is _xlfn.SINGLE(x); Excel can't open "A1#"
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type !== 'op' || (t.text !== '#' && t.text !== '@')) continue;
    if (t.text === '#') {
      let p = i - 1;
      while (p >= 0 && toks[p].type === 'ws') p--;
      const start = operandStart(toks, p);
      if (start < 0) continue;
      out[start] = '_xlfn.ANCHORARRAY(' + out[start];
      out[p] += ')';
      out[i] = '';
    } else {
      let n = i + 1;
      while (n < toks.length && toks[n].type === 'ws') n++;
      const end = operandEnd(toks, n);
      if (end < 0) continue;
      out[i] = '_xlfn.SINGLE(';
      out[end] += ')';
    }
  }
  return out.join('');
}

/** Index of the first token of the operand that ends at token p (a reference, or a parenthesized call). */
function operandStart(toks: Token[], p: number): number {
  if (p < 0) return -1;
  if (toks[p].type === 'ref') return p;
  if (toks[p].type !== 'rparen') return -1;
  let depth = 0;
  for (let i = p; i >= 0; i--) {
    if (toks[i].type === 'rparen') depth++;
    else if (toks[i].type === 'lparen' && --depth === 0) {
      let k = i - 1;
      while (k >= 0 && toks[k].type === 'ws') k--;
      return k >= 0 && toks[k].type === 'func' ? k : i;
    }
  }
  return -1;
}

/** Index of the last token of the operand that starts at token n. */
function operandEnd(toks: Token[], n: number): number {
  const t = toks[n];
  if (!t) return -1;
  if (t.type === 'ref' || t.type === 'name' || t.type === 'num') return n;
  let i = n;
  if (t.type === 'func') {
    i++;
    while (i < toks.length && toks[i].type === 'ws') i++;
  }
  if (toks[i]?.type !== 'lparen') return -1;
  let depth = 0;
  for (; i < toks.length; i++) {
    if (toks[i].type === 'lparen') depth++;
    else if (toks[i].type === 'rparen' && --depth === 0) return i;
  }
  return -1;
}

/** Does this formula do array math that older Excel would evaluate differently unless entered as an array formula? */
function needsArrayEntry(ast: Node | null): boolean {
  if (!ast) return false;
  const multi = (n: Node) => (n.t === 'ref' && n.ref.kind !== 'cell') || n.t === 'array' || n.t === 'range';
  const walk = (n: Node): boolean => {
    switch (n.t) {
      case 'bin':
        return multi(n.l) || multi(n.r) || walk(n.l) || walk(n.r);
      case 'unary':
      case 'pct':
        return multi(n.e) || walk(n.e);
      case 'func': {
        const def = FUNCTIONS.get(n.name);
        if (def?.lift) {
          const idx = def.lift === true ? n.args.map((_, i) => i) : def.lift;
          if (idx.some((i) => n.args[i] && multi(n.args[i]))) return true;
        }
        return n.args.some(walk);
      }
      default:
        return false;
    }
  };
  return walk(ast);
}

// ---------- values ----------

function fromResult(r: unknown): Scalar | undefined {
  if (r === undefined || r === null) return undefined;
  if (typeof r === 'number' || typeof r === 'string' || typeof r === 'boolean') return r;
  if (r instanceof Date) return dateToSerial1900(r);
  if (typeof r === 'object' && 'error' in (r as object)) {
    const code = parseErrorCode(String((r as { error: string }).error));
    return code ? new ErrorValue(code, 'This error was saved in the file.') : undefined;
  }
  return undefined;
}

function dateToSerial1900(d: Date): number {
  return d.getTime() / 86_400_000 + 25569;
}

function richTextOf(v: { richText?: { text: string }[] }): string {
  return (v.richText ?? []).map((t) => t.text).join('');
}

function cfValueFrom(c: ExcelJSTypes.Cvfo | undefined, fallback: CFValue): CFValue {
  if (!c) return fallback;
  switch (c.type) {
    case 'min':
    case 'autoMin':
      return { type: 'min' };
    case 'max':
    case 'autoMax':
      return { type: 'max' };
    case 'num':
      return { type: 'num', value: Number(c.value ?? 0) };
    case 'percent':
      return { type: 'percent', value: Number(c.value ?? 0) };
    case 'percentile':
      return { type: 'percentile', value: Number(c.value ?? 0) };
    case 'formula':
      return { type: 'formula', value: String(c.value ?? '') };
  }
  return fallback;
}

function cfStyleFrom(st: Partial<ExcelJSTypes.Style> | undefined): CFStyle {
  const out: CFStyle = {};
  if (!st) return out;
  const color = fromColor(st.font?.color as XColor, 'border');
  if (color) out.color = color;
  if (st.font?.bold) out.bold = true;
  if (st.font?.italic) out.italic = true;
  if (st.font?.strike) out.strike = true;
  if (st.font?.underline) out.underline = true;
  const fill = st.fill;
  if (fill && fill.type === 'pattern') {
    const c = fromColor(fill.bgColor as XColor, 'fill') ?? fromColor(fill.fgColor as XColor, 'fill');
    if (c) out.fill = c;
  }
  return out;
}

let ruleCounter = 0;

function cfRulesFrom(cf: { ref: string; rules: ExcelJSTypes.ConditionalFormattingRule[] }): CFRule[] {
  const ranges = cf.ref
    .split(/\s+/)
    .map((r) => parseRange(r))
    .filter((r): r is Range => !!r);
  if (!ranges.length) return [];
  const out: CFRule[] = [];
  const sorted = [...cf.rules].sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
  for (const r of sorted) {
    const id = 'cf' + ++ruleCounter;
    const base = { id, ranges };
    const anyR = r as unknown as { formulae?: string[]; operator?: string; text?: string; type: string };
    const style = cfStyleFrom(r.style);
    switch (anyR.type) {
      case 'expression':
        if (anyR.formulae?.[0]) out.push({ ...base, type: 'expression', formula: fromFileFormula(String(anyR.formulae[0])), style });
        break;
      case 'cellIs': {
        const values = (anyR.formulae ?? []).map((f) => {
          const s = String(f);
          return /^-?[\d.]+$/.test(s) ? s : '=' + fromFileFormula(s);
        });
        const op = (anyR.operator ?? 'equal') as never;
        out.push({ ...base, type: 'cellIs', operator: op, values, style });
        break;
      }
      case 'containsText': {
        const op = anyR.operator;
        if (op === 'containsBlanks') out.push({ ...base, type: 'blanks', style });
        else if (op === 'notContainsBlanks') out.push({ ...base, type: 'noBlanks', style });
        else if (op === 'containsErrors') out.push({ ...base, type: 'errors', style });
        else if (op === 'notContainsErrors') out.push({ ...base, type: 'noErrors', style });
        else out.push({ ...base, type: 'text', operator: op === 'notContains' ? 'notContains' : op === 'beginsWith' ? 'beginsWith' : op === 'endsWith' ? 'endsWith' : 'contains', text: anyR.text ?? '', style });
        break;
      }
      case 'containsBlanks':
        out.push({ ...base, type: 'blanks', style });
        break;
      case 'notContainsBlanks':
        out.push({ ...base, type: 'noBlanks', style });
        break;
      case 'containsErrors':
        out.push({ ...base, type: 'errors', style });
        break;
      case 'notContainsErrors':
        out.push({ ...base, type: 'noErrors', style });
        break;
      case 'beginsWith':
      case 'endsWith':
      case 'notContainsText':
        out.push({ ...base, type: 'text', operator: anyR.type === 'notContainsText' ? 'notContains' : (anyR.type as 'beginsWith' | 'endsWith'), text: anyR.text ?? '', style });
        break;
      case 'duplicateValues':
        out.push({ ...base, type: 'duplicate', style });
        break;
      case 'uniqueValues':
        out.push({ ...base, type: 'unique', style });
        break;
      case 'top10': {
        const t = r as ExcelJSTypes.Top10RuleType;
        out.push({ ...base, type: 'top', rank: t.rank ?? 10, percent: !!t.percent, bottom: !!t.bottom, style });
        break;
      }
      case 'aboveAverage': {
        const t = r as ExcelJSTypes.AboveAverageRuleType;
        out.push({ ...base, type: 'average', below: t.aboveAverage === false, style });
        break;
      }
      case 'colorScale': {
        const t = r as ExcelJSTypes.ColorScaleRuleType;
        const cfvo = t.cfvo ?? [];
        const colors = t.color ?? [];
        const stops = cfvo.map((c, i) => ({ at: cfValueFrom(c, { type: 'min' }), color: fromColor(colors[i] as XColor, 'fill') ?? '#888888' }));
        if (stops.length >= 2) out.push({ ...base, type: 'colorScale', stops });
        break;
      }
      case 'dataBar': {
        const t = r as ExcelJSTypes.DataBarRuleType & { color?: XColor };
        out.push({
          ...base,
          type: 'dataBar',
          color: fromColor(t.color, 'fill') ?? '#5b9bd5',
          min: cfValueFrom(t.cfvo?.[0], { type: 'min' }),
          max: cfValueFrom(t.cfvo?.[1], { type: 'max' }),
          gradient: t.gradient !== false,
          showValue: t.showValue !== false,
        });
        break;
      }
      case 'iconSet': {
        const t = r as ExcelJSTypes.IconSetRuleType;
        const set = t.iconSet ?? '3TrafficLights1';
        const n = parseInt(set, 10) || 3;
        const thresholds = (t.cfvo ?? []).map((c) => cfValueFrom(c, { type: 'percent', value: 0 }));
        while (thresholds.length < n) thresholds.push({ type: 'percent', value: Math.round((thresholds.length * 100) / n) });
        out.push({ ...base, type: 'iconSet', iconSet: set, thresholds: thresholds.slice(0, n), reverse: !!t.reverse, showValue: t.showValue !== false });
        break;
      }
    }
  }
  return out;
}

// ---------- read ----------

const pxFromChars = (w: number) => Math.round(w * 7 + 5);
const charsFromPx = (px: number) => Math.max(0, (px - 5) / 7);
const pxFromPt = (pt: number) => Math.round((pt * 4) / 3);
const ptFromPx = (px: number) => Math.round(px * 0.75 * 100) / 100;

export interface ReadResult {
  wb: Workbook;
  warnings: string[];
  /** The file contains VBA code. */
  hasVba: boolean;
}

export async function readXlsx(data: Uint8Array): Promise<ReadResult> {
  const ExcelJS = await loadExcelJS();
  const xwb = new ExcelJS.Workbook();
  const forExcelJS = (await prepareForExcelJS(data).catch(() => null)) ?? data;
  await xwb.xlsx.load(forExcelJS as unknown as ExcelJSTypes.Buffer);
  const wb = new Workbook(true);
  const warnings = new Set<string>();
  const styleCache = new Map<string, number>();
  const styleId = (st: Partial<ExcelJSTypes.Style> | undefined): number | undefined => {
    if (!st || !Object.keys(st).length) return undefined;
    const key = JSON.stringify(st);
    let id = styleCache.get(key);
    if (id === undefined) {
      id = wb.styles.intern(fromExcelStyle(st));
      styleCache.set(key, id);
    }
    return id || undefined;
  };

  xwb.eachSheet((ws) => {
    const sheet = new Sheet(ws.name);
    sheet.hidden = ws.state === 'hidden' || ws.state === 'veryHidden';
    const tab = fromColor(ws.properties?.tabColor as XColor, 'fill');
    if (tab) sheet.tabColor = tab;
    if (ws.properties?.defaultColWidth) sheet.defaultColWidth = pxFromChars(ws.properties.defaultColWidth);
    if (ws.properties?.defaultRowHeight) sheet.defaultRowHeight = Math.max(18, pxFromPt(ws.properties.defaultRowHeight) + 2);
    const view = ws.views?.[0] as (Partial<ExcelJSTypes.WorksheetViewFrozen> & Partial<ExcelJSTypes.WorksheetViewCommon>) | undefined;
    if (view) {
      if (view.state === 'frozen') sheet.freeze = { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 };
      if (view.showGridLines === false) sheet.showGrid = false;
      if (view.zoomScale) sheet.view.zoom = Math.max(0.25, Math.min(4, view.zoomScale / 100));
    }

    // Columns
    const cols = (ws.columns ?? []) as ExcelJSTypes.Column[];
    cols.forEach((col, i) => {
      if (!col) return;
      if (col.width) {
        const px = pxFromChars(col.width);
        if (px !== sheet.defaultColWidth) sheet.colWidths.set(i, px);
      }
      if (col.hidden) sheet.hiddenCols.add(i);
      const sid = styleId(col.style as Partial<ExcelJSTypes.Style>);
      if (sid) sheet.colStyles.set(i, sid);
    });

    const arrayMasters: Range[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const r = rowNumber - 1;
      if (r >= MAX_ROWS) return;
      if (row.height) {
        const px = pxFromPt(row.height);
        if (Math.abs(px - sheet.defaultRowHeight) > 1) sheet.rowHeights.set(r, px);
      }
      if (row.hidden) sheet.hiddenRows.add(r);
      const rsid = styleId((row as unknown as { style: Partial<ExcelJSTypes.Style> }).style);
      if (rsid) sheet.rowStyles.set(r, rsid);
      row.eachCell({ includeEmpty: true }, (xc, colNumber) => {
        const c = colNumber - 1;
        if (c >= MAX_COLS) return;
        const cell: Cell = {};
        const sid = styleId(xc.style as Partial<ExcelJSTypes.Style>);
        const inherited = rsid ?? sheet.colStyles.get(c);
        if (sid !== undefined && sid !== inherited) cell.s = sid;
        const v = xc.value as unknown;
        const type = xc.type as number;
        if (type === 1 /* Merge */) {
          // value belongs to the merge master
        } else if (v === null || v === undefined) {
          // empty
        } else if (typeof v === 'number' || typeof v === 'boolean') cell.v = v;
        else if (typeof v === 'string') cell.v = v;
        else if (v instanceof Date) cell.v = dateToSerial1900(v);
        else if (typeof v === 'object') {
          const o = v as Record<string, unknown>;
          if ('formula' in o || 'sharedFormula' in o) {
            const ftext = xc.formula ?? (o.formula as string | undefined);
            if (ftext) {
              cell.f = fromFileFormula(ftext);
              const cached = fromResult(o.result);
              if (cached !== undefined) cell.cached = cached;
              if (o.shareType === 'array' && typeof o.ref === 'string') {
                const ar = parseRange(o.ref);
                if (ar && (ar.r2 > ar.r1 || ar.c2 > ar.c1)) arrayMasters.push(ar);
              }
            } else {
              const cached = fromResult(o.result);
              if (cached !== undefined) cell.v = cached;
            }
          } else if ('richText' in o) cell.v = richTextOf(o as { richText: { text: string }[] });
          else if ('hyperlink' in o) {
            const t = o.text as unknown;
            cell.v = typeof t === 'string' ? t : t && typeof t === 'object' ? richTextOf(t as { richText: { text: string }[] }) : String(o.hyperlink);
            if (typeof o.hyperlink === 'string' && !o.hyperlink.startsWith('#')) cell.link = o.hyperlink;
          } else if ('error' in o) {
            const code = parseErrorCode(String(o.error));
            if (code) cell.v = new ErrorValue(code as ErrorCode, 'This error was saved in the file.');
          }
        }
        if (Object.keys(cell).length) sheet.cells.set(cellKey(r, c), cell);
      });
    });
    // Cells covered by an array formula hold cached values; they'll be filled by spilling.
    for (const ar of arrayMasters) {
      for (let r = ar.r1; r <= ar.r2; r++)
        for (let c = ar.c1; c <= ar.c2; c++) {
          if (r === ar.r1 && c === ar.c1) continue;
          const k = cellKey(r, c);
          const cell = sheet.cells.get(k);
          if (!cell || cell.f !== undefined) continue;
          if (cell.s !== undefined) sheet.cells.set(k, { s: cell.s });
          else sheet.cells.delete(k);
        }
    }

    const merges = (ws.model as { merges?: string[] }).merges ?? [];
    for (const m of merges) {
      const r = parseRange(m);
      if (r) sheet.merges.push(r);
    }
    const cfs = (ws as unknown as { conditionalFormattings?: { ref: string; rules: ExcelJSTypes.ConditionalFormattingRule[] }[] }).conditionalFormattings ?? [];
    for (const cf of cfs) sheet.cf.push(...cfRulesFrom(cf));
    const af = ws.autoFilter as unknown;
    if (af) {
      let range: Range | null = null;
      if (typeof af === 'string') range = parseRange(af);
      else if (typeof af === 'object') {
        const o = af as { from: unknown; to: unknown };
        const pos = (p: unknown) => (typeof p === 'string' ? p : p && typeof p === 'object' ? formatCell(((p as { row: number }).row ?? 1) - 1, ((p as { column: number }).column ?? 1) - 1) : '');
        range = parseRange(`${pos(o.from)}:${pos(o.to)}`);
      }
      if (range) sheet.filter = { range, conds: {} };
    }
    sheet.invalidateBounds();
    wb.sheets.push(sheet);
  });

  // Defined names
  try {
    const model = (xwb.definedNames as unknown as { model: { name: string; ranges: string[] }[] }).model ?? [];
    for (const dn of model) {
      if (!dn.name || dn.name.startsWith('_xlnm')) continue;
      const ref = dn.ranges?.[0];
      if (ref) wb.names.set(dn.name.toUpperCase(), { name: dn.name, ref: ref.replace(/^=/, '') });
    }
  } catch {
    // ignore malformed names
  }

  if (!wb.sheets.length) wb.sheets.push(new Sheet('Sheet1'));
  const visible = wb.sheets.find((s) => !s.hidden) ?? wb.sheets[0];
  const activeTab = (xwb.views?.[0] as { activeTab?: number } | undefined)?.activeTab;
  const active = activeTab !== undefined ? wb.sheets[activeTab] : undefined;
  wb.activeSheetId = active && !active.hidden ? active.id : visible.id;

  // Notes and charts are read directly from the package
  try {
    await readNotes(data, wb);
  } catch {
    warnings.add("Some cell notes couldn't be read.");
  }
  try {
    const charts = await readCharts(data, wb);
    if (charts.unsupported) warnings.add("Some chart types in this file aren't supported and were skipped.");
    if (charts.pictures) warnings.add("Pictures in this file aren't supported yet and won't be kept when you save.");
  } catch {
    warnings.add('Charts in this file couldn\'t be read.');
  }

  try {
    await readValidations(data, wb);
  } catch {
    warnings.add("Some data validation rules couldn't be read.");
  }

  // Macros: VBA code, sheet code names and Gridline's step macros/buttons
  let hasVba = false;
  try {
    const macros = await readMacros(data, wb);
    hasVba = macros.hasVba;
    for (const w of macros.warnings) warnings.add(w);
  } catch {
    warnings.add("The macros in this file couldn't be read.");
  }

  // Note any functions we can't calculate (their saved values are kept).
  const unknown = new Set<string>();
  for (const s of wb.sheets)
    for (const cell of s.cells.values()) {
      if (cell.f === undefined) continue;
      for (const f of parseFormula(cell.f).funcs) if (!FUNCTIONS.has(f)) unknown.add(f);
    }
  if (unknown.size) {
    const list = [...unknown].slice(0, 6).join(', ');
    warnings.add(`Some functions aren't supported yet (${list}); their last saved values are shown.`);
  }
  return { wb, warnings: [...warnings], hasVba };
}

// ---------- write ----------

/** With `macroEnabled`, VBA code is included and the result must be saved as .xlsm. */
export async function writeXlsx(doc: Doc, opts: { macroEnabled?: boolean } = {}): Promise<Uint8Array> {
  const ExcelJS = await loadExcelJS();
  const xwb = new ExcelJS.Workbook();
  xwb.creator = 'Gridline';
  xwb.lastModifiedBy = 'Gridline';
  xwb.created = new Date();
  xwb.modified = new Date();
  // Gridline-only functions (NFL.*, STOCK…) don't exist in Excel: keep our cached results instead of forcing a recalc
  (xwb as unknown as { calcProperties: { fullCalcOnLoad: boolean } }).calcProperties = { fullCalcOnLoad: !doc.engine.usesCustomFunctions() };
  const wb = doc.wb;
  const styleCache = new Map<number, Partial<ExcelJSTypes.Style>>();
  const styleOf = (id: number | undefined) => {
    if (id === undefined || id === 0) return undefined;
    let s = styleCache.get(id);
    if (!s) styleCache.set(id, (s = toExcelStyle(wb.styles.get(id))));
    return Object.keys(s).length ? s : undefined;
  };
  const activeIndex = Math.max(0, wb.sheets.findIndex((s) => s.id === wb.activeSheetId));
  xwb.views = [{ x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: activeIndex, visibility: 'visible' }];

  for (const sheet of wb.sheets) {
    const views: Partial<ExcelJSTypes.WorksheetView>[] = [];
    const common = { showGridLines: sheet.showGrid, zoomScale: Math.round(sheet.view.zoom * 100) };
    if (sheet.freeze.rows || sheet.freeze.cols) {
      views.push({ state: 'frozen', xSplit: sheet.freeze.cols, ySplit: sheet.freeze.rows, ...common } as Partial<ExcelJSTypes.WorksheetView>);
    } else views.push({ state: 'normal', ...common } as Partial<ExcelJSTypes.WorksheetView>);
    const ws = xwb.addWorksheet(sheet.name, {
      state: sheet.hidden ? 'hidden' : 'visible',
      properties: {
        tabColor: sheet.tabColor ? { argb: toArgb(sheet.tabColor) } : undefined,
        defaultRowHeight: ptFromPx(sheet.defaultRowHeight - 2),
        defaultColWidth: sheet.defaultColWidth !== 88 ? charsFromPx(sheet.defaultColWidth) : undefined,
      } as Partial<ExcelJSTypes.WorksheetProperties>,
      views: views as ExcelJSTypes.WorksheetView[],
    });

    const bounds = doc.engine.usedBounds(sheet);
    const colSet = new Set<number>([...sheet.colWidths.keys(), ...sheet.hiddenCols]);
    for (const c of sheet.colStyles.keys()) if (c <= Math.max(bounds.maxCol + 20, 50)) colSet.add(c);
    for (const c of [...colSet].sort((a, b) => a - b)) {
      if (c >= MAX_COLS) continue;
      const col = ws.getColumn(c + 1);
      col.width = charsFromPx(sheet.colWidth(c));
      if (sheet.hiddenCols.has(c)) col.hidden = true;
      const st = styleOf(sheet.colStyles.get(c));
      if (st) col.style = st as ExcelJSTypes.Style;
    }
    const rowSet = new Set<number>([...sheet.rowHeights.keys(), ...sheet.hiddenRows, ...sheet.filterHidden, ...sheet.rowStyles.keys()]);
    for (const r of [...rowSet].sort((a, b) => a - b)) {
      const row = ws.getRow(r + 1);
      if (sheet.rowHeights.has(r)) row.height = ptFromPx(sheet.rowHeight(r));
      if (sheet.hiddenRows.has(r) || sheet.filterHidden.has(r)) row.hidden = true;
      const st = styleOf(sheet.rowStyles.get(r));
      if (st) (row as unknown as { style: Partial<ExcelJSTypes.Style> }).style = st;
    }

    const keys = [...sheet.cells.keys()].sort((a, b) => a - b);
    for (const key of keys) {
      const cell = sheet.cells.get(key)!;
      const r = keyRow(key);
      const c = keyCol(key);
      const xc = ws.getCell(r + 1, c + 1);
      const st = styleOf(cell.s);
      if (st) xc.style = st as ExcelJSTypes.Style;
      else if (cell.s === 0 && (sheet.rowStyles.has(r) || sheet.colStyles.has(c))) xc.style = {} as ExcelJSTypes.Style;
      if (cell.f !== undefined) {
        const value = doc.engine.getValue(sheet, r, c);
        const result = toResult(value);
        const spill = doc.engine.spillAt(sheet, r, c);
        const formula = toFileFormula(cell.f);
        if (spill) {
          xc.value = { formula, result, shareType: 'array', ref: formatRange(spill) } as unknown as ExcelJSTypes.CellValue;
        } else if (needsArrayEntry(parseFormula(cell.f).ast)) {
          xc.value = { formula, result, shareType: 'array', ref: formatCell(r, c) } as unknown as ExcelJSTypes.CellValue;
        } else {
          xc.value = { formula, result } as ExcelJSTypes.CellFormulaValue;
        }
      } else if (cell.v !== undefined && cell.v !== null) {
        if (cell.link) xc.value = { text: String(scalarText(cell.v)), hyperlink: cell.link };
        else if (cell.v instanceof ErrorValue) xc.value = { error: excelErrorCode(cell.v.code) } as ExcelJSTypes.CellErrorValue;
        else xc.value = cell.v;
      }
      if (cell.note) xc.note = cell.note;
    }
    // Values produced by spilled arrays are stored as cached values of the array formula.
    for (const sp of doc.engine.spillsOn(sheet)) {
      const m = doc.engine.materialize(new RangeRef(sheet, sp.r1, sp.c1, sp.r2, sp.c2), sheet);
      if (!(m instanceof Matrix)) continue;
      for (let i = 0; i < m.rows; i++)
        for (let j = 0; j < m.cols; j++) {
          const r = sp.r1 + i;
          const c = sp.c1 + j;
          const xc = ws.getCell(r + 1, c + 1);
          // Table functions (AMORTIZATION, NFL.SCHEDULE…) suggest formats for their columns
          const fmt = doc.engine.spillFormatAt(sheet, r, c);
          if (fmt && typeof m.get(i, j) === 'number' && !wb.styles.get(sheet.styleIdAt(r, c)).numFmt) xc.numFmt = fmt;
          if (i === 0 && j === 0) continue;
          const v = toResult(m.get(i, j));
          if (v !== undefined) xc.value = v as ExcelJSTypes.CellValue;
        }
    }

    for (const m of sheet.merges) {
      try {
        ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
      } catch {
        // overlapping merges are skipped
      }
    }
    for (const rule of sheet.cf) {
      const xr = toExcelRule(rule);
      if (xr) ws.addConditionalFormatting({ ref: rule.ranges.map((r) => formatRange(r)).join(' '), rules: [xr] });
    }
    if (sheet.filter) ws.autoFilter = formatRange(sheet.filter.range);
  }

  for (const dn of wb.names.values()) {
    try {
      const ref = dn.ref.replace(/^=/, '');
      if (/!/.test(ref)) (xwb.definedNames as unknown as { add: (loc: string, name: string) => void }).add(ref, dn.name);
    } catch {
      // skip names ExcelJS can't represent
    }
  }

  let buf: Uint8Array = new Uint8Array((await xwb.xlsx.writeBuffer()) as ArrayBuffer);
  if (wb.sheets.some((s) => s.charts.length)) buf = await writeCharts(buf, doc);
  buf = await writeValidations(buf, wb);
  return writeMacros(buf, wb, !!opts.macroEnabled);
}

function scalarText(v: Scalar): string {
  if (v === null) return '';
  if (v instanceof ErrorValue) return v.code;
  return String(v);
}

function excelErrorCode(code: ErrorCode): ExcelJSTypes.CellErrorValue['error'] {
  const ok = ['#N/A', '#REF!', '#NAME?', '#DIV/0!', '#NULL!', '#VALUE!', '#NUM!'];
  return (ok.includes(code) ? code : '#VALUE!') as ExcelJSTypes.CellErrorValue['error'];
}

function toResult(v: Scalar | Matrix): number | string | boolean | ExcelJSTypes.CellErrorValue | undefined {
  if (v instanceof Matrix) return toResult(v.get(0, 0));
  if (v === null) return undefined;
  if (v instanceof ErrorValue) return { error: excelErrorCode(v.code) };
  return v;
}

function cfStyleTo(s: CFStyle): Partial<ExcelJSTypes.Style> {
  const out: Partial<ExcelJSTypes.Style> = {};
  const font: Partial<ExcelJSTypes.Font> = {};
  if (s.color) font.color = { argb: toArgb(s.color) };
  if (s.bold) font.bold = true;
  if (s.italic) font.italic = true;
  if (s.strike) font.strike = true;
  if (s.underline) font.underline = true;
  if (Object.keys(font).length) out.font = font;
  if (s.fill) out.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: toArgb(s.fill) }, fgColor: { argb: toArgb(s.fill) } };
  return out;
}

function cfvoTo(v: CFValue): ExcelJSTypes.Cvfo {
  if (v.type === 'formula') return { type: 'formula', value: String(v.value ?? '').replace(/^=/, '') as unknown as number };
  if (v.type === 'min' || v.type === 'max') return { type: v.type };
  return { type: v.type, value: Number(v.value ?? 0) };
}

function toExcelRule(rule: CFRule): ExcelJSTypes.ConditionalFormattingRule | null {
  const tl = formatCell(rule.ranges[0].r1, rule.ranges[0].c1);
  const abs = formatRange(rule.ranges[0]).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
  const priority = 1;
  switch (rule.type) {
    case 'expression':
      return { type: 'expression', priority, formulae: [toFileFormula(rule.formula.replace(/^=/, ''))], style: cfStyleTo(rule.style) };
    case 'cellIs': {
      const vals = rule.values.map((v) => (v.startsWith('=') ? toFileFormula(v.slice(1)) : /^-?[\d.]+$/.test(v) ? v : `"${v.replace(/"/g, '""')}"`));
      if (rule.operator === 'notBetween' || rule.operator === 'notEqual' || rule.operator === 'greaterThanOrEqual' || rule.operator === 'lessThanOrEqual') {
        const opExpr: Record<string, string> = {
          notEqual: `${tl}<>${vals[0]}`,
          greaterThanOrEqual: `${tl}>=${vals[0]}`,
          lessThanOrEqual: `${tl}<=${vals[0]}`,
          notBetween: `OR(${tl}<MIN(${vals[0]},${vals[1]}),${tl}>MAX(${vals[0]},${vals[1]}))`,
        };
        return { type: 'expression', priority, formulae: [opExpr[rule.operator]], style: cfStyleTo(rule.style) };
      }
      return { type: 'cellIs', priority, operator: rule.operator as ExcelJSTypes.CellIsOperators, formulae: vals, style: cfStyleTo(rule.style) };
    }
    case 'text': {
      const q = `"${rule.text.replace(/"/g, '""')}"`;
      const expr: Record<string, string> = {
        contains: `NOT(ISERROR(SEARCH(${q},${tl})))`,
        notContains: `ISERROR(SEARCH(${q},${tl}))`,
        beginsWith: `LEFT(${tl},LEN(${q}))=${q}`,
        endsWith: `RIGHT(${tl},LEN(${q}))=${q}`,
      };
      return { type: 'expression', priority, formulae: [expr[rule.operator]], style: cfStyleTo(rule.style) };
    }
    case 'duplicate':
      return { type: 'expression', priority, formulae: [`AND(${tl}<>"",COUNTIF(${abs},${tl})>1)`], style: cfStyleTo(rule.style) };
    case 'unique':
      return { type: 'expression', priority, formulae: [`AND(${tl}<>"",COUNTIF(${abs},${tl})=1)`], style: cfStyleTo(rule.style) };
    case 'top':
      return { type: 'top10', priority, rank: rule.rank, percent: rule.percent, bottom: rule.bottom, style: cfStyleTo(rule.style) };
    case 'average':
      return { type: 'aboveAverage', priority, aboveAverage: !rule.below, style: cfStyleTo(rule.style) };
    case 'blanks':
      return { type: 'expression', priority, formulae: [`LEN(TRIM(${tl}))=0`], style: cfStyleTo(rule.style) };
    case 'noBlanks':
      return { type: 'expression', priority, formulae: [`LEN(TRIM(${tl}))>0`], style: cfStyleTo(rule.style) };
    case 'errors':
      return { type: 'expression', priority, formulae: [`ISERROR(${tl})`], style: cfStyleTo(rule.style) };
    case 'noErrors':
      return { type: 'expression', priority, formulae: [`NOT(ISERROR(${tl}))`], style: cfStyleTo(rule.style) };
    case 'colorScale':
      return { type: 'colorScale', priority, cfvo: rule.stops.map((s) => cfvoTo(s.at)), color: rule.stops.map((s) => ({ argb: toArgb(s.color) })) };
    case 'dataBar':
      return {
        type: 'dataBar',
        priority,
        gradient: rule.gradient,
        showValue: rule.showValue,
        cfvo: [cfvoTo(rule.min), cfvoTo(rule.max)],
        color: { argb: toArgb(rule.color) },
      } as unknown as ExcelJSTypes.ConditionalFormattingRule;
    case 'iconSet':
      return {
        type: 'iconSet',
        priority,
        iconSet: rule.iconSet as ExcelJSTypes.IconSetTypes,
        reverse: rule.reverse,
        showValue: rule.showValue,
        cfvo: rule.thresholds.map(cfvoTo),
      };
  }
  return null;
}

export { pxFromChars, charsFromPx };
