// Chart data: building series from a range and resolving them to numbers.
import { formatCell, quoteSheetName, type Range } from './address';
import type { Doc } from './doc';
import { tokenize } from './formula/tokenizer';
import { formatValue, isDateFormat } from './numfmt';
import type { Scalar } from './values';
import type { ChartDef, ChartSeries, ChartType, Sheet } from './workbook';

const PALETTES: Record<string, string[]> = {
  gridline: ['#4ea1ff', '#ff9f43', '#3fcf8e', '#e8618c', '#a78bfa', '#ffd43b', '#22b8cf', '#ff6b6b', '#94d82d', '#c4a7ff'],
  office: ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e', '#636363', '#997300'],
  colorblind: ['#0072b2', '#e69f00', '#009e73', '#cc79a7', '#56b4e9', '#d55e00', '#f0e442', '#000000', '#999999', '#88ccee'],
  pastel: ['#a0c4ff', '#ffc6a5', '#b9fbc0', '#ffadad', '#cdb4db', '#fdffb6', '#9bf6ff', '#ffd6e0', '#caffbf', '#bdb2ff'],
  vivid: ['#0066ff', '#ff3d00', '#00c853', '#d500f9', '#ffab00', '#00b8d4', '#ff1744', '#64dd17', '#651fff', '#ff6d00'],
  mono: ['#1e4e8c', '#2f6fbf', '#4a8fe0', '#73abed', '#9cc5f4', '#c3dcf8', '#163a68', '#5b9be6', '#86b8f0', '#b0d1f6'],
};

/** Series colors, in order. Changed in place by setChartPalette so every importer sees the update. */
export const CHART_COLORS = [...PALETTES.gridline];

export function setChartPalette(name: string): void {
  CHART_COLORS.splice(0, CHART_COLORS.length, ...(PALETTES[name] ?? PALETTES.gridline));
}

function absRange(sheet: Sheet, r1: number, c1: number, r2: number, c2: number): string {
  const a = formatCell(r1, c1, true, true);
  const b = formatCell(r2, c2, true, true);
  return `${quoteSheetName(sheet.name)}!${a === b ? a : `${a}:${b}`}`;
}

export interface ChartSource {
  range: string;
  byRows: boolean;
  headerRow: boolean;
  headerCol: boolean;
}

/** Guesses the layout of a data range: header row, category column, series orientation. */
export function guessSource(doc: Doc, sheet: Sheet, range: Range): ChartSource {
  const get = (r: number, c: number) => doc.engine.getValue(sheet, r, c);
  const rows = range.r2 - range.r1 + 1;
  const cols = range.c2 - range.c1 + 1;
  const isText = (v: Scalar) => typeof v === 'string' && v !== '';
  const topLeftEmpty = get(range.r1, range.c1) === null || get(range.r1, range.c1) === '';
  let headerRow = false;
  if (rows > 1) {
    let text = 0;
    for (let c = range.c1; c <= range.c2; c++) if (isText(get(range.r1, c))) text++;
    let numsBelow = 0;
    for (let c = range.c1; c <= range.c2; c++) if (typeof get(range.r1 + 1, c) === 'number') numsBelow++;
    headerRow = (text > 0 && numsBelow > 0) || topLeftEmpty;
  }
  let headerCol = false;
  if (cols > 1) {
    let text = 0;
    let dates = 0;
    const start = headerRow ? range.r1 + 1 : range.r1;
    for (let r = start; r <= range.r2; r++) {
      const v = get(r, range.c1);
      if (isText(v)) text++;
      else if (typeof v === 'number' && isDateFormat(doc.wb.styles.get(sheet.styleIdAt(r, range.c1)).numFmt)) dates++;
    }
    headerCol = text + dates >= Math.max(1, (range.r2 - start + 1) / 2) || topLeftEmpty;
  }
  const dataRows = rows - (headerRow ? 1 : 0);
  const dataCols = cols - (headerCol ? 1 : 0);
  return {
    range: absRange(sheet, range.r1, range.c1, range.r2, range.c2),
    byRows: dataCols > dataRows * 2 && dataRows <= 6,
    headerRow,
    headerCol,
  };
}

/** Turns a source description into explicit series references. */
export function seriesFromSource(doc: Doc, src: ChartSource, type: ChartType): ChartSeries[] {
  const resolved = resolveRef(doc, src.range, doc.wb.activeSheet);
  if (!resolved) return [];
  const { sheet, range } = resolved;
  const dr1 = range.r1 + (src.headerRow ? 1 : 0);
  const dc1 = range.c1 + (src.headerCol ? 1 : 0);
  const out: ChartSeries[] = [];
  if (!src.byRows) {
    const cats = src.headerCol ? absRange(sheet, dr1, range.c1, range.r2, range.c1) : undefined;
    for (let c = dc1; c <= range.c2; c++) {
      if (dr1 > range.r2) break;
      out.push({
        name: src.headerRow ? '=' + absRange(sheet, range.r1, c, range.r1, c) : undefined,
        categories: cats,
        values: absRange(sheet, dr1, c, range.r2, c),
      });
    }
  } else {
    const cats = src.headerRow ? absRange(sheet, range.r1, dc1, range.r1, range.c2) : undefined;
    for (let r = dr1; r <= range.r2; r++) {
      if (dc1 > range.c2) break;
      out.push({
        name: src.headerCol ? '=' + absRange(sheet, r, range.c1, r, range.c1) : undefined,
        categories: cats,
        values: absRange(sheet, r, dc1, r, range.c2),
      });
    }
  }
  if ((type === 'pie' || type === 'doughnut') && out.length > 1) return out.slice(0, 1);
  return out;
}

export function createChart(doc: Doc, sheet: Sheet, range: Range, type: ChartType, anchorRow: number, anchorCol: number): ChartDef {
  const source = guessSource(doc, sheet, range);
  return {
    id: 'chart-' + Math.random().toString(36).slice(2, 10),
    type,
    title: undefined,
    series: seriesFromSource(doc, source, type),
    source,
    legend: 'bottom',
    anchor: { row: anchorRow, col: anchorCol, dx: 8, dy: 8, w: 520, h: 320 },
  };
}

/** Resolves "Sheet1!$A$1:$B$5" to a sheet and range. */
export function resolveRef(doc: Doc, text: string, defaultSheet: Sheet): { sheet: Sheet; range: Range } | null {
  const toks = tokenize(text.replace(/^=/, '').trim()).filter((t) => t.type !== 'ws');
  if (toks.length !== 1 || toks[0].type !== 'ref' || !toks[0].ref) return null;
  const ref = toks[0].ref;
  const sheet = ref.sheet !== undefined ? doc.wb.sheetByName(ref.sheet) : defaultSheet;
  if (!sheet) return null;
  return { sheet, range: { r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 } };
}

function refCells(doc: Doc, text: string | undefined, defaultSheet: Sheet): { values: Scalar[]; labels: string[] } {
  if (!text) return { values: [], labels: [] };
  const res = resolveRef(doc, text, defaultSheet);
  if (!res) {
    // A literal (e.g. a typed series name)
    return { values: [text], labels: [text.replace(/^"(.*)"$/, '$1')] };
  }
  const { sheet, range } = res;
  const b = doc.engine.usedBounds(sheet);
  const r2 = Math.min(range.r2, Math.max(range.r1, b.maxRow));
  const c2 = Math.min(range.c2, Math.max(range.c1, b.maxCol));
  const values: Scalar[] = [];
  const labels: string[] = [];
  for (let r = range.r1; r <= r2; r++)
    for (let c = range.c1; c <= c2; c++) {
      if (sheet.isRowHidden(r) || sheet.isColHidden(c)) continue;
      const v = doc.engine.getValue(sheet, r, c);
      values.push(v);
      labels.push(v === null ? '' : formatValue(v, doc.wb.styles.get(sheet.styleIdAt(r, c)).numFmt, 40).text);
    }
  return { values, labels };
}

export interface ResolvedChart {
  categories: string[];
  /** Numeric x values for scatter charts. */
  xs: (number | null)[];
  series: { name: string; values: (number | null)[]; color: string; format?: string }[];
}

export function chartData(doc: Doc, chart: ChartDef, sheet: Sheet): ResolvedChart {
  let categories: string[] = [];
  let xs: (number | null)[] = [];
  const series = chart.series.map((se, i) => {
    const vals = refCells(doc, se.values, sheet);
    const name = se.name
      ? se.name.startsWith('=')
        ? refCells(doc, se.name.slice(1), sheet).labels.join(' ')
        : se.name
      : `Series ${i + 1}`;
    if (i === 0 || !categories.length) {
      const cats = refCells(doc, se.categories, sheet);
      categories = cats.labels;
      xs = cats.values.map((v) => (typeof v === 'number' ? v : null));
    }
    const res = resolveRef(doc, se.values, sheet);
    const format = res ? doc.wb.styles.get(res.sheet.styleIdAt(res.range.r1, res.range.c1)).numFmt : undefined;
    return {
      name,
      values: vals.values.map((v) => (typeof v === 'number' ? v : null)),
      color: se.color ?? CHART_COLORS[i % CHART_COLORS.length],
      format,
    };
  });
  const n = Math.max(0, ...series.map((s) => s.values.length));
  if (!categories.length) categories = Array.from({ length: n }, (_, i) => String(i + 1));
  if (!xs.length || xs.every((x) => x === null)) xs = Array.from({ length: n }, (_, i) => i + 1);
  return { categories, xs, series };
}
