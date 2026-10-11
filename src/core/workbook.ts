// The workbook data model: sheets, cells, and sheet-level features.
import { cellKey, keyCol, keyRow, MAX_COLS, MAX_ROWS, type Range } from './address';
import type { MacroProject, SheetButton } from './macros/project';
import { StyleTable } from './styles';
import type { DataValidation } from './validation';
import type { Scalar } from './values';

export interface Cell {
  /** Literal value for non-formula cells. */
  v?: Scalar;
  /** Formula text without the leading '='. */
  f?: string;
  /** Style id in the workbook's StyleTable. */
  s?: number;
  /** Value saved in the file for formulas we can't evaluate (unsupported functions). */
  cached?: Scalar;
  note?: string;
  link?: string;
}

export function isEmptyCell(c: Cell | undefined): boolean {
  return !c || (c.f === undefined && (c.v === undefined || c.v === null) && c.s === undefined && !c.note && !c.link);
}

export function hasContent(c: Cell | undefined): boolean {
  return !!c && (c.f !== undefined || (c.v !== undefined && c.v !== null && c.v !== ''));
}

// ---- Conditional formatting ----

export interface CFStyle {
  color?: string;
  fill?: string;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: boolean;
}

export type CFValueType = 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula';
export interface CFValue {
  type: CFValueType;
  value?: number | string;
}

export type CFOperator =
  | 'greaterThan'
  | 'lessThan'
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';

export type CFRuleSpec =
  | { type: 'cellIs'; operator: CFOperator; values: string[]; style: CFStyle }
  | { type: 'text'; operator: 'contains' | 'notContains' | 'beginsWith' | 'endsWith'; text: string; style: CFStyle }
  | { type: 'duplicate' | 'unique'; style: CFStyle }
  | { type: 'top'; rank: number; percent: boolean; bottom: boolean; style: CFStyle }
  | { type: 'average'; below: boolean; style: CFStyle }
  | { type: 'blanks' | 'noBlanks' | 'errors' | 'noErrors'; style: CFStyle }
  | { type: 'expression'; formula: string; style: CFStyle }
  | { type: 'colorScale'; stops: { at: CFValue; color: string }[] }
  | { type: 'dataBar'; color: string; min: CFValue; max: CFValue; gradient: boolean; showValue: boolean }
  | { type: 'iconSet'; iconSet: string; thresholds: CFValue[]; reverse: boolean; showValue: boolean };

export type CFRule = CFRuleSpec & { id: string; ranges: Range[]; stopIfTrue?: boolean };

// ---- Filters ----

export interface FilterCond {
  /** Allowed display values ('' = blanks). undefined = all values allowed. */
  values?: string[];
  op?: { type: FilterOp; value: string; value2?: string };
}

export type FilterOp =
  | 'contains'
  | 'notContains'
  | 'beginsWith'
  | 'endsWith'
  | 'equals'
  | 'notEquals'
  | 'gt'
  | 'gte'
  | 'lt'
  | 'lte'
  | 'between'
  | 'top10'
  | 'aboveAvg'
  | 'belowAvg';

export interface FilterState {
  range: Range;
  /** Keyed by absolute column index. */
  conds: Record<number, FilterCond>;
}

// ---- Charts ----

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';

export interface ChartSeries {
  /** Formula-style reference or literal for the series name, e.g. "Sheet1!$B$1". */
  name?: string;
  categories?: string; // e.g. "Sheet1!$A$2:$A$10"
  values: string; // e.g. "Sheet1!$B$2:$B$10"
  color?: string;
}

export interface ChartDef {
  id: string;
  type: ChartType;
  title?: string;
  series: ChartSeries[];
  /** When created from a range, remembers how so the editor can regenerate series. */
  source?: { range: string; byRows: boolean; headerRow: boolean; headerCol: boolean };
  stacked?: boolean;
  legend: 'right' | 'bottom' | 'top' | 'none';
  xTitle?: string;
  yTitle?: string;
  /** Anchor: top-left cell plus pixel offset, and size in px (at 100% zoom). */
  anchor: { row: number; col: number; dx: number; dy: number; w: number; h: number };
}

export interface DefinedName {
  name: string;
  /** Formula text without '=', usually a reference like Sheet1!$A$1:$B$5 */
  ref: string;
}

export const DEFAULT_COL_WIDTH = 88;
export const DEFAULT_ROW_HEIGHT = 22;

let nextSheetId = 1;

export class Sheet {
  readonly id: number;
  name: string;
  cells = new Map<number, Cell>();
  colWidths = new Map<number, number>();
  rowHeights = new Map<number, number>();
  hiddenCols = new Set<number>();
  hiddenRows = new Set<number>();
  filterHidden = new Set<number>();
  colStyles = new Map<number, number>();
  rowStyles = new Map<number, number>();
  merges: Range[] = [];
  freeze = { rows: 0, cols: 0 };
  cf: CFRule[] = [];
  filter: FilterState | null = null;
  charts: ChartDef[] = [];
  /** Buttons that run macros. */
  buttons: SheetButton[] = [];
  /** Data validation rules (allowed values, dropdown lists). */
  validations: DataValidation[] = [];
  /** The sheet's name in VBA code (Sheet1), which stays the same when the tab is renamed. */
  codeName?: string;
  tabColor?: string;
  showGrid = true;
  hidden = false;
  defaultColWidth = DEFAULT_COL_WIDTH;
  defaultRowHeight = DEFAULT_ROW_HEIGHT;

  /** UI state, not part of undo. */
  view = {
    activeRow: 0,
    activeCol: 0,
    scrollX: 0,
    scrollY: 0,
    zoom: 1,
  };

  private boundsCache: { maxRow: number; maxCol: number } | null = null;

  constructor(name: string, id?: number) {
    this.id = id ?? nextSheetId++;
    if (this.id >= nextSheetId) nextSheetId = this.id + 1;
    this.name = name;
  }

  getCell(row: number, col: number): Cell | undefined {
    return this.cells.get(cellKey(row, col));
  }

  /** Low-level setter; use a Transaction for undoable edits. */
  rawSet(key: number, cell: Cell | undefined): void {
    if (cell && !isEmptyCell(cell)) this.cells.set(key, cell);
    else this.cells.delete(key);
    this.boundsCache = null;
  }

  invalidateBounds(): void {
    this.boundsCache = null;
  }

  /** Last row/column containing any cell (content or formatting). -1 if empty. */
  bounds(): { maxRow: number; maxCol: number } {
    if (this.boundsCache) return this.boundsCache;
    let maxRow = -1;
    let maxCol = -1;
    for (const key of this.cells.keys()) {
      const r = keyRow(key);
      const c = keyCol(key);
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    }
    for (const m of this.merges) {
      if (m.r2 > maxRow) maxRow = m.r2;
      if (m.c2 > maxCol) maxCol = m.c2;
    }
    this.boundsCache = { maxRow, maxCol };
    return this.boundsCache;
  }

  /** Bounds considering only cells with values or formulas. */
  contentBounds(): { maxRow: number; maxCol: number } {
    let maxRow = -1;
    let maxCol = -1;
    for (const [key, c] of this.cells) {
      if (!hasContent(c)) continue;
      const r = keyRow(key);
      const col = keyCol(key);
      if (r > maxRow) maxRow = r;
      if (col > maxCol) maxCol = col;
    }
    return { maxRow, maxCol };
  }

  isRowHidden(r: number): boolean {
    return this.hiddenRows.has(r) || this.filterHidden.has(r);
  }

  isColHidden(c: number): boolean {
    return this.hiddenCols.has(c);
  }

  colWidth(c: number): number {
    return this.colWidths.get(c) ?? this.defaultColWidth;
  }

  rowHeight(r: number): number {
    return this.rowHeights.get(r) ?? this.defaultRowHeight;
  }

  mergeAt(row: number, col: number): Range | undefined {
    for (const m of this.merges) {
      if (row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2) return m;
    }
    return undefined;
  }

  /** Effective style id: cell style, else row style, else column style. */
  styleIdAt(row: number, col: number): number | undefined {
    const c = this.cells.get(cellKey(row, col));
    if (c?.s !== undefined) return c.s;
    const rs = this.rowStyles.get(row);
    if (rs !== undefined) return rs;
    return this.colStyles.get(col);
  }
}

export class Workbook {
  sheets: Sheet[] = [];
  styles = new StyleTable();
  names = new Map<string, DefinedName>();
  activeSheetId = 0;
  /** Macros saved in the workbook (null when it has none). */
  macros: MacroProject | null = null;

  constructor(empty = false) {
    if (!empty) {
      const s = new Sheet('Sheet1');
      this.sheets.push(s);
      this.activeSheetId = s.id;
    }
  }

  get activeSheet(): Sheet {
    return this.sheets.find((s) => s.id === this.activeSheetId) ?? this.sheets[0];
  }

  sheetById(id: number): Sheet | undefined {
    return this.sheets.find((s) => s.id === id);
  }

  sheetByName(name: string): Sheet | undefined {
    const lower = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === lower);
  }

  uniqueSheetName(base = 'Sheet'): string {
    let i = this.sheets.length + 1;
    const exists = (n: string) => this.sheets.some((s) => s.name.toLowerCase() === n.toLowerCase());
    if (base !== 'Sheet' && !exists(base)) return base;
    let name = `${base}${i}`;
    while (exists(name)) name = `${base}${++i}`;
    return name;
  }

  getName(name: string): DefinedName | undefined {
    return this.names.get(name.toUpperCase());
  }
}

export function validateSheetName(name: string, wb: Workbook, exceptId?: number): string | null {
  const n = name.trim();
  if (!n) return 'Sheet names can\'t be empty.';
  if (n.length > 31) return 'Sheet names can be at most 31 characters.';
  if (/[\\/?*[\]:]/.test(n)) return 'Sheet names can\'t contain \\ / ? * [ ] or :';
  if (n.startsWith("'") || n.endsWith("'")) return 'Sheet names can\'t start or end with an apostrophe.';
  if (wb.sheets.some((s) => s.id !== exceptId && s.name.toLowerCase() === n.toLowerCase()))
    return `There's already a sheet called "${n}".`;
  return null;
}

export const SHEET_LIMITS = { rows: MAX_ROWS, cols: MAX_COLS };
