// The spreadsheet grid: a virtualized canvas with native scrolling, frozen panes and mouse interaction.
import { cellKey, colToName, MAX_COLS, MAX_ROWS, normalize, type Range } from '../../core/address';
import { formatValue } from '../../core/numfmt';
import { luminance, parseHex, toHex, type BorderSide, type Style } from '../../core/styles';
import { sortIsStale, TABLE_COLORS } from '../../core/tables';
import { ErrorValue, type Scalar } from '../../core/values';
import type { Sheet, TableDef } from '../../core/workbook';
import type { App } from '../app';
import { Axis } from './axis';
import { displayFill, displayText, readTheme, type GridTheme } from './theme';

/** A click on one of a table's header buttons: the sort arrow or the menu. */
export interface TableButtonHit {
  table: TableDef;
  col: number;
  part: 'sort' | 'menu';
  /** The button, in grid coordinates. */
  rect: { x: number; y: number; w: number; h: number };
}

export type Hit =
  | { kind: 'corner' }
  | { kind: 'colHeader'; col: number; edge: number | null; tableButton?: TableButtonHit | null }
  | { kind: 'rowHeader'; row: number; edge: number | null }
  | { kind: 'cell'; row: number; col: number; fillHandle: boolean; filterButton: boolean; tableButton: TableButtonHit | null }
  | { kind: 'none' };

type Box = { x: number; y: number; w: number; h: number };

interface RefHighlight {
  sheetId: number;
  range: Range;
  color: string;
}

type Drag =
  | { kind: 'select'; additive: boolean }
  | { kind: 'cols'; from: number }
  | { kind: 'rows'; from: number }
  | { kind: 'fill'; src: Range }
  | { kind: 'resizeCol'; col: number; startX: number; startW: number; cols: number[] }
  | { kind: 'resizeRow'; row: number; startY: number; startH: number; rows: number[] }
  | { kind: 'point'; anchor: { row: number; col: number } };

/** `a` mixed into `b` (weight 0–1), for hex colors. */
function mixHex(a: string, b: string, w: number): string {
  const [ar, ag, ab] = parseHex(a);
  const [br, bg, bb] = parseHex(b);
  return toHex(Math.round(ar * w + br * (1 - w)), Math.round(ag * w + bg * (1 - w)), Math.round(ab * w + bb * (1 - w)));
}

export const REF_COLORS = ['#4ea1ff', '#ff7a6b', '#b38cff', '#3fcf8e', '#ffb547', '#2fc6d6', '#ff6fb1', '#a3d65c'];

const PAD = 4;

export class GridView {
  readonly el: HTMLDivElement;
  readonly scroller: HTMLDivElement;
  readonly overlay: HTMLDivElement;
  readonly focusTarget: HTMLTextAreaElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private spacer: HTMLDivElement;
  private dpr = 1;
  width = 0;
  height = 0;
  rows!: Axis;
  cols!: Axis;
  headerW = 46;
  headerH = 24;
  frozenW = 0;
  frozenH = 0;
  theme!: GridTheme;
  private drawQueued = false;
  private layoutDirty = true;
  private drag: Drag | null = null;
  private autoScroll: { dx: number; dy: number; x: number; y: number } | null = null;
  private autoScrollTimer: number | null = null;
  fillPreview: Range | null = null;
  copyMarquee: { sheetId: number; range: Range } | null = null;
  refHighlights: RefHighlight[] = [];
  private marqueeOffset = 0;
  private marqueeTimer: number | null = null;
  private measureCache = new Map<string, number>();
  /** Row/col extents of the scrollable area (grows as you scroll). */
  private extentRows = 200;
  private extentCols = 52;
  private lastScroll = { x: 0, y: 0 };
  private dragMoved = false;

  constructor(private app: App) {
    this.canvas = document.createElement('canvas');
    this.ctx = this.canvas.getContext('2d', { alpha: false })!;
    this.spacer = document.createElement('div');
    this.spacer.className = 'grid-spacer';
    const sticky = document.createElement('div');
    sticky.className = 'grid-sticky';
    sticky.append(this.canvas);
    this.scroller = document.createElement('div');
    this.scroller.className = 'grid-scroller';
    this.scroller.append(sticky, this.spacer);
    this.overlay = document.createElement('div');
    this.overlay.className = 'grid-overlay';
    this.focusTarget = document.createElement('textarea');
    this.focusTarget.className = 'grid-focus';
    this.focusTarget.setAttribute('aria-label', 'Spreadsheet grid');
    this.focusTarget.spellcheck = false;
    this.el = document.createElement('div');
    this.el.className = 'grid-host';
    this.el.append(this.scroller, this.overlay, this.focusTarget);

    new ResizeObserver(() => this.resize()).observe(this.el);
    this.scroller.addEventListener('scroll', () => this.onScroll(), { passive: true });
    this.scroller.addEventListener('pointerdown', (e) => this.onPointerDown(e));
    this.scroller.addEventListener('pointermove', (e) => this.onPointerMove(e));
    this.scroller.addEventListener('pointerup', (e) => this.onPointerUp(e));
    this.scroller.addEventListener('dblclick', (e) => this.onDoubleClick(e));
    this.scroller.addEventListener('contextmenu', (e) => this.onContextMenu(e));
    this.scroller.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
    this.scroller.addEventListener('mouseleave', () => this.app.hideCellTip());
    this.refreshTheme();
  }

  get sheet(): Sheet {
    return this.app.sheet;
  }

  get zoom(): number {
    return this.sheet.view.zoom;
  }

  refreshTheme(): void {
    this.theme = readTheme(this.app.settings.bool('appearance.adaptColors'));
    this.invalidate();
  }

  // ---------- layout ----------

  /** Rebuilds row/column geometry (after resizing, hiding, zoom or sheet change). */
  relayout(): void {
    this.layoutDirty = true;
    this.draw();
  }

  private ensureLayout(): void {
    if (!this.layoutDirty) return;
    this.layoutDirty = false;
    const s = this.sheet;
    const z = this.zoom;
    this.rows = new Axis(MAX_ROWS, s.defaultRowHeight, s.rowHeights, [s.hiddenRows, s.filterHidden], z);
    this.cols = new Axis(MAX_COLS, s.defaultColWidth, s.colWidths, [s.hiddenCols], z);
    const headings = this.app.settings.bool('grid.headings');
    this.headerH = headings ? Math.round(24 * Math.max(0.8, Math.min(z, 1.4))) : 0;
    const digits = String(Math.max(this.extentRows, 1000)).length;
    this.headerW = headings ? Math.round((14 + digits * 8) * Math.max(0.85, Math.min(z, 1.4))) : 0;
    this.frozenH = this.rows.offset(s.freeze.rows);
    this.frozenW = this.cols.offset(s.freeze.cols);
    this.updateExtent();
  }

  private updateExtent(): void {
    const b = this.app.doc.engine.usedBounds(this.sheet);
    const sel = this.app.sel.range;
    const visRows = Math.ceil(this.height / Math.max(8, this.rows.def)) + 1;
    const visCols = Math.ceil(this.width / Math.max(8, this.cols.def)) + 1;
    const lastRowShown = this.rows.indexAt(this.scroller.scrollTop + this.height);
    const lastColShown = this.cols.indexAt(this.scroller.scrollLeft + this.width);
    const chartBottom = Math.max(0, ...this.sheet.charts.map((c) => c.anchor.row + 30));
    const chartRight = Math.max(0, ...this.sheet.charts.map((c) => c.anchor.col + 12));
    const rows = Math.max(b.maxRow + 1 + visRows, lastRowShown + visRows, Math.min(sel.r2, MAX_ROWS - 1) + 1 + visRows, chartBottom, 100);
    const cols = Math.max(b.maxCol + 1 + visCols, lastColShown + visCols, Math.min(sel.c2, MAX_COLS - 1) + 1 + visCols, chartRight, 26);
    this.extentRows = Math.min(MAX_ROWS, rows);
    this.extentCols = Math.min(MAX_COLS, cols);
    const w = this.headerW + this.cols.offset(this.extentCols);
    const h = this.headerH + this.rows.offset(this.extentRows);
    this.spacer.style.width = `${Math.min(w, 30_000_000)}px`;
    this.spacer.style.height = `${Math.min(h, 30_000_000)}px`;
  }

  resize(): void {
    const rect = this.el.getBoundingClientRect();
    // The scroller reserves space for its scrollbars; draw only in the client area.
    this.width = this.scroller.clientWidth || rect.width;
    this.height = this.scroller.clientHeight || rect.height;
    this.dpr = window.devicePixelRatio || 1;
    this.canvas.width = Math.max(1, Math.floor(this.width * this.dpr));
    this.canvas.height = Math.max(1, Math.floor(this.height * this.dpr));
    this.canvas.style.width = `${this.width}px`;
    this.canvas.style.height = `${this.height}px`;
    this.layoutDirty = true;
    this.draw();
    this.app.onGridScroll();
  }

  private onScroll(): void {
    const x = this.scroller.scrollLeft;
    const y = this.scroller.scrollTop;
    if (x === this.lastScroll.x && y === this.lastScroll.y) return;
    this.lastScroll = { x, y };
    this.sheet.view.scrollX = x;
    this.sheet.view.scrollY = y;
    // Grow the scrollable area when nearing the end, like Excel.
    if (this.rows && (y + this.height * 2 > this.rows.offset(this.extentRows) || x + this.width * 2 > this.cols.offset(this.extentCols))) {
      this.updateExtent();
    }
    this.app.hideCellTip();
    this.draw();
    this.app.onGridScroll();
  }

  get scrollX(): number {
    return this.scroller.scrollLeft;
  }

  get scrollY(): number {
    return this.scroller.scrollTop;
  }

  restoreScroll(): void {
    this.ensureLayout();
    this.scroller.scrollLeft = this.sheet.view.scrollX;
    this.scroller.scrollTop = this.sheet.view.scrollY;
    this.lastScroll = { x: this.scroller.scrollLeft, y: this.scroller.scrollTop };
  }

  // ---------- coordinates ----------

  colX(c: number): number {
    return this.headerW + this.cols.offset(c) - (c >= this.sheet.freeze.cols ? this.scrollX : 0);
  }

  rowY(r: number): number {
    return this.headerH + this.rows.offset(r) - (r >= this.sheet.freeze.rows ? this.scrollY : 0);
  }

  /** Rect of a cell (or merged block) in grid-host coordinates. */
  cellRect(row: number, col: number): { x: number; y: number; w: number; h: number } {
    this.ensureLayout();
    const m = this.sheet.mergeAt(row, col);
    const r1 = m ? m.r1 : row;
    const c1 = m ? m.c1 : col;
    const r2 = m ? m.r2 : row;
    const c2 = m ? m.c2 : col;
    const x = this.colX(c1);
    const y = this.rowY(r1);
    return { x, y, w: this.cols.offset(c2 + 1) - this.cols.offset(c1), h: this.rows.offset(r2 + 1) - this.rows.offset(r1) };
  }

  rangeRect(r: Range): { x: number; y: number; w: number; h: number } {
    const x1 = this.colX(r.c1);
    const y1 = this.rowY(r.r1);
    const x2 = this.colX(Math.min(r.c2, MAX_COLS - 1)) + this.cols.size(Math.min(r.c2, MAX_COLS - 1));
    const y2 = this.rowY(Math.min(r.r2, MAX_ROWS - 1)) + this.rows.size(Math.min(r.r2, MAX_ROWS - 1));
    return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
  }

  colAt(x: number): number {
    const sx = x - this.headerW;
    if (sx < this.frozenW) return this.cols.indexAt(Math.max(0, sx));
    return this.cols.indexAt(sx + this.scrollX);
  }

  rowAt(y: number): number {
    const sy = y - this.headerH;
    if (sy < this.frozenH) return this.rows.indexAt(Math.max(0, sy));
    return this.rows.indexAt(sy + this.scrollY);
  }

  private local(e: MouseEvent): { x: number; y: number } {
    const r = this.el.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  hitTest(x: number, y: number): Hit {
    this.ensureLayout();
    if (x > this.width || y > this.height) return { kind: 'none' };
    if (x < this.headerW && y < this.headerH) return { kind: 'corner' };
    if (y < this.headerH) {
      const col = this.colAt(x);
      const right = this.colX(col) + this.cols.size(col);
      const left = this.colX(col);
      let edge: number | null = null;
      if (Math.abs(x - right) <= 4) edge = col;
      else if (Math.abs(x - left) <= 4 && col > 0) edge = this.cols.step(col, -1);
      // A table's names (and buttons) show here once its header row scrolls away
      const sticky = edge === null ? this.stickyTable() : null;
      const tableButton = sticky && col >= sticky.range.c1 && col <= sticky.range.c2 ? this.tableButtonAt(sticky, col, { x: left, y: 0, w: right - left, h: this.headerH }, x, y) : null;
      return { kind: 'colHeader', col, edge, tableButton };
    }
    if (x < this.headerW) {
      const row = this.rowAt(y);
      const bottom = this.rowY(row) + this.rows.size(row);
      const top = this.rowY(row);
      let edge: number | null = null;
      if (Math.abs(y - bottom) <= 3) edge = row;
      else if (Math.abs(y - top) <= 3 && row > 0) edge = this.rows.step(row, -1);
      return { kind: 'rowHeader', row, edge };
    }
    const row = this.rowAt(y);
    const col = this.colAt(x);
    // Fill handle at the bottom-right of the selection
    const sel = this.app.sel;
    let fillHandle = false;
    if (sel.ranges.length === 1 && !this.app.isEditing && this.app.settings.bool('grid.fillHandle')) {
      const rr = this.rangeRect(sel.range);
      const hx = rr.x + rr.w;
      const hy = rr.y + rr.h;
      if (Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5) fillHandle = true;
    }
    let filterButton = false;
    const f = this.sheet.filter;
    if (f && row === f.range.r1 && col >= f.range.c1 && col <= f.range.c2) {
      const cr = this.cellRect(row, col);
      const bs = this.filterButtonSize();
      if (x >= cr.x + cr.w - bs - 3 && x <= cr.x + cr.w - 1) filterButton = true;
    }
    let tableButton: TableButtonHit | null = null;
    const table = fillHandle ? undefined : this.sheet.tables.find((t) => t.range.r1 === row && col >= t.range.c1 && col <= t.range.c2);
    if (table) tableButton = this.tableButtonAt(table, col, this.cellRect(row, col), x, y);
    return { kind: 'cell', row, col, fillHandle, filterButton, tableButton };
  }

  // ---------- tables ----------

  /** Where a table header cell's buttons go: the sort arrow, then the menu (left out when narrow). */
  private headerButtons(rect: Box): { sort: Box; menu: Box | null } {
    const bs = this.filterButtonSize();
    const y = rect.y + (rect.h - bs) / 2;
    const menu = rect.w >= bs * 3.6 ? { x: rect.x + rect.w - bs - 3, y, w: bs, h: bs } : null;
    const sort = { x: (menu ? menu.x - 2 : rect.x + rect.w - 3) - bs, y, w: bs, h: bs };
    return { sort, menu };
  }

  /** Width a table header cell needs at 100% zoom: its name in bold, plus the buttons. */
  tableHeaderWidth(name: string): number {
    const font = `700 ${((11 * 96) / 72).toFixed(2)}px "Calibri", Calibri, "Segoe UI", Arial, sans-serif`;
    return Math.ceil(this.measure(font, name) + PAD * 2 + 16 * 2 + 12);
  }

  /** Width the header buttons take from a header cell's text. */
  tableButtonsWidth(cellW: number): number {
    const bs = this.filterButtonSize();
    return cellW >= bs * 3.6 ? bs * 2 + 8 : bs + 6;
  }

  private tableButtonAt(t: TableDef, col: number, rect: Box, x: number, y: number): TableButtonHit | null {
    const { sort, menu } = this.headerButtons(rect);
    const inside = (b: Box) => x >= b.x - 1 && x <= b.x + b.w + 1 && y >= b.y - 2 && y <= b.y + b.h + 2;
    if (inside(sort)) return { table: t, col, part: 'sort', rect: sort };
    if (menu && inside(menu)) return { table: t, col, part: 'menu', rect: menu };
    return null;
  }

  /** The table whose header row has scrolled out of view (its names then replace the column letters). */
  stickyTable(): TableDef | null {
    const s = this.sheet;
    if (!s.tables.length) return null;
    const top = this.rows.indexAt(this.frozenH + this.scrollY);
    const list = s.tables.filter((t) => t.range.r1 >= s.freeze.rows && t.range.r1 < top && t.range.r2 >= top);
    if (!list.length) return null;
    const a = this.app.sel.active;
    return list.find((t) => a.col >= t.range.c1 && a.col <= t.range.c2 && a.row >= t.range.r1 && a.row <= t.range.r2) ?? list[0];
  }

  private staleCache = new Map<string, { key: string; stale: boolean }>();

  /** Whether a sorted table's rows changed since (cached until the next edit). */
  isTableStale(t: TableDef): boolean {
    if (!t.sort.length) return false;
    const key = `${this.app.doc.history.version}|${t.range.r2}|${t.sort.map((k) => `${k.col}${k.desc ? 'd' : 'a'}`).join(',')}`;
    const hit = this.staleCache.get(t.id);
    if (hit && hit.key === key) return hit.stale;
    const stale = sortIsStale(this.app.doc, this.sheet, t);
    this.staleCache.set(t.id, { key, stale });
    return stale;
  }

  private drawTableButtons(t: TableDef, col: number, rect: Box, bg: string): void {
    const ctx = this.ctx;
    const { sort, menu } = this.headerButtons(rect);
    const onDark = luminance(bg) < 0.45;
    const ink = onDark ? '#ffffff' : this.theme.headerText;
    const plate = (b: Box, strong: boolean) => {
      ctx.fillStyle = onDark ? `rgba(255,255,255,${strong ? 0.3 : 0.13})` : strong ? this.theme.accent + '33' : this.theme.dark ? '#2a2f38' : '#e8eaee';
      ctx.beginPath();
      ctx.roundRect(b.x, b.y, b.w, b.h, 3);
      ctx.fill();
    };
    const rank = t.sort.findIndex((k) => k.col === col);
    plate(sort, rank >= 0);
    const cx = sort.x + sort.w / 2;
    const cy = sort.y + sort.h / 2;
    ctx.fillStyle = ink;
    const tri = (x: number, y: number, up: boolean, size: number) => {
      ctx.beginPath();
      ctx.moveTo(x - size, up ? y + size * 0.55 : y - size * 0.55);
      ctx.lineTo(x + size, up ? y + size * 0.55 : y - size * 0.55);
      ctx.lineTo(x, up ? y - size * 0.75 : y + size * 0.75);
      ctx.closePath();
      ctx.fill();
    };
    if (rank < 0) {
      // Not sorted: a faint up-and-down arrow
      ctx.globalAlpha = 0.7;
      tri(cx, cy - 3, true, 2.6);
      tri(cx, cy + 3.2, false, 2.6);
      ctx.globalAlpha = 1;
    } else {
      const multi = t.sort.length > 1;
      tri(multi ? cx + 2 : cx, cy, !t.sort[rank].desc, 3.6);
      if (multi) {
        ctx.font = `600 ${Math.round(8 * Math.max(1, this.zoom))}px "Segoe UI", system-ui, sans-serif`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(String(rank + 1), sort.x + 4, cy + 0.5);
      }
      if (rank === 0 && this.isTableStale(t)) {
        // Rows changed since sorting: a dot says "click to sort again"
        ctx.fillStyle = '#ffc247';
        ctx.beginPath();
        ctx.arc(sort.x + sort.w - 1.5, sort.y + 1.5, 3, 0, Math.PI * 2);
        ctx.fill();
      }
    }
    if (menu) {
      const filtered = !!t.filter[col];
      plate(menu, filtered);
      const mx = menu.x + menu.w / 2;
      const my = menu.y + menu.h / 2;
      ctx.strokeStyle = ink;
      ctx.fillStyle = ink;
      ctx.lineWidth = 1.4;
      ctx.beginPath();
      if (filtered) {
        ctx.moveTo(mx - 4, my - 3.5);
        ctx.lineTo(mx + 4, my - 3.5);
        ctx.lineTo(mx + 1, my + 0.5);
        ctx.lineTo(mx + 1, my + 4);
        ctx.lineTo(mx - 1, my + 3);
        ctx.lineTo(mx - 1, my + 0.5);
        ctx.closePath();
        ctx.fill();
      } else {
        ctx.moveTo(mx - 3.5, my - 1.5);
        ctx.lineTo(mx, my + 2);
        ctx.lineTo(mx + 3.5, my - 1.5);
        ctx.stroke();
      }
    }
  }

  /** A grid rectangle in window coordinates (to open a menu next to it). */
  screenRect(r: Box): DOMRect {
    const host = this.el.getBoundingClientRect();
    return new DOMRect(host.left + r.x, host.top + r.y, r.w, r.h);
  }

  private filterButtonSize(): number {
    return Math.max(14, Math.min(18, Math.round(16 * this.zoom)));
  }

  // ---------- scrolling helpers ----------

  /** Scrolls so the cell is visible. */
  scrollToCell(row: number, col: number): void {
    // The viewport may not have been measured yet (right after startup)
    if (!this.width || !this.height) {
      this.width = this.scroller.clientWidth;
      this.height = this.scroller.clientHeight;
      if (!this.width || !this.height) return;
    }
    this.ensureLayout();
    const fr = this.sheet.freeze.rows;
    const fc = this.sheet.freeze.cols;
    const viewW = this.width - this.headerW - this.frozenW;
    const viewH = this.height - this.headerH - this.frozenH;
    if (col >= fc) {
      const left = this.cols.offset(col) - this.frozenW;
      const right = left + this.cols.size(col);
      if (left < this.scrollX) this.scroller.scrollLeft = left;
      else if (right > this.scrollX + viewW) this.scroller.scrollLeft = Math.max(0, right - viewW);
    }
    if (row >= fr) {
      const top = this.rows.offset(row) - this.frozenH;
      const bottom = top + this.rows.size(row);
      if (top < this.scrollY) this.scroller.scrollTop = top;
      else if (bottom > this.scrollY + viewH) this.scroller.scrollTop = Math.max(0, bottom - viewH);
    }
    if (row + 50 > this.extentRows || col + 20 > this.extentCols) {
      this.updateExtent();
      this.scrollToCellNoExtent(row, col);
    }
  }

  private scrollToCellNoExtent(row: number, col: number): void {
    const viewW = this.width - this.headerW - this.frozenW;
    const viewH = this.height - this.headerH - this.frozenH;
    if (col >= this.sheet.freeze.cols) {
      const right = this.cols.offset(col) - this.frozenW + this.cols.size(col);
      if (right > this.scrollX + viewW) this.scroller.scrollLeft = Math.max(0, right - viewW);
    }
    if (row >= this.sheet.freeze.rows) {
      const bottom = this.rows.offset(row) - this.frozenH + this.rows.size(row);
      if (bottom > this.scrollY + viewH) this.scroller.scrollTop = Math.max(0, bottom - viewH);
    }
  }

  /** Scrolls so the cell appears near the top-left (used by Go To). */
  scrollToCellTop(row: number, col: number): void {
    if (!this.width || !this.height) {
      this.width = this.scroller.clientWidth;
      this.height = this.scroller.clientHeight;
    }
    this.ensureLayout();
    const s = this.sheet;
    if (row + 60 > this.extentRows || col + 20 > this.extentCols) {
      this.extentRows = Math.max(this.extentRows, row + 120);
      this.extentCols = Math.max(this.extentCols, col + 30);
      this.updateExtentFrom();
    }
    if (row >= s.freeze.rows) this.scroller.scrollTop = Math.max(0, this.rows.offset(Math.max(s.freeze.rows, row - 2)) - this.frozenH);
    if (col >= s.freeze.cols && this.cols.offset(col + 1) - this.scrollX > this.width - this.headerW) {
      this.scroller.scrollLeft = Math.max(0, this.cols.offset(Math.max(s.freeze.cols, col - 1)) - this.frozenW);
    }
  }

  private updateExtentFrom(): void {
    const w = this.headerW + this.cols.offset(this.extentCols);
    const h = this.headerH + this.rows.offset(this.extentRows);
    this.spacer.style.width = `${Math.min(w, 30_000_000)}px`;
    this.spacer.style.height = `${Math.min(h, 30_000_000)}px`;
  }

  /** Number of rows that fit in the scrollable area (for PageUp/PageDown). */
  pageRows(): number {
    this.ensureLayout();
    return Math.max(1, Math.floor((this.height - this.headerH - this.frozenH) / Math.max(1, this.rows.def)) - 1);
  }

  pageCols(): number {
    this.ensureLayout();
    return Math.max(1, Math.floor((this.width - this.headerW - this.frozenW) / Math.max(1, this.cols.def)) - 1);
  }

  visibleRange(): Range {
    this.ensureLayout();
    const r1 = this.rows.indexAt(this.frozenH + this.scrollY);
    const c1 = this.cols.indexAt(this.frozenW + this.scrollX);
    const r2 = this.rows.indexAt(this.scrollY + this.height - this.headerH);
    const c2 = this.cols.indexAt(this.scrollX + this.width - this.headerW);
    return { r1, c1, r2, c2 };
  }

  // ---------- drawing ----------

  invalidate(): void {
    this.layoutDirty = true;
    this.draw();
  }

  draw(): void {
    if (this.drawQueued) return;
    this.drawQueued = true;
    requestAnimationFrame(() => {
      this.drawQueued = false;
      this.render();
    });
  }

  /** Draws immediately (used after layout changes that must be measured right away). */
  render(): void {
    if (!this.width || !this.height) return;
    this.ensureLayout();
    const ctx = this.ctx;
    const t = this.theme;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = t.cellBg;
    ctx.fillRect(0, 0, this.width, this.height);

    const s = this.sheet;
    const fr = s.freeze.rows;
    const fc = s.freeze.cols;
    const W = this.width;
    const H = this.height;
    const hx = this.headerW;
    const hy = this.headerH;
    const scrollR1 = this.rows.indexAt(this.frozenH + this.scrollY);
    const scrollC1 = this.cols.indexAt(this.frozenW + this.scrollX);
    const lastR = Math.min(MAX_ROWS - 1, this.rows.indexAt(this.scrollY + H - hy) + 1);
    const lastC = Math.min(MAX_COLS - 1, this.cols.indexAt(this.scrollX + W - hx) + 1);

    const quads: { r: [number, number]; c: [number, number]; clip: [number, number, number, number] }[] = [];
    const fx = hx + this.frozenW;
    const fy = hy + this.frozenH;
    quads.push({ r: [scrollR1, lastR], c: [scrollC1, lastC], clip: [fx, fy, W - fx, H - fy] });
    if (fr > 0) quads.push({ r: [0, fr - 1], c: [scrollC1, lastC], clip: [fx, hy, W - fx, this.frozenH] });
    if (fc > 0) quads.push({ r: [scrollR1, lastR], c: [0, fc - 1], clip: [hx, fy, this.frozenW, H - fy] });
    if (fr > 0 && fc > 0) quads.push({ r: [0, fr - 1], c: [0, fc - 1], clip: [hx, hy, this.frozenW, this.frozenH] });

    for (const q of quads) {
      if (q.clip[2] <= 0 || q.clip[3] <= 0) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(q.clip[0], q.clip[1], q.clip[2], q.clip[3]);
      ctx.clip();
      this.drawCells(q.r[0], q.r[1], q.c[0], q.c[1], q.clip);
      this.drawSelection();
      ctx.restore();
    }
    // Frozen pane separators
    if (this.app.settings.bool('grid.frozenLine')) {
      ctx.fillStyle = t.frozenLine;
      if (fr > 0) ctx.fillRect(hx, fy - 1, W - hx, 1.5);
      if (fc > 0) ctx.fillRect(fx - 1, hy, 1.5, H - hy);
    }
    this.drawHeaders(scrollR1, lastR, scrollC1, lastC);
    this.app.charts?.position();
    this.app.buttons?.position();
    this.app.validation?.position();
  }

  private font(st: Style): string {
    const size = ((st.fontSize ?? 11) * 96) / 72 * this.zoom;
    const fam = st.fontName ?? 'Calibri';
    return `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}${size.toFixed(2)}px "${fam}", Calibri, "Segoe UI", Arial, sans-serif`;
  }

  private measure(font: string, text: string): number {
    const key = font + '\u0000' + text;
    let w = this.measureCache.get(key);
    if (w === undefined) {
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      if (this.measureCache.size > 20000) this.measureCache.clear();
      this.measureCache.set(key, w);
    }
    return w;
  }

  private cellValue(row: number, col: number): Scalar {
    return this.app.doc.engine.getValue(this.sheet, row, col);
  }

  private drawCells(r1: number, r2: number, c1: number, c2: number, clip: [number, number, number, number]): void {
    const ctx = this.ctx;
    const t = this.theme;
    const s = this.sheet;
    const wb = this.app.doc.wb;
    const cf = this.app.cf;
    const showGrid = s.showGrid;
    const clipR = clip[0] + clip[2];
    const clipB = clip[1] + clip[3];

    // Pre-compute visible column positions
    const colPos: { c: number; x: number; w: number }[] = [];
    for (let c = c1; c <= c2; c++) {
      const w = this.cols.size(c);
      if (w === 0) continue;
      const x = this.colX(c);
      if (x > clipR) break;
      colPos.push({ c, x, w });
    }
    const rowPos: { r: number; y: number; h: number }[] = [];
    for (let r = r1; r <= r2; r++) {
      const h = this.rows.size(r);
      if (h === 0) continue;
      const y = this.rowY(r);
      if (y > clipB) break;
      rowPos.push({ r, y, h });
    }

    // Gridlines
    if (showGrid) {
      ctx.fillStyle = t.gridLine;
      for (const cp of colPos) ctx.fillRect(Math.round(cp.x + cp.w) - 1, clip[1], 1, clip[3]);
      for (const rp of rowPos) ctx.fillRect(clip[0], Math.round(rp.y + rp.h) - 1, clip[2], 1);
    }

    const merges = s.merges.filter((m) => m.r2 >= r1 && m.r1 <= r2 && m.c2 >= c1 && m.c1 <= c2);
    const inMerge = (r: number, c: number) => merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
    const tables = s.tables.filter((tb) => tb.range.r2 >= r1 && tb.range.r1 <= r2 && tb.range.c2 >= c1 && tb.range.c1 <= c2);
    const tableIn = (r: number, c: number) => (tables.length ? tables.find((tb) => r >= tb.range.r1 && r <= tb.range.r2 && c >= tb.range.c1 && c <= tb.range.c2) : undefined);
    const headerBgs = new Map<string, string>();
    // Bands count the rows you can see, so a filtered table still alternates
    const bands = new Map<string, number>();
    for (const tb of tables) {
      let idx = -1;
      for (const rp of rowPos) {
        if (rp.r <= tb.range.r1 || rp.r > tb.range.r2) continue;
        if (idx < 0) {
          idx = rp.r - tb.range.r1 - 1;
          for (const hr of s.filterHidden) if (hr > tb.range.r1 && hr < rp.r) idx--;
          for (const hr of s.hiddenRows) if (hr > tb.range.r1 && hr < rp.r && !s.filterHidden.has(hr)) idx--;
        }
        bands.set(`${tb.id}:${rp.r}`, idx++);
      }
    }

    interface TextJob {
      r: number;
      c: number;
      x: number;
      y: number;
      w: number;
      h: number;
      st: Style;
      v: Scalar;
      bg: string;
      merged: boolean;
      cfStyle?: { color?: string; bold?: boolean; italic?: boolean; strike?: boolean; underline?: boolean };
      hideValue?: boolean;
      icon?: { set: string; index: number; count: number };
    }
    const jobs: TextJob[] = [];
    const borders: { x: number; y: number; w: number; h: number; st: Style }[] = [];
    const marks: { x: number; y: number; w: number; h: number; note: boolean; error: boolean }[] = [];
    const showNotes = this.app.settings.bool('grid.noteMarkers');
    const showErrors = this.app.settings.bool('grid.errorMarkers');

    const visit = (r: number, c: number, x: number, y: number, w: number, h: number, merged: boolean) => {
      const cell = s.cells.get(cellKey(r, c));
      const sid = cell?.s ?? s.rowStyles.get(r) ?? s.colStyles.get(c);
      const st = wb.styles.get(sid);
      const res = cf.evaluate(s, r, c);
      // Light fills (the cell's own or from conditional formatting) become dark tints in dark mode
      const cfFill = res?.style?.fill ?? res?.fill;
      const fillHex = cfFill ?? st.fill;
      // Tables: a colored header row and banded rows, under any fill of the cells' own
      const tb = tableIn(r, c);
      const isHead = !!tb && r === tb.range.r1;
      const band = tb && !isHead ? (bands.get(`${tb.id}:${r}`) ?? r - tb.range.r1 - 1) : 0;
      const tableFill = tb && !fillHex ? (isHead ? TABLE_COLORS[tb.color].header : tb.banded && band % 2 === 0 ? TABLE_COLORS[tb.color].band : undefined) : undefined;
      let bg = t.cellBg;
      if (fillHex || tableFill) {
        // In dark mode bands are a faint wash of the table's color
        bg = isHead && tableFill ? tableFill : !fillHex && t.dark ? mixHex(TABLE_COLORS[tb!.color].header, t.cellBg, 0.14) : displayFill(t, (fillHex ?? tableFill)!);
        ctx.fillStyle = bg;
        ctx.fillRect(x, y, w - (merged ? 0 : 1), h - (merged ? 0 : 1));
      } else if (merged) {
        ctx.fillStyle = t.cellBg;
        ctx.fillRect(x, y, w - 1, h - 1);
      }
      // The column a table is sorted by is tinted, so you can tell at a glance
      if (tb && !isHead && tb.sort[0]?.col === c) {
        ctx.fillStyle = t.accent;
        ctx.globalAlpha = t.dark ? 0.1 : 0.07;
        ctx.fillRect(x, y, w - 1, h - 1);
        ctx.globalAlpha = 1;
      }
      if (isHead) headerBgs.set(`${tb!.id}:${c}`, bg);
      if (res?.bar) {
        const b = res.bar;
        const bx = x + 2 + (w - 5) * b.from;
        const bw = Math.max(1, (w - 5) * (b.to - b.from));
        const color = b.negative ? '#ff6b6b' : b.color;
        if (b.gradient) {
          const g = ctx.createLinearGradient(bx, 0, bx + bw, 0);
          g.addColorStop(0, color);
          g.addColorStop(1, color + '33');
          ctx.fillStyle = g;
        } else ctx.fillStyle = color;
        ctx.globalAlpha = 0.85;
        ctx.fillRect(bx, y + 3, bw, h - 7);
        ctx.globalAlpha = 1;
      }
      if (st.borderTop || st.borderBottom || st.borderLeft || st.borderRight) borders.push({ x, y, w, h, st });
      const v = this.app.showFormulas && cell?.f !== undefined ? '=' + cell.f : this.cellValue(r, c);
      if (v !== null && v !== '' || res?.icon) {
        // Table functions (AMORTIZATION, STOCKHISTORY…) suggest formats for their columns
        const spillFmt = !st.numFmt && typeof v === 'number' ? this.app.doc.engine.spillFormatAt(s, r, c) : undefined;
        let jst = spillFmt ? { ...st, numFmt: spillFmt } : st;
        // Header names are bold (white on the table's color), and leave room for the buttons
        if (isHead) jst = { ...jst, bold: true, color: tableFill ? (st.color ?? '#ffffff') : st.color };
        const tw = isHead ? Math.max(8, w - this.tableButtonsWidth(w)) : w;
        jobs.push({ r, c, x, y, w: tw, h, st: jst, v, bg, merged, cfStyle: res?.style, hideValue: res?.hideValue, icon: res?.icon });
      }
      if (cell?.note && showNotes) marks.push({ x, y, w, h, note: true, error: false });
      if (v instanceof ErrorValue && v.code !== '#BUSY!' && showErrors) marks.push({ x, y, w, h, note: false, error: true });
    };

    for (const rp of rowPos) {
      for (const cp of colPos) {
        const m = inMerge(rp.r, cp.c);
        if (m) {
          if (rp.r !== Math.max(m.r1, r1) || cp.c !== Math.max(m.c1, c1)) continue;
          // Draw the merged block once, anchored at its real top-left
          const rect = this.cellRect(m.r1, m.c1);
          visit(m.r1, m.c1, rect.x, rect.y, rect.w, rect.h, true);
          continue;
        }
        visit(rp.r, cp.c, cp.x, rp.y, cp.w, rp.h, false);
      }
    }

    // Text that overflows from cells left of the visible area
    if (colPos.length && c1 > 0) {
      for (const rp of rowPos) {
        for (let c = c1 - 1; c >= Math.max(0, c1 - 12); c--) {
          const cell = s.cells.get(cellKey(rp.r, c));
          if (!cell) continue;
          const v = this.cellValue(rp.r, c);
          if (typeof v !== 'string' || !v) break;
          const st = wb.styles.get(cell.s ?? s.rowStyles.get(rp.r) ?? s.colStyles.get(c));
          if (st.wrap || (st.hAlign && st.hAlign !== 'left' && st.hAlign !== 'general')) break;
          if (inMerge(rp.r, c) || s.mergeAt(rp.r, c)) break;
          const x = this.colX(c);
          jobs.unshift({ r: rp.r, c, x, y: rp.y, w: this.cols.size(c), h: rp.h, st, v, bg: t.cellBg, merged: false });
          break;
        }
      }
    }

    // Borders
    for (const b of borders) this.drawBorders(b.x, b.y, b.w, b.h, b.st);

    // Text
    for (const j of jobs) this.drawText(j, clip, (r, c) => {
      if (inMerge(r, c)) return false;
      const cell = s.cells.get(cellKey(r, c));
      if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''))) return false;
      if (this.app.doc.engine.spillAnchorOf(s, r, c)) return false;
      return true;
    });

    // Note & error markers
    for (const m of marks) {
      ctx.fillStyle = m.note ? t.noteMark : t.errorMark;
      ctx.beginPath();
      const sz = Math.max(5, 6 * this.zoom);
      if (m.note) {
        ctx.moveTo(m.x + m.w - 1 - sz, m.y);
        ctx.lineTo(m.x + m.w - 1, m.y);
        ctx.lineTo(m.x + m.w - 1, m.y + sz);
      } else {
        ctx.moveTo(m.x, m.y);
        ctx.lineTo(m.x + sz, m.y);
        ctx.lineTo(m.x, m.y + sz);
      }
      ctx.closePath();
      ctx.fill();
    }

    // Tables: an outline, and the buttons in the header row
    for (const tb of tables) {
      const look = TABLE_COLORS[tb.color];
      const rr = this.rangeRect(tb.range);
      ctx.strokeStyle = t.dark ? look.header : look.line;
      ctx.lineWidth = 1;
      ctx.strokeRect(Math.round(rr.x) - 0.5, Math.round(rr.y) - 0.5, Math.round(rr.w), Math.round(rr.h));
      if (tb.range.r1 < r1 || tb.range.r1 > r2) continue;
      for (let c = Math.max(tb.range.c1, c1); c <= Math.min(tb.range.c2, c2); c++) {
        if (this.cols.size(c) === 0) continue;
        this.drawTableButtons(tb, c, this.cellRect(tb.range.r1, c), headerBgs.get(`${tb.id}:${c}`) ?? TABLE_COLORS[tb.color].header);
      }
    }

    // Filter buttons
    const f = s.filter;
    if (f && f.range.r1 >= r1 && f.range.r1 <= r2) {
      for (let c = Math.max(f.range.c1, c1); c <= Math.min(f.range.c2, c2); c++) {
        if (this.cols.size(c) === 0) continue;
        const rect = this.cellRect(f.range.r1, c);
        this.drawFilterButton(rect, !!f.conds[c]);
      }
    }
  }

  private drawFilterButton(rect: { x: number; y: number; w: number; h: number }, active: boolean): void {
    const ctx = this.ctx;
    const t = this.theme;
    const bs = this.filterButtonSize();
    const bx = rect.x + rect.w - bs - 3;
    const by = rect.y + (rect.h - bs) / 2;
    ctx.fillStyle = active ? t.accent : t.dark ? '#2a2f38' : '#e8eaee';
    ctx.beginPath();
    ctx.roundRect(bx, by, bs, bs, 3);
    ctx.fill();
    ctx.strokeStyle = active ? (t.dark ? '#0b1220' : '#ffffff') : t.headerText;
    ctx.fillStyle = ctx.strokeStyle;
    ctx.lineWidth = 1.4;
    const cx = bx + bs / 2;
    const cy = by + bs / 2;
    if (active) {
      // funnel
      ctx.beginPath();
      ctx.moveTo(cx - 4, cy - 3.5);
      ctx.lineTo(cx + 4, cy - 3.5);
      ctx.lineTo(cx + 1, cy + 0.5);
      ctx.lineTo(cx + 1, cy + 4);
      ctx.lineTo(cx - 1, cy + 3);
      ctx.lineTo(cx - 1, cy + 0.5);
      ctx.closePath();
      ctx.fill();
    } else {
      ctx.beginPath();
      ctx.moveTo(cx - 3.5, cy - 1.5);
      ctx.lineTo(cx, cy + 2);
      ctx.lineTo(cx + 3.5, cy - 1.5);
      ctx.stroke();
    }
  }

  private borderColor(side: BorderSide, bg: string): string {
    return displayText(this.theme, side.color, bg);
  }

  private drawBorders(x: number, y: number, w: number, h: number, st: Style): void {
    const ctx = this.ctx;
    const bg = this.theme.cellBg;
    const line = (side: BorderSide, x1: number, y1: number, x2: number, y2: number) => {
      const width = side.style === 'thick' ? 3 : side.style.startsWith('medium') ? 2 : 1;
      ctx.strokeStyle = this.borderColor(side, bg);
      ctx.lineWidth = width;
      ctx.setLineDash(
        side.style.includes('dashDotDot') ? [6, 2, 2, 2, 2, 2] : side.style.includes('dashDot') ? [6, 2, 2, 2] : side.style.includes('ashed') ? [5, 3] : side.style === 'dotted' || side.style === 'hair' ? [1, 2] : [],
      );
      const off = width % 2 ? 0.5 : 0;
      if (side.style === 'double') {
        ctx.lineWidth = 1;
        const horizontal = y1 === y2;
        ctx.beginPath();
        if (horizontal) {
          ctx.moveTo(x1, Math.round(y1) - 1.5);
          ctx.lineTo(x2, Math.round(y2) - 1.5);
          ctx.moveTo(x1, Math.round(y1) + 0.5);
          ctx.lineTo(x2, Math.round(y2) + 0.5);
        } else {
          ctx.moveTo(Math.round(x1) - 1.5, y1);
          ctx.lineTo(Math.round(x2) - 1.5, y2);
          ctx.moveTo(Math.round(x1) + 0.5, y1);
          ctx.lineTo(Math.round(x2) + 0.5, y2);
        }
        ctx.stroke();
        return;
      }
      ctx.beginPath();
      if (y1 === y2) {
        ctx.moveTo(x1, Math.round(y1) - off);
        ctx.lineTo(x2, Math.round(y2) - off);
      } else {
        ctx.moveTo(Math.round(x1) - off, y1);
        ctx.lineTo(Math.round(x2) - off, y2);
      }
      ctx.stroke();
    };
    const R = x + w;
    const B = y + h;
    if (st.borderTop) line(st.borderTop, x - 1, y, R, y);
    if (st.borderBottom) line(st.borderBottom, x - 1, B, R, B);
    if (st.borderLeft) line(st.borderLeft, x, y - 1, x, B);
    if (st.borderRight) line(st.borderRight, R, y - 1, R, B);
    ctx.setLineDash([]);
  }

  private drawText(
    j: {
      r: number;
      c: number;
      x: number;
      y: number;
      w: number;
      h: number;
      st: Style;
      v: Scalar;
      bg: string;
      merged: boolean;
      cfStyle?: { color?: string; bold?: boolean; italic?: boolean; strike?: boolean; underline?: boolean };
      hideValue?: boolean;
      icon?: { set: string; index: number; count: number };
    },
    clip: [number, number, number, number],
    isEmpty: (r: number, c: number) => boolean,
  ): void {
    const ctx = this.ctx;
    const t = this.theme;
    const st0 = j.st;
    const st: Style = j.cfStyle
      ? { ...st0, bold: j.cfStyle.bold || st0.bold, italic: j.cfStyle.italic || st0.italic, strike: j.cfStyle.strike || st0.strike, underline: j.cfStyle.underline || st0.underline }
      : st0;
    const font = this.font(st);
    const fontPx = ((st.fontSize ?? 11) * 96) / 72 * this.zoom;
    const v = j.v;
    const isFormulaText = this.app.showFormulas && typeof v === 'string' && v.startsWith('=') && this.sheet.getCell(j.r, j.c)?.f !== undefined;
    let iconW = 0;
    if (j.icon) {
      iconW = Math.round(fontPx * 1.1) + 4;
      this.drawIcon(j.icon, j.x + 3, j.y + j.h / 2, fontPx * 0.95);
    }
    if (j.hideValue || v === null) return;

    let text: string;
    let color: string | undefined;
    const isNum = typeof v === 'number' && !isFormulaText;
    const avail = j.w - PAD * 2 - iconW;
    if (isNum) {
      if (v === 0 && !this.app.settings.bool('grid.showZeros')) return;
      const digitW = this.measure(font, '0');
      const maxChars = Math.max(1, Math.floor(avail / Math.max(1, digitW)));
      const f = formatValue(v, st.numFmt, Math.min(11, maxChars));
      text = f.text;
      color = f.color;
      // Negative numbers in red, unless the cell or its format already chooses a color
      if (v < 0 && !color && !st.color && !j.cfStyle?.color && this.app.settings.bool('grid.negativeRed')) color = t.dark ? '#ff7b7b' : '#d93025';
    } else if (v instanceof ErrorValue) {
      text = v.code === '#BUSY!' ? (this.app.settings.str('grid.loadingText') === 'busy' ? '#BUSY!' : 'Loading…') : v.code;
      color = v.code === '#BUSY!' ? t.headerText : t.errorMark;
    } else if (typeof v === 'boolean') text = v ? 'TRUE' : 'FALSE';
    else {
      const f = isFormulaText ? { text: String(v) } : formatValue(v, st.numFmt);
      text = f.text;
      color = (f as { color?: string }).color;
    }
    const cell = this.sheet.getCell(j.r, j.c);
    const isLink = !!cell?.link;
    const textColor = j.cfStyle?.color
      ? displayText(t, j.cfStyle.color, j.bg)
      : color && !isNum
        ? color
        : color
          ? displayText(t, color, j.bg)
          : isLink && !st.color
            ? t.linkText
            : displayText(t, st.color, j.bg);

    // Horizontal alignment
    let align: 'left' | 'center' | 'right' = 'left';
    const ha = st.hAlign ?? 'general';
    if (ha === 'center' || ha === 'centerContinuous') align = 'center';
    else if (ha === 'right') align = 'right';
    else if (ha === 'left' || ha === 'fill' || ha === 'justify' || ha === 'distributed') align = 'left';
    else if (isFormulaText) align = 'left';
    else if (typeof v === 'number') align = 'right';
    else if (typeof v === 'boolean' || v instanceof ErrorValue) align = 'center';
    const indent = (st.indent ?? 0) * 9 * this.zoom;

    ctx.font = font;
    ctx.fillStyle = textColor;
    const lineH = Math.round(fontPx * 1.22);
    const va = st.vAlign ?? 'bottom';

    if (st.wrap && !isNum) {
      const lines = this.wrapLines(font, text, Math.max(4, avail - indent));
      const blockH = lines.length * lineH;
      let ty = va === 'top' ? j.y + 2 : va === 'middle' ? j.y + (j.h - blockH) / 2 : j.y + j.h - blockH - 2;
      ctx.save();
      ctx.beginPath();
      ctx.rect(j.x, j.y, j.w - 1, j.h - 1);
      ctx.clip();
      ctx.textBaseline = 'top';
      for (const line of lines) {
        const lw = this.measure(font, line);
        const tx = align === 'left' ? j.x + PAD + indent + iconW : align === 'right' ? j.x + j.w - PAD - lw - indent : j.x + (j.w - lw) / 2;
        ctx.fillText(line, tx, ty + (lineH - fontPx) / 2);
        this.decorate(st, tx, ty + (lineH - fontPx) / 2, lw, fontPx, isLink);
        ty += lineH;
      }
      ctx.restore();
      return;
    }

    let tw = this.measure(font, text);
    if (isNum && tw > avail) {
      // Numbers never spill into neighbours — show #### like Excel
      text = '#'.repeat(Math.max(1, Math.floor(avail / Math.max(1, this.measure(font, '#')))));
      tw = this.measure(font, text);
    }
    if (text.includes('\n')) text = text.replace(/\n/g, ' ');

    // Overflow into empty neighbours (text only, unmerged)
    let left = j.x;
    let right = j.x + j.w;
    if (!isNum && !j.merged && tw > avail) {
      if (align === 'left' || align === 'center') {
        let c = j.c + 1;
        let guard = 0;
        while (right - j.x - PAD * 2 < tw && c < MAX_COLS && guard++ < 50 && isEmpty(j.r, c)) {
          right += this.cols.size(c);
          c++;
        }
      }
      if (align === 'right' || align === 'center') {
        let c = j.c - 1;
        let guard = 0;
        while (j.x + j.w - left - PAD * 2 < tw && c >= 0 && guard++ < 50 && isEmpty(j.r, c)) {
          left -= this.cols.size(c);
          c--;
        }
      }
    }
    let tx: number;
    if (align === 'left') tx = j.x + PAD + indent + iconW;
    else if (align === 'right') tx = j.x + j.w - PAD - tw - indent;
    else tx = j.x + (j.w - tw) / 2;
    if (align === 'right' && iconW && tx < j.x + iconW + PAD) tx = j.x + iconW + PAD;

    let ty: number;
    if (va === 'top') {
      ctx.textBaseline = 'top';
      ty = j.y + 3;
    } else if (va === 'middle') {
      ctx.textBaseline = 'middle';
      ty = j.y + j.h / 2;
    } else {
      ctx.textBaseline = 'alphabetic';
      ty = j.y + j.h - Math.max(3, Math.round(fontPx * 0.28));
    }
    const needsClip = tw > right - left - PAD * 2 || left < clip[0] || right > clip[0] + clip[2];
    if (needsClip) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(left, j.y, right - left - 1, j.h);
      ctx.clip();
    }
    // Paint the overflow area so gridlines don't cut through the text
    if (right > j.x + j.w || left < j.x) {
      ctx.fillStyle = j.bg;
      if (right > j.x + j.w) ctx.fillRect(j.x + j.w - 1, j.y, right - j.x - j.w, j.h - 1);
      if (left < j.x) ctx.fillRect(left, j.y, j.x - left, j.h - 1);
      ctx.fillStyle = textColor;
    }
    ctx.fillText(text, tx, ty);
    const topY = va === 'top' ? ty : va === 'middle' ? ty - fontPx / 2 : ty - fontPx * 0.8;
    this.decorate(st, tx, topY, tw, fontPx, isLink);
    if (needsClip) ctx.restore();
  }

  private decorate(st: Style, x: number, top: number, w: number, fontPx: number, link: boolean): void {
    const ctx = this.ctx;
    if (st.underline || link) ctx.fillRect(x, Math.round(top + fontPx * 0.95), w, Math.max(1, fontPx / 14));
    if (st.strike) ctx.fillRect(x, Math.round(top + fontPx * 0.5), w, Math.max(1, fontPx / 14));
  }

  private wrapLines(font: string, text: string, width: number): string[] {
    const out: string[] = [];
    for (const para of text.split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        const test = line + w;
        if (this.measure(font, test) <= width || !line) {
          if (!line && this.measure(font, w) > width) {
            // Break long words
            let chunk = '';
            for (const ch of w) {
              if (this.measure(font, chunk + ch) > width && chunk) {
                out.push(chunk);
                chunk = ch;
              } else chunk += ch;
            }
            line = chunk;
          } else line = test;
        } else {
          out.push(line.trimEnd());
          line = w.trimStart();
        }
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  private drawIcon(icon: { set: string; index: number; count: number }, x: number, cy: number, size: number): void {
    const ctx = this.ctx;
    const n = icon.count;
    const i = icon.index;
    const frac = n <= 1 ? 1 : i / (n - 1);
    const colors = ['#ff5c5c', '#ffb547', '#3fcf8e'];
    const color = frac < 0.34 ? colors[0] : frac < 0.67 ? colors[1] : colors[2];
    const s = size;
    const x0 = x;
    const y0 = cy - s / 2;
    ctx.save();
    const set = icon.set;
    if (/Arrows/.test(set)) {
      const gray = /Gray/.test(set);
      ctx.fillStyle = gray ? '#9aa1ad' : color;
      ctx.translate(x0 + s / 2, cy);
      const angle = frac >= 0.99 ? -Math.PI / 2 : frac <= 0.01 ? Math.PI / 2 : n >= 4 && frac > 0.5 ? -Math.PI / 4 : n >= 4 && frac < 0.5 && frac > 0 ? Math.PI / 4 : 0;
      ctx.rotate(angle);
      ctx.beginPath();
      ctx.moveTo(s * 0.45, 0);
      ctx.lineTo(0, -s * 0.4);
      ctx.lineTo(0, -s * 0.15);
      ctx.lineTo(-s * 0.42, -s * 0.15);
      ctx.lineTo(-s * 0.42, s * 0.15);
      ctx.lineTo(0, s * 0.15);
      ctx.lineTo(0, s * 0.4);
      ctx.closePath();
      ctx.fill();
    } else if (/Flags/.test(set)) {
      ctx.fillStyle = color;
      ctx.fillRect(x0 + s * 0.15, y0 + s * 0.05, s * 0.1, s * 0.9);
      ctx.beginPath();
      ctx.moveTo(x0 + s * 0.25, y0 + s * 0.05);
      ctx.lineTo(x0 + s * 0.9, y0 + s * 0.28);
      ctx.lineTo(x0 + s * 0.25, y0 + s * 0.52);
      ctx.fill();
    } else if (/Rating|Boxes|Quarters/.test(set)) {
      const bars = 4;
      for (let k = 0; k < bars; k++) {
        ctx.fillStyle = k < Math.round(frac * bars) || (frac > 0 && k === 0) ? '#4ea1ff' : '#5a6170';
        const bh = (s * (k + 1)) / bars;
        ctx.fillRect(x0 + k * (s / bars), y0 + s - bh, s / bars - 1, bh);
      }
    } else if (/Stars/.test(set)) {
      ctx.fillStyle = frac > 0.66 ? '#ffd43b' : frac > 0.33 ? '#ffd43b88' : '#5a6170';
      ctx.beginPath();
      for (let k = 0; k < 10; k++) {
        const rad = k % 2 ? s * 0.22 : s * 0.5;
        const a = (Math.PI / 5) * k - Math.PI / 2;
        ctx.lineTo(x0 + s / 2 + Math.cos(a) * rad, cy + Math.sin(a) * rad);
      }
      ctx.closePath();
      ctx.fill();
    } else if (/Triangles/.test(set)) {
      ctx.fillStyle = color;
      ctx.beginPath();
      if (frac > 0.66) {
        ctx.moveTo(x0 + s / 2, y0 + s * 0.15);
        ctx.lineTo(x0 + s * 0.9, y0 + s * 0.85);
        ctx.lineTo(x0 + s * 0.1, y0 + s * 0.85);
      } else if (frac < 0.34) {
        ctx.moveTo(x0 + s / 2, y0 + s * 0.85);
        ctx.lineTo(x0 + s * 0.9, y0 + s * 0.15);
        ctx.lineTo(x0 + s * 0.1, y0 + s * 0.15);
      } else ctx.rect(x0 + s * 0.15, cy - s * 0.12, s * 0.7, s * 0.24);
      ctx.fill();
    } else {
      // Traffic lights, symbols, signs
      const fill = /RedToBlack/.test(set) ? ['#ff5c5c', '#c25555', '#8a5a5a', '#3a3f48'][Math.min(3, n - 1 - i)] : color;
      ctx.fillStyle = fill;
      ctx.beginPath();
      ctx.arc(x0 + s / 2, cy, s * 0.42, 0, Math.PI * 2);
      ctx.fill();
      if (/Symbols/.test(set)) {
        ctx.strokeStyle = '#ffffff';
        ctx.lineWidth = Math.max(1.5, s * 0.12);
        ctx.beginPath();
        if (frac > 0.66) {
          ctx.moveTo(x0 + s * 0.28, cy);
          ctx.lineTo(x0 + s * 0.45, cy + s * 0.16);
          ctx.lineTo(x0 + s * 0.72, cy - s * 0.16);
        } else if (frac < 0.34) {
          ctx.moveTo(x0 + s * 0.32, cy - s * 0.18);
          ctx.lineTo(x0 + s * 0.68, cy + s * 0.18);
          ctx.moveTo(x0 + s * 0.68, cy - s * 0.18);
          ctx.lineTo(x0 + s * 0.32, cy + s * 0.18);
        } else {
          ctx.moveTo(x0 + s / 2, cy - s * 0.22);
          ctx.lineTo(x0 + s / 2, cy + s * 0.05);
          ctx.moveTo(x0 + s / 2, cy + s * 0.16);
          ctx.lineTo(x0 + s / 2, cy + s * 0.2);
        }
        ctx.stroke();
      }
    }
    ctx.restore();
  }

  private drawSelection(): void {
    const ctx = this.ctx;
    const t = this.theme;
    const sel = this.app.sel;
    const s = this.sheet;

    // Spilled array border for the active cell
    const anchor = this.app.doc.engine.spillAnchorOf(s, sel.active.row, sel.active.col) ?? sel.active;
    const spill = this.app.settings.bool('grid.spillBorder') ? this.app.doc.engine.spillAt(s, anchor.row, anchor.col) : null;
    // Reading guide: a faint band across the active row and down the active column
    if (this.app.settings.bool('grid.crosshair') && !this.app.isEditing) {
      const a = this.cellRect(sel.active.row, sel.active.col);
      ctx.save();
      ctx.globalAlpha = 0.55;
      ctx.fillStyle = t.selFill;
      ctx.fillRect(this.headerW, a.y, this.width - this.headerW, a.h - 1);
      ctx.fillRect(a.x, this.headerH, a.w - 1, this.height - this.headerH);
      ctx.restore();
    }
    if (spill) {
      const r = this.rangeRect(spill);
      ctx.strokeStyle = t.spillBorder;
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
      ctx.setLineDash([]);
    }

    // Formula reference highlights while editing
    for (const h of this.refHighlights) {
      if (h.sheetId !== s.id) continue;
      const r = this.rangeRect(h.range);
      ctx.fillStyle = h.color + '1f';
      ctx.fillRect(r.x, r.y, r.w - 1, r.h - 1);
      ctx.strokeStyle = h.color;
      ctx.lineWidth = 2;
      ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
    }

    const editingRefs = this.refHighlights.length > 0 && this.app.isEditing;
    for (let i = 0; i < sel.ranges.length; i++) {
      const range = sel.ranges[i];
      const r = this.rangeRect(range);
      if (r.w <= 0 || r.h <= 0) continue;
      const single = range.r1 === range.r2 && range.c1 === range.c2;
      const merge = s.mergeAt(range.r1, range.c1);
      const isActiveOnly = single || (merge && merge.r1 === range.r1 && merge.c1 === range.c1 && merge.r2 === range.r2 && merge.c2 === range.c2);
      if (!isActiveOnly) {
        // Shade the range but keep the active cell clear, like Excel
        ctx.fillStyle = t.selFill;
        ctx.beginPath();
        ctx.rect(r.x, r.y, r.w - 1, r.h - 1);
        if (i === sel.ranges.length - 1) {
          const a = this.cellRect(sel.active.row, sel.active.col);
          ctx.rect(a.x, a.y, a.w - 1, a.h - 1);
        }
        ctx.fill('evenodd');
      }
      if (i === sel.ranges.length - 1 && !editingRefs) {
        ctx.strokeStyle = t.selBorder;
        ctx.lineWidth = 2;
        ctx.strokeRect(r.x - 0.5, r.y - 0.5, r.w, r.h);
      } else if (!editingRefs) {
        ctx.strokeStyle = t.selBorder + '99';
        ctx.lineWidth = 1;
        ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
      }
    }
    // Active cell
    if (!this.app.isEditing) {
      const a = this.cellRect(sel.active.row, sel.active.col);
      ctx.strokeStyle = t.selBorder;
      ctx.lineWidth = 2;
      ctx.strokeRect(a.x - 0.5, a.y - 0.5, a.w, a.h);
      // Fill handle
      if (sel.ranges.length === 1 && this.app.settings.bool('grid.fillHandle')) {
        const r = this.rangeRect(sel.range);
        const hs = 7;
        ctx.fillStyle = t.selBorder;
        ctx.strokeStyle = t.cellBg;
        ctx.lineWidth = 1.5;
        ctx.fillRect(r.x + r.w - hs / 2 - 1, r.y + r.h - hs / 2 - 1, hs, hs);
        ctx.strokeRect(r.x + r.w - hs / 2 - 1, r.y + r.h - hs / 2 - 1, hs, hs);
      }
    }

    if (this.fillPreview) {
      const r = this.rangeRect(this.fillPreview);
      ctx.strokeStyle = t.selBorder;
      ctx.lineWidth = 1.5;
      ctx.setLineDash([5, 3]);
      ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1);
      ctx.setLineDash([]);
    }

    if (this.copyMarquee && this.copyMarquee.sheetId === s.id) {
      const r = this.rangeRect(this.copyMarquee.range);
      ctx.strokeStyle = t.selBorder;
      ctx.lineWidth = 2;
      ctx.setLineDash([6, 4]);
      ctx.lineDashOffset = -this.marqueeOffset;
      ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 2, r.h - 2);
      ctx.setLineDash([]);
      ctx.lineDashOffset = 0;
    }
  }

  private drawHeaders(scrollR1: number, lastR: number, scrollC1: number, lastC: number): void {
    if (!this.headerW || !this.headerH) return;
    const ctx = this.ctx;
    const t = this.theme;
    const s = this.sheet;
    const sel = this.app.sel;
    const W = this.width;
    const H = this.height;
    const hx = this.headerW;
    const hy = this.headerH;
    const fullCols = new Set(sel.fullColumns());
    const fullRows = new Set(sel.fullRows());
    const highlight = this.app.settings.bool('grid.highlightHeadings');
    const colSelected = (c: number) => highlight && sel.ranges.some((r) => c >= r.c1 && c <= r.c2);
    const rowSelected = (r: number) => highlight && sel.ranges.some((rg) => r >= rg.r1 && r <= rg.r2);
    const fontPx = Math.max(10, Math.min(13, 11.5 * this.zoom));
    ctx.font = `${fontPx}px "Segoe UI", system-ui, sans-serif`;

    // Column header
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(0, 0, W, hy);
    const drawCol = (c: number) => {
      const w = this.cols.size(c);
      if (w === 0) return;
      const x = this.colX(c);
      const isSel = colSelected(c);
      if (isSel) {
        ctx.fillStyle = fullCols.has(c) ? t.selBorder + '40' : t.headerSelBg;
        pill(ctx, x + 2, 3, w - 4, hy - 6);
      }
      ctx.fillStyle = isSel ? t.headerSelText : t.headerText;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      if (w > 12) ctx.fillText(colToName(c), x + w / 2, hy / 2 + 0.5);
      ctx.fillStyle = t.headerLine;
      ctx.fillRect(Math.round(x + w) - 1, 4, 1, hy - 8);
    };
    const sticky = this.stickyTable();
    const drawSticky = (c: number) => {
      if (!sticky || c < sticky.range.c1 || c > sticky.range.c2) return;
      const w = this.cols.size(c);
      if (w === 0) return;
      const x = this.colX(c);
      const color = TABLE_COLORS[sticky.color].header;
      ctx.fillStyle = color;
      ctx.fillRect(x, 0, w - 1, hy);
      const name = sticky.columns[c - sticky.range.c1] ?? '';
      ctx.save();
      ctx.beginPath();
      ctx.rect(x + 4, 0, Math.max(0, w - this.tableButtonsWidth(w) - 6), hy);
      ctx.clip();
      ctx.fillStyle = '#ffffff';
      ctx.font = `600 ${fontPx}px "Segoe UI", system-ui, sans-serif`;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      ctx.fillText(name, x + 6, hy / 2 + 0.5);
      ctx.restore();
      ctx.font = `${fontPx}px "Segoe UI", system-ui, sans-serif`;
      this.drawTableButtons(sticky, c, { x, y: 0, w, h: hy }, color);
    };
    ctx.save();
    ctx.beginPath();
    ctx.rect(hx + this.frozenW, 0, W - hx - this.frozenW, hy);
    ctx.clip();
    for (let c = scrollC1; c <= lastC; c++) {
      drawCol(c);
      drawSticky(c);
    }
    ctx.restore();
    for (let c = 0; c < s.freeze.cols; c++) {
      drawCol(c);
      drawSticky(c);
    }
    // Hidden column markers
    ctx.fillStyle = t.selBorder;
    for (const c of s.hiddenCols) {
      if (c < scrollC1 - 1 || c > lastC + 1) continue;
      const x = this.colX(c);
      if (x >= hx) ctx.fillRect(x - 2, 4, 3, hy - 8);
    }

    // Row header
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(0, hy, hx, H - hy);
    const drawRow = (r: number) => {
      const h = this.rows.size(r);
      if (h === 0) return;
      const y = this.rowY(r);
      const isSel = rowSelected(r);
      if (isSel) {
        ctx.fillStyle = fullRows.has(r) ? t.selBorder + '40' : t.headerSelBg;
        pill(ctx, 3, y + 1, hx - 6, h - 2);
      }
      ctx.fillStyle = s.filterHidden.size && this.isFiltered(r) ? t.accent : isSel ? t.headerSelText : t.headerText;
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      if (h > 9) ctx.fillText(String(r + 1), hx / 2, y + h / 2 + 0.5);
      ctx.fillStyle = t.headerLine;
      ctx.fillRect(4, Math.round(y + h) - 1, hx - 8, 1);
    };
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, hy + this.frozenH, hx, H - hy - this.frozenH);
    ctx.clip();
    for (let r = scrollR1; r <= lastR; r++) drawRow(r);
    ctx.restore();
    for (let r = 0; r < s.freeze.rows; r++) drawRow(r);
    ctx.fillStyle = t.selBorder;
    for (const r of s.hiddenRows) {
      if (r < scrollR1 - 1 || r > lastR + 1) continue;
      const y = this.rowY(r);
      if (y >= hy) ctx.fillRect(4, y - 2, hx - 8, 3);
    }

    // Header borders & corner
    ctx.fillStyle = t.headerLine;
    ctx.fillRect(0, hy - 1, W, 1);
    ctx.fillRect(hx - 1, 0, 1, H);
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(0, 0, hx - 1, hy - 1);
    ctx.strokeStyle = t.headerText;
    ctx.globalAlpha = 0.45;
    ctx.lineWidth = 1.25;
    ctx.beginPath();
    ctx.roundRect(Math.round(hx / 2 - 5) + 0.5, Math.round(hy / 2 - 5) + 0.5, 10, 10, 2.5);
    ctx.stroke();
    ctx.globalAlpha = 1;
    ctx.textAlign = 'left';
  }

  private isFiltered(r: number): boolean {
    const f = this.sheet.filter;
    if (f && r > f.range.r1 && r <= f.range.r2) return true;
    return this.sheet.tables.some((t) => r > t.range.r1 && r <= t.range.r2 && Object.keys(t.filter).length > 0);
  }

  // ---------- copy marquee animation ----------

  setCopyMarquee(m: { sheetId: number; range: Range } | null): void {
    this.copyMarquee = m;
    if (this.marqueeTimer) {
      clearInterval(this.marqueeTimer);
      this.marqueeTimer = null;
    }
    if (m && !this.app.settings.bool('appearance.reduceMotion')) {
      this.marqueeTimer = window.setInterval(() => {
        this.marqueeOffset = (this.marqueeOffset + 1) % 10;
        this.draw();
      }, 60);
    }
    this.draw();
  }

  // ---------- mouse ----------

  private onWheel(e: WheelEvent): void {
    if (e.ctrlKey) {
      e.preventDefault();
      if (this.app.settings.bool('grid.ctrlWheelZoom')) this.app.zoomBy(e.deltaY < 0 ? 0.1 : -0.1);
      return;
    }
    const speed = this.app.settings.num('grid.scrollSpeed');
    if (speed !== 1) {
      e.preventDefault();
      const unit = e.deltaMode === 1 ? 20 : e.deltaMode === 2 ? this.height : 1;
      const dx = e.shiftKey && !e.deltaX ? e.deltaY : e.deltaX;
      const dy = e.shiftKey && !e.deltaX ? 0 : e.deltaY;
      this.scroller.scrollBy({ left: dx * unit * speed, top: dy * unit * speed });
    }
  }

  private onPointerDown(e: PointerEvent): void {
    if (e.button !== 0) return;
    const p = this.local(e);
    // Clicks on scrollbars
    if (p.x > this.width || p.y > this.height) return;
    const hit = this.hitTest(p.x, p.y);
    const sel = this.app.sel;
    const s = this.sheet;
    this.dragMoved = false;

    // Formula point mode: clicking cells inserts references
    if (this.app.isEditing && hit.kind === 'cell' && this.app.editor.canInsertRef()) {
      e.preventDefault();
      this.drag = { kind: 'point', anchor: { row: hit.row, col: hit.col } };
      this.app.editor.insertRef({ r1: hit.row, c1: hit.col, r2: hit.row, c2: hit.col }, s, e.shiftKey);
      this.scroller.setPointerCapture(e.pointerId);
      return;
    }
    if (this.app.isEditing) {
      if (!this.app.commitEdit()) {
        e.preventDefault();
        return;
      }
    }
    e.preventDefault();
    this.app.focusGrid();

    switch (hit.kind) {
      case 'corner':
        sel.selectAll();
        this.app.onSelectionChanged();
        return;
      case 'colHeader': {
        if (hit.tableButton) {
          this.app.tableButton(hit.tableButton, e.shiftKey, this.screenRect(hit.tableButton.rect));
          return;
        }
        if (hit.edge !== null) {
          const cols = sel.fullColumns().includes(hit.edge) ? sel.fullColumns() : [hit.edge];
          this.drag = { kind: 'resizeCol', col: hit.edge, startX: p.x, startW: this.cols.size(hit.edge) / this.zoom, cols };
          break;
        }
        if (e.shiftKey) {
          const from = sel.anchor.col;
          sel.ranges[sel.ranges.length - 1] = { r1: 0, c1: Math.min(from, hit.col), r2: MAX_ROWS - 1, c2: Math.max(from, hit.col) };
        } else {
          sel.selectColumns(hit.col, hit.col, e.ctrlKey || e.metaKey);
        }
        this.drag = { kind: 'cols', from: e.shiftKey ? sel.anchor.col : hit.col };
        this.app.onSelectionChanged();
        break;
      }
      case 'rowHeader': {
        if (hit.edge !== null) {
          const rows = sel.fullRows().includes(hit.edge) ? sel.fullRows() : [hit.edge];
          this.drag = { kind: 'resizeRow', row: hit.edge, startY: p.y, startH: this.rows.size(hit.edge) / this.zoom, rows };
          break;
        }
        if (e.shiftKey) {
          const from = sel.anchor.row;
          sel.ranges[sel.ranges.length - 1] = { r1: Math.min(from, hit.row), c1: 0, r2: Math.max(from, hit.row), c2: MAX_COLS - 1 };
        } else sel.selectRows(hit.row, hit.row, e.ctrlKey || e.metaKey);
        this.drag = { kind: 'rows', from: e.shiftKey ? sel.anchor.row : hit.row };
        this.app.onSelectionChanged();
        break;
      }
      case 'cell': {
        if (hit.tableButton) {
          this.app.tableButton(hit.tableButton, e.shiftKey, this.screenRect(hit.tableButton.rect));
          return;
        }
        if (hit.filterButton) {
          const rect = this.cellRect(hit.row, hit.col);
          const host = this.el.getBoundingClientRect();
          this.app.openFilterMenu(hit.col, new DOMRect(host.left + rect.x, host.top + rect.y, rect.w, rect.h));
          return;
        }
        if (hit.fillHandle) {
          this.drag = { kind: 'fill', src: { ...sel.range } };
          break;
        }
        const cell = s.getCell(hit.row, hit.col);
        if ((e.ctrlKey || e.metaKey) && cell?.link) {
          this.app.openLink(cell.link);
          return;
        }
        if (e.ctrlKey || e.metaKey) {
          const fv = s.getCell(hit.row, hit.col)?.f;
          if (fv && /^\s*HYPERLINK\s*\(/i.test(fv)) {
            this.app.openHyperlinkFormula(hit.row, hit.col);
            return;
          }
          sel.addCell(s, hit.row, hit.col);
        } else if (e.shiftKey) sel.extendTo(s, hit.row, hit.col);
        else sel.setCell(s, hit.row, hit.col);
        this.drag = { kind: 'select', additive: e.ctrlKey || e.metaKey };
        this.app.onSelectionChanged();
        break;
      }
    }
    if (this.drag) this.scroller.setPointerCapture(e.pointerId);
  }

  private onPointerMove(e: PointerEvent): void {
    const p = this.local(e);
    if (!this.drag) {
      this.updateCursor(p.x, p.y);
      if (e.buttons === 0) this.app.hoverCell(p.x, p.y, e);
      return;
    }
    this.dragMoved = true;
    this.handleDrag(p.x, p.y);
    // Auto-scroll near the edges
    const margin = 24;
    let dx = 0;
    let dy = 0;
    if (this.drag.kind !== 'resizeCol' && this.drag.kind !== 'resizeRow') {
      if (p.x > this.width - margin) dx = Math.min(40, (p.x - this.width + margin) / 2);
      else if (p.x < this.headerW + this.frozenW && this.drag.kind !== 'rows' && this.scrollX > 0) dx = -Math.min(40, (this.headerW + this.frozenW - p.x) / 2 + 2);
      if (p.y > this.height - margin) dy = Math.min(40, (p.y - this.height + margin) / 2);
      else if (p.y < this.headerH + this.frozenH && this.drag.kind !== 'cols' && this.scrollY > 0) dy = -Math.min(40, (this.headerH + this.frozenH - p.y) / 2 + 2);
    }
    if (dx || dy) {
      this.autoScroll = { dx, dy, x: p.x, y: p.y };
      if (!this.autoScrollTimer) {
        this.autoScrollTimer = window.setInterval(() => {
          if (!this.autoScroll || !this.drag) return;
          this.scroller.scrollLeft += this.autoScroll.dx;
          this.scroller.scrollTop += this.autoScroll.dy;
          this.handleDrag(this.autoScroll.x, this.autoScroll.y);
        }, 30);
      }
    } else this.stopAutoScroll();
  }

  private stopAutoScroll(): void {
    this.autoScroll = null;
    if (this.autoScrollTimer) {
      clearInterval(this.autoScrollTimer);
      this.autoScrollTimer = null;
    }
  }

  private handleDrag(x: number, y: number): void {
    const d = this.drag!;
    const sel = this.app.sel;
    const s = this.sheet;
    const row = this.rowAt(Math.max(this.headerH, y));
    const col = this.colAt(Math.max(this.headerW, x));
    switch (d.kind) {
      case 'select':
        sel.extendTo(s, row, col);
        this.app.onSelectionChanged(false);
        break;
      case 'cols':
        sel.ranges[sel.ranges.length - 1] = { r1: 0, c1: Math.min(d.from, col), r2: MAX_ROWS - 1, c2: Math.max(d.from, col) };
        this.app.onSelectionChanged(false);
        break;
      case 'rows':
        sel.ranges[sel.ranges.length - 1] = { r1: Math.min(d.from, row), c1: 0, r2: Math.max(d.from, row), c2: MAX_COLS - 1 };
        this.app.onSelectionChanged(false);
        break;
      case 'point': {
        const r = normalize({ r1: d.anchor.row, c1: d.anchor.col, r2: row, c2: col });
        this.app.editor.insertRef(r, s, false, true);
        break;
      }
      case 'fill': {
        const src = d.src;
        // Extend along the axis with the larger distance
        const downBy = row - src.r2;
        const upBy = src.r1 - row;
        const rightBy = col - src.c2;
        const leftBy = src.c1 - col;
        const vert = Math.max(downBy, upBy);
        const horz = Math.max(rightBy, leftBy);
        let target: Range | null = null;
        if (vert <= 0 && horz <= 0) target = null;
        else if (vert >= horz) target = downBy > 0 ? { ...src, r2: row } : { ...src, r1: row };
        else target = rightBy > 0 ? { ...src, c2: col } : { ...src, c1: col };
        this.fillPreview = target;
        this.draw();
        break;
      }
      case 'resizeCol': {
        const w = Math.max(4, d.startW + (x - d.startX) / this.zoom);
        this.app.previewColWidth(d.cols, w);
        break;
      }
      case 'resizeRow': {
        const h = Math.max(4, d.startH + (y - d.startY) / this.zoom);
        this.app.previewRowHeight(d.rows, h);
        break;
      }
    }
  }

  private onPointerUp(e: PointerEvent): void {
    this.stopAutoScroll();
    const d = this.drag;
    this.drag = null;
    if (this.scroller.hasPointerCapture(e.pointerId)) this.scroller.releasePointerCapture(e.pointerId);
    if (!d) return;
    const p = this.local(e);
    switch (d.kind) {
      case 'fill':
        if (this.fillPreview) this.app.fill(d.src, this.fillPreview, e.ctrlKey);
        this.fillPreview = null;
        this.draw();
        break;
      case 'resizeCol':
        if (this.dragMoved) this.app.commitColWidth(d.cols, Math.max(4, d.startW + (p.x - d.startX) / this.zoom));
        break;
      case 'resizeRow':
        if (this.dragMoved) this.app.commitRowHeight(d.rows, Math.max(4, d.startH + (p.y - d.startY) / this.zoom));
        break;
      case 'select':
      case 'cols':
      case 'rows':
        this.app.onSelectionChanged();
        this.app.onSelectionFinished();
        break;
      case 'point':
        this.app.editor.focus();
        break;
    }
  }

  private onDoubleClick(e: MouseEvent): void {
    const p = this.local(e);
    const hit = this.hitTest(p.x, p.y);
    if (hit.kind === 'colHeader' && hit.edge !== null) {
      const sel = this.app.sel;
      const cols = sel.fullColumns().includes(hit.edge) ? sel.fullColumns() : [hit.edge];
      this.app.autofitColumns(cols);
    } else if (hit.kind === 'rowHeader' && hit.edge !== null) {
      const sel = this.app.sel;
      const rows = sel.fullRows().includes(hit.edge) ? sel.fullRows() : [hit.edge];
      this.app.autofitRows(rows);
    } else if (hit.kind === 'cell') {
      if (hit.fillHandle) {
        this.app.fillDownDoubleClick();
        return;
      }
      if (!this.app.settings.bool('grid.doubleClickEdit')) return;
      if (hit.filterButton || hit.tableButton) return;
      this.app.startEdit({ mode: 'edit' });
    }
  }

  private onContextMenu(e: MouseEvent): void {
    e.preventDefault();
    const p = this.local(e);
    const hit = this.hitTest(p.x, p.y);
    const sel = this.app.sel;
    if (this.app.isEditing) return;
    if (hit.kind === 'cell' && !sel.contains(hit.row, hit.col)) {
      sel.setCell(this.sheet, hit.row, hit.col);
      this.app.onSelectionChanged();
    } else if (hit.kind === 'colHeader' && !sel.fullColumns().includes(hit.col)) {
      sel.selectColumns(hit.col, hit.col);
      this.app.onSelectionChanged();
    } else if (hit.kind === 'rowHeader' && !sel.fullRows().includes(hit.row)) {
      sel.selectRows(hit.row, hit.row);
      this.app.onSelectionChanged();
    }
    this.app.showContextMenu(e, hit.kind === 'colHeader' ? 'col' : hit.kind === 'rowHeader' ? 'row' : 'cell');
  }

  private updateCursor(x: number, y: number): void {
    const hit = this.hitTest(x, y);
    let cursor = 'cell';
    if (hit.kind === 'colHeader') {
      cursor = hit.tableButton ? 'pointer' : hit.edge !== null ? 'col-resize' : 's-resize';
    } else if (hit.kind === 'rowHeader') {
      cursor = hit.edge !== null ? 'row-resize' : 'e-resize';
    } else if (hit.kind === 'cell') {
      if (hit.fillHandle) cursor = 'crosshair';
      else if (hit.filterButton || hit.tableButton) cursor = 'pointer';
      else if (this.app.isEditing && this.app.editor.canInsertRef()) cursor = 'copy';
    } else if (hit.kind === 'corner') cursor = 'default';
    this.scroller.style.cursor = cursor;
  }

  /** Measures the width needed for a column's content (for auto-fit). */
  measureColumn(col: number): number {
    const s = this.sheet;
    const wb = this.app.doc.wb;
    let max = 0;
    const b = this.app.doc.engine.usedBounds(s);
    for (let r = 0; r <= b.maxRow; r++) {
      const cell = s.cells.get(cellKey(r, col));
      const v = this.app.doc.engine.getValue(s, r, col);
      if (v === null && !cell) continue;
      if (s.mergeAt(r, col)) continue;
      const st = wb.styles.get(cell?.s ?? s.rowStyles.get(r) ?? s.colStyles.get(col));
      const text = v === null ? '' : formatValue(v, st.numFmt, 255).text;
      const font = this.font({ ...st, fontSize: st.fontSize ?? 11 }).replace(/[\d.]+px/, `${(((st.fontSize ?? 11) * 96) / 72).toFixed(2)}px`);
      const longest = text.split('\n').reduce((m, l) => Math.max(m, this.measure(font, l)), 0);
      const filterPad =
        s.filter && r === s.filter.range.r1 && col >= s.filter.range.c1 && col <= s.filter.range.c2
          ? 20
          : s.tables.some((t) => t.range.r1 === r && col >= t.range.c1 && col <= t.range.c2)
            ? this.filterButtonSize() * 2 + 10
            : 0;
      max = Math.max(max, longest + PAD * 2 + 6 + filterPad);
    }
    return max;
  }

  measureRow(row: number): number {
    const s = this.sheet;
    const wb = this.app.doc.wb;
    let max = s.defaultRowHeight;
    const b = this.app.doc.engine.usedBounds(s);
    for (let c = 0; c <= b.maxCol; c++) {
      const cell = s.cells.get(cellKey(row, c));
      if (!cell) continue;
      const st = wb.styles.get(cell.s ?? s.rowStyles.get(row) ?? s.colStyles.get(c));
      const fontPx = ((st.fontSize ?? 11) * 96) / 72;
      let lines = 1;
      if (st.wrap) {
        const v = this.app.doc.engine.getValue(s, row, c);
        const text = v === null ? '' : formatValue(v, st.numFmt, 255).text;
        const font = `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}${fontPx}px "${st.fontName ?? 'Calibri'}", Calibri, sans-serif`;
        lines = this.wrapLines(font, text, s.colWidth(c) - PAD * 2).length;
      }
      max = Math.max(max, Math.ceil(lines * fontPx * 1.22 + 6));
    }
    return max;
  }
}

/** A rounded highlight (selected headings). */
function pill(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number): void {
  if (w <= 0 || h <= 0) return;
  ctx.beginPath();
  ctx.roundRect(x, y, w, h, Math.min(6, h / 2, w / 2));
  ctx.fill();
}
