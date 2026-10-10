// The application controller: owns the document and wires every UI component together.
import { cellKey, formatCell, formatRange, isFullColumn, isFullRow, MAX_COLS, MAX_ROWS, parseCell, parseRange, rangeContains, type Range } from '../core/address';
import { CFEvaluator } from '../core/cf';
import { external, FetchError } from '../core/data/external';
import { Doc, type DocEvent } from '../core/doc';
import { buildTemplate, TEMPLATES } from '../core/templates';
import { tokenize } from '../core/formula/tokenizer';
import { RangeRef } from '../core/formula/types';
import type { Transaction } from '../core/history';
import { csvToWorkbook, decodeText, encodeCSV, sheetToCSV } from '../core/io/csv';
import { readXlsx, writeXlsx } from '../core/io/xlsx';
import { applyStyle, clearRanges, fillSelectionWithInput, setCellInput, type ClearWhat, type PatchFn } from '../core/ops/cells';
import { copyRange, pasteClip, pasteText, type ClipData, type PasteMode } from '../core/ops/clipboard';
import { currentRegion } from '../core/ops/data';
import { autofill, fillDownExtent, type FillMode } from '../core/ops/fill';
import { setColumnWidth, setRowHeight } from '../core/ops/structure';
import { ERROR_HELP, ErrorValue } from '../core/values';
import { Workbook, type Sheet } from '../core/workbook';
import type { OpenedFile, SaveKind } from '../shared/api';
import { ChartLayer } from './charts/ChartLayer';
import { Commands } from './commands';
import { showContextMenu } from './contextMenu';
import { debounce, h } from './dom';
import { EditController, type Move } from './editor';
import { FilterMenu } from './filterMenu';
import { FindBar } from './findBar';
import { FormulaBar } from './formulaBar';
import { GridView } from './grid/GridView';
import { clearColorCaches } from './grid/theme';
import { SidePanel } from './panels/SidePanel';
import { isDesktop, platform } from './platform';
import { Selection } from './selection';
import { SheetTabs } from './sheetTabs';
import { StatusBar } from './statusBar';
import { pickTemplate } from './templatePicker';
import { TitleBar } from './titleBar';
import { Toolbar } from './toolbar';
import { dialog, hideTooltip, toast } from './widgets';

export interface Prefs {
  theme: 'dark' | 'light';
  adaptColors: boolean;
  recentCommands: string[];
  showFormulaBarExpanded: boolean;
}

const PREFS_KEY = 'gridline.prefs';

function loadPrefs(): Prefs {
  const base: Prefs = { theme: 'dark', adaptColors: true, recentCommands: [], showFormulaBarExpanded: false };
  try {
    const raw = localStorage.getItem(PREFS_KEY);
    if (raw) return { ...base, ...(JSON.parse(raw) as Partial<Prefs>) };
  } catch {
    // ignore
  }
  return base;
}

export class App {
  doc: Doc;
  sel = new Selection();
  prefs: Prefs = loadPrefs();
  cf: CFEvaluator;
  grid: GridView;
  editor: EditController;
  formulaBar: FormulaBar;
  toolbar: Toolbar;
  tabs: SheetTabs;
  statusbar: StatusBar;
  titlebar: TitleBar;
  panel: SidePanel;
  charts: ChartLayer;
  commands: Commands;
  findBar: FindBar;
  showFormulas = false;
  clipboard: { clip: ClipData; text: string } | null = null;
  formatPainter: { sheetId: number; range: Range; sticky: boolean } | null = null;
  lastFill: { src: Range; target: Range; mode: FillMode } | null = null;
  private tabStartCol: number | null = null;
  private cellTip: HTMLDivElement | null = null;
  private cellTipTimer: ReturnType<typeof setTimeout> | undefined;
  private cellTipKey = '';
  private csvChoice: 'xlsx' | 'csv' | null = null;
  readonly root: HTMLElement;
  private explainerEl: HTMLDivElement;

  constructor(root: HTMLElement) {
    this.root = root;
    this.applyTheme();
    this.doc = new Doc();
    this.cf = new CFEvaluator(this.doc);
    this.commands = new Commands(this);
    this.grid = new GridView(this);
    this.editor = new EditController(this);
    this.titlebar = new TitleBar(this);
    this.toolbar = new Toolbar(this);
    this.formulaBar = new FormulaBar(this);
    this.explainerEl = h('div', { class: 'explainer', hidden: true });
    this.statusbar = new StatusBar(this);
    this.tabs = new SheetTabs(this);
    this.panel = new SidePanel(this);
    this.charts = new ChartLayer(this);
    this.findBar = new FindBar(this);
    this.grid.overlay.append(this.charts.el, this.editor.cellBox);
    this.grid.el.append(this.findBar.el);

    const main = h('div', { class: 'main' }, this.grid.el, this.panel.el);
    const bottom = h('footer', { class: 'bottombar' }, this.tabs.el, this.statusbar.el);
    this.grid.el.append(this.explainerEl);
    root.append(this.titlebar.el, this.toolbar.el, this.formulaBar.el, main, bottom);

    external.fetcher = async (url, opts) => {
      const r = await platform.fetchData(url, opts);
      if (!r.ok || r.text === undefined) throw new FetchError(r.error ?? 'Download failed', r.status);
      return r.text;
    };
    this.attachDoc();
    this.selSheetId = this.sheet.id;
    this.bindKeys();
    this.bindDrop();
    platform.onRequestClose(() => this.requestClose());
    window.addEventListener('beforeunload', (e) => {
      if (!isDesktop && this.doc.dirty) e.preventDefault();
    });
    this.refreshAll();
    this.focusGrid();
    this.openStartupFile();
  }

  // ---------- basics ----------

  get sheet(): Sheet {
    return this.doc.wb.activeSheet;
  }

  get isEditing(): boolean {
    return this.editor.isOpen;
  }

  savePrefs(): void {
    try {
      localStorage.setItem(PREFS_KEY, JSON.stringify(this.prefs));
    } catch {
      // ignore
    }
  }

  applyTheme(): void {
    document.documentElement.dataset.theme = this.prefs.theme;
    const dark = this.prefs.theme === 'dark';
    platform.setTheme({ dark, bar: dark ? '#16181d' : '#f6f7f9', symbol: dark ? '#c9ced8' : '#3c4350' });
    clearColorCaches();
    this.grid?.refreshTheme();
    this.charts?.renderAll();
  }

  private attachDoc(): void {
    this.cf = new CFEvaluator(this.doc);
    this.doc.selectionProvider = () => this.sel.snapshot(this.sheet.id);
    this.doc.on((e) => this.onDocEvent(e));
  }

  private onDocEvent(e: DocEvent): void {
    if (e.type === 'change' && e.direction !== 'do') {
      const snap = e.direction === 'undo' ? e.tx.selectionBefore : e.tx.selectionAfter;
      if (snap && this.doc.wb.sheetById(snap.sheetId)) {
        if (this.doc.wb.activeSheetId !== snap.sheetId) {
          this.doc.wb.activeSheetId = snap.sheetId;
          this.grid.relayout();
          this.grid.restoreScroll();
        }
        this.sel.restore(snap);
        this.grid.scrollToCell(this.sel.active.row, this.sel.active.col);
      }
    }
    if (e.type === 'reset') {
      this.cf = new CFEvaluator(this.doc);
    }
    if (!this.doc.wb.sheetById(this.doc.wb.activeSheetId)) this.doc.wb.activeSheetId = this.doc.wb.sheets[0].id;
    if (e.type === 'change' && e.direction !== 'do') {
      this.selSheetId = this.doc.wb.activeSheetId;
    } else if (this.doc.wb.activeSheetId !== this.selSheetId) {
      const target = this.doc.wb.activeSheetId;
      this.doc.wb.activeSheetId = this.selSheetId;
      if (this.doc.wb.sheetById(this.selSheetId)) this.setActiveSheet(target);
      else {
        this.doc.wb.activeSheetId = target;
        this.selSheetId = target;
        this.sel.setCell(this.sheet, this.sheet.view.activeRow, this.sheet.view.activeCol);
        this.grid.relayout();
        this.grid.restoreScroll();
      }
    }
    this.cf.invalidate();
    this.grid.invalidate();
    this.refreshAll();
  }

  /** Runs an undoable action; shows a friendly message if it fails. */
  transact(label: string, fn: (tx: Transaction) => void): Transaction | null {
    try {
      return this.doc.transact(label, fn);
    } catch (err) {
      toast((err as Error).message || String(err), { kind: 'error' });
      return null;
    }
  }

  refreshAll(): void {
    this.tabs?.render();
    this.toolbar?.refresh();
    this.refreshFormulaBar();
    this.statusbar?.refresh();
    this.titlebar?.refresh();
    this.charts?.sync();
    this.panel?.refresh();
    this.findBar?.refresh();
    platform.setDirty(this.doc.dirty);
  }

  refreshFormulaBar(): void {
    this.formulaBar?.refresh();
    this.updateExplainer();
  }

  focusGrid(): void {
    if (this.isEditing) return;
    this.grid.focusTarget.focus({ preventScroll: true });
  }

  // ---------- selection ----------

  onSelectionChanged(scroll = true): void {
    const a = this.sel.active;
    if (scroll) this.grid.scrollToCell(this.sel.cursor?.row ?? a.row, this.sel.cursor?.col ?? a.col);
    this.sheet.view.activeRow = a.row;
    this.sheet.view.activeCol = a.col;
    this.grid.draw();
    this.refreshFormulaBar();
    this.toolbar.refresh();
    this.updateStats();
    this.charts.deselect();
    this.hideCellTip();
    this.panel.selectionChanged();
  }

  private updateStats = debounce(() => this.statusbar.refresh(), 60);

  /** Called when a mouse selection ends (applies the format painter if active). */
  onSelectionFinished(): void {
    const fp = this.formatPainter;
    if (!fp) return;
    this.formatPainter = null;
    const src = this.doc.wb.sheetById(fp.sheetId);
    if (!src) return;
    const s = this.sheet;
    const h2 = fp.range.r2 - fp.range.r1 + 1;
    const w2 = fp.range.c2 - fp.range.c1 + 1;
    const target = this.sel.range;
    const t = {
      r1: target.r1,
      c1: target.c1,
      r2: target.r1 === target.r2 && target.c1 === target.c2 ? target.r1 + h2 - 1 : Math.min(target.r2, target.r1 + 100000),
      c2: target.r1 === target.r2 && target.c1 === target.c2 ? target.c1 + w2 - 1 : Math.min(target.c2, target.c1 + 1000),
    };
    this.transact('Paint formatting', (tx) => {
      for (let r = t.r1; r <= t.r2; r++)
        for (let c = t.c1; c <= t.c2; c++) {
          const sr = fp.range.r1 + ((r - t.r1) % h2);
          const sc = fp.range.c1 + ((c - t.c1) % w2);
          const sid = src.styleIdAt(sr, sc);
          const own = s.getCell(r, c)?.s;
          if (own !== sid) tx.patchCell(s, r, c, { s: sid ?? 0 });
        }
    });
    this.sel.setRange(t, { row: t.r1, col: t.c1 });
    this.onSelectionChanged(false);
    this.toolbar.refresh();
  }

  select(range: Range, active?: { row: number; col: number }): void {
    this.sel.setRange(range, active);
    this.onSelectionChanged();
  }

  /** The sheet the current selection belongs to (each sheet remembers its own cursor). */
  private selSheetId = -1;

  setActiveSheet(id: number, resetSelection = true): void {
    if (this.isEditing && !this.editor.canInsertRef()) {
      if (!this.commitEdit()) return;
    }
    const wb = this.doc.wb;
    const s = wb.sheetById(id);
    if (!s) return;
    const prev = wb.sheetById(this.selSheetId);
    if (prev && prev !== s) {
      prev.view.activeRow = this.sel.active.row;
      prev.view.activeCol = this.sel.active.col;
      prev.view.scrollX = this.grid.scrollX;
      prev.view.scrollY = this.grid.scrollY;
    }
    this.selSheetId = id;
    wb.activeSheetId = id;
    if (resetSelection) this.sel.setCell(s, s.view.activeRow, s.view.activeCol);
    this.grid.relayout();
    this.grid.restoreScroll();
    this.cf.invalidate();
    this.editor.position();
    this.refreshAll();
    this.grid.draw();
  }

  // ---------- editing ----------

  startEdit(opts: { mode: 'enter' | 'edit'; initial?: string }): void {
    if (this.isEditing) return;
    const a = this.sel.active;
    const anchor = this.doc.engine.spillAnchorOf(this.sheet, a.row, a.col);
    if (anchor && opts.initial === undefined) {
      toast(`This value comes from the formula in ${formatCell(anchor.row, anchor.col)}. Edit that cell instead.`, {
        action: { label: 'Go there', run: () => this.select({ r1: anchor.row, c1: anchor.col, r2: anchor.row, c2: anchor.col }) },
      });
      return;
    }
    this.editor.start({ mode: opts.mode, initial: opts.initial, source: 'cell' });
  }

  commitEdit(): boolean {
    return this.editor.commit('none');
  }

  enterInput(sheet: Sheet, row: number, col: number, text: string, fillSelection: boolean): void {
    const tx = this.transact(fillSelection ? 'Fill selection' : 'Edit cell', (tx) => {
      if (fillSelection) fillSelectionWithInput(tx, this.doc, sheet, this.sel.ranges, row, col, text);
      else setCellInput(tx, this.doc, sheet, row, col, text);
    });
    if (!tx) return;
    // Auto-fit the row height when text with line breaks is entered
    if (text.includes('\n') && !sheet.rowHeights.has(row)) this.autofitRows([row], true);
    const v = this.doc.engine.getValue(sheet, row, col);
    if (v instanceof ErrorValue && v.origin && v.origin.row === row && v.origin.col === col && v.code !== '#N/A') {
      // The error explainer shows details once the user returns to the cell
    }
  }

  moveAfterEnter(move: Move): void {
    const sel = this.sel;
    const s = this.sheet;
    const r = sel.range;
    const multi = sel.ranges.length === 1 && (r.r1 !== r.r2 || r.c1 !== r.c2) && !sel.isSingleCell && !s.mergeAt(r.r1, r.c1);
    if (multi) {
      // Move within the selection, wrapping around
      let { row, col } = sel.active;
      if (move === 'down') {
        row++;
        if (row > r.r2) {
          row = r.r1;
          col = col + 1 > r.c2 ? r.c1 : col + 1;
        }
      } else if (move === 'up') {
        row--;
        if (row < r.r1) {
          row = r.r2;
          col = col - 1 < r.c1 ? r.c2 : col - 1;
        }
      } else if (move === 'right') {
        col++;
        if (col > r.c2) {
          col = r.c1;
          row = row + 1 > r.r2 ? r.r1 : row + 1;
        }
      } else if (move === 'left') {
        col--;
        if (col < r.c1) {
          col = r.c2;
          row = row - 1 < r.r1 ? r.r2 : row - 1;
        }
      }
      sel.active = { row, col };
      this.grid.scrollToCell(row, col);
      this.onSelectionChanged(false);
      return;
    }
    if (move === 'right' || move === 'left') {
      if (this.tabStartCol === null) this.tabStartCol = sel.active.col;
      this.moveActive(0, move === 'right' ? 1 : -1, false, false);
      return;
    }
    if (move === 'down' && this.tabStartCol !== null) {
      // After Tab, Tab, Enter go back to the starting column, like Excel
      const row = this.grid.rows.step(sel.active.row, 1);
      sel.setCell(s, row, this.tabStartCol);
      this.tabStartCol = null;
      this.onSelectionChanged();
      return;
    }
    this.tabStartCol = null;
    this.moveActive(move === 'down' ? 1 : -1, 0, false, false);
  }

  /** Moves the active cell (or extends the selection with Shift). */
  moveActive(dr: number, dc: number, extend: boolean, jump: boolean): void {
    const s = this.sheet;
    const sel = this.sel;
    const grid = this.grid;
    const from = extend ? sel.cursor ?? sel.active : sel.active;
    let row = from.row;
    let col = from.col;
    // Step out of merged blocks
    const m = s.mergeAt(row, col);
    if (jump) {
      [row, col] = this.jumpTarget(row, col, dr, dc);
    } else if (Math.abs(dr) > 1 || Math.abs(dc) > 1) {
      row = Math.max(0, Math.min(MAX_ROWS - 1, row + dr));
      col = Math.max(0, Math.min(MAX_COLS - 1, col + dc));
      while (row < MAX_ROWS - 1 && grid.rows.isHidden(row)) row++;
      while (col < MAX_COLS - 1 && grid.cols.isHidden(col)) col++;
    } else {
      if (dr > 0 && m && !extend) row = m.r2;
      if (dc > 0 && m && !extend) col = m.c2;
      if (dr) row = grid.rows.step(row, dr as 1 | -1);
      if (dc) col = grid.cols.step(col, dc as 1 | -1);
    }
    if (extend) {
      sel.extendTo(s, row, col);
      sel.cursor = { row, col };
    } else {
      sel.setCell(s, row, col);
      sel.cursor = null;
    }
    this.onSelectionChanged();
  }

  private jumpTarget(row: number, col: number, dr: number, dc: number): [number, number] {
    const s = this.sheet;
    const filled = (r: number, c: number) => this.doc.engine.getValue(s, r, c) !== null;
    const inBounds = (r: number, c: number) => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS;
    const b = this.doc.engine.usedBounds(s);
    let r = row + dr;
    let c = col + dc;
    if (!inBounds(r, c)) return [row, col];
    if (filled(row, col) && filled(r, c)) {
      while (inBounds(r + dr, c + dc) && filled(r + dr, c + dc)) {
        r += dr;
        c += dc;
      }
      return [r, c];
    }
    // Skip empty cells to the next filled one, or the sheet edge
    while (inBounds(r, c) && !filled(r, c)) {
      if ((dr > 0 && r > b.maxRow) || (dc > 0 && c > b.maxCol)) {
        return [dr > 0 ? MAX_ROWS - 1 : r, dc > 0 ? MAX_COLS - 1 : c];
      }
      r += dr;
      c += dc;
    }
    if (!inBounds(r, c)) return [Math.max(0, Math.min(MAX_ROWS - 1, r - dr)), Math.max(0, Math.min(MAX_COLS - 1, c - dc))];
    return [r, c];
  }

  // ---------- keyboard ----------

  private bindKeys(): void {
    const ft = this.grid.focusTarget;
    ft.addEventListener('keydown', (e) => this.onGridKey(e));
    ft.addEventListener('input', () => {
      const text = ft.value;
      ft.value = '';
      if (!text || this.isEditing) return;
      if (text === '\n') return;
      this.startEdit({ mode: 'enter', initial: text });
    });
    ft.addEventListener('paste', (e) => e.preventDefault());
    // Global shortcuts that work anywhere (outside text inputs)
    document.addEventListener('keydown', (e) => {
      if (e.defaultPrevented) return;
      const target = e.target as HTMLElement;
      const inField = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target.isContentEditable;
      if (target === ft) return;
      const cmd = this.commands.match(e);
      if (!cmd) return;
      if (inField && !cmd.global) return;
      if (this.isEditing && !cmd.whileEditing) return;
      e.preventDefault();
      this.commands.run(cmd.id);
    });
  }

  private onGridKey(e: KeyboardEvent): void {
    if (e.isComposing) return;
    const cmd = this.commands.match(e);
    if (cmd) {
      e.preventDefault();
      this.commands.run(cmd.id);
      return;
    }
    const ctrl = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;
    const key = e.key;
    const nav: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (nav[key]) {
      e.preventDefault();
      if (this.charts.selected) {
        this.charts.nudge(nav[key][1] * (shift ? 10 : 1), nav[key][0] * (shift ? 10 : 1));
        return;
      }
      this.tabStartCol = null;
      this.moveActive(nav[key][0], nav[key][1], shift, ctrl);
      return;
    }
    switch (key) {
      case 'Enter':
        e.preventDefault();
        if (this.charts.selected) {
          this.panel.open('chart');
          return;
        }
        this.moveAfterEnter(shift ? 'up' : 'down');
        return;
      case 'Tab':
        e.preventDefault();
        this.moveAfterEnter(shift ? 'left' : 'right');
        return;
      case 'PageDown':
      case 'PageUp': {
        e.preventDefault();
        const dir = key === 'PageDown' ? 1 : -1;
        if (e.altKey) this.moveActive(0, dir * this.grid.pageCols(), shift, false);
        else {
          const rows = this.grid.pageRows();
          this.grid.scroller.scrollTop += dir * rows * this.grid.rows.def;
          this.moveActive(dir * rows, 0, shift, false);
        }
        return;
      }
      case 'Home':
        e.preventDefault();
        if (ctrl) {
          this.sel.setCell(this.sheet, this.sheet.freeze.rows, this.sheet.freeze.cols);
          this.onSelectionChanged();
        } else {
          const row = this.sel.active.row;
          if (shift) this.sel.extendTo(this.sheet, row, 0);
          else this.sel.setCell(this.sheet, row, 0);
          this.onSelectionChanged();
        }
        return;
      case 'End':
        if (ctrl) {
          e.preventDefault();
          const b = this.sheet.contentBounds();
          this.sel.setCell(this.sheet, Math.max(0, b.maxRow), Math.max(0, b.maxCol));
          this.onSelectionChanged();
        }
        return;
      case 'Delete':
        e.preventDefault();
        if (this.charts.selected) {
          this.charts.deleteSelected();
          return;
        }
        this.clear('contents');
        return;
      case 'Backspace':
        e.preventDefault();
        if (this.charts.selected) {
          this.charts.deleteSelected();
          return;
        }
        this.transact('Clear cell', (tx) => clearRanges(tx, this.sheet, [{ r1: this.sel.active.row, c1: this.sel.active.col, r2: this.sel.active.row, c2: this.sel.active.col }], 'contents'));
        this.startEdit({ mode: 'enter', initial: '' });
        return;
      case 'F2':
        e.preventDefault();
        this.startEdit({ mode: 'edit' });
        return;
      case 'Escape':
        if (this.clipboard) {
          this.clipboard = null;
          this.grid.setCopyMarquee(null);
        }
        if (this.formatPainter) {
          this.formatPainter = null;
          this.toolbar.refresh();
        }
        this.charts.deselect();
        this.findBar.close();
        return;
    }
  }

  // ---------- clearing, styling helpers ----------

  clear(what: ClearWhat): void {
    const labels: Record<ClearWhat, string> = { contents: 'Clear contents', formats: 'Clear formats', all: 'Clear all', notes: 'Clear notes', links: 'Remove links' };
    // Typing over a spilled array area would block it; clearing its anchor clears the whole spill.
    this.transact(labels[what], (tx) => clearRanges(tx, this.sheet, this.sel.ranges, what));
  }

  styleAtActive() {
    const a = this.sel.active;
    return this.doc.wb.styles.get(this.sheet.styleIdAt(a.row, a.col));
  }

  applyStyle(label: string, patch: PatchFn): void {
    this.transact(label, (tx) => applyStyle(tx, this.doc.wb, this.sheet, this.sel.ranges, patch));
  }

  // ---------- clipboard ----------

  async copy(cut = false): Promise<void> {
    if (this.sel.ranges.length > 1) {
      toast('Copying multiple separate ranges isn\'t supported — select one block of cells.', { kind: 'error' });
      return;
    }
    const clip = copyRange(this.doc, this.sheet, this.sel.range, cut);
    this.clipboard = { clip, text: clip.text };
    await platform.writeClipboard({ text: clip.text, html: clip.html });
    this.grid.setCopyMarquee({ sheetId: this.sheet.id, range: clip.range });
    const n = clip.cells.length * (clip.cells[0]?.length ?? 0);
    this.statusbar.flash(cut ? 'Cut — select a destination and paste' : `Copied ${n.toLocaleString()} cell${n === 1 ? '' : 's'}`);
  }

  async paste(mode: PasteMode = 'all'): Promise<void> {
    const { text } = await platform.readClipboard();
    const sel = this.sel.range;
    const internal = this.clipboard && (text === this.clipboard.text || (!text && this.clipboard));
    let out: Range | undefined;
    if (internal) {
      const clip = this.clipboard!.clip;
      const tx = this.transact(clip.cut ? 'Move cells' : 'Paste', (tx) => {
        out = pasteClip(tx, this.doc, clip, this.sheet, this.isWholeSheet(sel) ? { r1: sel.r1, c1: sel.c1, r2: sel.r1, c2: sel.c1 } : sel, mode);
      });
      if (tx && clip.cut) {
        this.clipboard = null;
        this.grid.setCopyMarquee(null);
      }
    } else if (text) {
      if (mode === 'formats') return;
      this.transact('Paste', (tx) => {
        out = pasteText(tx, this.doc, text, this.sheet, sel.r1, sel.c1);
      });
    }
    if (out) {
      this.sel.setRange(out, { row: out.r1, col: out.c1 });
      this.onSelectionChanged();
    }
  }

  private isWholeSheet(r: Range): boolean {
    return isFullColumn(r) || isFullRow(r);
  }

  // ---------- fill ----------

  fill(src: Range, target: Range, ctrl: boolean): void {
    const singleNumber = src.r1 === src.r2 && src.c1 === src.c2 && typeof this.doc.engine.getValue(this.sheet, src.r1, src.c1) === 'number';
    const mode: FillMode = ctrl ? (singleNumber ? 'series' : 'copy') : 'auto';
    const tx = this.transact('Auto fill', (tx) => autofill(tx, this.doc, this.sheet, src, target, mode));
    if (!tx) return;
    this.lastFill = { src, target, mode };
    this.sel.setRange(target, { ...this.sel.active });
    this.onSelectionChanged(false);
    this.offerFillOptions(singleNumber);
  }

  private offerFillOptions(singleNumber: boolean): void {
    if (!this.lastFill) return;
    const { mode } = this.lastFill;
    const alt: FillMode = mode === 'copy' ? 'series' : singleNumber && mode === 'auto' ? 'series' : 'copy';
    const label = alt === 'series' ? 'Fill as a series instead' : 'Copy cells instead';
    this.statusbar.flash('Filled', { label, run: () => this.refill(alt) });
  }

  refill(mode: FillMode): void {
    const lf = this.lastFill;
    if (!lf) return;
    this.doc.undo();
    this.transact('Auto fill', (tx) => autofill(tx, this.doc, this.sheet, lf.src, lf.target, mode));
    this.lastFill = { ...lf, mode };
  }

  fillDownDoubleClick(): void {
    const src = this.sel.range;
    const end = fillDownExtent(this.doc, this.sheet, src);
    if (end === null) return;
    this.fill(src, { ...src, r2: end }, false);
  }

  fillDirection(dir: 'down' | 'right'): void {
    const r = this.sel.range;
    if (dir === 'down') {
      if (r.r1 === r.r2) {
        if (r.r1 === 0) return;
        this.transact('Fill down', (tx) => autofill(tx, this.doc, this.sheet, { ...r, r1: r.r1 - 1, r2: r.r1 - 1 }, { ...r, r1: r.r1 - 1 }, 'copy'));
      } else this.transact('Fill down', (tx) => autofill(tx, this.doc, this.sheet, { ...r, r2: r.r1 }, r, 'copy'));
    } else {
      if (r.c1 === r.c2) {
        if (r.c1 === 0) return;
        this.transact('Fill right', (tx) => autofill(tx, this.doc, this.sheet, { ...r, c1: r.c1 - 1, c2: r.c1 - 1 }, { ...r, c1: r.c1 - 1 }, 'copy'));
      } else this.transact('Fill right', (tx) => autofill(tx, this.doc, this.sheet, { ...r, c2: r.c1 }, r, 'copy'));
    }
  }

  // ---------- sizes ----------

  previewColWidth(cols: number[], w: number): void {
    const s = this.sheet;
    if (!this.preResizeCols) this.preResizeCols = new Map(cols.map((c) => [c, s.colWidths.get(c)]));
    for (const c of cols) s.colWidths.set(c, Math.round(w));
    this.grid.relayout();
    this.statusbar.flash(`Width: ${Math.round(w)} px`);
  }

  previewRowHeight(rows: number[], hgt: number): void {
    const s = this.sheet;
    if (!this.preResizeRows) this.preResizeRows = new Map(rows.map((r) => [r, s.rowHeights.get(r)]));
    for (const r of rows) s.rowHeights.set(r, Math.round(hgt));
    this.grid.relayout();
    this.statusbar.flash(`Height: ${Math.round(hgt)} px`);
  }

  /** Restores the values changed during a live resize preview, then records the real change. */
  commitColWidth(cols: number[], w: number): void {
    const s = this.sheet;
    for (const [c, p] of this.preResizeCols ?? []) {
      if (p === undefined) s.colWidths.delete(c);
      else s.colWidths.set(c, p);
    }
    this.preResizeCols = null;
    this.transact('Resize columns', (tx) => setColumnWidth(tx, s, cols, w));
  }

  commitRowHeight(rows: number[], hgt: number): void {
    const s = this.sheet;
    for (const [r, p] of this.preResizeRows ?? []) {
      if (p === undefined) s.rowHeights.delete(r);
      else s.rowHeights.set(r, p);
    }
    this.preResizeRows = null;
    this.transact('Resize rows', (tx) => setRowHeight(tx, s, rows, hgt));
  }

  private preResizeCols: Map<number, number | undefined> | null = null;
  private preResizeRows: Map<number, number | undefined> | null = null;

  autofitColumns(cols: number[]): void {
    const s = this.sheet;
    this.transact('Auto-fit columns', (tx) => {
      for (const c of cols) {
        const w = this.grid.measureColumn(c);
        setColumnWidth(tx, s, [c], w > 0 ? Math.max(24, Math.min(600, Math.ceil(w))) : undefined);
      }
    });
  }

  autofitRows(rows: number[], _quiet = false): void {
    const s = this.sheet;
    this.transact('Auto-fit rows', (tx) => {
      for (const r of rows) {
        const hgt = this.grid.measureRow(r);
        setRowHeight(tx, s, [r], hgt <= s.defaultRowHeight ? undefined : hgt);
      }
    });
  }

  // ---------- zoom ----------

  setZoom(z: number): void {
    const zoom = Math.max(0.25, Math.min(4, Math.round(z * 100) / 100));
    this.sheet.view.zoom = zoom;
    this.grid.relayout();
    this.editor.position();
    this.statusbar.refresh();
    this.charts.position();
  }

  zoomBy(delta: number): void {
    this.setZoom(this.sheet.view.zoom + delta);
  }

  // ---------- notifications from the grid ----------

  onGridScroll(): void {
    this.editor.position();
    this.charts?.position();
  }

  hoverCell(x: number, y: number, e: MouseEvent): void {
    const hit = this.grid.hitTest(x, y);
    if (hit.kind !== 'cell') {
      this.hideCellTip();
      return;
    }
    const s = this.sheet;
    const cell = s.getCell(hit.row, hit.col);
    const v = this.doc.engine.getValue(s, hit.row, hit.col);
    const key = `${s.id}:${hit.row}:${hit.col}`;
    if (key === this.cellTipKey) return;
    this.hideCellTip();
    this.cellTipKey = key;
    let text = '';
    let cls = 'cell-tip';
    if (cell?.note) text = cell.note;
    else if (v instanceof ErrorValue) {
      text = `${ERROR_HELP[v.code]?.title ?? v.code}: ${v.message ?? ''}`;
      cls += ' error';
    } else if (cell?.link) text = `${cell.link}\nCtrl+click to open`;
    if (!text) return;
    this.cellTipTimer = setTimeout(() => {
      this.cellTip = h('div', { class: cls }, text);
      document.body.append(this.cellTip);
      const rect = this.grid.cellRect(hit.row, hit.col);
      const host = this.grid.el.getBoundingClientRect();
      this.cellTip.style.left = `${Math.min(window.innerWidth - this.cellTip.offsetWidth - 8, host.left + rect.x + rect.w + 6)}px`;
      this.cellTip.style.top = `${host.top + rect.y}px`;
    }, 350);
    void e;
  }

  hideCellTip(): void {
    clearTimeout(this.cellTipTimer);
    this.cellTip?.remove();
    this.cellTip = null;
    this.cellTipKey = '';
    hideTooltip();
  }

  showContextMenu(e: MouseEvent, kind: 'cell' | 'row' | 'col'): void {
    showContextMenu(this, e, kind);
  }

  openFilterMenu(col: number, rect: DOMRect): void {
    new FilterMenu(this, col).open(rect);
  }

  openLink(url: string): void {
    if (/^(https?:\/\/|mailto:)/i.test(url)) platform.openExternal(url);
    else if (/^www\./i.test(url)) platform.openExternal('https://' + url);
    else {
      // Internal link like #Sheet2!A1 or Sheet2!A1
      this.goTo(url.replace(/^#/, ''));
    }
  }

  openHyperlinkFormula(row: number, col: number): void {
    const f = this.sheet.getCell(row, col)?.f;
    if (!f) return;
    const toks = tokenize(f).filter((t) => t.type !== 'ws');
    const arg = toks[2];
    let url: string | null = null;
    if (arg?.type === 'str') url = String(arg.value);
    else if (arg?.type === 'ref' && arg.ref) {
      const v = this.doc.engine.getValue(this.sheet, arg.ref.r1, arg.ref.c1);
      if (typeof v === 'string') url = v;
    }
    if (url) this.openLink(url);
  }

  /** Jumps to an address ("B12", "Sheet2!A1", "A1:C9") or a defined name. */
  goTo(text: string): boolean {
    const t = text.trim();
    if (!t) return false;
    let sheet = this.sheet;
    let addr = t;
    const bang = t.lastIndexOf('!');
    if (bang > 0) {
      const name = t.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
      const s = this.doc.wb.sheetByName(name);
      if (!s) return false;
      sheet = s;
      addr = t.slice(bang + 1);
    }
    let range = parseRange(addr.replace(/\$/g, ''));
    if (!range) {
      const dn = this.doc.wb.getName(t);
      if (dn) return this.goTo(dn.ref);
      return false;
    }
    if (sheet.id !== this.sheet.id) this.setActiveSheet(sheet.id, false);
    range = { r1: range.r1, c1: range.c1, r2: Math.min(range.r2, MAX_ROWS - 1), c2: Math.min(range.c2, MAX_COLS - 1) };
    this.sel.setRange(range, { row: range.r1, col: range.c1 });
    this.grid.scrollToCellTop(range.r1, range.c1);
    this.onSelectionChanged(false);
    this.focusGrid();
    return true;
  }

  // ---------- error explainer ----------

  updateExplainer(): void {
    const el = this.explainerEl;
    if (!el) return;
    const a = this.sel.active;
    const s = this.sheet;
    const v = this.doc.engine.getValue(s, a.row, a.col);
    const unsupported = this.doc.engine.isUnsupported(s, a.row, a.col);
    el.innerHTML = '';
    if (unsupported.length && s.getCell(a.row, a.col)?.cached !== undefined) {
      el.className = 'explainer warn';
      el.hidden = false;
      el.append(
        h('span', { class: 'icon' }, '⚠'),
        h('div', { class: 'body' }, h('div', { class: 'title' }, 'Function not supported yet'), h('div', { class: 'msg' }, `${unsupported.join(', ')} can't be calculated here, so the value saved in the file is shown. It won't update if its inputs change.`)),
      );
      return;
    }
    if (!(v instanceof ErrorValue) || v.code === '#BUSY!') {
      el.hidden = true;
      return;
    }
    const help = ERROR_HELP[v.code];
    el.className = 'explainer';
    el.hidden = false;
    const actions = h('div', { class: 'actions' });
    const origin = v.origin;
    if (origin && (origin.sheet !== s.name || origin.row !== a.row || origin.col !== a.col)) {
      const label = `${origin.sheet !== s.name ? origin.sheet + '!' : ''}${formatCell(origin.row, origin.col)}`;
      actions.append(h('button', { class: 'btn outline', onclick: () => this.goTo(`'${origin.sheet.replace(/'/g, "''")}'!${formatCell(origin.row, origin.col)}`) }, `Go to source (${label})`));
    }
    const sug = v.code === '#NAME?' && v.message ? /Did you mean (\w[\w.]*)\?/.exec(v.message) : null;
    const cell = s.getCell(a.row, a.col);
    if (sug && cell?.f) {
      const bad = /There's no function called ([\w.]+)\./.exec(v.message ?? '')?.[1];
      if (bad) {
        actions.append(
          h('button', {
            class: 'btn primary',
            onclick: () => {
              const f = cell.f!.replace(new RegExp(`\\b${bad.replace('.', '\\.')}\\s*\\(`, 'i'), sug[1] + '(');
              this.transact('Fix formula', (tx) => setCellInput(tx, this.doc, s, a.row, a.col, '=' + f));
            },
          }, `Change to ${sug[1]}`),
        );
      }
    }
    if (v.code === '#SPILL!') {
      const blocked = this.doc.engine.blockedSpillAt(s, a.row, a.col);
      if (blocked) actions.append(h('button', { class: 'btn outline', onclick: () => this.select(blocked, a) }, 'Show the spill area'));
    }
    if (cell?.f && !/^IFERROR\(/i.test(cell.f) && v.code !== '#CIRC!') {
      actions.append(
        h('button', {
          class: 'btn',
          title: 'Wrap the formula in IFERROR so it shows a blank instead of an error',
          onclick: () => this.transact('Wrap in IFERROR', (tx) => setCellInput(tx, this.doc, s, a.row, a.col, `=IFERROR(${cell.f}, "")`)),
        }, 'Hide this error'),
      );
    }
    actions.append(h('button', { class: 'btn', title: 'Dismiss', onclick: () => (el.hidden = true) }, '✕'));
    el.append(
      h('span', { class: 'icon' }, '⛔'),
      h(
        'div',
        { class: 'body' },
        h('div', null, h('span', { class: 'title' }, help?.title ?? 'Error'), h('span', { class: 'code' }, v.code)),
        h('div', { class: 'msg' }, v.message ?? ''),
        help ? h('div', { class: 'fix' }, help.fix) : null,
      ),
      actions,
    );
  }

  // ---------- files ----------

  private isPristine(): boolean {
    return !this.doc.dirty && !this.doc.filePath && this.doc.wb.sheets.length === 1 && this.doc.wb.sheets[0].cells.size === 0;
  }

  private async openStartupFile(): Promise<void> {
    const params = new URLSearchParams(location.search);
    const template = params.get('template');
    if (template) {
      this.loadTemplate(template);
      return;
    }
    const p = params.get('file');
    if (!p) return;
    try {
      const f = await platform.openPath(p);
      if (f) await this.loadFile(f);
    } catch (err) {
      toast(`Couldn't open ${p}: ${(err as Error).message}`, { kind: 'error' });
    }
  }

  async newWorkbook(): Promise<void> {
    if (isDesktop) {
      platform.newWindow();
      return;
    }
    if (this.doc.dirty && !(await this.confirmDiscard())) return;
    this.loadWorkbook(new Workbook(), null, 'Untitled', null);
  }

  /** File → New from template. Opens in this window when it's empty, otherwise in a new one. */
  async newFromTemplate(id?: string): Promise<void> {
    const pick = id ?? (await pickTemplate());
    if (!pick) return;
    if (!this.isPristine()) {
      if (isDesktop) {
        platform.newWindow(undefined, pick);
        return;
      }
      if (this.doc.dirty && !(await this.confirmDiscard())) return;
    }
    this.loadTemplate(pick);
  }

  private loadTemplate(id: string): void {
    const info = TEMPLATES.find((t) => t.id === id);
    const wb = buildTemplate(id);
    if (!info || !wb) return;
    this.loadWorkbook(wb, null, info.name, null);
    this.statusbar.flash(`New ${info.name.toLowerCase()} — the yellow cells are yours to change`);
  }

  async openFile(): Promise<void> {
    const f = await platform.openFileDialog();
    if (!f) return;
    await this.openInBestWindow(f);
  }

  async openRecent(path: string): Promise<void> {
    if (!this.isPristine() && isDesktop) {
      platform.newWindow(path);
      return;
    }
    try {
      const f = await platform.openPath(path);
      if (f) await this.loadFile(f);
    } catch (err) {
      toast((err as Error).message, { kind: 'error' });
    }
  }

  private async openInBestWindow(f: OpenedFile): Promise<void> {
    if (!this.isPristine() && f.path && isDesktop) {
      platform.newWindow(f.path);
      return;
    }
    if (!this.isPristine() && this.doc.dirty && !(await this.confirmDiscard())) return;
    await this.loadFile(f);
  }

  private async confirmDiscard(): Promise<boolean> {
    return (await dialog('Discard unsaved changes?', `"${this.doc.fileName}" has changes that haven't been saved.`, [
      { label: 'Discard changes', value: true, danger: true },
      { label: 'Cancel', value: false },
    ])) as boolean;
  }

  async loadFile(f: OpenedFile): Promise<void> {
    const name = f.name;
    const ext = (/\.([^.]+)$/.exec(name)?.[1] ?? '').toLowerCase();
    this.statusbar.flash(`Opening ${name}…`);
    try {
      let wb: Workbook;
      let type: 'xlsx' | 'csv' | 'tsv';
      let warnings: string[] = [];
      if (ext === 'xlsx' || ext === 'xlsm') {
        const res = await readXlsx(f.data);
        wb = res.wb;
        warnings = res.warnings;
        type = 'xlsx';
        if (ext === 'xlsm') warnings.push('Macros (VBA) aren\'t supported. Saving creates a regular .xlsx without macros.');
      } else if (ext === 'xls') {
        throw new Error('Old .xls files aren\'t supported. Open it in another app and save as .xlsx first.');
      } else {
        const text = decodeText(f.data);
        type = ext === 'tsv' || (ext === 'txt' && text.includes('\t')) ? 'tsv' : 'csv';
        wb = csvToWorkbook(text, name, type === 'tsv' ? '\t' : undefined);
      }
      this.loadWorkbook(wb, f.path, name, ext === 'xlsm' ? 'xlsx' : type);
      if (ext === 'xlsm' && f.path) this.doc.filePath = null;
      if (warnings.length) toast(warnings.join(' '), { ms: 9000 });
      this.statusbar.flash(`Opened ${name}`);
    } catch (err) {
      toast(`Couldn't open "${name}": ${(err as Error).message}`, { kind: 'error' });
    }
  }

  private loadWorkbook(wb: Workbook, path: string | null, name: string, type: 'xlsx' | 'csv' | 'tsv' | null): void {
    this.doc.load(wb);
    this.doc.filePath = path;
    this.doc.fileName = name;
    this.doc.fileType = type;
    this.csvChoice = null;
    this.cf = new CFEvaluator(this.doc);
    this.clipboard = null;
    this.grid.setCopyMarquee(null);
    const s = this.sheet;
    this.selSheetId = s.id;
    this.sel.setCell(s, s.view.activeRow, s.view.activeCol);
    this.panel.close();
    this.grid.relayout();
    this.grid.restoreScroll();
    this.refreshAll();
    this.focusGrid();
  }

  /** Saves to the current file (or asks where). Returns true when saved. */
  async save(): Promise<boolean> {
    if (this.isEditing && !this.commitEdit()) return false;
    const d = this.doc;
    if (!d.filePath || !isDesktop) return this.saveAs();
    if (d.fileType === 'csv' || d.fileType === 'tsv') {
      const choice = this.csvChoice ?? (this.hasNonCsvFeatures() ? await this.askCsvChoice() : 'csv');
      if (choice === null) return false;
      this.csvChoice = choice;
      if (choice === 'xlsx') {
        const path = d.filePath.replace(/\.(csv|tsv|txt)$/i, '') + '.xlsx';
        const chosen = await platform.saveFileDialog(path, 'xlsx');
        if (!chosen) return false;
        return this.writeTo(chosen, 'xlsx');
      }
      return this.writeTo(d.filePath, d.fileType);
    }
    return this.writeTo(d.filePath, 'xlsx');
  }

  async saveAs(kind?: SaveKind): Promise<boolean> {
    if (this.isEditing && !this.commitEdit()) return false;
    const base = this.doc.fileName.replace(/\.[^.]+$/, '') || 'Untitled';
    const k: SaveKind = kind ?? 'xlsx';
    const dir = this.doc.filePath ? this.doc.filePath.replace(/[\\/][^\\/]*$/, '') : '';
    const suggested = (dir ? dir + (dir.includes('\\') ? '\\' : '/') : '') + `${base}.${k}`;
    const path = await platform.saveFileDialog(suggested, k);
    if (!path) return false;
    const ext = (/\.([^.]+)$/.exec(path)?.[1] ?? k).toLowerCase();
    const type: SaveKind = ext === 'csv' ? 'csv' : ext === 'tsv' || ext === 'txt' ? 'tsv' : 'xlsx';
    if (type !== 'xlsx' && this.hasNonCsvFeatures()) {
      const ok = await dialog(
        'Save as CSV?',
        'CSV files keep only the values of the current sheet. Formatting, formulas, charts and other sheets will not be saved in this file.',
        [
          { label: 'Save CSV anyway', value: true, primary: true },
          { label: 'Cancel', value: false },
        ],
      );
      if (!ok) return false;
    }
    return this.writeTo(path, type);
  }

  async exportCsv(kind: 'csv' | 'tsv'): Promise<void> {
    const base = this.doc.fileName.replace(/\.[^.]+$/, '') || 'Untitled';
    const name = this.doc.wb.sheets.length > 1 ? `${base} - ${this.sheet.name}.${kind}` : `${base}.${kind}`;
    const path = await platform.saveFileDialog(name, kind);
    if (!path) return;
    const data = encodeCSV(sheetToCSV(this.doc, this.sheet, kind === 'tsv' ? '\t' : ','));
    const res = await platform.saveFile(path, data);
    if (res.ok) toast(`Exported "${this.sheet.name}" to ${path.split(/[\\/]/).pop()}`, { kind: 'success' });
    else toast(res.error ?? 'Export failed.', { kind: 'error' });
  }

  private hasNonCsvFeatures(): boolean {
    const wb = this.doc.wb;
    if (wb.sheets.length > 1) return true;
    const s = wb.sheets[0];
    if (s.charts.length || s.cf.length || s.merges.length) return true;
    for (const c of s.cells.values()) {
      if (c.f !== undefined) return true;
      if (c.s) {
        const st = wb.styles.get(c.s);
        if (Object.keys(st).some((k) => k !== 'numFmt')) return true;
      }
    }
    return false;
  }

  private askCsvChoice(): Promise<'xlsx' | 'csv' | null> {
    return dialog<'xlsx' | 'csv' | null>(
      'This file is a CSV',
      'CSV files only keep the values of one sheet. Your formatting, formulas or extra sheets would be lost. Save a copy as an Excel workbook instead?',
      [
        { label: 'Save as .xlsx', value: 'xlsx', primary: true },
        { label: 'Keep CSV (values only)', value: 'csv' },
        { label: 'Cancel', value: null },
      ],
      { cancelValue: null },
    );
  }

  private async writeTo(path: string, type: SaveKind): Promise<boolean> {
    try {
      this.statusbar.flash('Saving…');
      const data = type === 'xlsx' ? await writeXlsx(this.doc) : encodeCSV(sheetToCSV(this.doc, this.sheet, type === 'tsv' ? '\t' : ','));
      const res = await platform.saveFile(path, data);
      if (!res.ok) {
        toast(res.error ?? 'Saving failed.', { kind: 'error' });
        return false;
      }
      const name = path.split(/[\\/]/).pop() ?? path;
      this.doc.filePath = isDesktop ? path : null;
      this.doc.fileName = name;
      this.doc.fileType = type;
      this.doc.markSaved();
      this.refreshAll();
      this.statusbar.flash(`Saved ${name}`);
      return true;
    } catch (err) {
      toast(`Couldn't save: ${(err as Error).message}`, { kind: 'error' });
      return false;
    }
  }

  private async requestClose(): Promise<void> {
    if (!this.doc.dirty) {
      platform.confirmClose();
      return;
    }
    const choice = await dialog<'save' | 'discard' | 'cancel'>(
      `Save changes to "${this.doc.fileName}"?`,
      'Your changes will be lost if you don\'t save them.',
      [
        { label: 'Save', value: 'save', primary: true },
        { label: 'Don\'t save', value: 'discard', danger: true },
        { label: 'Cancel', value: 'cancel' },
      ],
      { cancelValue: 'cancel' },
    );
    if (choice === 'cancel') return;
    if (choice === 'save' && !(await this.save())) return;
    platform.confirmClose();
  }

  private bindDrop(): void {
    let overlay: HTMLDivElement | null = null;
    const host = this.grid.el;
    host.addEventListener('dragover', (e) => {
      if (!e.dataTransfer?.types.includes('Files')) return;
      e.preventDefault();
      if (!overlay) {
        overlay = h('div', { class: 'drop-overlay' }, 'Drop to open the spreadsheet');
        host.append(overlay);
      }
    });
    host.addEventListener('dragleave', (e) => {
      if (e.target === host || !host.contains(e.relatedTarget as Node)) {
        overlay?.remove();
        overlay = null;
      }
    });
    host.addEventListener('drop', async (e) => {
      e.preventDefault();
      overlay?.remove();
      overlay = null;
      const file = e.dataTransfer?.files?.[0];
      if (!file) return;
      if (!/\.(xlsx|xlsm|csv|tsv|txt)$/i.test(file.name)) {
        toast('Gridline opens .xlsx, .csv and .tsv files.', { kind: 'error' });
        return;
      }
      const f = await platform.readDroppedFile(file);
      await this.openInBestWindow(f);
    });
  }

  // ---------- misc helpers used by commands ----------

  /** The block of data around the active cell (for sorting, filtering, charts). */
  dataRegion(): Range {
    const r = this.sel.range;
    if (r.r1 !== r.r2 || r.c1 !== r.c2) {
      // Clamp whole columns/rows to the used area
      const b = this.doc.engine.usedBounds(this.sheet);
      return { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, b.maxRow)), c2: Math.min(r.c2, Math.max(r.c1, b.maxCol)) };
    }
    return currentRegion(this.doc, this.sheet, r.r1, r.c1);
  }

  selectionLabel(): string {
    const r = this.sel.range;
    if (this.sel.ranges.length > 1) return `${this.sel.ranges.length} ranges`;
    return formatRange(r);
  }

  valueAt(row: number, col: number) {
    return this.doc.engine.getValue(this.sheet, row, col);
  }

  rangeRef(r: Range): RangeRef {
    return new RangeRef(this.sheet, r.r1, r.c1, r.r2, r.c2);
  }

  cellExists(row: number, col: number): boolean {
    return this.sheet.cells.has(cellKey(row, col));
  }

  isActiveIn(r: Range): boolean {
    return rangeContains(r, this.sel.active.row, this.sel.active.col);
  }

  parseAddress(s: string) {
    return parseCell(s);
  }
}
