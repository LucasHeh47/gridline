// Every action in the app, with keyboard shortcuts. The toolbar, menus and command palette all use these.
import * as L from 'lucide';
import type { IconNode } from 'lucide';
import { formatCell, isFullColumn, isFullRow, MAX_COLS, MAX_ROWS, type Range } from '../core/address';
import { createChart } from '../core/charts';
import { external } from '../core/data/external';
import { applyBorders, changeDecimals, mergeCells, setLink, setNote, type BorderPreset, type MergeMode } from '../core/ops/cells';
import { applyFilter, detectHeader, removeDuplicates, sortRange } from '../core/ops/data';
import { addSheet, deleteSheet, duplicateSheet, renameSheet, setSheetHidden } from '../core/ops/sheets';
import { restructure, setHidden } from '../core/ops/structure';
import { validateSheetName, type ChartType } from '../core/workbook';
import type { App } from './app';
import { platform } from './platform';
import { dialog, promptDialog, toast } from './widgets';
import { openPalette } from './palette';
import { showShortcuts, showAbout } from './help';

export interface Command {
  id: string;
  title: string;
  category: string;
  keys?: string[];
  icon?: IconNode;
  keywords?: string;
  /** Allowed while a cell is being edited. */
  whileEditing?: boolean;
  /** Runs even when focus is in a text field. */
  global?: boolean;
  hidden?: boolean;
  run: () => unknown;
  enabled?: () => boolean;
  checked?: () => boolean;
}

const CODE_KEYS: Record<string, string> = {
  Semicolon: ';',
  Equal: '=',
  Minus: '-',
  Backquote: '`',
  Comma: ',',
  Period: '.',
  Slash: '/',
  Backslash: '\\',
  BracketLeft: '[',
  BracketRight: ']',
  Quote: "'",
  NumpadAdd: '+',
  NumpadSubtract: '-',
  Space: 'Space',
};

/** Turns a keyboard event into a string like "Ctrl+Shift+L". */
export function eventKey(e: KeyboardEvent): string {
  let key: string;
  if (/^Key[A-Z]$/.test(e.code)) key = e.code.slice(3);
  else if (/^Digit\d$/.test(e.code)) key = e.code.slice(5);
  else if (CODE_KEYS[e.code]) key = CODE_KEYS[e.code];
  else key = e.key.length === 1 ? e.key.toUpperCase() : e.key;
  const parts: string[] = [];
  if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
  if (e.altKey) parts.push('Alt');
  if (e.shiftKey) parts.push('Shift');
  parts.push(key);
  return parts.join('+');
}

export class Commands {
  private list: Command[] = [];
  private byId = new Map<string, Command>();
  private byKey = new Map<string, Command>();

  constructor(private app: App) {
    this.register();
  }

  all(): Command[] {
    return this.list.filter((c) => !c.hidden);
  }

  get(id: string): Command | undefined {
    return this.byId.get(id);
  }

  match(e: KeyboardEvent): Command | undefined {
    if (['Control', 'Shift', 'Alt', 'Meta'].includes(e.key)) return undefined;
    return this.byKey.get(eventKey(e));
  }

  run(id: string): void {
    const c = this.byId.get(id);
    if (!c) return;
    if (c.enabled && !c.enabled()) return;
    if (this.app.isEditing && !c.whileEditing) {
      if (!this.app.commitEdit()) return;
    }
    if (this.app.settings.bool('keyboard.remember')) {
      const recent = this.app.prefs.recentCommands.filter((x) => x !== id);
      recent.unshift(id);
      this.app.prefs.recentCommands = recent.slice(0, 8);
      this.app.savePrefs();
    }
    try {
      const r = c.run();
      if (r instanceof Promise) r.catch((err: Error) => toast(String(err?.message ?? err), { kind: 'error' }));
    } catch (err) {
      toast(String((err as Error).message ?? err), { kind: 'error' });
    }
  }

  private add(c: Command): void {
    this.list.push(c);
    this.byId.set(c.id, c);
    for (const k of c.keys ?? []) this.byKey.set(k, c);
  }

  private register(): void {
    const app = this.app;
    const a = (c: Command) => this.add(c);
    const sel = () => app.sel;
    const sheet = () => app.sheet;
    const style = () => app.styleAtActive();

    // ---------- File ----------
    a({ id: 'file.new', title: 'New workbook', category: 'File', keys: ['Ctrl+N'], icon: L.FilePlus, global: true, run: () => app.newWorkbook() });
    a({ id: 'file.home', title: 'Home — recent files and presets', category: 'File', keys: ['Ctrl+Shift+H'], icon: L.House, keywords: 'start page recent backstage file', global: true, run: () => app.openHome('home') });
    a({ id: 'file.template', title: 'New from a preset…', category: 'File', icon: L.LayoutTemplate, keywords: 'template budget loan mortgage debt savings net worth stocks crypto nfl fantasy football trip travel project tracker', global: true, run: () => app.openHome('new') });
    a({ id: 'file.open', title: 'Open…', category: 'File', keys: ['Ctrl+O'], icon: L.FolderOpen, global: true, run: () => app.openFile() });
    a({ id: 'file.save', title: 'Save', category: 'File', keys: ['Ctrl+S'], icon: L.Save, global: true, run: () => void app.save() });
    a({ id: 'file.saveAs', title: 'Save as…', category: 'File', keys: ['Ctrl+Shift+S'], icon: L.Save, global: true, run: () => void app.saveAs() });
    a({ id: 'file.exportCsv', title: 'Export current sheet as CSV…', category: 'File', icon: L.Download, keywords: 'comma separated', run: () => app.exportCsv('csv') });
    a({ id: 'file.exportTsv', title: 'Export current sheet as TSV…', category: 'File', icon: L.Download, keywords: 'tab separated', run: () => app.exportCsv('tsv') });
    a({ id: 'file.showInFolder', title: 'Show file in folder', category: 'File', icon: L.FolderOpen, enabled: () => !!app.doc.filePath, run: () => app.doc.filePath && platform.showInFolder(app.doc.filePath) });

    // ---------- Edit ----------
    a({ id: 'edit.undo', title: 'Undo', category: 'Edit', keys: ['Ctrl+Z'], icon: L.Undo2, enabled: () => app.doc.history.canUndo, run: () => app.doc.undo() });
    a({ id: 'edit.redo', title: 'Redo', category: 'Edit', keys: ['Ctrl+Y', 'Ctrl+Shift+Z'], icon: L.Redo2, enabled: () => app.doc.history.canRedo, run: () => app.doc.redo() });
    a({ id: 'edit.cut', title: 'Cut', category: 'Edit', keys: ['Ctrl+X'], icon: L.Scissors, run: () => app.copy(true) });
    a({ id: 'edit.copy', title: 'Copy', category: 'Edit', keys: ['Ctrl+C'], icon: L.Copy, run: () => app.copy(false) });
    a({ id: 'edit.paste', title: 'Paste', category: 'Edit', keys: ['Ctrl+V'], icon: L.ClipboardPaste, run: () => app.paste('all') });
    a({ id: 'edit.pasteValues', title: 'Paste values only', category: 'Edit', keys: ['Ctrl+Shift+V'], icon: L.ClipboardPaste, keywords: 'paste special', run: () => app.paste('values') });
    a({ id: 'edit.pasteFormats', title: 'Paste formatting only', category: 'Edit', icon: L.Paintbrush, keywords: 'paste special', run: () => app.paste('formats') });
    a({ id: 'edit.pasteFormulas', title: 'Paste formulas only', category: 'Edit', icon: L.SquareFunction, keywords: 'paste special', run: () => app.paste('formulas') });
    a({ id: 'edit.pasteTranspose', title: 'Paste transposed (swap rows and columns)', category: 'Edit', icon: L.ArrowLeftRight, keywords: 'paste special transpose', run: () => app.paste('transpose') });
    a({ id: 'edit.pasteNoBorders', title: 'Paste without borders', category: 'Edit', icon: L.ClipboardPaste, keywords: 'paste special', run: () => app.paste('noBorders') });
    a({ id: 'edit.clearContents', title: 'Clear contents', category: 'Edit', icon: L.Eraser, run: () => app.clear('contents') });
    a({ id: 'edit.clearFormats', title: 'Clear formatting', category: 'Edit', icon: L.RemoveFormatting, run: () => app.clear('formats') });
    a({ id: 'edit.clearAll', title: 'Clear everything', category: 'Edit', icon: L.Trash2, run: () => app.clear('all') });
    a({ id: 'edit.fillDown', title: 'Fill down', category: 'Edit', keys: ['Ctrl+D'], icon: L.ArrowDownToLine, run: () => app.fillDirection('down') });
    a({ id: 'edit.fillRight', title: 'Fill right', category: 'Edit', keys: ['Ctrl+R'], icon: L.ArrowRightToLine, run: () => app.fillDirection('right') });
    a({ id: 'edit.find', title: 'Find', category: 'Edit', keys: ['Ctrl+F'], icon: L.Search, global: true, run: () => app.findBar.open(false) });
    a({ id: 'edit.replace', title: 'Find and replace', category: 'Edit', keys: ['Ctrl+H'], icon: L.Replace, global: true, run: () => app.findBar.open(true) });
    a({ id: 'edit.goto', title: 'Go to cell…', category: 'Edit', keys: ['Ctrl+G', 'F5'], icon: L.Crosshair, keywords: 'jump navigate address', run: () => app.formulaBar.focusNameBox() });
    a({
      id: 'edit.selectAll',
      title: 'Select all (data region, then whole sheet)',
      category: 'Edit',
      keys: ['Ctrl+A'],
      icon: L.SquareDashed,
      run: () => {
        const r = sel().range;
        const region = app.dataRegion();
        const isRegion = r.r1 === region.r1 && r.c1 === region.c1 && r.r2 === region.r2 && r.c2 === region.c2;
        if (!isRegion && (region.r1 !== region.r2 || region.c1 !== region.c2)) sel().setRange(region, { ...sel().active });
        else sel().selectAll();
        app.onSelectionChanged(false);
      },
    });
    a({ id: 'edit.selectColumn', title: 'Select entire column', category: 'Edit', keys: ['Ctrl+Space'], icon: L.Columns3, run: () => {
      const r = sel().range;
      sel().selectColumns(r.c1, r.c2);
      sel().active = { row: app.sel.active.row, col: r.c1 };
      app.onSelectionChanged(false);
    } });
    a({ id: 'edit.selectRow', title: 'Select entire row', category: 'Edit', keys: ['Shift+Space'], icon: L.Rows3, run: () => {
      const r = sel().range;
      sel().selectRows(r.r1, r.r2);
      app.onSelectionChanged(false);
    } });
    a({ id: 'edit.insertDate', title: 'Insert today\'s date', category: 'Edit', keys: ['Ctrl+;'], icon: L.CalendarDays, whileEditing: true, run: () => insertNow(app, 'date') });
    a({ id: 'edit.insertTime', title: 'Insert current time', category: 'Edit', keys: ['Ctrl+Shift+;'], icon: L.Clock, whileEditing: true, run: () => insertNow(app, 'time') });
    a({ id: 'edit.editCell', title: 'Edit cell', category: 'Edit', icon: L.Pencil, run: () => app.startEdit({ mode: 'edit' }) });
    a({ id: 'edit.formatPainter', title: 'Copy formatting (format painter)', category: 'Edit', icon: L.Paintbrush, keywords: 'format painter brush', checked: () => !!app.formatPainter, run: () => toggleFormatPainter(app) });
    a({ id: 'edit.note', title: 'Add or edit note', category: 'Edit', keys: ['Shift+F2'], icon: L.MessageSquare, keywords: 'comment', run: () => editNote(app) });
    a({ id: 'edit.deleteNote', title: 'Delete note', category: 'Edit', icon: L.MessageSquareX, enabled: () => !!sheet().getCell(sel().active.row, sel().active.col)?.note, run: () => app.transact('Delete note', (tx) => setNote(tx, sheet(), sel().active.row, sel().active.col, undefined)) });
    a({ id: 'edit.link', title: 'Insert or edit link', category: 'Edit', keys: ['Ctrl+Alt+K'], icon: L.Link, keywords: 'hyperlink url', run: () => editLink(app) });
    a({ id: 'edit.removeLink', title: 'Remove link', category: 'Edit', icon: L.Unlink, run: () => app.clear('links') });

    // ---------- Format ----------
    const toggle = (key: 'bold' | 'italic' | 'underline' | 'strike', label: string) => () => app.applyStyle(label, { [key]: !style()[key] || null });
    a({ id: 'fmt.bold', title: 'Bold', category: 'Format', keys: ['Ctrl+B', 'Ctrl+2'], icon: L.Bold, checked: () => !!style().bold, run: toggle('bold', 'Bold') });
    a({ id: 'fmt.italic', title: 'Italic', category: 'Format', keys: ['Ctrl+I', 'Ctrl+3'], icon: L.Italic, checked: () => !!style().italic, run: toggle('italic', 'Italic') });
    a({ id: 'fmt.underline', title: 'Underline', category: 'Format', keys: ['Ctrl+U', 'Ctrl+4'], icon: L.Underline, checked: () => !!style().underline, run: toggle('underline', 'Underline') });
    a({ id: 'fmt.strike', title: 'Strikethrough', category: 'Format', keys: ['Ctrl+5'], icon: L.Strikethrough, checked: () => !!style().strike, run: toggle('strike', 'Strikethrough') });
    a({ id: 'fmt.cells', title: 'Format cells…', category: 'Format', keys: ['Ctrl+1'], icon: L.Settings2, keywords: 'number font border fill alignment', run: () => app.panel.open('format') });
    const align = (h: string | null, label: string) => () => app.applyStyle(label, { hAlign: (style().hAlign === h ? null : h) as never });
    a({ id: 'fmt.alignLeft', title: 'Align left', category: 'Format', icon: L.AlignLeft, checked: () => style().hAlign === 'left', run: align('left', 'Align left') });
    a({ id: 'fmt.alignCenter', title: 'Align center', category: 'Format', keys: ['Ctrl+E'], icon: L.AlignCenter, checked: () => style().hAlign === 'center', run: align('center', 'Align center') });
    a({ id: 'fmt.alignRight', title: 'Align right', category: 'Format', icon: L.AlignRight, checked: () => style().hAlign === 'right', run: align('right', 'Align right') });
    a({ id: 'fmt.alignTop', title: 'Align to top', category: 'Format', icon: L.AlignVerticalJustifyStart, checked: () => style().vAlign === 'top', run: () => app.applyStyle('Align top', { vAlign: 'top' }) });
    a({ id: 'fmt.alignMiddle', title: 'Align to middle', category: 'Format', icon: L.AlignVerticalJustifyCenter, checked: () => style().vAlign === 'middle', run: () => app.applyStyle('Align middle', { vAlign: 'middle' }) });
    a({ id: 'fmt.alignBottom', title: 'Align to bottom', category: 'Format', icon: L.AlignVerticalJustifyEnd, checked: () => !style().vAlign || style().vAlign === 'bottom', run: () => app.applyStyle('Align bottom', { vAlign: null }) });
    a({ id: 'fmt.wrap', title: 'Wrap text', category: 'Format', icon: L.WrapText, checked: () => !!style().wrap, run: () => {
      const on = !style().wrap;
      app.applyStyle('Wrap text', { wrap: on || null });
      if (on) {
        const rows = new Set<number>();
        for (const r of sel().ranges) for (let i = r.r1; i <= Math.min(r.r2, r.r1 + 500); i++) if (!sheet().rowHeights.has(i)) rows.add(i);
        if (rows.size && rows.size <= 500) app.autofitRows([...rows]);
      }
    } });
    a({ id: 'fmt.indentInc', title: 'Increase indent', category: 'Format', icon: L.IndentIncrease, run: () => app.applyStyle('Indent', (s) => ({ indent: Math.min(15, (s.indent ?? 0) + 1), hAlign: s.hAlign === 'right' ? 'right' : 'left' })) });
    a({ id: 'fmt.indentDec', title: 'Decrease indent', category: 'Format', icon: L.IndentDecrease, run: () => app.applyStyle('Indent', (s) => ({ indent: Math.max(0, (s.indent ?? 0) - 1) || null })) });
    const merge = (mode: MergeMode, title: string, icon: IconNode) =>
      a({
        id: `fmt.merge.${mode}`,
        title,
        category: 'Format',
        icon,
        keywords: 'merge cells combine',
        run: () => {
          let lost = false;
          const tx = app.transact(title, (tx) => {
            for (const r of sel().ranges) lost = mergeCells(tx, app.doc, sheet(), clampRange(app, r), mode) || lost;
          });
          if (tx && lost) toast('Merged — only the top-left value was kept.', { action: { label: 'Undo', run: () => app.doc.undo() } });
          if (tx) {
            sel().setRange(sel().range, { row: sel().range.r1, col: sel().range.c1 });
            app.onSelectionChanged(false);
          }
        },
      });
    merge('center', 'Merge and center', L.TableCellsMerge);
    merge('merge', 'Merge cells', L.TableCellsMerge);
    merge('across', 'Merge across (each row)', L.TableCellsMerge);
    merge('unmerge', 'Unmerge cells', L.TableCellsSplit);
    const fmt = (id: string, title: string, f: string, keys?: string[], icon?: IconNode) =>
      a({ id, title, category: 'Number format', keys, icon: icon ?? L.Hash, keywords: 'number format', run: () => app.applyStyle(title, { numFmt: f === 'General' ? null : f }) });
    fmt('fmt.general', 'Format as automatic (General)', 'General', ['Ctrl+Shift+`']);
    fmt('fmt.number', 'Format as number (1,234.56)', '#,##0.00', ['Ctrl+Shift+1']);
    fmt('fmt.currency', 'Format as currency ($1,234.56)', '"$"#,##0.00', ['Ctrl+Shift+4'], L.DollarSign);
    fmt('fmt.accounting', 'Format as accounting', '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)', undefined, L.DollarSign);
    fmt('fmt.percent', 'Format as percent (12%)', '0%', ['Ctrl+Shift+5'], L.Percent);
    fmt('fmt.date', 'Format as date', 'm/d/yyyy', ['Ctrl+Shift+3'], L.CalendarDays);
    fmt('fmt.dateIso', 'Format as date (2025-03-14)', 'yyyy-mm-dd', undefined, L.CalendarDays);
    fmt('fmt.time', 'Format as time', 'h:mm AM/PM', ['Ctrl+Shift+2'], L.Clock);
    fmt('fmt.scientific', 'Format as scientific (1.23E+04)', '0.00E+00', ['Ctrl+Shift+6']);
    fmt('fmt.text', 'Format as plain text (keep exactly as typed)', '@', undefined, L.Type);
    a({ id: 'fmt.thousands', title: 'Thousands separator (1,234)', category: 'Number format', icon: L.Hash, run: () => app.applyStyle('Comma style', { numFmt: '#,##0.00' }) });
    a({ id: 'fmt.decInc', title: 'Increase decimals', category: 'Number format', icon: L.DecimalsArrowRight, run: () => app.transact('Increase decimals', (tx) => changeDecimals(tx, app.doc, sheet(), sel().ranges, 1, sel().active.row, sel().active.col)) });
    a({ id: 'fmt.decDec', title: 'Decrease decimals', category: 'Number format', icon: L.DecimalsArrowLeft, run: () => app.transact('Decrease decimals', (tx) => changeDecimals(tx, app.doc, sheet(), sel().ranges, -1, sel().active.row, sel().active.col)) });
    a({ id: 'fmt.fontBigger', title: 'Increase font size', category: 'Format', keys: ['Ctrl+Shift+.'], icon: L.AArrowUp, run: () => app.applyStyle('Font size', (s) => ({ fontSize: nextSize(s.fontSize ?? 11, 1) })) });
    a({ id: 'fmt.fontSmaller', title: 'Decrease font size', category: 'Format', keys: ['Ctrl+Shift+,'], icon: L.AArrowDown, run: () => app.applyStyle('Font size', (s) => ({ fontSize: nextSize(s.fontSize ?? 11, -1) })) });
    const border = (p: BorderPreset, title: string, icon: IconNode) =>
      a({ id: `fmt.border.${p}`, title, category: 'Borders', icon, keywords: 'border lines', run: () => app.transact(title, (tx) => {
        for (const r of sel().ranges) applyBorders(tx, app.doc.wb, sheet(), clampRange(app, r), p, { style: 'thin', color: app.ribbon.borderColor });
      }) });
    border('all', 'All borders', L.Grid3x3);
    border('outer', 'Outside borders', L.Square);
    border('thickOuter', 'Thick outside border', L.Square);
    border('inner', 'Inside borders', L.Plus);
    border('top', 'Top border', L.PanelTop);
    border('bottom', 'Bottom border', L.PanelBottom);
    border('bottomDouble', 'Double bottom border', L.PanelBottom);
    border('left', 'Left border', L.PanelLeft);
    border('right', 'Right border', L.PanelRight);
    border('horizontal', 'Inner horizontal borders', L.Rows3);
    border('vertical', 'Inner vertical borders', L.Columns3);
    border('none', 'Remove borders', L.SquareDashed);
    a({ id: 'fmt.advanced', title: 'Advanced format — format this table automatically', category: 'Format', keys: ['Ctrl+Alt+F'], icon: L.Sparkles, keywords: 'smart auto format table magic ai style', run: () => app.advancedFormat() });
    a({ id: 'fmt.cf', title: 'Conditional formatting…', category: 'Format', icon: L.Highlighter, keywords: 'color scale data bars highlight rules icon sets', run: () => app.panel.open('cf') });
    a({ id: 'fmt.clearCf', title: 'Clear conditional formatting from selection', category: 'Format', icon: L.Highlighter, run: () => clearCf(app) });

    // ---------- Insert / delete ----------
    a({ id: 'ins.rowsAbove', title: 'Insert rows above', category: 'Insert', keys: ['Ctrl+Shift+='], icon: L.BetweenHorizontalStart, keywords: 'add row', run: () => insertRows(app, 'above') });
    a({ id: 'ins.rowsBelow', title: 'Insert rows below', category: 'Insert', icon: L.BetweenHorizontalEnd, keywords: 'add row', run: () => insertRows(app, 'below') });
    a({ id: 'ins.colsLeft', title: 'Insert columns left', category: 'Insert', icon: L.BetweenVerticalStart, keywords: 'add column', run: () => insertCols(app, 'left') });
    a({ id: 'ins.colsRight', title: 'Insert columns right', category: 'Insert', icon: L.BetweenVerticalEnd, keywords: 'add column', run: () => insertCols(app, 'right') });
    a({ id: 'del.rows', title: 'Delete rows', category: 'Insert', keys: ['Ctrl+-'], icon: L.Trash2, keywords: 'remove row', run: () => deleteRowsOrCols(app, 'row') });
    a({ id: 'del.cols', title: 'Delete columns', category: 'Insert', icon: L.Trash2, keywords: 'remove column', run: () => deleteRowsOrCols(app, 'col') });
    a({ id: 'ins.autosum', title: 'AutoSum', category: 'Insert', keys: ['Alt+='], icon: L.Sigma, keywords: 'sum total', run: () => autoSum(app, 'SUM') });
    for (const fn of ['AVERAGE', 'COUNT', 'MAX', 'MIN']) a({ id: `ins.auto.${fn}`, title: `Auto ${fn[0] + fn.slice(1).toLowerCase()}`, category: 'Insert', icon: L.Sigma, run: () => autoSum(app, fn) });
    a({ id: 'ins.function', title: 'Insert function…', category: 'Insert', keys: ['Shift+F3'], icon: L.SquareFunction, keywords: 'formula browse functions', run: () => app.panel.open('functions') });
    const chart = (type: ChartType, title: string, icon: IconNode) => a({ id: `ins.chart.${type}`, title, category: 'Insert', icon, keywords: 'chart graph plot', run: () => insertChart(app, type) });
    a({ id: 'ins.chart.quick', title: 'Quick chart', category: 'Insert', keys: ['F11', 'Alt+F1'], icon: L.ChartColumnIncreasing, keywords: 'chart graph plot', run: () => insertChart(app, app.settings.str('charts.type') as ChartType) });
    const startFormula = (id: string, title: string, icon: IconNode, text: string, keywords: string) =>
      a({ id, title, category: 'Insert', icon, keywords, run: () => {
        if (app.isEditing) app.editor.insertText(app.editor.text.startsWith('=') ? text.slice(1) : text);
        else app.editor.start({ mode: 'enter', initial: text, source: 'cell' });
      } });
    startFormula('ins.stock', 'Stock price (=STOCK)', L.ChartCandlestick, '=STOCK("', 'quote ticker share price live');
    startFormula('ins.crypto', 'Crypto price (=CRYPTO)', L.Bitcoin, '=CRYPTO("', 'bitcoin ethereum coin price live');
    startFormula('ins.fx', 'Currency conversion (=FX)', L.ArrowLeftRight, '=FX(', 'exchange rate money convert');
    startFormula('ins.stockHistory', 'Stock price history (=STOCKHISTORY)', L.ChartLine, '=STOCKHISTORY("', 'chart past prices ticker');
    startFormula('ins.fxrate', 'Exchange rate (=FXRATE)', L.Landmark, '=FXRATE("', 'currency rate money');
    startFormula('ins.weather', 'Weather (=WEATHER)', L.CloudSun, '=WEATHER("', 'temperature forecast rain');
    startFormula('ins.forecast', 'Weather forecast table (=WEATHER.FORECAST)', L.CalendarDays, '=WEATHER.FORECAST("', 'temperature days rain');
    startFormula('ins.holidays', 'Public holidays (=HOLIDAYS)', L.PartyPopper, '=HOLIDAYS("', 'calendar days off country');
    startFormula('ins.country', 'Country statistics (=COUNTRYDATA)', L.Globe, '=COUNTRYDATA("', 'population gdp inflation world bank');
    chart('column', 'Column chart', L.ChartColumn);
    chart('bar', 'Bar chart', L.ChartBar);
    chart('line', 'Line chart', L.ChartLine);
    chart('area', 'Area chart', L.ChartArea);
    chart('pie', 'Pie chart', L.ChartPie);
    chart('doughnut', 'Doughnut chart', L.Donut);
    chart('scatter', 'Scatter chart', L.ChartScatter);

    // ---------- Data ----------
    a({ id: 'data.sortAsc', title: 'Sort A → Z (smallest to largest)', category: 'Data', icon: L.ArrowDownAZ, keywords: 'sort ascending', run: () => quickSort(app, false) });
    a({ id: 'data.sortDesc', title: 'Sort Z → A (largest to smallest)', category: 'Data', icon: L.ArrowUpZA, keywords: 'sort descending', run: () => quickSort(app, true) });
    a({ id: 'data.sortCustom', title: 'Custom sort…', category: 'Data', icon: L.ArrowUpDown, keywords: 'sort multiple columns levels', run: () => app.panel.open('sort') });
    a({ id: 'data.filter', title: 'Filter (toggle)', category: 'Data', keys: ['Ctrl+Shift+L'], icon: L.Funnel, keywords: 'autofilter', checked: () => !!sheet().filter, run: () => toggleFilter(app) });
    a({ id: 'data.clearFilters', title: 'Clear all filters', category: 'Data', icon: L.FunnelX, enabled: () => !!sheet().filter && Object.keys(sheet().filter!.conds).length > 0, run: () => app.transact('Clear filters', (tx) => applyFilter(tx, app.doc, sheet(), { ...sheet().filter!, conds: {} })) });
    a({ id: 'data.reapplyFilter', title: 'Reapply filter', category: 'Data', icon: L.ListFilter, enabled: () => !!sheet().filter, run: () => app.transact('Reapply filter', (tx) => applyFilter(tx, app.doc, sheet(), sheet().filter)) });
    a({ id: 'data.dedupe', title: 'Remove duplicate rows…', category: 'Data', icon: L.CopyMinus, keywords: 'duplicates unique', run: () => dedupe(app) });
    a({ id: 'data.recalc', title: 'Recalculate everything', category: 'Data', keys: ['F9'], icon: L.RefreshCw, run: () => app.doc.recalcAll() });
    a({
      id: 'data.refreshLive',
      title: 'Refresh live data (NFL stats, stock prices, exchange rates)',
      category: 'Data',
      keys: ['Ctrl+Alt+F5'],
      icon: L.CloudDownload,
      keywords: 'update download nfl stocks fx reload',
      run: () => {
        external.refresh();
        app.statusbar.flash('Refreshing live data…');
      },
    });
    a({ id: 'data.clearCache', title: 'Clear downloaded data cache', category: 'Data', icon: L.Trash2, keywords: 'nfl offline cache', run: async () => {
      await platform.clearDataCache();
      external.refresh();
      toast('Cleared the downloaded data. It will be downloaded again when needed.');
    } });
    a({ id: 'data.trace', title: 'Show cells this formula uses', category: 'Data', icon: L.Workflow, keywords: 'trace precedents', run: () => tracePrecedents(app) });

    // ---------- Automate (macros) ----------
    a({ id: 'macro.editor', title: 'Macro editor', category: 'Automate', keys: ['Alt+F11'], icon: L.SquareCode, keywords: 'vba code visual basic macros script automate', global: true, run: () => (app.macroEditor.isOpen ? app.closeMacroEditor() : app.openMacroEditor()) });
    a({ id: 'macro.run', title: 'Run a macro…', category: 'Automate', keys: ['Alt+F8'], icon: L.Play, keywords: 'vba execute automate', global: true, run: async () => {
      const m = await app.macros.pickMacro('Run a macro');
      if (m) await app.macros.run(m);
    } });
    a({ id: 'macro.new', title: 'New macro (VBA code)', category: 'Automate', icon: L.FileCode, keywords: 'vba sub write code', run: () => {
      app.openMacroEditor();
      app.macroEditor.newModule();
    } });
    a({ id: 'macro.newSteps', title: 'New step macro (no code) — Beta', category: 'Automate', icon: L.ListChecks, keywords: 'no code blocks easy builder steps automate', enabled: () => app.macros.stepsAvailable, run: () => {
      app.openMacroEditor();
      void app.macroEditor.newStepMacro();
    } });
    a({ id: 'macro.button', title: 'Add a button that runs a macro', category: 'Automate', icon: L.MousePointerClick, keywords: 'macro button sheet click', run: () => void app.macros.addButtonFor(null) });
    a({ id: 'macro.shortcut', title: 'Give a macro a keyboard shortcut', category: 'Automate', icon: L.Keyboard, keywords: 'hotkey key macro', run: () => void app.macros.assignShortcutFor(null) });
    a({ id: 'macro.triggers', title: 'Run a macro automatically…', category: 'Automate', icon: L.Zap, keywords: 'trigger event open change automatic', run: () => void app.macros.triggersDialog() });
    a({ id: 'macro.stop', title: 'Stop the running macro', category: 'Automate', icon: L.Square, keywords: 'cancel break', global: true, enabled: () => app.macros.running, run: () => app.macros.stop() });
    a({ id: 'macro.settings', title: 'Macro settings', category: 'Automate', icon: L.ShieldCheck, keywords: 'security trust enable disable macros', global: true, run: () => app.openSettings('macros') });
    a({ id: 'nfl.panel', title: 'NFL data — insert player, team, schedule or standings stats', category: 'Insert', icon: L.Trophy, keywords: 'football stats nflverse player team qb fantasy', run: () => app.panel.open('nfl') });
    const nfl = (mode: string, title: string, icon: IconNode) => a({ id: `nfl.${mode}`, title: `NFL: ${title}`, category: 'Insert', icon, keywords: 'football nfl', run: () => app.panel.open('nfl', mode) });
    nfl('player', 'player stats', L.User);
    nfl('team', 'team stats', L.Users);
    nfl('schedule', 'team schedule', L.CalendarDays);
    nfl('standings', 'standings', L.ListOrdered);
    nfl('leaders', 'league leaders', L.Medal);

    // ---------- View ----------
    a({ id: 'view.freezeTop', title: 'Freeze top row', category: 'View', icon: L.Snowflake, run: () => setFreeze(app, 1, 0) });
    a({ id: 'view.freezeFirstCol', title: 'Freeze first column', category: 'View', icon: L.Snowflake, run: () => setFreeze(app, 0, 1) });
    a({ id: 'view.freezeHere', title: 'Freeze rows and columns above/left of the selected cell', category: 'View', icon: L.Snowflake, keywords: 'freeze panes', run: () => setFreeze(app, sel().active.row, sel().active.col) });
    a({ id: 'view.unfreeze', title: 'Unfreeze', category: 'View', icon: L.Snowflake, enabled: () => !!(sheet().freeze.rows || sheet().freeze.cols), run: () => setFreeze(app, 0, 0) });
    a({ id: 'view.gridlines', title: 'Show gridlines', category: 'View', icon: L.Grid3x3, checked: () => sheet().showGrid, run: () => app.transact('Gridlines', (tx) => tx.setProp(sheet(), 'showGrid', !sheet().showGrid)) });
    a({ id: 'view.formulas', title: 'Show formulas instead of results', category: 'View', keys: ['Ctrl+`'], icon: L.SquareFunction, checked: () => app.showFormulas, run: () => {
      app.showFormulas = !app.showFormulas;
      app.grid.draw();
    } });
    a({ id: 'view.zoomIn', title: 'Zoom in', category: 'View', keys: ['Ctrl+Alt+='], icon: L.ZoomIn, run: () => app.zoomBy(0.1) });
    a({ id: 'view.zoomOut', title: 'Zoom out', category: 'View', keys: ['Ctrl+Alt+-'], icon: L.ZoomOut, run: () => app.zoomBy(-0.1) });
    a({ id: 'view.zoomReset', title: 'Zoom to 100%', category: 'View', keys: ['Ctrl+0'], icon: L.Search, run: () => app.setZoom(1) });
    a({ id: 'view.theme', title: 'Switch light / dark theme', category: 'View', icon: L.SunMoon, keywords: 'dark mode light mode appearance', run: () => app.settings.set('appearance.theme', app.isDark ? 'light' : 'dark') });
    a({ id: 'view.adaptColors', title: 'Dark-friendly cell colors', category: 'View', icon: L.Palette, keywords: 'true colors contrast', checked: () => app.settings.bool('appearance.adaptColors'), run: () => {
      const on = !app.settings.bool('appearance.adaptColors');
      app.settings.set('appearance.adaptColors', on);
      toast(on ? 'Light cell colors are softened in dark mode for readability (files are unchanged).' : 'Showing cell colors exactly as stored.');
    } });
    const settingToggle = (id: string, setting: string, title: string, icon: IconNode, keywords?: string) =>
      a({ id, title, category: 'View', icon, keywords, checked: () => app.settings.bool(setting), run: () => app.settings.set(setting, !app.settings.bool(setting)) });
    settingToggle('view.headings', 'grid.headings', 'Show row and column headings', L.PanelTopDashed, 'headers A B C 1 2 3');
    settingToggle('view.formulaBar', 'appearance.showFormulaBar', 'Show formula bar', L.SquareFunction);
    settingToggle('view.statusBar', 'appearance.showStatusBar', 'Show status bar', L.PanelBottom);
    settingToggle('grid.crosshair', 'grid.crosshair', 'Highlight the active row and column', L.Crosshair, 'reading guide crosshair');
    a({ id: 'view.zoomSelection', title: 'Zoom to selection', category: 'View', icon: L.Scan, run: () => app.zoomToSelection() });
    a({ id: 'view.ribbonFull', title: 'Ribbon: full', category: 'View', icon: L.PanelTop, keywords: 'toolbar size', checked: () => app.settings.str('ribbon.size') === 'full', run: () => app.ribbon.setSize('full') });
    a({ id: 'view.ribbonCompact', title: 'Ribbon: compact (one row)', category: 'View', icon: L.Rows2, keywords: 'toolbar size', checked: () => app.settings.str('ribbon.size') === 'compact', run: () => app.ribbon.setSize('compact') });
    a({ id: 'view.ribbonTabs', title: 'Ribbon: tabs only', category: 'View', keys: ['Ctrl+F1'], icon: L.PanelTopClose, keywords: 'toolbar collapse hide', checked: () => app.settings.str('ribbon.size') === 'tabs', run: () => app.ribbon.setSize(app.settings.str('ribbon.size') === 'tabs' ? 'full' : 'tabs') });

    // ---------- Gridline ----------
    a({ id: 'app.settings', title: 'Settings', category: 'Gridline', keys: ['Ctrl+,'], icon: L.Settings, keywords: 'preferences options customize', global: true, run: () => app.openSettings() });
    a({ id: 'app.settingsLive', title: 'Live data settings', category: 'Gridline', icon: L.Settings2, keywords: 'nfl stocks offline refresh', run: () => app.openSettings('live') });
    a({ id: 'help.updates', title: 'Check for updates', category: 'Gridline', icon: L.RefreshCw, keywords: 'new version upgrade', run: () => app.updates.check(false) });
    a({ id: 'help.whatsNew', title: "What's new in this version", category: 'Gridline', icon: L.Gift, keywords: 'release notes changelog', run: () => app.updates.showWhatsNew() });
    a({ id: 'file.newWindow', title: 'New window', category: 'File', keys: ['Ctrl+Shift+N'], icon: L.AppWindow, global: true, run: () => platform.newWindow() });
    a({ id: 'view.palette', title: 'Search commands', category: 'View', keys: ['Ctrl+K', 'Ctrl+Shift+P', 'F1'], icon: L.Command, global: true, hidden: true, run: () => openPalette(app) });

    // ---------- Rows & columns ----------
    a({ id: 'rows.hide', title: 'Hide rows', category: 'Rows & columns', keys: ['Ctrl+9'], icon: L.EyeOff, run: () => hideRowsCols(app, 'row', true) });
    a({ id: 'rows.unhide', title: 'Unhide rows', category: 'Rows & columns', keys: ['Ctrl+Shift+9'], icon: L.Eye, run: () => hideRowsCols(app, 'row', false) });
    a({ id: 'cols.hide', title: 'Hide columns', category: 'Rows & columns', icon: L.EyeOff, run: () => hideRowsCols(app, 'col', true) });
    a({ id: 'cols.unhide', title: 'Unhide columns', category: 'Rows & columns', keys: ['Ctrl+Shift+0'], icon: L.Eye, run: () => hideRowsCols(app, 'col', false) });
    a({ id: 'cols.autofit', title: 'Auto-fit column width', category: 'Rows & columns', icon: L.MoveHorizontal, run: () => app.autofitColumns(selectedCols(app)) });
    a({ id: 'rows.autofit', title: 'Auto-fit row height', category: 'Rows & columns', icon: L.MoveVertical, run: () => app.autofitRows(selectedRows(app)) });
    a({ id: 'cols.width', title: 'Set column width…', category: 'Rows & columns', icon: L.MoveHorizontal, run: () => setSize(app, 'col') });
    a({ id: 'rows.height', title: 'Set row height…', category: 'Rows & columns', icon: L.MoveVertical, run: () => setSize(app, 'row') });

    // ---------- Sheets ----------
    a({ id: 'sheet.add', title: 'New sheet', category: 'Sheet', keys: ['Shift+F11'], icon: L.Plus, run: () => {
      const idx = app.doc.wb.sheets.indexOf(sheet()) + 1;
      const tx = app.transact('Add sheet', (tx) => app.applySheetDefaults(addSheet(tx, app.doc.wb, app.nextSheetName(), idx)));
      if (tx) app.setActiveSheet(app.doc.wb.activeSheetId);
    } });
    a({ id: 'sheet.next', title: 'Next sheet', category: 'Sheet', keys: ['Ctrl+PageDown'], icon: L.ArrowRight, global: true, run: () => cycleSheet(app, 1) });
    a({ id: 'sheet.prev', title: 'Previous sheet', category: 'Sheet', keys: ['Ctrl+PageUp'], icon: L.ArrowLeft, global: true, run: () => cycleSheet(app, -1) });
    a({ id: 'sheet.rename', title: 'Rename sheet…', category: 'Sheet', icon: L.Pencil, run: () => app.tabs.startRename(sheet().id) });
    a({ id: 'sheet.duplicate', title: 'Duplicate sheet', category: 'Sheet', icon: L.CopyPlus, run: () => {
      const tx = app.transact('Duplicate sheet', (tx) => duplicateSheet(tx, app.doc.wb, sheet()));
      if (tx) app.setActiveSheet(app.doc.wb.activeSheetId);
    } });
    a({ id: 'sheet.delete', title: 'Delete sheet', category: 'Sheet', icon: L.Trash2, enabled: () => app.doc.wb.sheets.length > 1, run: () => deleteCurrentSheet(app) });
    a({ id: 'sheet.hide', title: 'Hide sheet', category: 'Sheet', icon: L.EyeOff, run: () => {
      if (app.transact('Hide sheet', (tx) => setSheetHidden(tx, app.doc.wb, sheet(), true))) app.setActiveSheet(app.doc.wb.activeSheetId);
    } });
    a({ id: 'sheet.unhide', title: 'Unhide sheets…', category: 'Sheet', icon: L.Eye, enabled: () => app.doc.wb.sheets.some((s) => s.hidden), run: () => app.tabs.showUnhideMenu() });

    // ---------- Help ----------
    a({ id: 'help.shortcuts', title: 'Keyboard shortcuts', category: 'Help', keys: ['Ctrl+/'], icon: L.Keyboard, global: true, run: () => showShortcuts(app) });
    a({ id: 'help.functions', title: 'Browse functions', category: 'Help', icon: L.BookOpen, run: () => app.panel.open('functions') });
    a({ id: 'help.about', title: 'About Gridline', category: 'Help', icon: L.Info, run: () => showAbout() });
    a({ id: 'dev.tools', title: 'Toggle developer tools', category: 'Help', keys: ['Ctrl+Shift+I'], icon: L.Bug, global: true, hidden: true, run: () => platform.toggleDevTools() });
  }
}

// ---------- implementations ----------

export function clampRange(app: App, r: Range): Range {
  const b = app.doc.engine.usedBounds(app.sheet);
  return {
    r1: r.r1,
    c1: r.c1,
    r2: isFullColumn(r) ? Math.max(r.r1, b.maxRow) : r.r2,
    c2: isFullRow(r) ? Math.max(r.c1, b.maxCol) : r.c2,
  };
}

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];
function nextSize(cur: number, dir: 1 | -1): number {
  if (dir > 0) return SIZES.find((s) => s > cur) ?? cur + 4;
  return [...SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
}

function insertNow(app: App, what: 'date' | 'time'): void {
  const now = new Date();
  const text =
    what === 'date'
      ? `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`
      : `${((now.getHours() + 11) % 12) + 1}:${String(now.getMinutes()).padStart(2, '0')} ${now.getHours() >= 12 ? 'PM' : 'AM'}`;
  if (app.isEditing) {
    const ta = document.activeElement as HTMLTextAreaElement;
    if (ta && ta.tagName === 'TEXTAREA') {
      const s = ta.selectionStart;
      ta.setRangeText(text, s, ta.selectionEnd, 'end');
      ta.dispatchEvent(new Event('input'));
    }
    return;
  }
  const { row, col } = app.sel.active;
  app.enterInput(app.sheet, row, col, text, false);
}

function toggleFormatPainter(app: App): void {
  if (app.formatPainter) {
    app.formatPainter = null;
    app.ribbon.refresh();
    return;
  }
  app.formatPainter = { sheetId: app.sheet.id, range: { ...app.sel.range }, sticky: false };
  app.statusbar.flash('Select the cells to paint the formatting onto');
  app.ribbon.refresh();
}

async function editNote(app: App): Promise<void> {
  const { row, col } = app.sel.active;
  const cur = app.sheet.getCell(row, col)?.note ?? '';
  const body = document.createElement('div');
  const ta = document.createElement('textarea');
  ta.className = 'input';
  ta.rows = 5;
  ta.style.width = '100%';
  ta.value = cur;
  body.append(ta);
  setTimeout(() => ta.focus(), 0);
  const ok = await dialog(`Note for ${formatCell(row, col)}`, body, [
    { label: 'Save note', value: true, primary: true },
    { label: 'Cancel', value: false },
  ]);
  if (ok) app.transact('Edit note', (tx) => setNote(tx, app.sheet, row, col, ta.value));
}

async function editLink(app: App): Promise<void> {
  const { row, col } = app.sel.active;
  const cur = app.sheet.getCell(row, col)?.link ?? '';
  const url = await promptDialog('Insert link', 'Web address, email (mailto:) or cell (e.g. Sheet2!A1)', cur || 'https://');
  if (url === null) return;
  app.transact('Insert link', (tx) => {
    setLink(tx, app.sheet, row, col, url);
    const cell = app.sheet.getCell(row, col);
    if (!cell?.v && cell?.f === undefined) tx.patchCell(app.sheet, row, col, { v: url });
  });
}

function clearCf(app: App): void {
  const s = app.sheet;
  const ranges = app.sel.ranges;
  const remaining = s.cf.filter((rule) => !rule.ranges.some((rr) => ranges.some((r) => rr.r1 <= r.r2 && r.r1 <= rr.r2 && rr.c1 <= r.c2 && r.c1 <= rr.c2)));
  if (remaining.length === s.cf.length) {
    toast('There\'s no conditional formatting in the selection.');
    return;
  }
  app.transact('Clear conditional formatting', (tx) => tx.setProp(s, 'cf', remaining));
}

function selectedRows(app: App): number[] {
  const rows = new Set<number>();
  for (const r of app.sel.ranges) {
    const b = app.doc.engine.usedBounds(app.sheet);
    const end = isFullColumn(r) ? Math.max(r.r1, b.maxRow) : r.r2;
    for (let i = r.r1; i <= Math.min(end, r.r1 + 100000); i++) rows.add(i);
  }
  return [...rows].sort((x, y) => x - y);
}

function selectedCols(app: App): number[] {
  const cols = new Set<number>();
  for (const r of app.sel.ranges) {
    const b = app.doc.engine.usedBounds(app.sheet);
    const end = isFullRow(r) ? Math.max(r.c1, b.maxCol) : r.c2;
    for (let i = r.c1; i <= end; i++) cols.add(i);
  }
  return [...cols].sort((x, y) => x - y);
}

function insertRows(app: App, where: 'above' | 'below'): void {
  const r = app.sel.range;
  const count = Math.min(r.r2 - r.r1 + 1, 10000);
  const at = where === 'above' ? r.r1 : Math.min(r.r2, MAX_ROWS - 1) + 1;
  const tx = app.transact(`Insert ${count} row${count > 1 ? 's' : ''}`, (tx) => restructure(tx, app.doc.wb, app.sheet, 'row', at, count));
  if (tx) {
    app.sel.selectRows(at, at + count - 1);
    app.sel.active = { row: at, col: app.sel.active.col };
    app.grid.relayout();
    app.onSelectionChanged();
  }
}

function insertCols(app: App, where: 'left' | 'right'): void {
  const r = app.sel.range;
  const count = Math.min(r.c2 - r.c1 + 1, 1000);
  const at = where === 'left' ? r.c1 : Math.min(r.c2, MAX_COLS - 1) + 1;
  const tx = app.transact(`Insert ${count} column${count > 1 ? 's' : ''}`, (tx) => restructure(tx, app.doc.wb, app.sheet, 'col', at, count));
  if (tx) {
    app.sel.selectColumns(at, at + count - 1);
    app.grid.relayout();
    app.onSelectionChanged();
  }
}

function deleteRowsOrCols(app: App, axis: 'row' | 'col'): void {
  const sel = app.sel;
  const ranges = [...sel.ranges];
  const idx = new Set<number>();
  for (const r of ranges) {
    if (axis === 'row') for (let i = r.r1; i <= Math.min(r.r2, MAX_ROWS - 1); i++) idx.add(i);
    else for (let i = r.c1; i <= Math.min(r.c2, MAX_COLS - 1); i++) idx.add(i);
  }
  if (axis === 'row' && idx.size >= MAX_ROWS) {
    toast('You can\'t delete every row.', { kind: 'error' });
    return;
  }
  // Delete from the bottom up as contiguous runs
  const sorted = [...idx].sort((x, y) => y - x);
  const runs: [number, number][] = [];
  for (const i of sorted) {
    const last = runs[runs.length - 1];
    if (last && last[0] - 1 === i) last[0] = i;
    else runs.push([i, i]);
  }
  const n = idx.size;
  const tx = app.transact(`Delete ${n} ${axis === 'row' ? 'row' : 'column'}${n > 1 ? 's' : ''}`, (tx) => {
    for (const [start, end] of runs) restructure(tx, app.doc.wb, app.sheet, axis, start, -(end - start + 1));
  });
  if (tx) {
    const first = Math.min(...idx);
    if (axis === 'row') sel.setCell(app.sheet, Math.min(first, MAX_ROWS - 1), sel.active.col);
    else sel.setCell(app.sheet, sel.active.row, Math.min(first, MAX_COLS - 1));
    app.grid.relayout();
    app.onSelectionChanged();
  }
}

function autoSum(app: App, fn: string): void {
  const s = app.sheet;
  const sel = app.sel;
  const r = sel.range;
  const get = (row: number, col: number) => app.doc.engine.getValue(s, row, col);
  const isNum = (row: number, col: number) => typeof get(row, col) === 'number';
  // Multi-cell selection of numbers: put totals below each column
  if (r.r1 !== r.r2 || r.c1 !== r.c2) {
    const target = r.r2 + 1;
    app.transact(`Auto ${fn}`, (tx) => {
      for (let c = r.c1; c <= r.c2; c++) {
        tx.setCell(s, target, c, { ...(s.getCell(target, c) ?? {}), f: `${fn}(${formatCell(r.r1, c)}:${formatCell(r.r2, c)})`, v: undefined });
      }
    });
    return;
  }
  const { row, col } = sel.active;
  // Look up for numbers, else left
  let range: Range | null = null;
  if (row > 0 && isNum(row - 1, col)) {
    let top = row - 1;
    while (top > 0 && (isNum(top - 1, col) || (get(top - 1, col) === null && top - 2 >= 0 && isNum(top - 2, col)))) top--;
    range = { r1: top, c1: col, r2: row - 1, c2: col };
  } else if (col > 0 && isNum(row, col - 1)) {
    let left = col - 1;
    while (left > 0 && isNum(row, left - 1)) left--;
    range = { r1: row, c1: left, r2: row, c2: col - 1 };
  }
  const arg = range ? (range.r1 === range.r2 && range.c1 === range.c2 ? formatCell(range.r1, range.c1) : `${formatCell(range.r1, range.c1)}:${formatCell(range.r2, range.c2)}`) : '';
  app.editor.start({ mode: 'enter', initial: `=${fn}(${arg})`, source: 'cell' });
  if (!range) toast('Select the cells to add up, then press Enter.');
}

function insertChart(app: App, type: ChartType): void {
  const region = app.dataRegion();
  const vals = app.doc.engine;
  let hasNumber = false;
  for (let r = region.r1; r <= Math.min(region.r2, region.r1 + 200) && !hasNumber; r++)
    for (let c = region.c1; c <= region.c2; c++) if (typeof vals.getValue(app.sheet, r, c) === 'number') {
      hasNumber = true;
      break;
    }
  if (!hasNumber) {
    toast('Select some numbers to chart first (for example a table with labels and values).', { kind: 'error' });
    return;
  }
  const anchorCol = Math.min(MAX_COLS - 10, region.c2 + 2);
  // Place the chart next to the data, below any frozen rows so it scrolls with the sheet
  const anchorRow = Math.max(region.r1, app.sheet.freeze.rows);
  const chart = createChart(app.doc, app.sheet, region, type, anchorRow, anchorCol);
  chart.legend = app.settings.str('charts.legend') as typeof chart.legend;
  const size = { small: [400, 250], medium: [520, 320], large: [680, 420] }[app.settings.str('charts.size')] ?? [520, 320];
  chart.anchor = { ...chart.anchor, w: size[0], h: size[1] };
  const tx = app.transact('Insert chart', (tx) => tx.setProp(app.sheet, 'charts', [...app.sheet.charts, chart]));
  if (tx) {
    app.charts.select(chart.id);
    app.panel.open('chart');
  }
}

function quickSort(app: App, desc: boolean): void {
  const region = app.dataRegion();
  if (region.r1 === region.r2 && region.c1 === region.c2) return;
  const header = detectHeader(app.doc, app.sheet, region);
  const col = Math.max(region.c1, Math.min(region.c2, app.sel.active.col));
  app.transact(desc ? 'Sort Z → A' : 'Sort A → Z', (tx) => sortRange(tx, app.doc, app.sheet, region, [{ col, desc }], header));
  if (app.sel.isSingleCell) {
    app.sel.setRange(region, { ...app.sel.active });
    app.onSelectionChanged(false);
  }
  app.statusbar.flash(`Sorted ${formatCell(region.r1, region.c1)}:${formatCell(region.r2, region.c2)}${header ? ' (kept the header row in place)' : ''}`);
}

function toggleFilter(app: App): void {
  const s = app.sheet;
  if (s.filter) {
    app.transact('Remove filter', (tx) => applyFilter(tx, app.doc, s, null));
    return;
  }
  const region = app.dataRegion();
  if (region.r2 <= region.r1) {
    toast('Select a table with a header row to filter.', { kind: 'error' });
    return;
  }
  app.transact('Add filter', (tx) => applyFilter(tx, app.doc, s, { range: region, conds: {} }));
  app.statusbar.flash('Filter added — use the buttons in the header row');
}

async function dedupe(app: App): Promise<void> {
  const region = app.dataRegion();
  const header = detectHeader(app.doc, app.sheet, region);
  const cols: number[] = [];
  for (let c = region.c1; c <= region.c2; c++) cols.push(c);
  const ok = await dialog(
    'Remove duplicate rows',
    `Rows in ${formatCell(region.r1, region.c1)}:${formatCell(region.r2, region.c2)} that are identical in every column will be removed, keeping the first one.${header ? ' The first row is treated as a header.' : ''}`,
    [
      { label: 'Remove duplicates', value: true, primary: true },
      { label: 'Cancel', value: false },
    ],
  );
  if (!ok) return;
  let n = 0;
  app.transact('Remove duplicates', (tx) => {
    n = removeDuplicates(tx, app.doc, app.sheet, region, cols, header);
  });
  toast(n ? `Removed ${n} duplicate row${n === 1 ? '' : 's'}.` : 'No duplicate rows found.', { kind: n ? 'success' : 'info' });
}

function setFreeze(app: App, rows: number, cols: number): void {
  app.transact(rows || cols ? 'Freeze panes' : 'Unfreeze panes', (tx) => tx.setProp(app.sheet, 'freeze', { rows, cols }));
  app.grid.relayout();
}

function hideRowsCols(app: App, axis: 'row' | 'col', hide: boolean): void {
  const idx = axis === 'row' ? selectedRows(app) : selectedCols(app);
  const s = app.sheet;
  if (!hide) {
    // Unhide: include hidden neighbours of a single selected index
    const set = axis === 'row' ? s.hiddenRows : s.hiddenCols;
    const lo = Math.min(...idx);
    const hi = Math.max(...idx);
    const targets = [...set].filter((i) => i >= lo - 1 && i <= hi + 1 || (i >= lo && i <= hi));
    let i = lo - 1;
    while (set.has(i)) targets.push(i--);
    i = hi + 1;
    while (set.has(i)) targets.push(i++);
    if (!targets.length) {
      toast(`There are no hidden ${axis === 'row' ? 'rows' : 'columns'} in the selection. Select the ${axis === 'row' ? 'rows' : 'columns'} on both sides of the hidden ones.`);
      return;
    }
    app.transact(axis === 'row' ? 'Unhide rows' : 'Unhide columns', (tx) => setHidden(tx, s, axis, [...new Set(targets)], false));
  } else {
    app.transact(axis === 'row' ? 'Hide rows' : 'Hide columns', (tx) => setHidden(tx, s, axis, idx, true));
  }
  app.grid.relayout();
}

async function setSize(app: App, axis: 'row' | 'col'): Promise<void> {
  const s = app.sheet;
  const idx = axis === 'row' ? selectedRows(app) : selectedCols(app);
  const cur = axis === 'row' ? s.rowHeight(idx[0]) : s.colWidth(idx[0]);
  const v = await promptDialog(axis === 'row' ? 'Row height' : 'Column width', 'Size in pixels', String(cur), (x) => (/^\d+(\.\d+)?$/.test(x.trim()) && +x >= 2 && +x <= 2000 ? null : 'Enter a number between 2 and 2000.'));
  if (v === null) return;
  const n = parseFloat(v);
  app.transact(axis === 'row' ? 'Row height' : 'Column width', (tx) => {
    for (const i of idx) {
      if (axis === 'row') tx.setMap(s.rowHeights, i, Math.round(n), s);
      else tx.setMap(s.colWidths, i, Math.round(n), s);
    }
  });
  app.grid.relayout();
}

function cycleSheet(app: App, dir: number): void {
  const sheets = app.doc.wb.sheets.filter((s) => !s.hidden);
  const i = sheets.findIndex((s) => s.id === app.sheet.id);
  const next = sheets[(i + dir + sheets.length) % sheets.length];
  if (next) app.setActiveSheet(next.id);
}

export async function deleteCurrentSheet(app: App, id?: number): Promise<void> {
  const s = id !== undefined ? app.doc.wb.sheetById(id)! : app.sheet;
  if (app.doc.wb.sheets.length <= 1) {
    toast('A workbook needs at least one sheet.', { kind: 'error' });
    return;
  }
  const tx = app.transact(`Delete sheet "${s.name}"`, (tx) => deleteSheet(tx, app.doc.wb, s));
  if (tx) {
    app.setActiveSheet(app.doc.wb.activeSheetId);
    toast(`Deleted "${s.name}".`, { action: { label: 'Undo', run: () => app.doc.undo() } });
  }
}

export function renameSheetTo(app: App, id: number, name: string): boolean {
  const s = app.doc.wb.sheetById(id);
  if (!s) return false;
  const err = validateSheetName(name, app.doc.wb, id);
  if (err) {
    toast(err, { kind: 'error' });
    return false;
  }
  app.transact('Rename sheet', (tx) => renameSheet(tx, app.doc.wb, s, name.trim()));
  return true;
}

function tracePrecedents(app: App): void {
  const { row, col } = app.sel.active;
  const precs = app.doc.engine.precedentsOf(app.sheet, row, col).filter((p) => p.sheetId === app.sheet.id);
  if (!precs.length) {
    toast('This cell doesn\'t use any other cells.');
    return;
  }
  app.grid.refHighlights = precs.map((p, i) => ({ sheetId: p.sheetId, range: p.range, color: ['#4ea1ff', '#ff7a6b', '#b38cff', '#3fcf8e', '#ffb547'][i % 5] }));
  app.grid.draw();
  app.statusbar.flash('Highlighted the cells this formula uses — press Esc or click to clear');
  const clear = () => {
    app.grid.refHighlights = [];
    app.grid.draw();
    document.removeEventListener('pointerdown', clear, true);
    document.removeEventListener('keydown', clear, true);
  };
  setTimeout(() => {
    document.addEventListener('pointerdown', clear, true);
    document.addEventListener('keydown', clear, true);
  }, 0);
}

