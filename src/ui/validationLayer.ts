// Data validation on the grid: the dropdown arrow of list cells, the input message shown when a cell
// with a rule is selected, red circles around invalid entries, and the alert after typing a bad value.
import * as L from 'lucide';
import type { Sheet } from '../core/workbook';
import { cellIsValid, describeRule, invalidCells, listItems, validationAt } from '../core/validation';
import type { App } from './app';
import { h, icon } from './dom';
import { dialog, showMenu, toast, type MenuItem } from './widgets';

export class ValidationLayer {
  readonly el: HTMLDivElement;
  private arrow: HTMLButtonElement;
  private promptBox: HTMLDivElement;
  private circlesEl: HTMLDivElement;
  /** Sheet id → invalid cells being circled ("Circle invalid data"). */
  private circled = new Map<number, { row: number; col: number }[]>();

  constructor(private app: App) {
    this.arrow = h('button', { class: 'dv-arrow', 'aria-label': 'Choose from the list', tabindex: '-1', hidden: true }, icon(L.ChevronDown, 14)) as HTMLButtonElement;
    this.arrow.addEventListener('mousedown', (e) => e.preventDefault());
    this.arrow.addEventListener('click', () => this.openList());
    this.promptBox = h('div', { class: 'dv-prompt', hidden: true });
    this.circlesEl = h('div', { class: 'dv-circles' });
    this.el = h('div', { style: { position: 'absolute', left: '0', top: '0', width: '0', height: '0', overflow: 'visible' } }, this.circlesEl, this.arrow, this.promptBox);
  }

  /** Repositions everything (after scrolling, zooming or moving the selection). */
  position(): void {
    const app = this.app;
    const sheet = app.sheet;
    const a = app.sel.active;
    const dv = app.isEditing ? undefined : validationAt(sheet, a.row, a.col);
    const g = app.grid;
    const rect = g.cellRect(a.row, a.col);
    const visible = rect.x + rect.w > g.headerW && rect.y + rect.h > g.headerH && rect.x < g.width && rect.y < g.height;
    // Dropdown arrow for list rules
    const showArrow = !!dv && dv.type === 'list' && dv.dropdown && visible;
    this.arrow.hidden = !showArrow;
    if (showArrow) {
      const size = Math.max(16, Math.min(22, rect.h - 2));
      Object.assign(this.arrow.style, { left: `${rect.x + rect.w + 1}px`, top: `${rect.y + rect.h - size}px`, width: `${size}px`, height: `${size}px` });
    }
    // Input message
    const showPrompt = !!dv && dv.showPrompt && !!(dv.prompt || dv.promptTitle) && visible;
    this.promptBox.hidden = !showPrompt;
    if (showPrompt && dv) {
      this.promptBox.innerHTML = '';
      if (dv.promptTitle) this.promptBox.append(h('strong', null, dv.promptTitle));
      if (dv.prompt) this.promptBox.append(h('div', null, dv.prompt));
      Object.assign(this.promptBox.style, { left: `${rect.x + Math.min(rect.w, 24)}px`, top: `${rect.y + rect.h + 6}px` });
    }
    this.positionCircles();
  }

  private positionCircles(): void {
    const sheet = this.app.sheet;
    const list = this.circled.get(sheet.id) ?? [];
    this.circlesEl.innerHTML = '';
    if (!list.length) return;
    const g = this.app.grid;
    let shown = 0;
    for (const c of list) {
      const r = g.cellRect(c.row, c.col);
      if (r.x + r.w < g.headerW || r.y + r.h < g.headerH || r.x > g.width || r.y > g.height) continue;
      this.circlesEl.append(h('div', { class: 'dv-circle', style: { left: `${r.x - 4}px`, top: `${r.y - 3}px`, width: `${r.w + 8}px`, height: `${r.h + 6}px` } }));
      if (++shown > 600) break;
    }
  }

  /** Opens the list of choices for the active cell (also Alt+Down). Returns false if it has none. */
  openList(): boolean {
    const app = this.app;
    const sheet = app.sheet;
    const a = app.sel.active;
    const dv = validationAt(sheet, a.row, a.col);
    if (!dv || dv.type !== 'list') return false;
    const { items } = listItems(app.doc, sheet, dv, a.row, a.col);
    const current = app.doc.engine.getValue(sheet, a.row, a.col);
    const cur = current === null ? '' : String(current);
    const menu: MenuItem[] = items.length
      ? items.map((it) => ({ label: it, checked: it === cur, run: () => app.enterInput(sheet, a.row, a.col, it, false) }))
      : [{ label: 'The list is empty', disabled: true }];
    const r = this.arrow.hidden ? app.grid.cellRect(a.row, a.col) : null;
    const box = app.grid.el.getBoundingClientRect();
    showMenu(menu, this.arrow.hidden && r ? { left: box.left + r.x, top: box.top + r.y, right: box.left + r.x + r.w, bottom: box.top + r.y + r.h } : this.arrow);
    return true;
  }

  // ---------- circles ----------

  circleInvalid(): number {
    const app = this.app;
    const list = invalidCells(app.doc, app.sheet);
    this.circled.set(app.sheet.id, list);
    this.positionCircles();
    return list.length;
  }

  clearCircles(): void {
    this.circled.delete(this.app.sheet.id);
    this.positionCircles();
  }

  get circling(): boolean {
    return (this.circled.get(this.app.sheet.id)?.length ?? 0) > 0;
  }

  /** After edits: drop circles from cells that are valid now. */
  refresh(): void {
    const sheet = this.app.sheet;
    const list = this.circled.get(sheet.id);
    if (list) this.circled.set(sheet.id, list.filter((c) => !cellIsValid(this.app.doc, sheet, c.row, c.col)));
    this.position();
  }

  // ---------- typing a value that breaks a rule ----------

  /**
   * Called right after someone typed into a cell. When the value breaks the cell's rule and the rule
   * shows an alert, asks the same way Excel does. Returns 'keep', 'undo' or 'retry'.
   */
  async checkEntry(sheet: Sheet, row: number, col: number): Promise<'keep' | 'undo' | 'retry'> {
    const app = this.app;
    const dv = validationAt(sheet, row, col);
    if (!dv || !dv.showError || cellIsValid(app.doc, sheet, row, col)) return 'keep';
    const message = dv.error || `This value doesn't match the rule for this cell: it must be ${describeRule(dv)}.`;
    const title = dv.errorTitle || (dv.errorStyle === 'stop' ? 'Not allowed' : dv.errorStyle === 'warning' ? 'Check this value' : 'Note');
    if (dv.errorStyle === 'stop') {
      return dialog<'retry' | 'undo'>(title, message, [
        { label: 'Retry', value: 'retry', primary: true },
        { label: 'Cancel', value: 'undo' },
      ], { cancelValue: 'undo' });
    }
    if (dv.errorStyle === 'warning') {
      return dialog<'keep' | 'retry' | 'undo'>(title, `${message}\n\nKeep it anyway?`, [
        { label: 'Yes', value: 'keep', primary: true },
        { label: 'No', value: 'retry' },
        { label: 'Cancel', value: 'undo' },
      ], { cancelValue: 'undo' });
    }
    const ok = await dialog<boolean>(title, message, [
      { label: 'OK', value: true, primary: true },
      { label: 'Cancel', value: false },
    ], { cancelValue: false });
    if (!ok) toast('The value was not entered.');
    return ok ? 'keep' : 'undo';
  }
}
