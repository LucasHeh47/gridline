// Name box + formula input.
import * as L from 'lucide';
import { formatRange, isFullColumn, isFullRow, parseRange } from '../core/address';
import type { App } from './app';
import { h, icon } from './dom';
import { HlEditor } from './editor';
import { toast, tooltip } from './widgets';

export class FormulaBar {
  readonly el: HTMLElement;
  private nameBox: HTMLInputElement;
  private bar: HlEditor;
  private wrap: HTMLDivElement;

  constructor(private app: App) {
    this.nameBox = h('input', { class: 'name-box', spellcheck: false, 'aria-label': 'Cell address' }) as HTMLInputElement;
    tooltip(this.nameBox, 'Type a cell (B12), range (A1:C9), sheet (Sheet2!A1) or name to jump there. Type a new name to name the selection.', 'Ctrl+G');
    this.nameBox.addEventListener('focus', () => setTimeout(() => this.nameBox.select(), 0));
    this.nameBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.submitName();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.refresh();
        app.focusGrid();
      }
    });
    this.nameBox.addEventListener('blur', () => this.refresh());

    const fx = h('button', { class: 'btn fx-btn', onclick: () => app.panel.open('functions') }, 'fx');
    tooltip(fx, 'Browse functions', 'Shift+F3');

    this.bar = new HlEditor();
    this.bar.ta.setAttribute('aria-label', 'Formula bar');
    const expand = h('button', { class: 'btn fx-expand', onclick: () => this.toggleExpand() }, icon(L.ChevronsUpDown, 14));
    tooltip(expand, 'Expand the formula bar');
    this.wrap = h('div', { class: 'fx-input-wrap' }, this.bar.root, expand);
    this.bar.root.style.height = '100%';
    if (app.prefs.showFormulaBarExpanded) this.wrap.classList.add('expanded');
    app.editor.attachBar(this.bar);
    this.el = h('div', { class: 'formula-bar' }, this.nameBox, fx, this.wrap);
  }

  private toggleExpand(): void {
    this.wrap.classList.toggle('expanded');
    this.app.prefs.showFormulaBarExpanded = this.wrap.classList.contains('expanded');
    this.app.savePrefs();
  }

  focusNameBox(): void {
    this.nameBox.focus();
    this.nameBox.select();
  }

  refresh(): void {
    const app = this.app;
    if (document.activeElement !== this.nameBox) {
      const sel = app.sel;
      const r = sel.range;
      let label: string;
      if (sel.ranges.length > 1) label = formatRange(r);
      else if (isFullColumn(r) && isFullRow(r)) label = 'All cells';
      else if (r.r1 === r.r2 && r.c1 === r.c2) label = formatRange(r);
      else {
        const m = app.sheet.mergeAt(r.r1, r.c1);
        label = m && m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2 ? formatRange({ r1: r.r1, c1: r.c1, r2: r.r1, c2: r.c1 }) : formatRange(r);
      }
      // Show a defined name if the selection matches one exactly
      for (const dn of app.doc.wb.names.values()) {
        const ref = dn.ref.replace(/\$/g, '');
        const bang = ref.lastIndexOf('!');
        const sheetPart = bang > 0 ? ref.slice(0, bang).replace(/^'|'$/g, '') : app.sheet.name;
        if (sheetPart.toLowerCase() === app.sheet.name.toLowerCase() && parseRange(ref.slice(bang + 1)) && formatRange(parseRange(ref.slice(bang + 1))!) === formatRange(r)) {
          label = dn.name;
          break;
        }
      }
      this.nameBox.value = label;
    }
    if (!app.isEditing) this.bar.value = app.editor.describeActive();
  }

  private submitName(): void {
    const app = this.app;
    const text = this.nameBox.value.trim();
    if (!text) return;
    if (app.goTo(text)) return;
    // A new name for the selection
    if (/^[A-Za-z_\\][\w.]*$/.test(text) && !parseRange(text)) {
      const r = app.sel.range;
      const ref = `'${app.sheet.name.replace(/'/g, "''")}'!${formatRange(r).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`;
      app.transact('Define name', (tx) => {
        const names = new Map(app.doc.wb.names);
        names.set(text.toUpperCase(), { name: text, ref });
        tx.setProp(app.doc.wb, 'names', names, true);
      });
      toast(`Named the selection "${text}". You can now use it in formulas, like =SUM(${text}).`, { kind: 'success' });
      app.focusGrid();
      return;
    }
    toast(`"${text}" isn't a cell address or name.`, { kind: 'error' });
  }
}
