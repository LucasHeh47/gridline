// Multi-level sort.
import * as L from 'lucide';
import { colToName, formatRange, parseRange, type Range } from '../../core/address';
import { detectHeader, sortRange, type SortKey } from '../../core/ops/data';
import { formatValue } from '../../core/numfmt';
import type { App } from '../app';
import { h, icon } from '../dom';
import { toast } from '../widgets';
import type { PanelContent } from './SidePanel';

export class SortPanel implements PanelContent {
  title = 'Sort';
  readonly el: HTMLElement;
  private range: Range;
  private header: boolean;
  private keys: SortKey[];

  constructor(private app: App) {
    this.range = app.dataRegion();
    this.header = detectHeader(app.doc, app.sheet, this.range);
    const col = Math.max(this.range.c1, Math.min(this.range.c2, app.sel.active.col));
    this.keys = [{ col, desc: false }];
    this.el = h('div', { class: 'sp-body' });
    this.render();
  }

  private colLabel(c: number): string {
    if (this.header) {
      const v = this.app.valueAt(this.range.r1, c);
      if (v !== null && v !== '') return `${formatValue(v, undefined).text} (${colToName(c)})`;
    }
    return `Column ${colToName(c)}`;
  }

  private render(): void {
    const b = this.el;
    b.innerHTML = '';
    const rangeInput = h('input', { class: 'input mono', value: formatRange(this.range) }) as HTMLInputElement;
    rangeInput.addEventListener('change', () => {
      const r = parseRange(rangeInput.value);
      if (!r) {
        toast('Enter a range like A1:F200.', { kind: 'error' });
        return;
      }
      this.range = r;
      this.keys = this.keys.map((k) => ({ ...k, col: Math.max(r.c1, Math.min(r.c2, k.col)) }));
      this.render();
    });
    const headerBox = h('input', { type: 'checkbox', checked: this.header }) as HTMLInputElement;
    headerBox.addEventListener('change', () => {
      this.header = headerBox.checked;
      this.render();
    });
    b.append(h('div', { class: 'field' }, h('label', null, 'Range to sort'), rangeInput), h('label', { class: 'check' }, headerBox, 'First row is a header (keep it on top)'));
    b.append(h('div', { class: 'field-label', style: { margin: '14px 0 6px' } }, 'Sort by'));
    this.keys.forEach((k, i) => {
      const colSel = h('select', { class: 'input', style: { flex: '1' } }) as HTMLSelectElement;
      for (let c = this.range.c1; c <= this.range.c2; c++) colSel.append(h('option', { value: String(c) }, this.colLabel(c)));
      colSel.value = String(k.col);
      colSel.addEventListener('change', () => (k.col = +colSel.value));
      const order = h('select', { class: 'input' }) as HTMLSelectElement;
      order.append(h('option', { value: 'asc' }, 'A → Z, 1 → 9'), h('option', { value: 'desc' }, 'Z → A, 9 → 1'));
      order.value = k.desc ? 'desc' : 'asc';
      order.addEventListener('change', () => (k.desc = order.value === 'desc'));
      const remove = h('button', { class: 'btn', disabled: this.keys.length === 1, onclick: () => {
        this.keys.splice(i, 1);
        this.render();
      } }, icon(L.X, 14));
      b.append(h('div', { class: 'card' }, h('div', { class: 'hint', style: { marginBottom: '4px' } }, i === 0 ? 'Sort by' : 'Then by'), h('div', { class: 'row' }, colSel, order, remove)));
    });
    b.append(
      h('button', { class: 'btn outline', onclick: () => {
        this.keys.push({ col: this.range.c1, desc: false });
        this.render();
      } }, icon(L.Plus, 14), 'Add level'),
      h('div', { class: 'row', style: { justifyContent: 'flex-end', marginTop: '16px' } }, h('button', { class: 'btn primary', onclick: () => this.apply() }, 'Sort')),
    );
  }

  private apply(): void {
    const app = this.app;
    app.transact('Sort', (tx) => sortRange(tx, app.doc, app.sheet, this.range, this.keys, this.header));
    app.sel.setRange(this.range, { row: this.range.r1, col: this.range.c1 });
    app.onSelectionChanged(false);
    app.statusbar.flash('Sorted');
  }
}
