// Find & replace as a small floating bar (not a blocking dialog).
import * as L from 'lucide';
import { findAll, replaceIn, type FindHit, type FindOptions } from '../core/ops/data';
import type { App } from './app';
import { h, icon } from './dom';
import { toast, tooltip } from './widgets';

export class FindBar {
  readonly el: HTMLDivElement;
  private query: HTMLInputElement;
  private replace: HTMLInputElement;
  private replaceRow: HTMLDivElement;
  private count: HTMLSpanElement;
  private hits: FindHit[] = [];
  private index = -1;
  private opts: FindOptions = { matchCase: false, wholeCell: false, inFormulas: false, allSheets: false };
  private chips: Record<keyof FindOptions, HTMLButtonElement>;
  private dirty = true;

  constructor(private app: App) {
    this.query = h('input', { class: 'input', placeholder: 'Find', spellcheck: false }) as HTMLInputElement;
    this.replace = h('input', { class: 'input', placeholder: 'Replace with', spellcheck: false }) as HTMLInputElement;
    this.count = h('span', { class: 'count' });
    const prev = h('button', { class: 'btn', onclick: () => this.step(-1) }, icon(L.ChevronUp, 16));
    const next = h('button', { class: 'btn', onclick: () => this.step(1) }, icon(L.ChevronDown, 16));
    const close = h('button', { class: 'btn', onclick: () => this.close() }, icon(L.X, 16));
    tooltip(prev, 'Previous match', 'Shift+Enter');
    tooltip(next, 'Next match', 'Enter');
    const chip = (key: keyof FindOptions, label: string) => {
      const b = h('button', { class: 'btn', onclick: () => {
        this.opts[key] = !this.opts[key];
        b.classList.toggle('on', this.opts[key]);
        this.dirty = true;
        this.search();
      } }, label) as HTMLButtonElement;
      return b;
    };
    this.chips = {
      matchCase: chip('matchCase', 'Match case'),
      wholeCell: chip('wholeCell', 'Whole cell'),
      inFormulas: chip('inFormulas', 'Search formulas'),
      allSheets: chip('allSheets', 'All sheets'),
    };
    this.replaceRow = h(
      'div',
      { class: 'row' },
      this.replace,
      h('button', { class: 'btn outline', onclick: () => this.replaceOne() }, 'Replace'),
      h('button', { class: 'btn outline', onclick: () => this.replaceAll() }, 'All'),
    );
    this.el = h(
      'div',
      { class: 'findbar', hidden: true },
      h('div', { class: 'row' }, this.query, this.count, prev, next, close),
      this.replaceRow,
      h('div', { class: 'opts' }, ...Object.values(this.chips)),
    );
    this.query.addEventListener('input', () => {
      this.dirty = true;
      this.search();
    });
    const keys = (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (e.target === this.replace) this.replaceOne();
        else this.step(e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      }
    };
    this.query.addEventListener('keydown', keys);
    this.replace.addEventListener('keydown', keys);
  }

  open(withReplace: boolean): void {
    this.el.hidden = false;
    this.replaceRow.style.display = withReplace ? '' : 'none';
    const a = this.app.sel.active;
    if (!this.query.value) {
      const v = this.app.valueAt(a.row, a.col);
      if (typeof v === 'string' && v.length < 60) this.query.value = v;
    }
    this.query.focus();
    this.query.select();
    this.dirty = true;
    this.search(false);
  }

  close(): void {
    if (this.el.hidden) return;
    this.el.hidden = true;
    this.app.focusGrid();
  }

  /** Called when the document changes. */
  refresh(): void {
    this.dirty = true;
    if (!this.el.hidden) this.search(false);
  }

  private search(jump = true): void {
    if (!this.dirty) return;
    this.dirty = false;
    this.hits = findAll(this.app.doc, this.app.sheet, this.query.value, this.opts);
    const a = this.app.sel.active;
    this.index = this.hits.findIndex((h2) => h2.sheetId === this.app.sheet.id && (h2.row > a.row || (h2.row === a.row && h2.col >= a.col)));
    if (this.index < 0 && this.hits.length) this.index = 0;
    this.updateCount();
    if (jump && this.index >= 0) this.show();
  }

  private updateCount(): void {
    if (!this.query.value) this.count.textContent = '';
    else if (!this.hits.length) this.count.textContent = 'No results';
    else this.count.textContent = `${this.index + 1} of ${this.hits.length}`;
  }

  private step(dir: number): void {
    this.search(false);
    if (!this.hits.length) return;
    this.index = (this.index + dir + this.hits.length) % this.hits.length;
    this.updateCount();
    this.show();
  }

  private show(): void {
    const hit = this.hits[this.index];
    if (!hit) return;
    const app = this.app;
    if (hit.sheetId !== app.sheet.id) app.setActiveSheet(hit.sheetId, false);
    app.sel.setCell(app.sheet, hit.row, hit.col);
    app.onSelectionChanged();
  }

  private replaceOne(): void {
    this.search(false);
    const hit = this.hits[this.index];
    if (!hit) return;
    let n = 0;
    this.app.transact('Replace', (tx) => {
      n = replaceIn(tx, this.app.doc, [hit], this.query.value, this.replace.value, this.opts);
    });
    this.dirty = true;
    this.search(false);
    if (n && this.hits.length) {
      this.index = Math.min(this.index, this.hits.length - 1);
      this.updateCount();
      this.show();
    }
  }

  private replaceAll(): void {
    this.search(false);
    if (!this.hits.length) return;
    let n = 0;
    this.app.transact('Replace all', (tx) => {
      n = replaceIn(tx, this.app.doc, this.hits, this.query.value, this.replace.value, this.opts);
    });
    toast(`Replaced ${n} cell${n === 1 ? '' : 's'}.`, { kind: 'success', action: n ? { label: 'Undo', run: () => this.app.doc.undo() } : undefined });
    this.dirty = true;
    this.search(false);
  }
}
