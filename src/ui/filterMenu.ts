// The dropdown shown from a filter button: sort, filter by condition, filter by values. Tables use
// it too, with their own sorting choices on top (see tableMenu.ts).
import * as L from 'lucide';
import type { Range } from '../core/address';
import { applyFilter, columnValues, sortRange } from '../core/ops/data';
import type { FilterCond, FilterOp } from '../core/workbook';
import type { App } from './app';
import { h, icon } from './dom';
import { Popover } from './widgets';

const OPS: { value: FilterOp | ''; label: string; inputs: number }[] = [
  { value: '', label: 'No condition', inputs: 0 },
  { value: 'contains', label: 'Text contains', inputs: 1 },
  { value: 'notContains', label: 'Text does not contain', inputs: 1 },
  { value: 'beginsWith', label: 'Text begins with', inputs: 1 },
  { value: 'endsWith', label: 'Text ends with', inputs: 1 },
  { value: 'equals', label: 'Is equal to', inputs: 1 },
  { value: 'notEquals', label: 'Is not equal to', inputs: 1 },
  { value: 'gt', label: 'Greater than', inputs: 1 },
  { value: 'gte', label: 'Greater than or equal to', inputs: 1 },
  { value: 'lt', label: 'Less than', inputs: 1 },
  { value: 'lte', label: 'Less than or equal to', inputs: 1 },
  { value: 'between', label: 'Between', inputs: 2 },
  { value: 'top10', label: 'Top N values', inputs: 1 },
  { value: 'aboveAvg', label: 'Above average', inputs: 0 },
  { value: 'belowAvg', label: 'Below average', inputs: 0 },
];

/** What a filter menu filters: the sheet's filter, or a table's. */
export interface FilterTarget {
  /** Header row first. */
  range: Range;
  conds: Record<number, FilterCond>;
  apply(conds: Record<number, FilterCond>, label: string): void;
  /** Sorting choices and anything else above the filter. */
  top(close: () => void): HTMLElement[];
  /** Extra buttons at the bottom, left of Clear. */
  footer?(close: () => void): HTMLElement[];
}

export class FilterMenu {
  constructor(
    private app: App,
    private col: number,
    private target?: FilterTarget,
  ) {}

  /** The sheet's filter (Data → Filter). */
  private sheetTarget(): FilterTarget | null {
    const app = this.app;
    const s = app.sheet;
    const f = s.filter;
    if (!f) return null;
    const col = this.col;
    const sortBtn = (desc: boolean, close: () => void) =>
      h('button', {
        class: 'btn text',
        style: { justifyContent: 'flex-start' },
        onclick: () => {
          close();
          app.transact(desc ? 'Sort Z → A' : 'Sort A → Z', (tx) => sortRange(tx, app.doc, s, f.range, [{ col, desc }], true));
          app.transact('Reapply filter', (tx) => applyFilter(tx, app.doc, s, s.filter));
        },
      }, icon(desc ? L.ArrowUpZA : L.ArrowDownAZ, 16), desc ? 'Sort Z → A' : 'Sort A → Z');
    return {
      range: f.range,
      conds: f.conds,
      apply: (conds, label) => app.transact(label, (tx) => applyFilter(tx, app.doc, s, { range: f.range, conds })),
      top: (close) => [sortBtn(false, close), sortBtn(true, close)],
    };
  }

  open(rect: DOMRect): void {
    const app = this.app;
    const s = app.sheet;
    const f = this.target ?? this.sheetTarget();
    if (!f) return;
    const col = this.col;
    const existing: FilterCond = f.conds[col] ?? {};
    const values = columnValues(app.doc, s, f.range, col);
    const selected = new Set(existing.values ?? values.map((v) => v.text));

    const opSelect = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const o of OPS) opSelect.append(h('option', { value: o.value }, o.label));
    opSelect.value = existing.op?.type ?? '';
    const in1 = h('input', { class: 'input', placeholder: 'Value', value: existing.op?.value ?? '' }) as HTMLInputElement;
    const in2 = h('input', { class: 'input', placeholder: 'and', value: existing.op?.value2 ?? '' }) as HTMLInputElement;
    const syncInputs = () => {
      const n = OPS.find((o) => o.value === opSelect.value)?.inputs ?? 0;
      in1.style.display = n >= 1 ? '' : 'none';
      in2.style.display = n >= 2 ? '' : 'none';
      if (opSelect.value === 'top10' && !in1.value) in1.value = '10';
    };
    opSelect.addEventListener('change', syncInputs);
    syncInputs();

    const search = h('input', { class: 'input', placeholder: 'Search values', spellcheck: false }) as HTMLInputElement;
    const listEl = h('div', { class: 'filter-values' });
    const allBox = h('input', { type: 'checkbox' }) as HTMLInputElement;
    const renderList = () => {
      listEl.innerHTML = '';
      const q = search.value.toLowerCase();
      const shown = values.filter((v) => !q || v.text.toLowerCase().includes(q));
      allBox.checked = shown.every((v) => selected.has(v.text));
      for (const v of shown.slice(0, 2000)) {
        const cb = h('input', { type: 'checkbox', checked: selected.has(v.text) }) as HTMLInputElement;
        cb.addEventListener('change', () => {
          if (cb.checked) selected.add(v.text);
          else selected.delete(v.text);
          allBox.checked = shown.every((x) => selected.has(x.text));
        });
        listEl.append(h('label', { class: 'check' }, cb, h('span', null, v.text === '' ? '(Blanks)' : v.text), h('span', { class: 'count' }, String(v.count))));
      }
      if (shown.length > 2000) listEl.append(h('div', { class: 'hint' }, `Showing 2,000 of ${shown.length.toLocaleString()} values. Search to narrow down.`));
    };
    allBox.addEventListener('change', () => {
      const q = search.value.toLowerCase();
      for (const v of values) if (!q || v.text.toLowerCase().includes(q)) allBox.checked ? selected.add(v.text) : selected.delete(v.text);
      renderList();
    });
    search.addEventListener('input', renderList);
    renderList();

    const apply = () => {
      const conds = { ...f.conds };
      const cond: FilterCond = {};
      if (selected.size < values.length) cond.values = [...selected];
      if (opSelect.value) cond.op = { type: opSelect.value as FilterOp, value: in1.value, value2: in2.value };
      if (cond.values || cond.op) conds[col] = cond;
      else delete conds[col];
      pop.close();
      f.apply(conds, 'Filter');
      app.grid.relayout();
      app.focusGrid();
    };
    const clearBtn = h('button', {
      class: 'btn outline',
      onclick: () => {
        const conds = { ...f.conds };
        delete conds[col];
        pop.close();
        f.apply(conds, 'Clear filter');
        app.grid.relayout();
      },
    }, 'Clear');
    const close = () => pop.close();

    const content = h(
      'div',
      { class: 'filter-menu' },
      ...f.top(close),
      h('div', { class: 'menu-sep' }),
      h('div', { class: 'field', style: { marginBottom: '0' } }, h('label', null, 'Filter by condition'), opSelect, h('div', { class: 'row' }, in1, in2)),
      h('div', { class: 'field', style: { marginBottom: '0' } }, h('label', null, 'Filter by values'), search, h('label', { class: 'check' }, allBox, 'Select all'), listEl),
      h('div', { class: 'row', style: { justifyContent: 'flex-end' } }, ...(f.footer?.(close) ?? []), h('span', { style: { flex: '1' } }), clearBtn, h('button', { class: 'btn primary', onclick: apply }, 'Apply')),
    );
    content.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !(e.target instanceof HTMLSelectElement)) {
        e.preventDefault();
        apply();
      }
    });
    const pop = new Popover(content);
    pop.open({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom });
    search.focus();
  }
}
