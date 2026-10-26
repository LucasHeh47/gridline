// Tables in the app: Ctrl+T, the sort arrows in a table's header row, the ▾ menu (sorting and
// filtering) and the table options card (name, color, banded rows, back to normal cells).
import * as L from 'lucide';
import { formatRange } from '../core/address';
import { currentRegion } from '../core/ops/data';
import {
  canRestoreOrder,
  columnKind,
  createTable,
  MergedCellsError,
  nextSort,
  removeTable,
  renameTable,
  setTableFilter,
  setTableLook,
  sortLabel,
  sortTable,
  tableAt,
  tableById,
  TABLE_COLOR_IDS,
  TABLE_COLORS,
  tableNameProblem,
  type MadeTable,
} from '../core/tables';
import type { TableDef, TableSortKey } from '../core/workbook';
import type { App } from './app';
import { h, icon } from './dom';
import { FilterMenu } from './filterMenu';
import type { TableButtonHit } from './grid/GridView';
import { Popover, toast, tooltip } from './widgets';

/** The table the cursor is in, if any. */
export function activeTable(app: App): TableDef | undefined {
  const a = app.sel.active;
  return tableAt(app.sheet, a.row, a.col);
}

/** Ctrl+T: makes the selected cells (or the block of data around the cursor) a table. */
export function makeTable(app: App, opts: { header?: boolean; unmerge?: boolean } = {}): void {
  const s = app.sheet;
  const inside = activeTable(app);
  if (inside && app.sel.isSingleCell) {
    toast(`This is already a table (${inside.name}). Click the arrows in its header row to sort it.`);
    return;
  }
  const a = app.sel.active;
  const range = app.sel.isSingleCell ? currentRegion(app.doc, s, a.row, a.col) : app.dataRegion();
  if (range.r1 === range.r2 && range.c1 === range.c2 && app.doc.engine.getValue(s, range.r1, range.c1) === null) {
    toast('Click inside your data (or select it) first, then press Ctrl+T.');
    return;
  }
  let made: MadeTable | null = null;
  try {
    // Nothing in the cells changes: values, formats, column widths and rows stay as they are
    app.doc.transact('Make a table', (tx) => {
      made = createTable(tx, app.doc, s, range, opts);
    });
  } catch (err) {
    if (err instanceof MergedCellsError) toast(`${err.message} Unmerge them to make a table.`, { kind: 'error', ms: 10000, action: { label: 'Unmerge and make the table', run: () => makeTable(app, { ...opts, unmerge: true }) } });
    else toast((err as Error).message || String(err), { kind: 'error' });
    return;
  }
  if (!made) return;
  const { table, addedHeader, firstRowLooksLikeData } = made as MadeTable;
  app.sel.setRange(table.range, { row: table.range.r1, col: table.range.c1 });
  app.onSelectionChanged();
  const rows = table.range.r2 - table.range.r1;
  const first = `row ${table.range.r1 + 1}`;
  const message = addedHeader
    ? `Made ${table.name} with a new header row and ${rows.toLocaleString()} row${rows === 1 ? '' : 's'}.`
    : firstRowLooksLikeData
      ? `Made ${table.name}. Its column names come from ${first}, which looks like data: sorting keeps that row at the top.`
      : `Made ${table.name} with ${rows.toLocaleString()} row${rows === 1 ? '' : 's'}. Click a column’s arrow to sort it.`;
  toast(message, {
    kind: 'success',
    ms: firstRowLooksLikeData ? 12000 : 8000,
    // Data without a header row: add one above it (the only time making a table changes the sheet)
    action: firstRowLooksLikeData
      ? {
          label: 'Add a header row',
          run: () => {
            app.doc.undo();
            app.sel.setRange(range, { row: range.r1, col: range.c1 });
            makeTable(app, { ...opts, header: false });
          },
        }
      : undefined,
  });
}

/** Sorts a table and says what happened in the status bar. */
export function sortTableBy(app: App, table: TableDef, keys: TableSortKey[] | 'original'): void {
  const s = app.sheet;
  const label = keys === 'original' ? 'Original order' : 'Sort table';
  app.transact(label, (tx) => sortTable(tx, app.doc, s, table.id, keys));
  const t = tableById(s, table.id);
  if (!t) return;
  if (keys === 'original') {
    app.statusbar.flash(`${t.name} is back in its original order`);
    return;
  }
  const parts = t.sort.map((k, i) => {
    const name = t.columns[k.col - t.range.c1];
    return `${i ? 'then ' : ''}${name} (${sortLabel(columnKind(app.doc, s, t, k.col), k.desc).toLowerCase()})`;
  });
  if (parts.length) app.statusbar.flash(`Sorted ${t.name} by ${parts.join(', ')}`, canRestoreOrder(t) ? { label: 'Original order', run: () => sortTableBy(app, tableById(s, t.id)!, 'original') } : undefined);
}

/** A click on a header button: the arrow sorts (Shift adds a column), the ▾ opens the menu. */
export function tableButtonClick(app: App, hit: TableButtonHit, shift: boolean, rect: DOMRect): void {
  const t = tableById(app.sheet, hit.table.id);
  if (!t) return;
  if (hit.part === 'menu') {
    openTableMenu(app, t, hit.col, rect);
    return;
  }
  const keys = nextSort(t, hit.col, { additive: shift, stale: app.grid.isTableStale(t) });
  sortTableBy(app, t, keys);
}

/** What hovering a sort arrow says. */
export function tableButtonTip(app: App, hit: TableButtonHit): string {
  const t = hit.table;
  const name = t.columns[hit.col - t.range.c1];
  if (hit.part === 'menu') return `${name}: sort and filter options (Alt+↓)`;
  const kind = columnKind(app.doc, app.sheet, t, hit.col);
  const next = nextSort(t, hit.col, { stale: app.grid.isTableStale(t) });
  const what = next === 'original' ? 'put the rows back in their original order' : `sort by ${name}, ${sortLabel(kind, next[0]?.desc ?? false).toLowerCase()}`;
  const stale = t.sort[0]?.col === hit.col && app.grid.isTableStale(t) ? 'Rows changed since sorting. ' : '';
  return `${stale}Click to ${what}.\nShift+click to sort by this column too.`;
}

function menuButton(label: string, ic: L.IconNode, run: () => void, opts: { checked?: boolean; disabled?: boolean; hint?: string } = {}): HTMLButtonElement {
  const b = h(
    'button',
    { class: `btn text tm-item${opts.checked ? ' on' : ''}`, disabled: !!opts.disabled, onclick: run },
    icon(ic, 16),
    h('span', null, label),
    opts.checked ? h('span', { class: 'tm-check' }, icon(L.Check, 14)) : null,
  ) as HTMLButtonElement;
  if (opts.hint) tooltip(b, opts.hint);
  return b;
}

/** The ▾ menu of a table column: sorting choices, then the filter. */
export function openTableMenu(app: App, table: TableDef, col: number, rect: DOMRect): void {
  const s = app.sheet;
  const name = table.columns[col - table.range.c1];
  const kind = columnKind(app.doc, s, table, col);
  const rank = table.sort.findIndex((k) => k.col === col);
  const others = table.sort.filter((k) => k.col !== col);
  new FilterMenu(app, col, {
    range: table.range,
    conds: table.filter,
    apply: (conds, label) => app.transact(label, (tx) => setTableFilter(tx, app.doc, s, tableById(s, table.id)!, conds)),
    top: (close) => {
      const run = (keys: TableSortKey[] | 'original') => () => {
        close();
        sortTableBy(app, tableById(s, table.id)!, keys);
      };
      const items = [
        h('div', { class: 'tm-title' }, h('strong', null, name), h('span', null, `${table.name} · ${kind === 'number' ? 'numbers' : kind === 'date' ? 'dates' : 'text'}`)),
        menuButton(sortLabel(kind, false), L.ArrowUpNarrowWide, run([{ col, desc: false }]), { checked: rank === 0 && !table.sort[0].desc }),
        menuButton(sortLabel(kind, true), L.ArrowDownWideNarrow, run([{ col, desc: true }]), { checked: rank === 0 && table.sort[0].desc }),
      ];
      // Sort by this column too (like Shift+click on its arrow)
      if (others.length && rank !== 0) {
        const base = table.sort.filter((k) => k.col !== col);
        items.push(
          menuButton(`Then by ${name}: ${sortLabel(kind, false).toLowerCase()}`, L.ListOrdered, run([...base, { col, desc: false }]), { checked: rank > 0 && !table.sort[rank].desc, hint: 'Keeps the current sort and uses this column for ties (Shift+click its arrow)' }),
          menuButton(`Then by ${name}: ${sortLabel(kind, true).toLowerCase()}`, L.ListOrdered, run([...base, { col, desc: true }]), { checked: rank > 0 && table.sort[rank].desc }),
        );
      }
      items.push(menuButton('Original order', L.Undo2, run('original'), { disabled: !canRestoreOrder(table), hint: 'Puts the rows back the way they were before the first sort' }));
      return items;
    },
    footer: (close) => [
      h('button', {
        class: 'btn outline',
        onclick: () => {
          close();
          openTableOptions(app, tableById(s, table.id)!, rect);
        },
      }, icon(L.Settings2, 14), 'Table options'),
    ],
  }).open(rect);
}

/** The table options card: name, color, banded rows, select it, turn it back into normal cells. */
export function openTableOptions(app: App, table: TableDef, at: DOMRect | HTMLElement): void {
  const s = app.sheet;
  const current = () => tableById(s, table.id);
  const nameInput = h('input', { class: 'input', value: table.name, spellcheck: false, 'aria-label': 'Table name' }) as HTMLInputElement;
  const err = h('div', { class: 'hint tm-err' });
  const check = () => {
    const t = current();
    const problem = t ? tableNameProblem(nameInput.value, app.doc.wb, t) : null;
    err.textContent = problem ?? '';
    return !problem;
  };
  const saveName = () => {
    const t = current();
    if (!t || !check() || nameInput.value.trim() === t.name) return;
    app.transact('Rename table', (tx) => renameTable(tx, app.doc.wb, s, t, nameInput.value));
  };
  nameInput.addEventListener('input', check);
  nameInput.addEventListener('change', saveName);
  nameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      saveName();
      pop.close();
    }
  });
  const swatches = h(
    'div',
    { class: 'tm-swatches' },
    ...TABLE_COLOR_IDS.map((c) => {
      const b = h('button', {
        class: `tm-swatch${c === 'none' ? ' none' : ''}${table.color === c ? ' on' : ''}`,
        style: c === 'none' ? undefined : { background: TABLE_COLORS[c].header },
        'aria-label': TABLE_COLORS[c].label,
        onclick: () => {
          const t = current();
          if (!t) return;
          // Picking a color for a plain table turns on banded rows too
          const band = t.color === 'none' && c !== 'none' ? { banded: true } : {};
          app.transact('Table color', (tx) => setTableLook(tx, s, t, { color: c, ...band }));
          swatches.querySelectorAll('.tm-swatch').forEach((el) => el.classList.toggle('on', el === b));
          banded.checked = current()?.banded ?? banded.checked;
          banded.disabled = c === 'none';
        },
      });
      tooltip(b, TABLE_COLORS[c].label);
      return b;
    }),
  );
  const banded = h('input', { type: 'checkbox', checked: table.banded, disabled: table.color === 'none' }) as HTMLInputElement;
  banded.addEventListener('change', () => {
    const t = current();
    if (t) app.transact(banded.checked ? 'Banded rows' : 'No banded rows', (tx) => setTableLook(tx, s, t, { banded: banded.checked }));
  });
  const rows = table.range.r2 - table.range.r1;
  const content = h(
    'div',
    { class: 'table-options' },
    h('div', { class: 'tm-title' }, h('strong', null, 'Table'), h('span', null, `${formatRange(table.range)} · ${rows.toLocaleString()} row${rows === 1 ? '' : 's'} · ${table.columns.length} columns`)),
    h('div', { class: 'field' }, h('label', null, 'Name (for formulas like =SUM(' + table.name + '[' + (table.columns.find((c) => /^[\w]+$/.test(c)) ?? 'Column1') + ']))'), nameInput, err),
    h('div', { class: 'field' }, h('label', null, 'Color'), swatches),
    h('label', { class: 'check' }, banded, 'Banded rows'),
    h(
      'div',
      { class: 'row' },
      h('button', {
        class: 'btn outline',
        onclick: () => {
          pop.close();
          const t = current();
          if (!t) return;
          app.sel.setRange(t.range, { row: t.range.r1, col: t.range.c1 });
          app.onSelectionChanged();
        },
      }, icon(L.SquareDashed, 14), 'Select table'),
      h('span', { style: { flex: '1' } }),
      h('button', {
        class: 'btn outline',
        onclick: () => {
          pop.close();
          const t = current();
          if (t) convertTable(app, t);
        },
      }, icon(L.Table2, 14), 'Convert to normal cells'),
    ),
  );
  const pop = new Popover(content);
  pop.open(at);
  setTimeout(() => nameInput.focus(), 0);
}

/** Turns a table back into normal cells (data, values and formulas stay). */
export function convertTable(app: App, table: TableDef): void {
  app.transact('Convert table to cells', (tx) => removeTable(tx, app.doc, app.sheet, table));
  toast(`${table.name} is now normal cells. Its data and formulas are kept.`, { action: { label: 'Undo', run: () => app.doc.undo() } });
}
