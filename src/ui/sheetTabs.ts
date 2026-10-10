// Sheet tabs: switch, add, rename (double-click), reorder (drag), color, hide.
import * as L from 'lucide';
import { moveSheet, setSheetHidden } from '../core/ops/sheets';
import type { App } from './app';
import { deleteCurrentSheet, renameSheetTo } from './commands';
import { h, icon } from './dom';
import { colorPicker, Popover, showMenu, tooltip, type MenuItem } from './widgets';

export class SheetTabs {
  readonly el: HTMLElement;
  private list: HTMLDivElement;
  private renaming: number | null = null;
  private dragId: number | null = null;

  constructor(private app: App) {
    const add = h('button', { class: 'btn', onclick: () => app.commands.run('sheet.add') }, icon(L.Plus, 16));
    tooltip(add, 'New sheet', 'Shift+F11');
    const all = h('button', { class: 'btn' }, icon(L.Menu, 16)) as HTMLButtonElement;
    tooltip(all, 'All sheets');
    all.addEventListener('click', () =>
      showMenu(
        app.doc.wb.sheets.map((s) => ({
          label: s.name + (s.hidden ? ' (hidden)' : ''),
          checked: s.id === app.sheet.id,
          swatch: s.tabColor,
          run: () => {
            if (s.hidden) app.transact('Unhide sheet', (tx) => setSheetHidden(tx, app.doc.wb, s, false));
            app.setActiveSheet(s.id);
          },
        })),
        all,
      ),
    );
    this.list = h('div', { class: 'tabs-scroll' });
    this.list.addEventListener('wheel', (e) => {
      this.list.scrollLeft += e.deltaY;
    }, { passive: true });
    this.el = h('div', { class: 'tabs' }, add, all, this.list);
  }

  render(): void {
    const app = this.app;
    this.list.innerHTML = '';
    for (const s of app.doc.wb.sheets) {
      if (s.hidden) continue;
      const tab = h('div', { class: `tab${s.id === app.sheet.id ? ' on' : ''}`, draggable: 'true', role: 'tab', 'aria-selected': String(s.id === app.sheet.id) });
      if (this.renaming === s.id) {
        const input = h('input', { value: s.name, spellcheck: false }) as HTMLInputElement;
        const finish = (ok: boolean) => {
          if (this.renaming !== s.id) return;
          this.renaming = null;
          if (ok && input.value.trim() && input.value.trim() !== s.name) renameSheetTo(app, s.id, input.value);
          this.render();
          app.focusGrid();
        };
        input.addEventListener('keydown', (e) => {
          e.stopPropagation();
          if (e.key === 'Enter') finish(true);
          else if (e.key === 'Escape') finish(false);
        });
        input.addEventListener('blur', () => finish(true));
        tab.append(input);
        setTimeout(() => {
          input.focus();
          input.select();
        }, 0);
      } else tab.append(h('span', null, s.name));
      if (s.tabColor) tab.append(h('span', { class: 'color', style: { background: s.tabColor } }));
      tab.addEventListener('mousedown', (e) => {
        if (e.button === 0 && this.renaming !== s.id && s.id !== app.sheet.id) app.setActiveSheet(s.id);
      });
      tab.addEventListener('dblclick', () => this.startRename(s.id));
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        if (s.id !== app.sheet.id) app.setActiveSheet(s.id);
        this.tabMenu(s.id, { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY });
      });
      tab.addEventListener('dragstart', (e) => {
        this.dragId = s.id;
        e.dataTransfer?.setData('text/plain', s.name);
      });
      tab.addEventListener('dragover', (e) => {
        if (this.dragId === null) return;
        e.preventDefault();
        const r = tab.getBoundingClientRect();
        const left = e.clientX < r.left + r.width / 2;
        tab.classList.toggle('drag-over-left', left);
        tab.classList.toggle('drag-over-right', !left);
      });
      tab.addEventListener('dragleave', () => tab.classList.remove('drag-over-left', 'drag-over-right'));
      tab.addEventListener('drop', (e) => {
        e.preventDefault();
        tab.classList.remove('drag-over-left', 'drag-over-right');
        const id = this.dragId;
        this.dragId = null;
        if (id === null || id === s.id) return;
        const moving = app.doc.wb.sheetById(id)!;
        const r = tab.getBoundingClientRect();
        let to = app.doc.wb.sheets.indexOf(s) + (e.clientX < r.left + r.width / 2 ? 0 : 1);
        if (app.doc.wb.sheets.indexOf(moving) < to) to--;
        app.transact('Move sheet', (tx) => moveSheet(tx, app.doc.wb, moving, to));
      });
      tab.addEventListener('dragend', () => (this.dragId = null));
      this.list.append(tab);
    }
    const active = this.list.querySelector('.tab.on') as HTMLElement | null;
    active?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  startRename(id: number): void {
    this.renaming = id;
    this.render();
  }

  private tabMenu(id: number, at: { left: number; top: number; right: number; bottom: number }): void {
    const app = this.app;
    const s = app.doc.wb.sheetById(id)!;
    const idx = app.doc.wb.sheets.indexOf(s);
    const items: MenuItem[] = [
      { label: 'Rename', icon: L.Pencil, run: () => this.startRename(id) },
      { label: 'Duplicate', icon: L.CopyPlus, run: () => app.commands.run('sheet.duplicate') },
      {
        label: 'Tab color',
        icon: L.Palette,
        run: () => {
          const pop = new Popover(
            colorPicker((c) => {
              pop.close();
              app.transact('Tab color', (tx) => tx.setProp(s, 'tabColor', c ?? undefined));
            }, { noneLabel: 'No color', current: s.tabColor }),
          );
          pop.open(at);
        },
      },
      { separator: true },
      { label: 'Move left', icon: L.ArrowLeft, disabled: idx === 0, run: () => app.transact('Move sheet', (tx) => moveSheet(tx, app.doc.wb, s, idx - 1)) },
      { label: 'Move right', icon: L.ArrowRight, disabled: idx === app.doc.wb.sheets.length - 1, run: () => app.transact('Move sheet', (tx) => moveSheet(tx, app.doc.wb, s, idx + 1)) },
      { separator: true },
      { label: 'Hide', icon: L.EyeOff, disabled: app.doc.wb.sheets.filter((x) => !x.hidden).length <= 1, run: () => app.commands.run('sheet.hide') },
      { label: 'Unhide…', icon: L.Eye, disabled: !app.doc.wb.sheets.some((x) => x.hidden), run: () => this.showUnhideMenu(at) },
      { separator: true },
      { label: 'Delete', icon: L.Trash2, danger: true, disabled: app.doc.wb.sheets.length <= 1, run: () => deleteCurrentSheet(app, id) },
    ];
    showMenu(items, at);
  }

  showUnhideMenu(at?: { left: number; top: number; right: number; bottom: number }): void {
    const app = this.app;
    const hidden = app.doc.wb.sheets.filter((s) => s.hidden);
    const rect = at ?? this.el.getBoundingClientRect();
    showMenu(
      hidden.map((s) => ({
        label: s.name,
        icon: L.Eye,
        run: () => {
          app.transact('Unhide sheet', (tx) => setSheetHidden(tx, app.doc.wb, s, false));
          app.setActiveSheet(s.id);
        },
      })),
      rect,
    );
  }
}
