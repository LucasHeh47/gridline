// Workbook tabs in the title bar: several files open in one window, like browser tabs. Click to
// switch, middle-click or × to close, drag to reorder, right-click for more.
import * as L from 'lucide';
import type { Doc } from '../core/doc';
import type { SelectionSnapshot } from '../core/history';
import type { App } from './app';
import { h, icon } from './dom';
import type { MacroTabState } from './macros/MacroController';
import { isDesktop, platform } from './platform';
import { showMenu, toast, tooltip, type MenuItem } from './widgets';

/** One open workbook and what its tab remembers while another one is shown. */
export interface BookTab {
  id: number;
  doc: Doc;
  /** Selection when the tab was last shown. */
  sel?: SelectionSnapshot;
  /** Answer to "keep CSV or save as .xlsx?" for this file. */
  csvChoice: 'xlsx' | 'csv' | null;
  macros?: MacroTabState;
}

let seq = 0;
export const newBookTab = (doc: Doc): BookTab => ({ id: ++seq, doc, csvChoice: null });

function fileIcon(doc: Doc) {
  if (doc.fileType === 'csv' || doc.fileType === 'tsv') return L.FileText;
  if (doc.fileType === 'xlsm') return L.FileCode;
  return L.FileSpreadsheet;
}

export class BookTabStrip {
  readonly el: HTMLElement;
  private list: HTMLElement;
  private dragId: number | null = null;

  constructor(private app: App) {
    this.list = h('div', { class: 'book-tabs-list', role: 'tablist', 'aria-label': 'Open workbooks' });
    const add = h('button', { class: 'btn book-tab-new', 'aria-label': 'New workbook', onclick: () => app.commands.run('file.new') }, icon(L.Plus, 16));
    tooltip(add, 'New workbook', 'Ctrl+N');
    this.el = h('div', { class: 'book-tabs' }, this.list, add);
    // Scroll through many tabs with the mouse wheel
    this.list.addEventListener(
      'wheel',
      (e) => {
        if (this.list.scrollWidth <= this.list.clientWidth) return;
        e.preventDefault();
        this.list.scrollLeft += e.deltaY || e.deltaX;
      },
      { passive: false },
    );
    // Double-click empty space for a new tab, like a browser
    this.list.addEventListener('dblclick', (e) => {
      if (e.target === this.list) app.commands.run('file.new');
    });
    this.render();
  }

  render(): void {
    const app = this.app;
    const tabs = app.books;
    this.list.replaceChildren(
      ...tabs.map((t) => {
        const d = t.doc;
        const on = t === app.book;
        const m = /^(.*?)(\.[^.]+)?$/.exec(d.fileName)!;
        const close = h('button', { class: 'book-tab-close', 'aria-label': `Close ${d.fileName}`, tabindex: '-1', onclick: (e: Event) => {
          e.stopPropagation();
          void app.closeBook(t);
        } }, icon(L.X, 13));
        const el = h(
          'div',
          {
            class: `book-tab${on ? ' on' : ''}${d.dirty ? ' dirty' : ''}`,
            role: 'tab',
            'aria-selected': on ? 'true' : 'false',
            tabindex: on ? '0' : '-1',
            draggable: 'true',
            title: `${d.filePath ?? d.fileName}${d.dirty ? ' — unsaved changes' : ''}`,
          },
          h('span', { class: 'book-tab-icon' }, icon(fileIcon(d), 14)),
          h('span', { class: 'book-tab-name' }, m[1] || d.fileName),
          h('span', { class: 'book-tab-dot', 'aria-hidden': 'true' }),
          close,
        );
        el.addEventListener('mousedown', (e) => {
          if (e.button === 1) e.preventDefault();
        });
        el.addEventListener('click', () => app.switchBook(t));
        el.addEventListener('auxclick', (e) => {
          if (e.button === 1) void app.closeBook(t);
        });
        el.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          this.menu(t, { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY });
        });
        el.addEventListener('dragstart', (e) => {
          this.dragId = t.id;
          e.dataTransfer?.setData('text/x-gridline-book', String(t.id));
          if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
        });
        el.addEventListener('dragend', () => {
          this.dragId = null;
          this.list.querySelectorAll('.drop-left, .drop-right').forEach((x) => x.classList.remove('drop-left', 'drop-right'));
        });
        el.addEventListener('dragover', (e) => {
          if (this.dragId === null || this.dragId === t.id) return;
          e.preventDefault();
          const r = el.getBoundingClientRect();
          const left = e.clientX < r.left + r.width / 2;
          el.classList.toggle('drop-left', left);
          el.classList.toggle('drop-right', !left);
        });
        el.addEventListener('dragleave', () => el.classList.remove('drop-left', 'drop-right'));
        el.addEventListener('drop', (e) => {
          if (this.dragId === null) return;
          e.preventDefault();
          const from = app.books.findIndex((b) => b.id === this.dragId);
          const r = el.getBoundingClientRect();
          let to = app.books.indexOf(t) + (e.clientX < r.left + r.width / 2 ? 0 : 1);
          if (from < to) to--;
          app.moveBook(from, to);
        });
        return el;
      }),
    );
    // Keep the shown tab in view
    const active = this.list.querySelector<HTMLElement>('.book-tab.on');
    active?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  private menu(t: BookTab, at: { left: number; top: number; right: number; bottom: number }): void {
    const app = this.app;
    const i = app.books.indexOf(t);
    const path = t.doc.filePath;
    const items: MenuItem[] = [
      { label: 'New workbook', icon: L.Plus, hint: 'Ctrl+N', run: () => app.commands.run('file.new') },
      { separator: true },
      { label: 'Close', icon: L.X, hint: 'Ctrl+W', run: () => void app.closeBook(t) },
      { label: 'Close other tabs', disabled: app.books.length < 2, run: () => void app.closeBooks(app.books.filter((b) => b !== t)) },
      { label: 'Close tabs to the right', disabled: i === app.books.length - 1, run: () => void app.closeBooks(app.books.slice(i + 1)) },
    ];
    if (path && isDesktop) {
      items.push(
        { separator: true },
        { label: 'Show in folder', icon: L.FolderOpen, run: () => platform.showInFolder(path) },
        {
          label: 'Copy file path',
          icon: L.Copy,
          run: () => {
            void platform.writeClipboard({ text: path });
            toast('File path copied.');
          },
        },
        { label: 'Open in a new window', icon: L.AppWindow, disabled: t.doc.dirty, run: () => app.moveBookToWindow(t) },
      );
    }
    showMenu(items, at);
  }
}
