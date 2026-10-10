// The top bar: app menu, document name, command search and theme toggle.
import * as L from 'lucide';
import type { App } from './app';
import { h, icon, kbd } from './dom';
import { openPalette } from './palette';
import { isDesktop, platform } from './platform';
import { showMenu, tooltip, type MenuItem } from './widgets';

export class TitleBar {
  readonly el: HTMLElement;
  private nameEl: HTMLSpanElement;
  private extEl: HTMLSpanElement;
  private stateEl: HTMLSpanElement;
  private themeBtn: HTMLButtonElement;

  constructor(private app: App) {
    const fileBtn = h('button', { class: 'btn text', onclick: (e: MouseEvent) => this.openFileMenu(e.currentTarget as HTMLElement) }, 'File', icon(L.ChevronDown, 12, 'caret'));
    tooltip(fileBtn, 'New, open, save and export');
    this.nameEl = h('span', { class: 'name' }, 'Untitled');
    this.extEl = h('span', { class: 'ext' });
    this.stateEl = h('span', { class: 'saved' });
    const title = h('div', { class: 'doc-title', title: '' }, this.nameEl, this.extEl, this.stateEl);
    const search = h(
      'div',
      { class: 'cmd-search', role: 'button', tabindex: '0', onclick: () => openPalette(app) },
      icon(L.Search, 14),
      h('span', { class: 'label' }, 'Search commands, functions, cells…'),
      kbd('Ctrl+K'),
    );
    this.themeBtn = h('button', { class: 'btn', onclick: () => app.commands.run('view.theme') }) as HTMLButtonElement;
    tooltip(this.themeBtn, 'Switch light / dark theme');
    const help = h('button', { class: 'btn', onclick: () => app.commands.run('help.shortcuts') }, icon(L.Keyboard, 16));
    tooltip(help, 'Keyboard shortcuts', 'Ctrl+/');
    this.el = h(
      'header',
      { class: 'titlebar' },
      h('div', { class: 'brand' }, h('span', { class: 'brand-mark' }, icon(L.Grid2x2, 14)), 'Gridline'),
      fileBtn,
      title,
      h('div', { class: 'spacer' }),
      search,
      h('div', { class: 'spacer' }),
      help,
      this.themeBtn,
      isDesktop && platform.platform !== 'darwin' ? h('div', { class: 'controls-gap' }) : null,
    );
  }

  refresh(): void {
    const d = this.app.doc;
    const m = /^(.*?)(\.[^.]+)?$/.exec(d.fileName)!;
    this.nameEl.textContent = m[1] || d.fileName;
    this.extEl.textContent = m[2] ?? '';
    this.stateEl.innerHTML = '';
    if (d.dirty) {
      this.stateEl.className = 'dirty';
      this.stateEl.title = 'Unsaved changes';
    } else {
      this.stateEl.className = 'saved';
      this.stateEl.textContent = d.filePath ? 'Saved' : '';
    }
    this.themeBtn.innerHTML = '';
    this.themeBtn.append(icon(this.app.prefs.theme === 'dark' ? L.Sun : L.Moon, 16));
    platform.setTitle(`${d.fileName}${d.dirty ? ' •' : ''} — Gridline`);
  }

  private async openFileMenu(anchor: HTMLElement): Promise<void> {
    const app = this.app;
    const recent = await platform.recentFiles();
    const cmd = (id: string, label?: string): MenuItem => {
      const c = app.commands.get(id)!;
      return { label: label ?? c.title, icon: c.icon, hint: c.keys?.[0], disabled: c.enabled ? !c.enabled() : false, run: () => app.commands.run(id) };
    };
    const items: MenuItem[] = [
      cmd('file.new'),
      cmd('file.template'),
      cmd('file.open'),
      {
        label: 'Open recent',
        icon: L.History,
        disabled: !recent.length,
        submenu: () => [
          ...recent.map((r) => ({ label: r.name, hint: shortPath(r.path), icon: L.FileSpreadsheet, run: () => app.openRecent(r.path) })),
          { separator: true as const },
          { label: 'Clear recent files', icon: L.Trash2, run: () => platform.clearRecentFiles() },
        ],
      },
      { separator: true },
      cmd('file.save'),
      cmd('file.saveAs'),
      cmd('file.exportCsv'),
      cmd('file.exportTsv'),
      { separator: true },
      cmd('file.showInFolder'),
      { separator: true },
      cmd('help.shortcuts'),
      cmd('help.about'),
    ];
    showMenu(items, anchor);
  }
}

function shortPath(p: string): string {
  const parts = p.split(/[\\/]/);
  if (parts.length <= 3) return p;
  return '…\\' + parts.slice(-3, -1).join('\\');
}
