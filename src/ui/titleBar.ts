// The top bar: Gridline menu (logo), quick-access buttons, document name, command search and theme.
import * as L from 'lucide';
import type { App } from './app';
import { h, icon, kbd } from './dom';
import { openPalette } from './palette';
import { suiteButton } from './suiteMenu';
import { isDesktop, platform } from './platform';
import { showMenu, tooltip, type MenuItem } from './widgets';

export class TitleBar {
  readonly el: HTMLElement;
  private themeBtn!: HTMLButtonElement;
  private undoBtn!: HTMLButtonElement;
  private redoBtn!: HTMLButtonElement;

  constructor(private app: App) {
    this.el = h('header', { class: 'titlebar' });
    this.rebuild();
  }

  /** Rebuilds the bar (after settings change which buttons are shown). */
  rebuild(): void {
    const app = this.app;
    const s = app.settings;
    const brand = h('button', { class: 'brand', 'aria-haspopup': 'menu', onclick: (e: MouseEvent) => this.openAppMenu(e.currentTarget as HTMLElement) }, h('span', { class: 'brand-mark' }, icon(L.Grid2x2, 14)), 'Gridline', icon(L.ChevronDown, 12, 'caret'));
    tooltip(brand, 'Settings, updates and help', 'Ctrl+,');
    const quick = h('div', { class: 'quick-access' });
    if (s.bool('appearance.showQuickAccess')) {
      const save = h('button', { class: 'btn', 'aria-label': 'Save', onclick: () => app.commands.run('file.save') }, icon(L.Save, 16));
      tooltip(save, 'Save', 'Ctrl+S');
      this.undoBtn = h('button', { class: 'btn', 'aria-label': 'Undo', onclick: () => app.commands.run('edit.undo') }, icon(L.Undo2, 16)) as HTMLButtonElement;
      tooltip(this.undoBtn, () => (app.doc.history.peekUndoLabel() ? `Undo ${app.doc.history.peekUndoLabel()}` : 'Undo'), 'Ctrl+Z');
      this.redoBtn = h('button', { class: 'btn', 'aria-label': 'Redo', onclick: () => app.commands.run('edit.redo') }, icon(L.Redo2, 16)) as HTMLButtonElement;
      tooltip(this.redoBtn, () => (app.doc.history.peekRedoLabel() ? `Redo ${app.doc.history.peekRedoLabel()}` : 'Redo'), 'Ctrl+Y');
      quick.append(save, this.undoBtn, this.redoBtn);
    }

    const search = s.bool('appearance.showSearch')
      ? h('div', { class: 'cmd-search', role: 'button', tabindex: '0', onclick: () => openPalette(app) }, icon(L.Search, 14), h('span', { class: 'label' }, 'Search commands, functions, cells…'), kbd('Ctrl+K'))
      : null;
    this.themeBtn = h('button', { class: 'btn', 'aria-label': 'Switch light / dark theme', onclick: () => app.commands.run('view.theme') }) as HTMLButtonElement;
    tooltip(this.themeBtn, 'Switch light / dark theme');
    const help = h('button', { class: 'btn', 'aria-label': 'Keyboard shortcuts', onclick: () => app.commands.run('help.shortcuts') }, icon(L.Keyboard, 16));
    tooltip(help, 'Keyboard shortcuts', 'Ctrl+/');
    this.el.innerHTML = '';
    const parts: (HTMLElement | null)[] = [
      suiteButton(),
      brand,
      quick,
      app.bookStrip.el,
      h('div', { class: 'spacer' }),
      search ?? h('div', { class: 'spacer' }),
      h('div', { class: 'spacer' }),
      s.bool('appearance.showShortcutsButton') ? help : null,
      s.bool('appearance.showThemeButton') ? this.themeBtn : null,
      isDesktop && platform.platform !== 'darwin' ? h('div', { class: 'controls-gap' }) : null,
    ];
    this.el.append(...parts.filter((p): p is HTMLElement => p !== null));
    this.refresh();
  }

  refresh(): void {
    const d = this.app.doc;
    this.themeBtn.innerHTML = '';
    this.themeBtn.append(icon(this.app.isDark ? L.Sun : L.Moon, 16));
    if (this.undoBtn?.isConnected) {
      this.undoBtn.disabled = !d.history.canUndo;
      this.redoBtn.disabled = !d.history.canRedo;
    }
    const where = this.app.settings.bool('appearance.titlePath') && d.filePath ? d.filePath : d.fileName;
    platform.setTitle(`${where}${d.dirty ? ' •' : ''} — Gridline`);
  }

  /** The Gridline (logo) menu. */
  openAppMenu(anchor: HTMLElement): void {
    const app = this.app;
    const cmd = (id: string, label?: string): MenuItem => {
      const c = app.commands.get(id)!;
      return { label: label ?? c.title, icon: c.icon, hint: c.keys?.[0], run: () => app.commands.run(id) };
    };
    showMenu(
      [
        cmd('app.settings', 'Settings'),
        { separator: true },
        cmd('help.updates', 'Check for updates…'),
        cmd('help.whatsNew', "What's new"),
        { separator: true },
        cmd('file.home', 'Home'),
        cmd('file.newWindow', 'New window'),
        cmd('help.shortcuts'),
        cmd('help.functions'),
        { separator: true },
        cmd('help.about', 'About Gridline'),
      ],
      anchor,
    );
  }
}
