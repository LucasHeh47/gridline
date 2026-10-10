// The Home page: start something new (blank or a preset), reopen recent files, and the file actions
// (save, export…). Shown when Gridline starts and from the File tab.
import * as L from 'lucide';
import { cellKey } from '../core/address';
import { external } from '../core/data/external';
import { formatValue } from '../core/numfmt';
import { buildTemplate, TEMPLATE_CATEGORIES, TEMPLATES, type TemplateInfo } from '../core/templates';
import type { RecentFile } from '../shared/api';
import type { App } from './app';
import { h, icon } from './dom';
import { displayFill, readTheme } from './grid/theme';
import { isDesktop, platform } from './platform';
import { showMenu, toast, tooltip } from './widgets';

export type HomeSection = 'home' | 'new' | 'open';

const ICONS: Record<string, L.IconNode> = {
  Wallet: L.Wallet,
  House: L.House,
  TrendingDown: L.TrendingDown,
  PiggyBank: L.PiggyBank,
  ChartCandlestick: L.ChartCandlestick,
  Bitcoin: L.Bitcoin,
  Trophy: L.Trophy,
  Medal: L.Medal,
  Plane: L.Plane,
  ListChecks: L.ListChecks,
};

/** Presets shown on the Home section (the rest are under New). */
const FEATURED = ['budget', 'stocks', 'crypto', 'nfl', 'fantasy', 'travel'];

/** Small pictures of each preset's first sheet, built once without downloading anything. */
const previews = new Map<string, { rows: { text: string; bold: boolean; fill?: string; align: string }[][]; widths: number[] }>();

function previewData(id: string) {
  let p = previews.get(id);
  if (p) return p;
  external.paused++;
  try {
    const wb = buildTemplate(id);
    const rows: { text: string; bold: boolean; fill?: string; align: string }[][] = [];
    const widths: number[] = [];
    if (wb) {
      const s = wb.sheets[0];
      for (let c = 0; c < 6; c++) widths.push(Math.min(160, s.colWidth(c)));
      for (let r = 0; r < 11; r++) {
        const row = [];
        for (let c = 0; c < 6; c++) {
          const cell = s.cells.get(cellKey(r, c));
          const st = wb.styles.get(s.styleIdAt(r, c));
          let text = '';
          if (cell?.v !== undefined && cell.v !== null && cell.f === undefined) text = typeof cell.v === 'number' ? formatValue(cell.v, st.numFmt).text : String(cell.v);
          else if (cell?.f !== undefined) text = '·';
          row.push({ text, bold: !!st.bold, fill: st.fill, align: typeof cell?.v === 'number' || cell?.f !== undefined ? 'right' : 'left' });
        }
        rows.push(row);
      }
    }
    p = { rows, widths };
    previews.set(id, p);
    return p;
  } finally {
    external.paused--;
  }
}

function relativeTime(ms: number): string {
  const d = new Date(ms);
  const now = new Date();
  const mins = Math.round((now.getTime() - ms) / 60000);
  if (mins < 1) return 'Just now';
  if (mins < 60) return `${mins} minute${mins === 1 ? '' : 's'} ago`;
  const sameDay = d.toDateString() === now.toDateString();
  const time = d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
  if (sameDay) return `Today, ${time}`;
  const yest = new Date(now);
  yest.setDate(now.getDate() - 1);
  if (d.toDateString() === yest.toDateString()) return `Yesterday, ${time}`;
  const days = (now.getTime() - ms) / 86400000;
  if (days < 7) return d.toLocaleDateString(undefined, { weekday: 'long' });
  return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' });
}

function folderOf(p: string): string {
  const parts = p.split(/[\\/]/);
  parts.pop();
  if (parts.length <= 2) return parts.join('\\');
  return '…\\' + parts.slice(-2).join('\\');
}

function greeting(): string {
  const hr = new Date().getHours();
  return hr < 5 ? 'Good evening' : hr < 12 ? 'Good morning' : hr < 18 ? 'Good afternoon' : 'Good evening';
}

export class HomeView {
  readonly el: HTMLElement;
  private main: HTMLElement;
  private nav: HTMLElement;
  private section: HomeSection;
  private recent: RecentFile[] = [];
  private filter = '';
  private recentTab: 'recent' | 'pinned' = 'recent';
  private onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape' && !document.querySelector('.modal-backdrop, .popover')) {
      e.preventDefault();
      e.stopPropagation();
      this.close();
    }
  };

  constructor(
    private app: App,
    section: HomeSection,
    private onClosed: () => void,
  ) {
    this.section = section;
    this.nav = h('nav', { class: 'home-nav' });
    this.main = h('main', { class: 'home-main' });
    const close = h('button', { class: 'btn home-close', 'aria-label': 'Close Home', onclick: () => this.close() }, icon(L.X, 18));
    tooltip(close, 'Back to the workbook', 'Esc');
    this.el = h('div', { class: 'home-view', role: 'dialog', 'aria-label': 'Home' }, h('aside', { class: 'home-rail' }, this.nav), h('div', { class: 'home-body' }, h('div', { class: 'home-top' }, h('div', { class: 'spacer' }), close), this.main));
    document.addEventListener('keydown', this.onKey, true);
    this.renderNav();
    this.render();
    void this.loadRecent();
  }

  close(): void {
    document.removeEventListener('keydown', this.onKey, true);
    this.el.remove();
    this.onClosed();
  }

  show(section: HomeSection): void {
    this.section = section;
    this.renderNav();
    this.render();
    void this.loadRecent();
  }

  private async loadRecent(): Promise<void> {
    this.recent = await platform.recentFiles();
    this.render();
  }

  // ---------- navigation ----------

  private renderNav(): void {
    const app = this.app;
    const doc = app.doc;
    this.nav.innerHTML = '';
    const brand = h('div', { class: 'home-brand' }, h('span', { class: 'brand-mark' }, icon(L.Grid2x2, 16)), 'Gridline');
    const item = (id: HomeSection, label: string, ic: L.IconNode) =>
      h('button', { class: `home-nav-item${this.section === id ? ' on' : ''}`, onclick: () => this.show(id) }, icon(ic, 18), label);
    const action = (label: string, ic: L.IconNode, run: () => unknown, disabled = false) => {
      const b = h('button', { class: 'home-nav-item', onclick: () => {
        this.close();
        void run();
      } }, icon(ic, 18), label) as HTMLButtonElement;
      b.disabled = disabled;
      return b;
    };
    const hasWork = !app.isPristineDoc();
    const parts: (HTMLElement | null)[] = [
      brand,
      item('home', 'Home', L.House),
      item('new', 'New', L.FilePlus),
      item('open', 'Open', L.FolderOpen),
      h('div', { class: 'home-nav-sep' }),
      action('Save', L.Save, () => app.commands.run('file.save'), !hasWork),
      action('Save as', L.SaveAll, () => app.commands.run('file.saveAs'), !hasWork),
      action('Export as CSV', L.FileText, () => app.commands.run('file.exportCsv'), !hasWork),
      isDesktop && doc.filePath ? action('Show in folder', L.FolderSearch, () => app.commands.run('file.showInFolder')) : null,
      h('div', { class: 'spacer' }),
      action('Settings', L.Settings, () => app.openSettings()),
      action('About', L.Info, () => app.commands.run('help.about')),
    ];
    this.nav.append(...parts.filter((x): x is HTMLElement => x !== null));
  }

  // ---------- content ----------

  private render(): void {
    const scroll = this.main.scrollTop;
    this.main.innerHTML = '';
    if (this.section === 'home') this.renderHome();
    else if (this.section === 'new') this.renderNew();
    else this.renderOpen();
    this.main.scrollTop = scroll;
  }

  private renderHome(): void {
    this.main.append(h('h1', { class: 'home-h1' }, greeting()));
    const strip = h('div', { class: 'tpl-strip' }, this.blankCard());
    for (const id of FEATURED) {
      const t = TEMPLATES.find((x) => x.id === id);
      if (t) strip.append(this.templateCard(t));
    }
    this.main.append(
      h('div', { class: 'home-row-head' }, h('h2', null, 'Start something new'), h('button', { class: 'btn text home-link', onclick: () => this.show('new') }, 'More presets', icon(L.ArrowRight, 14))),
      strip,
    );
    this.main.append(this.recentBlock(false));
  }

  private renderNew(): void {
    this.main.append(h('h1', { class: 'home-h1' }, 'New'));
    this.main.append(h('div', { class: 'tpl-grid-home' }, this.blankCard()));
    for (const cat of TEMPLATE_CATEGORIES) {
      const list = TEMPLATES.filter((t) => t.category === cat);
      if (!list.length) continue;
      this.main.append(h('h2', { class: 'home-cat' }, cat));
      const grid = h('div', { class: 'tpl-grid-home' });
      for (const t of list) grid.append(this.templateCard(t));
      this.main.append(grid);
    }
    this.main.append(h('p', { class: 'home-note' }, icon(L.Wifi, 14), 'Presets marked "Live" download prices, scores or weather the first time you open them.'));
  }

  private renderOpen(): void {
    this.main.append(h('h1', { class: 'home-h1' }, 'Open'));
    const browse = h('button', { class: 'btn primary', onclick: () => {
      this.close();
      void this.app.openFile();
    } }, icon(L.FolderOpen, 16), 'Browse…');
    this.main.append(h('div', { class: 'home-open-actions' }, browse, h('span', { class: 'hint' }, 'Excel workbooks (.xlsx), CSV and tab-separated files. You can also drag a file onto the window.')));
    this.main.append(this.recentBlock(true));
  }

  private blankCard(): HTMLElement {
    const thumb = h('div', { class: 'tpl-thumb blank' }, icon(L.Plus, 30));
    const card = h('button', { class: 'tpl-card-home', onclick: () => this.newBlank() }, thumb, h('div', { class: 'tpl-meta' }, h('div', { class: 'tpl-title' }, 'Blank workbook'), h('div', { class: 'tpl-sub' }, 'Start from scratch')));
    tooltip(card, 'An empty workbook', 'Ctrl+N');
    return card;
  }

  private templateCard(t: TemplateInfo): HTMLElement {
    const thumb = h('div', { class: 'tpl-thumb' });
    thumb.style.setProperty('--tpl-color', t.color);
    const data = previewData(t.id);
    const theme = readTheme(this.app.settings.bool('appearance.adaptColors'));
    const table = h('div', { class: 'tpl-mini', style: { gridTemplateColumns: data.widths.map((w) => `${Math.max(14, w)}fr`).join(' ') } });
    for (const row of data.rows) {
      for (const c of row) {
        const cell = h('span', { class: `${c.bold ? 'b' : ''}${c.align === 'right' ? ' r' : ''}` }, c.text);
        if (c.fill) cell.style.background = displayFill(theme, c.fill);
        table.append(cell);
      }
    }
    thumb.append(table, h('span', { class: 'tpl-badge' }, icon(ICONS[t.icon] ?? L.FileSpreadsheet, 16)));
    const meta = h('div', { class: 'tpl-meta' }, h('div', { class: 'tpl-title' }, t.name, t.live ? h('span', { class: 'live-tag' }, 'Live') : null), h('div', { class: 'tpl-sub' }, t.category));
    const card = h('button', { class: 'tpl-card-home', onclick: () => this.newFrom(t.id) }, thumb, meta);
    tooltip(card, t.description);
    return card;
  }

  private recentBlock(full: boolean): HTMLElement {
    const wrap = h('section', { class: 'home-recent' });
    const pinned = this.recent.filter((r) => r.pinned);
    const tabs = h(
      'div',
      { class: 'home-tabs' },
      h('button', { class: `home-tab${this.recentTab === 'recent' ? ' on' : ''}`, onclick: () => {
        this.recentTab = 'recent';
        this.render();
      } }, 'Recent'),
      h('button', { class: `home-tab${this.recentTab === 'pinned' ? ' on' : ''}`, onclick: () => {
        this.recentTab = 'pinned';
        this.render();
      } }, `Pinned${pinned.length ? ` (${pinned.length})` : ''}`),
    );
    const search = h('input', { class: 'input home-search', type: 'search', placeholder: 'Search recent files', value: this.filter, 'aria-label': 'Search recent files' }) as HTMLInputElement;
    search.addEventListener('input', () => {
      this.filter = search.value;
      const pos = search.selectionStart;
      this.render();
      const again = this.main.querySelector('.home-search') as HTMLInputElement | null;
      again?.focus();
      again?.setSelectionRange(pos, pos);
    });
    wrap.append(h('div', { class: 'home-row-head' }, tabs, h('div', { class: 'spacer' }), search));

    const q = this.filter.trim().toLowerCase();
    let list = this.recentTab === 'pinned' ? pinned : [...pinned, ...this.recent.filter((r) => !r.pinned)];
    if (q) list = list.filter((r) => r.name.toLowerCase().includes(q) || r.path.toLowerCase().includes(q));
    if (!full && !q) list = list.slice(0, 12);

    if (!list.length) {
      const msg = q
        ? `No recent files match "${this.filter}".`
        : this.recentTab === 'pinned'
          ? 'Pin files you use often and they stay at the top. Hover a recent file and click the pin.'
          : isDesktop
            ? "Files you open or save will show up here."
            : 'Recent files are listed in the desktop app.';
      wrap.append(h('div', { class: 'home-empty' }, icon(this.recentTab === 'pinned' ? L.Pin : L.Clock, 26), h('p', null, msg)));
      return wrap;
    }
    const table = h('div', { class: 'recent-list', role: 'list' });
    table.append(h('div', { class: 'recent-row head' }, h('span'), h('span', null, 'Name'), h('span', null, 'Location'), h('span', null, 'Opened'), h('span')));
    for (const r of list) table.append(this.recentRow(r));
    wrap.append(table);
    return wrap;
  }

  private recentRow(r: RecentFile): HTMLElement {
    const ext = (/\.([^.]+)$/.exec(r.name)?.[1] ?? '').toLowerCase();
    const pin = h('button', { class: `btn icon-only${r.pinned ? ' on' : ''}`, 'aria-label': r.pinned ? 'Unpin' : 'Pin to the top', onclick: (e: MouseEvent) => {
      e.stopPropagation();
      void platform.pinRecentFile(r.path, !r.pinned).then(() => this.loadRecent());
    } }, icon(r.pinned ? L.PinOff : L.Pin, 15));
    tooltip(pin, r.pinned ? 'Unpin' : 'Pin to the top');
    const more = h('button', { class: 'btn icon-only', 'aria-label': 'More', onclick: (e: MouseEvent) => {
      e.stopPropagation();
      showMenu(
        [
          { label: 'Open', icon: L.FolderOpen, run: () => this.openRecent(r) },
          { label: 'Open in a new window', icon: L.AppWindow, run: () => platform.newWindow(r.path) },
          { label: 'Show in folder', icon: L.FolderSearch, run: () => platform.showInFolder(r.path) },
          { label: 'Copy path', icon: L.Copy, run: () => void platform.writeClipboard({ text: r.path }).then(() => toast('Copied the file path.')) },
          { separator: true },
          { label: 'Remove from list', icon: L.X, run: () => void platform.removeRecentFile(r.path).then(() => this.loadRecent()) },
        ],
        e.currentTarget as HTMLElement,
      );
    } }, icon(L.Ellipsis, 15));
    const row = h(
      'div',
      { class: 'recent-row', role: 'listitem', tabindex: '0', title: r.path },
      h('span', { class: `file-icon ${ext}` }, icon(ext === 'csv' || ext === 'tsv' || ext === 'txt' ? L.FileText : L.FileSpreadsheet, 18)),
      h('span', { class: 'name' }, r.name, r.pinned ? h('span', { class: 'pinned-mark', title: 'Pinned' }, icon(L.Pin, 12)) : null),
      h('span', { class: 'loc' }, folderOf(r.path)),
      h('span', { class: 'when' }, relativeTime(r.openedAt)),
      h('span', { class: 'acts' }, pin, more),
    );
    row.addEventListener('click', () => this.openRecent(r));
    row.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') this.openRecent(r);
    });
    return row;
  }

  // ---------- actions ----------

  private newBlank(): void {
    this.close();
    if (this.app.isPristineDoc()) return;
    void this.app.newWorkbook();
  }

  private newFrom(id: string): void {
    this.close();
    void this.app.newFromTemplate(id);
  }

  private openRecent(r: RecentFile): void {
    this.close();
    void this.app.openRecent(r.path);
  }
}
