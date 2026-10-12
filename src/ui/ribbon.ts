// The ribbon: tabs of small, clearly labelled groups. It has three sizes (full, compact, tabs only),
// chosen by dragging its bottom edge, double-clicking a tab, or in Settings. Groups that don't fit
// the window fold into a single button that opens them in a pop-up.
import * as L from 'lucide';
import type { IconNode } from 'lucide';
import { FUNCTIONS, type Category } from '../core/formula/functions';
import { FORMAT_PRESETS, formatValue } from '../core/numfmt';
import { DEFAULT_FONT, DEFAULT_FONT_SIZE } from '../core/styles';
import type { App } from './app';
import { h, icon } from './dom';
import { colorPicker, Popover, showMenu, tooltip, type MenuItem } from './widgets';

export const FONTS = [
  'Aptos',
  'Arial',
  'Calibri',
  'Cambria',
  'Candara',
  'Cascadia Mono',
  'Consolas',
  'Constantia',
  'Corbel',
  'Courier New',
  'Georgia',
  'Segoe UI',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
];

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

export type RibbonSize = 'full' | 'compact' | 'tabs';

interface GroupDef {
  id: string;
  label: string;
  icon: IconNode;
  build: () => HTMLElement[];
}

interface TabDef {
  id: string;
  label: string;
  groups: () => GroupDef[];
}

const TAB_KEY = 'gridline.ribbonTab';

/** Function categories offered as menus on the Formulas tab. */
const FN_MENUS: [string, Category[], IconNode][] = [
  ['Financial', ['Financial'], L.Landmark],
  ['Logical', ['Logical'], L.GitBranch],
  ['Text', ['Text'], L.Type],
  ['Date & time', ['Date & time'], L.CalendarDays],
  ['Lookup', ['Lookup'], L.Search],
  ['Math', ['Math'], L.Calculator],
  ['Statistics', ['Statistical'], L.ChartNoAxesColumn],
  ['Info', ['Information'], L.Info],
  ['Money', ['Personal finance'], L.Wallet],
  ['Sports', ['NFL'], L.Trophy],
  ['World', ['World'], L.Globe],
];

export class Ribbon {
  readonly el: HTMLElement;
  private tabsEl: HTMLElement;
  private body: HTMLElement;
  private active: string;
  private toggles: { el: HTMLElement; checked: () => boolean }[] = [];
  private updaters: (() => void)[] = [];
  private groupEls: HTMLElement[] = [];
  /** Tabs-only mode: the ribbon body is showing as a pop-up. */
  private floating = false;
  private dragPreview: RibbonSize | null = null;
  textColor = '#ff6b6b';
  fillColor = '#5a4c12';
  borderColor: string | undefined = undefined;

  constructor(private app: App) {
    let start = 'home';
    try {
      if (app.settings.str('ribbon.startTab') === 'last') start = localStorage.getItem(TAB_KEY) ?? 'home';
    } catch {
      // storage unavailable
    }
    this.active = start;
    this.tabsEl = h('div', { class: 'rb-tabs', role: 'tablist' });
    this.body = h('div', { class: 'rb-body', role: 'toolbar' });
    const grip = h('div', { class: 'rb-resize', title: 'Drag to resize the ribbon' });
    this.el = h('div', { class: 'ribbon' }, this.tabsEl, this.body, grip);
    this.bindResize(grip);
    this.tabsEl.addEventListener('wheel', (e) => {
      if (!app.settings.bool('ribbon.wheelTabs')) return;
      e.preventDefault();
      const ids = this.visibleTabs().map((t) => t.id);
      const i = ids.indexOf(this.active);
      this.select(ids[(i + (e.deltaY > 0 ? 1 : -1) + ids.length) % ids.length], false);
    }, { passive: false });
    // pointerdown: the grid cancels presses, so mousedown never arrives from there
    document.addEventListener('pointerdown', (e) => {
      if (!this.floating) return;
      const t = e.target;
      if (t instanceof Element && (this.el.contains(t) || t.closest('.popover, .modal-backdrop'))) return;
      this.setFloating(false);
    }, true);
    document.addEventListener('keydown', (e) => {
      if (this.floating && e.key === 'Escape' && !document.querySelector('.popover')) {
        this.setFloating(false);
        app.focusGrid();
      }
    });
    new ResizeObserver(() => this.layout()).observe(this.body);
    this.rebuild();
  }

  get size(): RibbonSize {
    return this.dragPreview ?? (this.app.settings.str('ribbon.size') as RibbonSize);
  }

  // ---------- tabs ----------

  private tabs(): TabDef[] {
    return [
      { id: 'home', label: 'Home', groups: () => this.homeGroups() },
      { id: 'insert', label: 'Insert', groups: () => this.insertGroups() },
      { id: 'formulas', label: 'Formulas', groups: () => this.formulaGroups() },
      { id: 'data', label: 'Data', groups: () => this.dataGroups() },
      { id: 'live', label: 'Live Data', groups: () => this.liveGroups() },
      { id: 'automate', label: 'Automate', groups: () => this.automateGroups() },
      { id: 'view', label: 'View', groups: () => this.viewGroups() },
      { id: 'help', label: 'Help', groups: () => this.helpGroups() },
    ];
  }

  private visibleTabs(): TabDef[] {
    const shown = new Set(this.app.settings.list('ribbon.tabs'));
    return this.tabs().filter((t) => t.id === 'home' || shown.has(t.id));
  }

  /** Rebuilds everything (after a settings change). */
  rebuild(): void {
    const s = this.app.settings;
    const size = this.size;
    this.el.className = `ribbon size-${size}`;
    this.el.classList.toggle('no-labels', !s.bool('ribbon.labels'));
    this.el.classList.toggle('no-group-labels', !s.bool('ribbon.groupLabels'));
    this.el.classList.toggle('large-icons', s.bool('ribbon.largeIcons'));
    this.el.classList.toggle('floating', this.floating);
    const tabs = this.visibleTabs();
    if (!tabs.some((t) => t.id === this.active)) this.active = 'home';
    this.tabsEl.innerHTML = '';
    const file = h('button', { class: 'rb-tab file', onclick: () => this.app.openHome('home') }, 'File');
    tooltip(file, 'Home: new workbooks, presets, recent files, save and export', 'Ctrl+Shift+H');
    this.tabsEl.append(file);
    for (const t of tabs) {
      const b = h('button', { class: `rb-tab${t.id === this.active ? ' active' : ''}`, role: 'tab', 'aria-selected': String(t.id === this.active) }, t.label);
      b.addEventListener('click', () => this.select(t.id, true));
      b.addEventListener('dblclick', () => this.toggleCollapsed());
      this.tabsEl.append(b);
    }
    this.tabsEl.append(h('div', { class: 'spacer' }));
    const sizeBtn = h('button', { class: 'btn rb-size-btn', 'aria-label': 'Ribbon size' }, icon(size === 'tabs' ? L.ChevronDown : L.ChevronUp, 14));
    tooltip(sizeBtn, 'Ribbon size');
    sizeBtn.addEventListener('click', () => showMenu(this.sizeMenu(), sizeBtn));
    this.tabsEl.append(sizeBtn);
    this.renderBody();
  }

  sizeMenu(): MenuItem[] {
    const cur = this.app.settings.str('ribbon.size');
    const item = (v: RibbonSize, label: string, ic: IconNode): MenuItem => ({ label, icon: ic, checked: cur === v, run: () => this.setSize(v) });
    return [
      item('full', 'Full ribbon', L.PanelTop),
      item('compact', 'Compact (one row)', L.Rows2),
      item('tabs', 'Tabs only', L.PanelTopClose),
      { separator: true },
      { label: 'Ribbon settings…', icon: L.Settings, run: () => this.app.openSettings('ribbon') },
    ];
  }

  setSize(size: RibbonSize): void {
    this.floating = false;
    this.app.settings.set('ribbon.size', size);
  }

  private toggleCollapsed(): void {
    if (this.size === 'tabs') {
      let prev: RibbonSize = 'full';
      try {
        prev = (localStorage.getItem('gridline.ribbonPrev') as RibbonSize) || 'full';
      } catch {
        // ignore
      }
      this.setSize(prev === 'tabs' ? 'full' : prev);
    } else {
      try {
        localStorage.setItem('gridline.ribbonPrev', this.size);
      } catch {
        // ignore
      }
      this.setSize('tabs');
    }
  }

  select(id: string, fromClick: boolean): void {
    if (this.size === 'tabs' && fromClick) {
      // Tabs-only: clicking the open tab closes the pop-up
      if (this.floating && this.active === id) {
        this.setFloating(false);
        return;
      }
      this.active = id;
      this.setFloating(true);
    } else this.active = id;
    try {
      localStorage.setItem(TAB_KEY, id);
    } catch {
      // ignore
    }
    for (const b of this.tabsEl.querySelectorAll('.rb-tab:not(.file)')) {
      const on = b.textContent === this.visibleTabs().find((t) => t.id === id)?.label;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', String(on));
    }
    this.renderBody();
  }

  private setFloating(on: boolean): void {
    this.floating = on;
    this.el.classList.toggle('floating', on);
    if (on) this.layout();
  }

  /** Called after a ribbon command runs. */
  private after(): void {
    if (this.floating && this.app.settings.bool('ribbon.autoClose')) this.setFloating(false);
  }

  private renderBody(): void {
    this.body.innerHTML = '';
    this.toggles = [];
    this.updaters = [];
    this.groupEls = [];
    const tab = this.tabs().find((t) => t.id === this.active) ?? this.tabs()[0];
    for (const g of tab.groups()) {
      const content = h('div', { class: 'rb-content' }, ...g.build());
      const label = h('div', { class: 'rb-label' }, g.label);
      const folded = h('button', { class: 'btn rb-folded', 'aria-label': g.label }, icon(g.icon, 20), h('span', { class: 'lbl' }, g.label), icon(L.ChevronDown, 12, 'caret'));
      folded.addEventListener('click', () => this.openFolded(group, content, folded));
      tooltip(folded, g.label);
      const group = h('div', { class: 'rb-group', 'data-group': g.id }, content, label, folded);
      this.groupEls.push(group);
      this.body.append(group);
    }
    this.refresh();
    requestAnimationFrame(() => this.layout());
  }

  // ---------- resizing ----------

  private bindResize(grip: HTMLElement): void {
    grip.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      grip.setPointerCapture(e.pointerId);
      const startY = e.clientY;
      const start = this.app.settings.str('ribbon.size') as RibbonSize;
      const order: RibbonSize[] = ['tabs', 'compact', 'full'];
      const move = (ev: PointerEvent) => {
        // Every ~32px of drag moves one size up or down
        const steps = Math.round((ev.clientY - startY) / 32);
        const idx = Math.max(0, Math.min(2, order.indexOf(start) + steps));
        const next = order[idx];
        if (next !== this.size) {
          this.dragPreview = next;
          this.floating = false;
          this.rebuild();
          this.app.grid.resize();
        }
      };
      const up = () => {
        grip.removeEventListener('pointermove', move);
        grip.removeEventListener('pointerup', up);
        const final = this.dragPreview;
        this.dragPreview = null;
        if (final && final !== start) this.app.settings.set('ribbon.size', final);
        else this.rebuild();
      };
      grip.addEventListener('pointermove', move);
      grip.addEventListener('pointerup', up);
    });
    grip.addEventListener('dblclick', () => this.toggleCollapsed());
  }

  // ---------- fitting groups into the window ----------

  private layout(): void {
    const groups = this.groupEls;
    for (const g of groups) g.classList.remove('folded');
    if (this.size === 'tabs' && !this.floating) return;
    // Fold groups from the right until everything fits
    for (let i = groups.length - 1; i >= 0 && this.body.scrollWidth > this.body.clientWidth + 1; i--) groups[i].classList.add('folded');
  }

  private openFolded(group: HTMLElement, content: HTMLElement, anchor: HTMLElement): void {
    const holder = h('div', { class: `rb-pop size-${this.size === 'compact' ? 'compact' : 'full'}` });
    const ph = document.createComment('rb');
    content.replaceWith(ph);
    holder.append(content);
    const pop = new Popover(holder, anchor);
    pop.onClose = () => {
      ph.replaceWith(content);
      if (!group.isConnected) return;
      this.layout();
    };
    pop.open(anchor);
  }

  // ---------- building blocks ----------

  private wrap(run: () => unknown): () => void {
    return () => {
      run();
      this.after();
    };
  }

  /** A command button. `big` buttons show a large icon above the label. */
  private cmd(id: string, opts: { label?: string; big?: boolean; toggle?: boolean; iconNode?: IconNode; tip?: string } = {}): HTMLButtonElement {
    const c = this.app.commands.get(id);
    if (!c) throw new Error(`Unknown command ${id}`);
    const label = opts.label ?? c.title;
    const b = this.button(opts.iconNode ?? c.icon ?? L.Square, label, () => this.app.commands.run(id), { big: opts.big, showLabel: opts.big || opts.label !== undefined });
    tooltip(b, opts.tip ?? c.title, c.keys?.[0]);
    if (opts.toggle) this.toggles.push({ el: b, checked: () => !!c.checked?.() });
    if (c.enabled) this.updaters.push(() => (b.disabled = !c.enabled!()));
    return b;
  }

  private button(ic: IconNode, label: string, run: () => unknown, opts: { big?: boolean; showLabel?: boolean; caret?: boolean } = {}): HTMLButtonElement {
    const b = h(
      'button',
      { class: `btn${opts.big ? ' rb-big' : ''}`, 'aria-label': label },
      icon(ic, opts.big ? 24 : 16),
      opts.showLabel ? h('span', { class: 'lbl' }, label) : null,
      opts.caret ? icon(L.ChevronDown, 12, 'caret') : null,
    ) as HTMLButtonElement;
    b.addEventListener('click', this.wrap(run));
    return b;
  }

  /** A button that opens a menu. */
  private menu(ic: IconNode, label: string, items: () => MenuItem[], opts: { big?: boolean; showLabel?: boolean; tip?: string } = {}): HTMLButtonElement {
    const b = h(
      'button',
      { class: `btn${opts.big ? ' rb-big' : ''}`, 'aria-label': label, 'aria-haspopup': 'menu' },
      icon(ic, opts.big ? 24 : 16),
      opts.showLabel !== false ? h('span', { class: 'lbl' }, label) : null,
      icon(L.ChevronDown, 12, 'caret'),
    ) as HTMLButtonElement;
    b.addEventListener('click', () => {
      const pop = showMenu(
        items().map((it) => ('run' in it && it.run ? { ...it, run: this.wrap(it.run) } : it)),
        b,
      );
      void pop;
    });
    tooltip(b, opts.tip ?? label);
    return b;
  }

  /** Main action plus a dropdown of related ones. */
  private split(id: string, items: () => MenuItem[], opts: { big?: boolean; label?: string } = {}): HTMLElement {
    const main = this.cmd(id, { big: opts.big, label: opts.label });
    const drop = h('button', { class: `btn split-drop${opts.big ? ' rb-big-drop' : ''}`, 'aria-label': 'More options' }, icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    drop.addEventListener('click', () => showMenu(items().map((it) => ('run' in it && it.run ? { ...it, run: this.wrap(it.run) } : it)), drop));
    tooltip(drop, 'More options');
    return h('div', { class: `split${opts.big ? ' big' : ''}` }, main, drop);
  }

  private item(id: string, label?: string): MenuItem {
    const c = this.app.commands.get(id)!;
    return { label: label ?? c.title, icon: c.icon, hint: c.keys?.[0], checked: c.checked?.(), disabled: c.enabled ? !c.enabled() : false, run: () => this.app.commands.run(id) };
  }

  /** Controls stacked in rows (one row each in the compact ribbon). */
  private stack(...rows: (HTMLElement | HTMLElement[])[]): HTMLElement {
    return h('div', { class: 'rb-stack' }, ...rows.map((r) => (Array.isArray(r) ? h('div', { class: 'rb-row' }, ...r) : r)));
  }

  private group(id: string, label: string, ic: IconNode, build: () => HTMLElement[]): GroupDef {
    return { id, label, icon: ic, build };
  }

  // ---------- rich controls (font, colors, borders, number format) ----------

  private fontPicker(): HTMLElement {
    const app = this.app;
    const btn = h('button', { class: 'btn font-pick', 'aria-label': 'Font' }, h('span', { class: 'label' }, DEFAULT_FONT), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    tooltip(btn, 'Font');
    btn.addEventListener('click', () => {
      const cur = app.styleAtActive().fontName ?? DEFAULT_FONT;
      const list = h('div', { class: 'menu', style: { maxHeight: '360px', overflow: 'auto' } });
      const pop = new Popover(list, btn);
      for (const f of FONTS) {
        const row = h('div', { class: 'menu-item' }, h('span', { class: 'mi-icon' }, f === cur ? icon(L.Check, 16) : null), h('span', { class: 'mi-label', style: { fontFamily: `"${f}"` } }, f));
        row.addEventListener('click', () => {
          pop.close();
          app.applyStyle('Font', { fontName: f === DEFAULT_FONT ? null : f });
          this.after();
          app.focusGrid();
        });
        list.append(row);
      }
      pop.open(btn);
    });
    this.updaters.push(() => {
      const f = app.styleAtActive().fontName ?? DEFAULT_FONT;
      const l = btn.firstElementChild as HTMLElement;
      l.textContent = f;
      l.style.fontFamily = `"${f}"`;
    });
    return btn;
  }

  private sizeBox(): HTMLElement {
    const app = this.app;
    const input = h('input', { value: String(DEFAULT_FONT_SIZE), 'aria-label': 'Font size' }) as HTMLInputElement;
    const applySize = () => {
      const n = parseFloat(input.value);
      if (n >= 1 && n <= 409) app.applyStyle('Font size', { fontSize: n === DEFAULT_FONT_SIZE ? null : n });
      app.focusGrid();
    };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        applySize();
      } else if (e.key === 'Escape') {
        this.refresh();
        app.focusGrid();
      }
    });
    input.addEventListener('change', applySize);
    const drop = h('button', { class: 'btn', 'aria-label': 'Font sizes' }, icon(L.ChevronDown, 12, 'caret'));
    drop.addEventListener('click', () =>
      showMenu(
        SIZES.map((s) => ({ label: String(s), checked: (app.styleAtActive().fontSize ?? DEFAULT_FONT_SIZE) === s, run: () => app.applyStyle('Font size', { fontSize: s === DEFAULT_FONT_SIZE ? null : s }) })),
        drop,
      ),
    );
    const box = h('div', { class: 'size-input' }, input, drop);
    tooltip(box, 'Font size');
    this.updaters.push(() => {
      if (document.activeElement !== input) input.value = String(app.styleAtActive().fontSize ?? DEFAULT_FONT_SIZE);
    });
    return box;
  }

  private colorSplit(kind: 'text' | 'fill'): HTMLElement {
    const app = this.app;
    const swatch = h('span', { class: 'swatch-bar' });
    const main = h('button', { class: 'btn has-swatch', 'aria-label': kind === 'text' ? 'Text color' : 'Fill color' }, icon(kind === 'text' ? L.Baseline : L.PaintBucket, 16), swatch) as HTMLButtonElement;
    const drop = h('button', { class: 'btn split-drop', 'aria-label': kind === 'text' ? 'Choose text color' : 'Choose fill color' }, icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    const apply = (c: string | null) => {
      if (kind === 'text') {
        if (c) this.textColor = c;
        app.applyStyle('Text color', { color: c });
      } else {
        if (c) this.fillColor = c;
        app.applyStyle('Fill color', { fill: c });
      }
      this.refresh();
      this.after();
      app.focusGrid();
    };
    main.addEventListener('click', () => apply(kind === 'text' ? this.textColor : this.fillColor));
    drop.addEventListener('click', () => {
      const cur = kind === 'text' ? app.styleAtActive().color : app.styleAtActive().fill;
      const pop = new Popover(
        colorPicker((c) => {
          pop.close();
          apply(c);
        }, { noneLabel: kind === 'text' ? 'Automatic' : 'No fill', current: cur }),
        drop,
      );
      pop.open(drop);
    });
    tooltip(main, kind === 'text' ? 'Text color' : 'Fill color');
    tooltip(drop, kind === 'text' ? 'Choose text color' : 'Choose fill color');
    this.updaters.push(() => (swatch.style.background = kind === 'text' ? this.textColor : this.fillColor));
    return h('div', { class: 'split' }, main, drop);
  }

  private borders(): HTMLElement {
    const app = this.app;
    const btn = h('button', { class: 'btn', 'aria-label': 'Borders' }, icon(L.Grid3x3, 16), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    tooltip(btn, 'Borders');
    btn.addEventListener('click', () => {
      const presets: [string, IconNode, string][] = [
        ['all', L.Grid3x3, 'All borders'],
        ['outer', L.Square, 'Outside borders'],
        ['thickOuter', L.SquareSquare, 'Thick outside'],
        ['inner', L.Plus, 'Inside borders'],
        ['top', L.PanelTop, 'Top'],
        ['bottom', L.PanelBottom, 'Bottom'],
        ['left', L.PanelLeft, 'Left'],
        ['right', L.PanelRight, 'Right'],
        ['horizontal', L.Rows3, 'Inner horizontal'],
        ['vertical', L.Columns3, 'Inner vertical'],
        ['bottomDouble', L.Equal, 'Double bottom'],
        ['none', L.SquareDashed, 'No borders'],
      ];
      const grid = h('div', { class: 'border-grid' });
      const pop = new Popover(h('div', null, grid), btn);
      for (const [p, ic, label] of presets) {
        const b = h('button', { class: 'btn', onclick: () => {
          pop.close();
          app.commands.run(`fmt.border.${p}`);
          this.after();
        } }, icon(ic, 18));
        tooltip(b, label);
        grid.append(b);
      }
      const colorRow = h('div', { class: 'color-actions' }, h('span', { class: 'hint' }, 'Line color'));
      const sw = h('button', { class: 'btn outline' }, h('span', { style: { width: '14px', height: '14px', borderRadius: '3px', background: this.borderColor ?? 'var(--text)', display: 'inline-block' } }), this.borderColor ? '' : 'Automatic');
      sw.addEventListener('click', () => {
        const p2 = new Popover(colorPicker((c) => {
          p2.close();
          this.borderColor = c ?? undefined;
          pop.close();
        }), sw);
        p2.open(sw, 'right', true);
      });
      colorRow.append(sw);
      pop.el.firstElementChild!.append(colorRow);
      pop.open(btn);
    });
    return btn;
  }

  private numberFormat(): HTMLElement {
    const app = this.app;
    const btn = h('button', { class: 'btn fmt-pick', 'aria-label': 'Number format' }, h('span', { class: 'label' }, 'Automatic'), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    tooltip(btn, 'Number format — shows a live preview of your value');
    btn.addEventListener('click', () => {
      const a = app.sel.active;
      const v = app.valueAt(a.row, a.col);
      const sample = typeof v === 'number' ? v : 1234.5678;
      const cur = app.styleAtActive().numFmt ?? 'General';
      const items: MenuItem[] = FORMAT_PRESETS.map((p) => ({
        label: p.label,
        hint: p.id === 'text' ? 'as typed' : formatValue(p.id.startsWith('date') || p.id.startsWith('time') || p.id === 'duration' || p.id === 'datetime' ? (typeof v === 'number' ? v : 45730.6) : sample, p.fmt).text,
        checked: cur === p.fmt,
        run: () => app.applyStyle(`Format: ${p.label}`, { numFmt: p.fmt === 'General' ? null : p.fmt }),
      }));
      items.push({ separator: true }, { label: 'More formats…', icon: L.Settings2, run: () => app.panel.open('format', 'number') });
      showMenu(items.map((it) => ('run' in it && it.run ? { ...it, run: this.wrap(it.run) } : it)), btn);
    });
    this.updaters.push(() => {
      const fmt = app.styleAtActive().numFmt ?? 'General';
      const preset = FORMAT_PRESETS.find((p) => p.fmt === fmt);
      (btn.firstElementChild as HTMLElement).textContent = preset?.label ?? (/[ymd]/i.test(fmt) ? 'Date' : 'Custom');
    });
    return btn;
  }

  private functionMenu(label: string, cats: Category[], ic: IconNode): HTMLElement {
    return this.menu(ic, label, () => {
      const defs = [...FUNCTIONS.values()].filter((d) => cats.includes(d.category)).sort((a, b) => a.name.localeCompare(b.name));
      return defs.map((d) => ({ label: d.name, hint: shortDesc(d.desc), run: () => this.app.insertFunction(d.name) }));
    }, { tip: `${label} functions` });
  }

  // ---------- tab contents ----------

  private homeGroups(): GroupDef[] {
    const app = this.app;
    return [
      this.group('undo', 'Undo', L.Undo2, () => [this.stack([this.cmd('edit.undo')], [this.cmd('edit.redo')])]),
      this.group('clipboard', 'Clipboard', L.Clipboard, () => [
        this.split('edit.paste', () => [this.item('edit.paste', 'Paste'), this.item('edit.pasteValues'), this.item('edit.pasteFormulas'), this.item('edit.pasteFormats'), this.item('edit.pasteTranspose'), this.item('edit.pasteNoBorders')], { big: true, label: 'Paste' }),
        this.stack([this.cmd('edit.cut', { label: 'Cut' })], [this.cmd('edit.copy', { label: 'Copy' })], [this.cmd('edit.formatPainter', { label: 'Painter', toggle: true })]),
      ]),
      this.group('font', 'Font', L.Type, () => [
        this.stack(
          [this.fontPicker(), this.sizeBox(), this.cmd('fmt.fontBigger'), this.cmd('fmt.fontSmaller')],
          [this.cmd('fmt.bold', { toggle: true }), this.cmd('fmt.italic', { toggle: true }), this.cmd('fmt.underline', { toggle: true }), this.cmd('fmt.strike', { toggle: true }), this.colorSplit('text'), this.colorSplit('fill'), this.borders()],
        ),
      ]),
      this.group('align', 'Alignment', L.AlignCenter, () => [
        this.stack(
          [this.cmd('fmt.alignTop', { toggle: true }), this.cmd('fmt.alignMiddle', { toggle: true }), this.cmd('fmt.alignBottom', { toggle: true }), this.cmd('fmt.wrap', { toggle: true, label: 'Wrap' })],
          [this.cmd('fmt.alignLeft', { toggle: true }), this.cmd('fmt.alignCenter', { toggle: true }), this.cmd('fmt.alignRight', { toggle: true }), this.cmd('fmt.indentDec'), this.cmd('fmt.indentInc'), this.menu(L.TableCellsMerge, 'Merge', () => [this.item('fmt.merge.center'), this.item('fmt.merge.merge'), this.item('fmt.merge.across'), { separator: true }, this.item('fmt.merge.unmerge')], { tip: 'Merge cells' })],
        ),
      ]),
      this.group('number', 'Number', L.Hash, () => [
        this.stack([this.numberFormat()], [this.cmd('fmt.currency'), this.cmd('fmt.percent'), this.cmd('fmt.thousands', { iconNode: L.Hash }), this.cmd('fmt.decDec'), this.cmd('fmt.decInc')]),
      ]),
      this.group('styles', 'Styles', L.Sparkles, () => [
        this.cmd('fmt.advanced', { big: true, label: 'Advanced format' }),
        this.stack(
          [this.menu(L.Highlighter, 'Conditional', () => [this.item('fmt.cf', 'New rule…'), this.item('fmt.clearCf')], { tip: 'Conditional formatting: color cells by their values' })],
          [this.menu(L.Brush, 'Cell styles', () => app.cellStyleMenu(), { tip: 'Ready-made styles: titles, totals, good/bad, input cells…' })],
          [this.cmd('fmt.cells', { label: 'Format cells' })],
        ),
      ]),
      this.group('cells', 'Cells', L.Rows3, () => [
        this.stack(
          [this.menu(L.BetweenHorizontalStart, 'Insert', () => [this.item('ins.rowsAbove'), this.item('ins.rowsBelow'), this.item('ins.colsLeft'), this.item('ins.colsRight'), { separator: true }, this.item('sheet.add')])],
          [this.menu(L.Trash2, 'Delete', () => [this.item('del.rows'), this.item('del.cols'), { separator: true }, this.item('sheet.delete')])],
          [this.menu(L.RulerDimensionLine, 'Size', () => [this.item('cols.autofit'), this.item('cols.width'), this.item('rows.autofit'), this.item('rows.height'), { separator: true }, this.item('rows.hide'), this.item('rows.unhide'), this.item('cols.hide'), this.item('cols.unhide')], { tip: 'Row height, column width, hide and unhide' })],
        ),
      ]),
      this.group('editing', 'Editing', L.Sigma, () => [
        this.stack(
          [this.split('ins.autosum', () => [this.item('ins.autosum', 'Sum'), this.item('ins.auto.AVERAGE', 'Average'), this.item('ins.auto.COUNT', 'Count numbers'), this.item('ins.auto.MAX', 'Max'), this.item('ins.auto.MIN', 'Min')], { label: 'AutoSum' })],
          [this.menu(L.ArrowDownToLine, 'Fill', () => [this.item('edit.fillDown'), this.item('edit.fillRight')])],
          [this.menu(L.Eraser, 'Clear', () => [this.item('edit.clearAll'), this.item('edit.clearContents'), this.item('edit.clearFormats'), this.item('fmt.clearCf'), this.item('edit.deleteNote'), this.item('edit.removeLink')])],
        ),
        this.stack(
          [this.menu(L.ArrowUpDown, 'Sort & filter', () => [this.item('data.sortAsc'), this.item('data.sortDesc'), this.item('data.sortCustom'), { separator: true }, this.item('data.filter'), this.item('data.clearFilters'), this.item('data.reapplyFilter')])],
          [this.menu(L.Search, 'Find', () => [this.item('edit.find'), this.item('edit.replace'), this.item('edit.goto'), { separator: true }, this.item('edit.selectColumn'), this.item('edit.selectRow')], { tip: 'Find, replace and go to' })],
        ),
      ]),
      ...(app.ai.available ? [this.group('ai', 'AI (Beta)', L.Sparkles, () => [this.cmd('ai.open', { big: true, label: 'Ask AI' })])] : []),
    ];
  }

  private insertGroups(): GroupDef[] {
    return [
      this.group('charts', 'Charts', L.ChartColumn, () => [
        this.cmd('ins.chart.quick', { big: true, label: 'Quick chart' }),
        this.stack(
          [this.cmd('ins.chart.column', { label: 'Column' }), this.cmd('ins.chart.bar', { label: 'Bar' }), this.cmd('ins.chart.line', { label: 'Line' }), this.cmd('ins.chart.area', { label: 'Area' })],
          [this.cmd('ins.chart.pie', { label: 'Pie' }), this.cmd('ins.chart.doughnut', { label: 'Doughnut' }), this.cmd('ins.chart.scatter', { label: 'Scatter' })],
        ),
      ]),
      this.group('functions', 'Formulas', L.SquareFunction, () => [
        this.cmd('ins.function', { big: true, label: 'Function' }),
        this.stack([this.cmd('ins.autosum', { label: 'AutoSum' })], [this.cmd('edit.insertDate', { label: "Today's date" })], [this.cmd('edit.insertTime', { label: 'Current time' })]),
      ]),
      this.group('live', 'Live data', L.CloudDownload, () => [
        this.stack([this.cmd('ins.stock', { label: 'Stock price' })], [this.cmd('ins.weather', { label: 'Weather' })], [this.cmd('nfl.panel', { label: 'NFL data' })]),
      ]),
      this.group('links', 'Links & notes', L.Link, () => [this.stack([this.cmd('edit.link', { label: 'Link' })], [this.cmd('edit.note', { label: 'Note' })])]),
      this.group('rowscols', 'Rows & columns', L.BetweenHorizontalStart, () => [
        this.stack([this.cmd('ins.rowsAbove', { label: 'Rows above' }), this.cmd('ins.rowsBelow', { label: 'Rows below' })], [this.cmd('ins.colsLeft', { label: 'Columns left' }), this.cmd('ins.colsRight', { label: 'Columns right' })]),
      ]),
      this.group('sheets', 'Sheets', L.Sheet, () => [
        this.cmd('sheet.add', { big: true, label: 'New sheet' }),
        this.stack([this.cmd('sheet.duplicate', { label: 'Duplicate sheet' })], [this.cmd('file.template', { label: 'From a preset' })]),
      ]),
    ];
  }

  private formulaGroups(): GroupDef[] {
    const half = Math.ceil(FN_MENUS.length / 2);
    const menus = FN_MENUS.map(([label, cats, ic]) => this.functionMenu(label, cats, ic));
    return [
      this.group('library', 'Function library', L.Library, () => [
        this.cmd('ins.function', { big: true, label: 'Insert function' }),
        this.split('ins.autosum', () => [this.item('ins.autosum', 'Sum'), this.item('ins.auto.AVERAGE', 'Average'), this.item('ins.auto.COUNT', 'Count numbers'), this.item('ins.auto.MAX', 'Max'), this.item('ins.auto.MIN', 'Min')], { big: true, label: 'AutoSum' }),
        this.stack(menus.slice(0, half), menus.slice(half)),
      ]),
      this.group('audit', 'Check formulas', L.Workflow, () => [
        this.stack([this.cmd('view.formulas', { label: 'Show formulas', toggle: true })], [this.cmd('data.trace', { label: 'Cells this formula uses' })]),
      ]),
      this.group('calc', 'Calculation', L.RefreshCw, () => [
        this.stack([this.cmd('data.recalc', { label: 'Recalculate' })], [this.cmd('data.refreshLive', { label: 'Refresh live data' })]),
      ]),
    ];
  }

  private dataGroups(): GroupDef[] {
    return [
      this.group('sort', 'Sort & filter', L.ArrowUpDown, () => [
        this.cmd('data.filter', { big: true, label: 'Filter', toggle: true }),
        this.stack([this.cmd('data.sortAsc', { label: 'Sort A → Z' })], [this.cmd('data.sortDesc', { label: 'Sort Z → A' })], [this.cmd('data.sortCustom', { label: 'Custom sort' })]),
        this.stack([this.cmd('data.clearFilters', { label: 'Clear filters' })], [this.cmd('data.reapplyFilter', { label: 'Reapply' })]),
      ]),
      this.group('validate', 'Validation', L.ListChecks, () => [
        this.cmd('data.validation', { big: true, label: 'Data validation' }),
        this.stack([this.cmd('data.circleInvalid', { label: 'Circle invalid data' })], [this.cmd('data.clearCircles', { label: 'Clear circles' })]),
      ]),
      this.group('tools', 'Data tools', L.Wrench, () => [
        this.cmd('fmt.advanced', { big: true, label: 'Advanced format' }),
        this.stack([this.cmd('data.dedupe', { label: 'Remove duplicates' })], [this.cmd('edit.fillDown', { label: 'Fill down' })], [this.cmd('edit.fillRight', { label: 'Fill right' })]),
      ]),
      this.group('live', 'Live data', L.CloudDownload, () => [
        this.cmd('data.refreshLive', { big: true, label: 'Refresh all' }),
        this.stack([this.cmd('nfl.panel', { label: 'NFL data' })], [this.cmd('data.clearCache', { label: 'Delete downloaded data' })], [this.cmd('app.settingsLive', { label: 'Live data settings' })]),
      ]),
    ];
  }

  private viewGroups(): GroupDef[] {
    const app = this.app;
    return [
      this.group('show', 'Show', L.Eye, () => [
        this.stack([this.cmd('view.gridlines', { label: 'Gridlines', toggle: true }), this.cmd('view.headings', { label: 'Headings', toggle: true })], [this.cmd('view.formulaBar', { label: 'Formula bar', toggle: true }), this.cmd('view.statusBar', { label: 'Status bar', toggle: true })], [this.cmd('grid.crosshair', { label: 'Row & column highlight', toggle: true })]),
      ]),
      this.group('zoom', 'Zoom', L.ZoomIn, () => [
        this.cmd('view.zoomReset', { big: true, label: '100%' }),
        this.stack([this.cmd('view.zoomIn', { label: 'Zoom in' })], [this.cmd('view.zoomOut', { label: 'Zoom out' })], [this.cmd('view.zoomSelection', { label: 'Zoom to selection' })]),
      ]),
      this.group('freeze', 'Freeze', L.Snowflake, () => [
        this.menu(L.Snowflake, 'Freeze panes', () => [this.item('view.freezeTop'), this.item('view.freezeFirstCol'), this.item('view.freezeHere', 'Freeze at selected cell'), { separator: true }, this.item('view.unfreeze')], { big: true }),
      ]),
      this.group('look', 'Appearance', L.Palette, () => [
        this.stack(
          [this.cmd('view.theme', { label: 'Light / dark' })],
          [this.cmd('view.adaptColors', { label: 'Dark-friendly colors', toggle: true })],
          [this.menu(L.Paintbrush, 'Accent color', () => app.accentMenu())],
        ),
      ]),
      this.group('window', 'Window', L.AppWindow, () => [
        this.stack([this.cmd('file.newWindow', { label: 'New window' })], [this.menu(L.PanelTop, 'Ribbon size', () => this.sizeMenu())], [this.cmd('app.settings', { label: 'Settings' })]),
      ]),
    ];
  }

  private liveGroups(): GroupDef[] {
    return [
      this.group('finance', 'Finance', L.Landmark, () => [
        this.cmd('ins.stock', { big: true, label: 'Stock price' }),
        this.stack([this.cmd('ins.stockHistory', { label: 'Price history' })], [this.cmd('ins.crypto', { label: 'Crypto' })], [this.cmd('ins.fx', { label: 'Convert currency' })]),
      ]),
      this.group('sports', 'Sports', L.Trophy, () => [
        this.cmd('nfl.panel', { big: true, label: 'NFL data' }),
        this.stack(
          [this.cmd('nfl.player', { label: 'Players' }), this.cmd('nfl.team', { label: 'Teams' })],
          [this.cmd('nfl.schedule', { label: 'Schedule' }), this.cmd('nfl.standings', { label: 'Standings' })],
          [this.cmd('nfl.leaders', { label: 'Leaders' })],
        ),
      ]),
      this.group('world', 'Weather & world', L.Globe, () => [
        this.cmd('ins.weather', { big: true, label: 'Weather' }),
        this.stack([this.cmd('ins.forecast', { label: 'Forecast' })], [this.cmd('ins.holidays', { label: 'Holidays' })], [this.cmd('ins.country', { label: 'Country data' })]),
      ]),
      this.group('refresh', 'Refresh', L.CloudDownload, () => [
        this.cmd('data.refreshLive', { big: true, label: 'Refresh all' }),
        this.stack([this.cmd('app.settingsLive', { label: 'Live data settings' })], [this.cmd('data.clearCache', { label: 'Delete downloaded data' })]),
      ]),
    ];
  }

  private automateGroups(): GroupDef[] {
    const app = this.app;
    const groups = [
      this.group('macros', 'Macros', L.SquareCode, () => [
        this.menu(L.Play, 'Run macro', () => [...app.macros.runMenu(), { separator: true }, this.item('macro.run', 'Choose a macro…')], { big: true }),
        this.cmd('macro.editor', { big: true, label: 'Macro editor' }),
        this.stack([this.cmd('macro.new', { label: 'New macro' })], [this.cmd('macro.stop', { label: 'Stop' })]),
      ]),
    ];
    if (app.macros.stepsAvailable) {
      groups.push(
        this.group('steps', 'No code (Beta)', L.ListChecks, () => [
          this.cmd('macro.newSteps', { big: true, label: 'Step macro' }),
          this.menu(L.Lightbulb, 'Examples', () => app.macroExamples(), { big: true }),
        ]),
      );
    }
    groups.push(
      this.group('triggers', 'Run it with', L.Zap, () => [
        this.cmd('macro.button', { big: true, label: 'Button' }),
        this.stack([this.cmd('macro.shortcut', { label: 'Shortcut' })], [this.cmd('macro.triggers', { label: 'Automatically' })]),
      ]),
      this.group('security', 'Security', L.ShieldCheck, () => [this.cmd('macro.settings', { big: true, label: 'Macro settings' })]),
    );
    return groups;
  }

  private helpGroups(): GroupDef[] {
    return [
      this.group('help', 'Help', L.CircleHelp, () => [
        this.cmd('view.palette', { big: true, label: 'Search commands' }),
        this.stack([this.cmd('help.shortcuts', { label: 'Keyboard shortcuts' })], [this.cmd('help.functions', { label: 'Browse functions' })]),
      ]),
      this.group('gridline', 'Gridline', L.Grid2x2, () => [
        this.cmd('app.settings', { big: true, label: 'Settings' }),
        this.stack([this.cmd('help.whatsNew', { label: "What's new" })], [this.cmd('help.updates', { label: 'Check for updates' })], [this.cmd('help.about', { label: 'About' })]),
      ]),
    ];
  }

  // ---------- state ----------

  refresh(): void {
    for (const t of this.toggles) t.el.classList.toggle('on', t.checked());
    for (const u of this.updaters) u();
  }
}

function shortDesc(d: string): string {
  const first = d.split(/(?<=[.!?])\s/)[0];
  return first.length > 60 ? first.slice(0, 57) + '…' : first;
}
