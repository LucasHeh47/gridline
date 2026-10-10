// A single, compact toolbar (instead of a ribbon). Groups that don't fit move into a "more" menu.
import * as L from 'lucide';
import type { IconNode } from 'lucide';
import { FORMAT_PRESETS, formatValue } from '../core/numfmt';
import { DEFAULT_FONT, DEFAULT_FONT_SIZE } from '../core/styles';
import type { App } from './app';
import { clear, h, icon } from './dom';
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

export class Toolbar {
  readonly el: HTMLElement;
  private toggles: { id: string; el: HTMLElement }[] = [];
  private fontBtn!: HTMLButtonElement;
  private sizeInput!: HTMLInputElement;
  private fmtBtn!: HTMLButtonElement;
  private textSwatch!: HTMLElement;
  private fillSwatch!: HTMLElement;
  private undoBtn!: HTMLButtonElement;
  private redoBtn!: HTMLButtonElement;
  private groups: HTMLElement[] = [];
  private moreBtn: HTMLButtonElement;
  textColor = '#ff6b6b';
  fillColor = '#5a4c12';
  borderColor: string | undefined = undefined;

  constructor(private app: App) {
    this.el = h('div', { class: 'toolbar', role: 'toolbar' });
    this.moreBtn = h('button', { class: 'btn', onclick: () => this.openOverflow() }, icon(L.Ellipsis, 16)) as HTMLButtonElement;
    tooltip(this.moreBtn, 'More tools');
    this.build();
    new ResizeObserver(() => this.layout()).observe(this.el);
  }

  private cmdBtn(id: string, opts: { label?: string; toggle?: boolean; iconNode?: IconNode } = {}): HTMLButtonElement {
    const c = this.app.commands.get(id)!;
    const b = h('button', { class: 'btn', 'aria-label': c.title, onclick: () => this.app.commands.run(id) }, icon(opts.iconNode ?? c.icon ?? L.Square, 16), opts.label ?? null) as HTMLButtonElement;
    tooltip(b, c.title, c.keys?.[0]);
    if (opts.toggle) this.toggles.push({ id, el: b });
    return b;
  }

  private menuBtn(iconNode: IconNode, title: string, items: () => MenuItem[], label?: string): HTMLButtonElement {
    const b = h('button', { class: 'btn' }, icon(iconNode, 16), label ?? null, icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    b.addEventListener('click', () => showMenu(items(), b));
    tooltip(b, title);
    return b;
  }

  private item(id: string, label?: string): MenuItem {
    const c = this.app.commands.get(id)!;
    return { label: label ?? c.title, icon: c.icon, hint: c.keys?.[0], checked: c.checked?.(), disabled: c.enabled ? !c.enabled() : false, run: () => this.app.commands.run(id) };
  }

  private group(...children: (HTMLElement | null)[]): HTMLElement {
    const g = h('div', { class: 'tb-group' }, ...children);
    this.groups.push(g);
    return g;
  }

  private build(): void {
    const app = this.app;
    this.undoBtn = this.cmdBtn('edit.undo');
    this.redoBtn = this.cmdBtn('edit.redo');
    const painter = this.cmdBtn('edit.formatPainter', { toggle: true });

    // Font family
    this.fontBtn = h('button', { class: 'btn font-pick' }, h('span', { class: 'label' }, DEFAULT_FONT), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    tooltip(this.fontBtn, 'Font');
    this.fontBtn.addEventListener('click', () => {
      const cur = app.styleAtActive().fontName ?? DEFAULT_FONT;
      const list = h('div', { class: 'menu', style: { maxHeight: '360px', overflow: 'auto' } });
      const pop = new Popover(list, this.fontBtn);
      for (const f of FONTS) {
        const row = h('div', { class: 'menu-item' }, h('span', { class: 'mi-icon' }, f === cur ? icon(L.Check, 16) : null), h('span', { class: 'mi-label', style: { fontFamily: `"${f}"` } }, f));
        row.addEventListener('click', () => {
          pop.close();
          app.applyStyle('Font', { fontName: f === DEFAULT_FONT ? null : f });
          app.focusGrid();
        });
        list.append(row);
      }
      pop.open(this.fontBtn);
    });

    // Font size
    this.sizeInput = h('input', { value: String(DEFAULT_FONT_SIZE), 'aria-label': 'Font size' }) as HTMLInputElement;
    const applySize = () => {
      const n = parseFloat(this.sizeInput.value);
      if (n >= 1 && n <= 409) app.applyStyle('Font size', { fontSize: n === DEFAULT_FONT_SIZE ? null : n });
      app.focusGrid();
    };
    this.sizeInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        applySize();
      } else if (e.key === 'Escape') {
        this.refresh();
        app.focusGrid();
      }
    });
    this.sizeInput.addEventListener('change', applySize);
    const sizeDrop = h('button', { class: 'btn' }, icon(L.ChevronDown, 12, 'caret'));
    sizeDrop.addEventListener('click', () =>
      showMenu(
        SIZES.map((s) => ({ label: String(s), checked: (app.styleAtActive().fontSize ?? 11) === s, run: () => app.applyStyle('Font size', { fontSize: s === 11 ? null : s }) })),
        sizeDrop,
      ),
    );
    const sizeBox = h('div', { class: 'size-input' }, this.sizeInput, sizeDrop);
    tooltip(sizeBox, 'Font size');

    // Colors
    const colorSplit = (kind: 'text' | 'fill') => {
      const swatch = h('span', { class: 'swatch-bar' });
      const main = h('button', { class: 'btn has-swatch' }, icon(kind === 'text' ? L.Baseline : L.PaintBucket, 16), swatch) as HTMLButtonElement;
      const drop = h('button', { class: 'btn' }, icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
      const apply = (c: string | null) => {
        if (kind === 'text') {
          if (c) this.textColor = c;
          app.applyStyle('Text color', { color: c });
        } else {
          if (c) this.fillColor = c;
          app.applyStyle('Fill color', { fill: c });
        }
        this.refresh();
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
      if (kind === 'text') this.textSwatch = swatch;
      else this.fillSwatch = swatch;
      return h('div', { class: 'split' }, main, drop);
    };

    // Borders
    const borderBtn = h('button', { class: 'btn' }, icon(L.Grid3x3, 16), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    tooltip(borderBtn, 'Borders');
    borderBtn.addEventListener('click', () => {
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
      const pop = new Popover(h('div', null, grid), borderBtn);
      for (const [p, ic, label] of presets) {
        const b = h('button', { class: 'btn', onclick: () => {
          pop.close();
          app.commands.run(`fmt.border.${p}`);
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
      pop.open(borderBtn);
    });

    // Alignment
    const alignV = this.menuBtn(L.AlignVerticalJustifyCenter, 'Vertical alignment', () => [this.item('fmt.alignTop', 'Top'), this.item('fmt.alignMiddle', 'Middle'), this.item('fmt.alignBottom', 'Bottom')]);
    const merge = this.menuBtn(L.TableCellsMerge, 'Merge cells', () => [
      this.item('fmt.merge.center'),
      this.item('fmt.merge.merge'),
      this.item('fmt.merge.across'),
      { separator: true },
      this.item('fmt.merge.unmerge'),
    ]);

    // Number format
    this.fmtBtn = h('button', { class: 'btn fmt-pick' }, h('span', { class: 'label' }, 'Automatic'), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    tooltip(this.fmtBtn, 'Number format — shows a live preview of your value');
    this.fmtBtn.addEventListener('click', () => this.openFormatMenu());

    // Data
    const sort = this.menuBtn(L.ArrowUpDown, 'Sort', () => [this.item('data.sortAsc'), this.item('data.sortDesc'), { separator: true }, this.item('data.sortCustom'), this.item('data.dedupe')]);
    const freeze = this.menuBtn(L.Snowflake, 'Freeze panes', () => [this.item('view.freezeTop'), this.item('view.freezeFirstCol'), this.item('view.freezeHere', 'Freeze at selected cell'), { separator: true }, this.item('view.unfreeze')]);
    const chart = this.menuBtn(L.ChartColumn, 'Insert chart', () => [
      this.item('ins.chart.column'),
      this.item('ins.chart.bar'),
      this.item('ins.chart.line'),
      this.item('ins.chart.area'),
      this.item('ins.chart.pie'),
      this.item('ins.chart.doughnut'),
      this.item('ins.chart.scatter'),
    ]);
    const sumMain = this.cmdBtn('ins.autosum');
    const sumDrop = h('button', { class: 'btn' }, icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    sumDrop.addEventListener('click', () => showMenu([this.item('ins.autosum', 'Sum'), this.item('ins.auto.AVERAGE', 'Average'), this.item('ins.auto.COUNT', 'Count numbers'), this.item('ins.auto.MAX', 'Max'), this.item('ins.auto.MIN', 'Min')], sumDrop));
    tooltip(sumDrop, 'More quick calculations');

    const sep = () => h('div', { class: 'tb-sep' });
    this.el.append(
      this.group(this.undoBtn, this.redoBtn, painter),
      sep(),
      this.group(this.fontBtn, sizeBox),
      sep(),
      this.group(this.cmdBtn('fmt.bold', { toggle: true }), this.cmdBtn('fmt.italic', { toggle: true }), this.cmdBtn('fmt.underline', { toggle: true }), this.cmdBtn('fmt.strike', { toggle: true })),
      sep(),
      this.group(colorSplit('text'), colorSplit('fill'), borderBtn),
      sep(),
      this.group(this.cmdBtn('fmt.alignLeft', { toggle: true }), this.cmdBtn('fmt.alignCenter', { toggle: true }), this.cmdBtn('fmt.alignRight', { toggle: true }), alignV, this.cmdBtn('fmt.wrap', { toggle: true }), merge),
      sep(),
      this.group(this.fmtBtn, this.cmdBtn('fmt.currency'), this.cmdBtn('fmt.percent'), this.cmdBtn('fmt.decDec'), this.cmdBtn('fmt.decInc')),
      sep(),
      this.group(sort, this.cmdBtn('data.filter', { toggle: true }), this.cmdBtn('fmt.cf'), freeze),
      sep(),
      this.group(chart, h('div', { class: 'split' }, sumMain, sumDrop), this.cmdBtn('ins.function'), this.cmdBtn('nfl.panel', { label: 'NFL' })),
      this.moreBtn,
    );
  }

  private openFormatMenu(): void {
    const app = this.app;
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
    showMenu(items, this.fmtBtn);
  }

  refresh(): void {
    const app = this.app;
    const st = app.styleAtActive();
    for (const t of this.toggles) {
      const c = app.commands.get(t.id);
      t.el.classList.toggle('on', !!c?.checked?.());
    }
    (this.fontBtn.firstElementChild as HTMLElement).textContent = st.fontName ?? DEFAULT_FONT;
    (this.fontBtn.firstElementChild as HTMLElement).style.fontFamily = `"${st.fontName ?? DEFAULT_FONT}"`;
    if (document.activeElement !== this.sizeInput) this.sizeInput.value = String(st.fontSize ?? DEFAULT_FONT_SIZE);
    const fmt = st.numFmt ?? 'General';
    const preset = FORMAT_PRESETS.find((p) => p.fmt === fmt);
    (this.fmtBtn.firstElementChild as HTMLElement).textContent = preset?.label ?? (/[ymd]/i.test(fmt) ? 'Date' : 'Custom');
    this.textSwatch.style.background = this.textColor;
    this.fillSwatch.style.background = this.fillColor;
    this.undoBtn.disabled = !app.doc.history.canUndo;
    this.redoBtn.disabled = !app.doc.history.canRedo;
    const ul = app.doc.history.peekUndoLabel();
    this.undoBtn.title = '';
    this.undoBtn.setAttribute('aria-label', ul ? `Undo ${ul}` : 'Undo');
  }

  /** Moves groups that don't fit into the overflow menu. */
  private layout(): void {
    for (const g of this.groups) {
      g.style.display = '';
      const prev = g.previousElementSibling as HTMLElement | null;
      if (prev?.classList.contains('tb-sep')) prev.style.display = '';
    }
    this.moreBtn.style.display = 'none';
    if (this.el.scrollWidth <= this.el.clientWidth) return;
    this.moreBtn.style.display = '';
    for (let i = this.groups.length - 1; i > 0 && this.el.scrollWidth > this.el.clientWidth; i--) {
      const g = this.groups[i];
      g.style.display = 'none';
      const prev = g.previousElementSibling as HTMLElement | null;
      if (prev?.classList.contains('tb-sep')) prev.style.display = 'none';
    }
  }

  private openOverflow(): void {
    const hidden = this.groups.filter((g) => g.style.display === 'none');
    const holder = h('div', { class: 'tb-overflow-holder' });
    const pop = new Popover(holder, this.moreBtn);
    const placeholders: [HTMLElement, Comment][] = [];
    for (const g of hidden) {
      const ph = document.createComment('tb');
      g.replaceWith(ph);
      g.style.display = '';
      holder.append(g);
      placeholders.push([g, ph]);
    }
    pop.onClose = () => {
      for (const [g, ph] of placeholders) {
        ph.replaceWith(g);
        g.style.display = 'none';
      }
      this.layout();
    };
    pop.open(this.moreBtn);
  }

  clearAll(): void {
    clear(this.el);
  }
}
