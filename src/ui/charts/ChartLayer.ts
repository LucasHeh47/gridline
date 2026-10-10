// Charts floating over the grid: positioning, selection, dragging and resizing.
import * as L from 'lucide';
import { chartData } from '../../core/charts';
import type { ChartDef } from '../../core/workbook';
import type { App } from '../app';
import { h, icon } from '../dom';
import { tooltip } from '../widgets';
import { renderChart, type ChartTheme, type HitInfo } from './ChartRenderer';

interface Box {
  def: ChartDef;
  el: HTMLDivElement;
  svg: SVGSVGElement;
  renderedKey: string;
}

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;

export class ChartLayer {
  readonly el: HTMLDivElement;
  private boxes = new Map<string, Box>();
  selected: string | null = null;
  private tip: HTMLDivElement | null = null;
  private dragging = false;

  constructor(private app: App) {
    this.el = h('div', { style: { position: 'absolute', left: '0', top: '0', width: '0', height: '0', overflow: 'visible' } });
  }

  private theme(): ChartTheme {
    const cs = getComputedStyle(document.documentElement);
    const v = (n: string) => cs.getPropertyValue(n).trim();
    const dark = this.app.isDark;
    const s = this.app.settings;
    return { dark, bg: v('--panel'), text: v('--text'), muted: v('--text-muted'), grid: dark ? '#2a2f38' : '#e6e8ec', axis: dark ? '#4a5262' : '#b6bcc7', gridlines: s.bool('charts.gridlines'), piePercent: s.bool('charts.piePercent') };
  }

  /** Creates/removes chart boxes to match the active sheet. */
  sync(): void {
    const sheet = this.app.sheet;
    const ids = new Set(sheet.charts.map((c) => c.id));
    for (const [id, box] of this.boxes) {
      if (!ids.has(id)) {
        box.el.remove();
        this.boxes.delete(id);
        if (this.selected === id) this.selected = null;
      }
    }
    for (const def of sheet.charts) {
      let box = this.boxes.get(def.id);
      if (!box) {
        box = this.createBox(def);
        this.boxes.set(def.id, box);
        this.el.append(box.el);
      }
      box.def = def;
    }
    this.position();
    this.renderAll(false);
  }

  private createBox(def: ChartDef): Box {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg') as SVGSVGElement;
    const editBtn = h('button', { class: 'btn', onclick: (e: MouseEvent) => {
      e.stopPropagation();
      this.select(def.id);
      this.app.panel.open('chart');
    } }, icon(L.Pencil, 14));
    tooltip(editBtn, 'Edit chart');
    const delBtn = h('button', { class: 'btn danger', onclick: (e: MouseEvent) => {
      e.stopPropagation();
      this.select(def.id);
      this.deleteSelected();
    } }, icon(L.Trash2, 14));
    tooltip(delBtn, 'Delete chart');
    const el = h('div', { class: 'chart-box', tabindex: '-1' }, svg as unknown as HTMLElement, h('div', { class: 'chart-tools' }, editBtn, delBtn));
    for (const hd of HANDLES) {
      const handle = h('div', { class: `handle h-${hd}`, style: handleStyle(hd) });
      handle.addEventListener('pointerdown', (e) => this.startDrag(e, def.id, hd));
      el.append(handle);
    }
    el.addEventListener('pointerdown', (e) => {
      if ((e.target as HTMLElement).closest('.chart-tools')) return;
      this.startDrag(e, def.id, 'move');
    });
    el.addEventListener('dblclick', () => {
      this.select(def.id);
      this.app.panel.open('chart');
    });
    el.addEventListener('mousemove', (e) => this.onHover(e));
    el.addEventListener('mouseleave', () => this.hideTip());
    el.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.select(def.id);
      this.app.panel.open('chart');
    });
    return { def, el, svg, renderedKey: '' };
  }

  position(): void {
    const g = this.app.grid;
    if (!g.rows) return;
    const z = g.zoom;
    // Clip charts to the grid area below the headers / frozen panes
    const clipX = g.headerW + g.frozenW;
    const clipY = g.headerH + g.frozenH;
    for (const box of this.boxes.values()) {
      const a = box.def.anchor;
      const x = g.colX(a.col) + a.dx * z;
      const y = g.rowY(a.row) + a.dy * z;
      const w = a.w * z;
      const hgt = a.h * z;
      Object.assign(box.el.style, { left: `${x}px`, top: `${y}px`, width: `${w}px`, height: `${hgt}px` });
      // Charts anchored inside frozen panes stay fixed and aren't clipped
      const sheet = this.app.sheet;
      const cl = a.col >= sheet.freeze.cols ? Math.max(0, clipX - x) : 0;
      const ct = a.row >= sheet.freeze.rows ? Math.max(0, clipY - y) : 0;
      box.el.style.clipPath = cl > 0 || ct > 0 ? `inset(${ct}px 0 0 ${cl}px)` : '';
      box.el.style.display = (cl && x + w < clipX) || (ct && y + hgt < clipY) || x > g.width || y > g.height ? 'none' : '';
    }
  }

  renderAll(force = true): void {
    const theme = this.theme();
    const version = this.app.doc.engine.version;
    for (const box of this.boxes.values()) {
      const a = box.def.anchor;
      const key = `${version}|${theme.dark}|${theme.gridlines}|${theme.piePercent}|${this.app.settings.str('charts.palette')}|${a.w}x${a.h}|${JSON.stringify({ ...box.def, anchor: null })}`;
      if (!force && key === box.renderedKey) continue;
      box.renderedKey = key;
      try {
        renderChart(box.svg, box.def, chartData(this.app.doc, box.def, this.app.sheet), a.w, a.h, theme);
      } catch (err) {
        console.error('chart render failed', err);
      }
    }
  }

  select(id: string): void {
    this.selected = id;
    for (const [bid, box] of this.boxes) box.el.classList.toggle('selected', bid === id);
    this.app.grid.focusTarget.focus({ preventScroll: true });
  }

  deselect(): void {
    if (!this.selected || this.dragging) return;
    this.selected = null;
    for (const box of this.boxes.values()) box.el.classList.remove('selected');
    if (this.app.panel.kind === 'chart') this.app.panel.close();
  }

  selectedDef(): ChartDef | null {
    return this.app.sheet.charts.find((c) => c.id === this.selected) ?? null;
  }

  deleteSelected(): void {
    const id = this.selected;
    if (!id) return;
    const s = this.app.sheet;
    this.selected = null;
    this.app.transact('Delete chart', (tx) => tx.setProp(s, 'charts', s.charts.filter((c) => c.id !== id)));
    if (this.app.panel.kind === 'chart') this.app.panel.close();
  }

  nudge(dx: number, dy: number): void {
    const def = this.selectedDef();
    if (!def) return;
    this.updateAnchor(def, { ...def.anchor, dx: def.anchor.dx + dx, dy: def.anchor.dy + dy }, true);
  }

  /** Writes a chart change as an undoable action. */
  update(def: ChartDef, patch: Partial<ChartDef>, label = 'Edit chart'): void {
    const s = this.app.sheet;
    const next = { ...def, ...patch };
    this.app.transact(label, (tx) => tx.setProp(s, 'charts', s.charts.map((c) => (c.id === def.id ? next : c))));
  }

  private updateAnchor(def: ChartDef, anchor: ChartDef['anchor'], normalize: boolean): void {
    const g = this.app.grid;
    let a = { ...anchor };
    if (normalize) {
      // Re-anchor to the cell under the top-left corner
      const z = g.zoom;
      const x = g.colX(def.anchor.col) + a.dx * z;
      const y = g.rowY(def.anchor.row) + a.dy * z;
      const col = Math.max(0, g.colAt(Math.max(g.headerW, x)));
      const row = Math.max(0, g.rowAt(Math.max(g.headerH, y)));
      a = { ...a, col, row, dx: Math.max(0, Math.round((x - g.colX(col)) / z)), dy: Math.max(0, Math.round((y - g.rowY(row)) / z)) };
    }
    this.update(def, { anchor: a }, 'Move chart');
  }

  private startDrag(e: PointerEvent, id: string, mode: 'move' | (typeof HANDLES)[number]): void {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    if (this.app.isEditing && !this.app.commitEdit()) return;
    this.select(id);
    const box = this.boxes.get(id)!;
    const def = box.def;
    const z = this.app.grid.zoom;
    const start = { x: e.clientX, y: e.clientY };
    const orig = { left: parseFloat(box.el.style.left), top: parseFloat(box.el.style.top), w: def.anchor.w * z, h: def.anchor.h * z };
    let moved = false;
    this.dragging = true;
    const target = e.currentTarget as HTMLElement;
    target.setPointerCapture(e.pointerId);
    const onMove = (ev: PointerEvent) => {
      const dx = ev.clientX - start.x;
      const dy = ev.clientY - start.y;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
      moved = true;
      let { left, top, w, h: hh } = orig;
      if (mode === 'move') {
        left += dx;
        top += dy;
      } else {
        if (mode.includes('e')) w = Math.max(160 * z, orig.w + dx);
        if (mode.includes('s')) hh = Math.max(110 * z, orig.h + dy);
        if (mode.includes('w')) {
          w = Math.max(160 * z, orig.w - dx);
          left = orig.left + orig.w - w;
        }
        if (mode.includes('n')) {
          hh = Math.max(110 * z, orig.h - dy);
          top = orig.top + orig.h - hh;
        }
      }
      Object.assign(box.el.style, { left: `${left}px`, top: `${top}px`, width: `${w}px`, height: `${hh}px`, clipPath: '' });
    };
    const onUp = () => {
      target.removeEventListener('pointermove', onMove);
      target.removeEventListener('pointerup', onUp);
      this.dragging = false;
      if (!moved) return;
      const g = this.app.grid;
      const left = parseFloat(box.el.style.left);
      const top = parseFloat(box.el.style.top);
      const w = parseFloat(box.el.style.width) / z;
      const hh = parseFloat(box.el.style.height) / z;
      const col = Math.max(0, g.colAt(Math.max(g.headerW + 1, left)));
      const row = Math.max(0, g.rowAt(Math.max(g.headerH + 1, top)));
      const anchor = { col, row, dx: Math.max(0, Math.round((left - g.colX(col)) / z)), dy: Math.max(0, Math.round((top - g.rowY(row)) / z)), w: Math.round(w), h: Math.round(hh) };
      this.update(def, { anchor }, mode === 'move' ? 'Move chart' : 'Resize chart');
    };
    target.addEventListener('pointermove', onMove);
    target.addEventListener('pointerup', onUp);
  }

  private onHover(e: MouseEvent): void {
    const target = e.target as Element;
    const data = target.getAttribute?.('data-hit');
    if (!data) {
      this.hideTip();
      return;
    }
    const info = JSON.parse(data) as HitInfo;
    if (!this.tip) {
      this.tip = h('div', { class: 'chart-tip' });
      document.body.append(this.tip);
    }
    this.tip.innerHTML = '';
    this.tip.append(
      h('div', null, h('span', { class: 'dot', style: { background: info.color } }), h('b', null, info.series)),
      h('div', { style: { color: 'var(--text-muted)' } }, `${info.category}: `, h('span', { style: { color: 'var(--text)' } }, info.value)),
    );
    this.tip.style.left = `${Math.min(window.innerWidth - this.tip.offsetWidth - 8, e.clientX + 12)}px`;
    this.tip.style.top = `${e.clientY + 12}px`;
  }

  private hideTip(): void {
    this.tip?.remove();
    this.tip = null;
  }
}

function handleStyle(hd: string): Partial<CSSStyleDeclaration> {
  const s: Partial<CSSStyleDeclaration> = {};
  const mid = 'calc(50% - 5px)';
  if (hd.includes('n')) s.top = '-5px';
  if (hd.includes('s')) s.bottom = '-5px';
  if (hd.includes('w')) s.left = '-5px';
  if (hd.includes('e')) s.right = '-5px';
  if (hd === 'n' || hd === 's') s.left = mid;
  if (hd === 'e' || hd === 'w') s.top = mid;
  s.cursor = hd === 'n' || hd === 's' ? 'ns-resize' : hd === 'e' || hd === 'w' ? 'ew-resize' : hd === 'nw' || hd === 'se' ? 'nwse-resize' : 'nesw-resize';
  return s;
}
