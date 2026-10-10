// Bottom-right status: mode, messages, selection statistics and zoom.
import * as L from 'lucide';
import { external } from '../core/data/external';
import { formatValue } from '../core/numfmt';
import { RangeRef } from '../core/formula/types';
import type { App } from './app';
import { h, icon } from './dom';
import { platform } from './platform';
import { tooltip } from './widgets';

export class StatusBar {
  readonly el: HTMLElement;
  private mode: HTMLSpanElement;
  private msg: HTMLSpanElement;
  private stats: HTMLDivElement;
  private zoomRange: HTMLInputElement;
  private zoomPct: HTMLSpanElement;
  private msgTimer: ReturnType<typeof setTimeout> | undefined;

  constructor(private app: App) {
    this.mode = h('span', { class: 'mode-pill' }, 'Ready');
    this.msg = h('span', { class: 'status-msg', style: { display: 'flex', alignItems: 'center', gap: '6px' } });
    this.stats = h('div', { class: 'status' });
    this.zoomRange = h('input', { type: 'range', min: '25', max: '300', step: '5', value: '100', 'aria-label': 'Zoom' }) as HTMLInputElement;
    this.zoomRange.addEventListener('input', () => app.setZoom(+this.zoomRange.value / 100));
    this.zoomPct = h('span', { class: 'pct', title: 'Reset zoom to 100%', onclick: () => app.setZoom(1) }, '100%');
    const zoomOut = h('button', { class: 'btn', onclick: () => app.zoomBy(-0.1) }, icon(L.Minus, 14));
    const zoomIn = h('button', { class: 'btn', onclick: () => app.zoomBy(0.1) }, icon(L.Plus, 14));
    tooltip(zoomOut, 'Zoom out', 'Ctrl+scroll');
    tooltip(zoomIn, 'Zoom in', 'Ctrl+scroll');
    this.loading = h('span', { class: 'loading-pill', hidden: true }, h('span', { class: 'spinner' }), 'Loading live data…');
    this.el = h('div', { class: 'status' }, this.msg, this.loading, this.mode, this.stats, h('div', { class: 'zoom' }, zoomOut, this.zoomRange, zoomIn, this.zoomPct));
    external.onActivity((n) => {
      this.loading.hidden = n === 0;
      (this.loading.lastChild as Text).textContent = n > 1 ? `Loading live data (${n})…` : 'Loading live data…';
    });
  }

  private loading: HTMLSpanElement;

  setMode(m: 'Ready' | 'Enter' | 'Edit' | 'Point'): void {
    this.mode.textContent = m;
    this.mode.classList.toggle('edit', m !== 'Ready');
  }

  /** Shows a short message (optionally with an action button). */
  flash(text: string, action?: { label: string; run: () => void }): void {
    clearTimeout(this.msgTimer);
    this.msg.innerHTML = '';
    this.msg.append(h('span', null, text));
    if (action) {
      this.msg.append(h('button', { class: 'btn text', style: { height: '22px', color: 'var(--accent)' }, onclick: () => {
        action.run();
        this.msg.innerHTML = '';
      } }, action.label));
    }
    this.msgTimer = setTimeout(() => (this.msg.innerHTML = ''), action ? 9000 : 3500);
  }

  refresh(): void {
    const app = this.app;
    const z = Math.round(app.sheet.view.zoom * 100);
    this.zoomRange.value = String(z);
    this.zoomPct.textContent = `${z}%`;
    this.stats.innerHTML = '';
    const sel = app.sel;
    if (sel.isSingleCell) return;
    let count = 0;
    let nums = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    let visited = 0;
    const engine = app.doc.engine;
    for (const r of sel.ranges) {
      engine.forEachInRange(new RangeRef(app.sheet, r.r1, r.c1, r.r2, r.c2), (v, row) => {
        if (app.sheet.isRowHidden(row)) return;
        if (++visited > 2_000_000) return false;
        count++;
        if (typeof v === 'number') {
          nums++;
          sum += v;
          if (v < min) min = v;
          if (v > max) max = v;
        }
      });
    }
    const fmt = app.styleAtActive().numFmt;
    const plain = new Intl.NumberFormat(undefined, { maximumFractionDigits: 4 });
    const show = (n: number) => (fmt && fmt !== 'General' && fmt !== '@' ? formatValue(Number(n.toPrecision(15)), fmt).text : plain.format(Number(n.toPrecision(15))));
    const stat = (label: string, value: string, raw: string) => {
      const el = h('span', { class: 'stat', title: 'Click to copy' }, `${label}: `, h('b', null, value));
      el.addEventListener('click', () => {
        void platform.writeClipboard({ text: raw });
        this.flash(`Copied ${label.toLowerCase()} (${value})`);
      });
      this.stats.append(el);
    };
    if (nums > 0) {
      stat('Sum', show(sum), String(sum));
      stat('Average', show(sum / nums), String(sum / nums));
      if (nums > 1) {
        stat('Min', show(min), String(min));
        stat('Max', show(max), String(max));
      }
    }
    if (count > 0) stat('Count', count.toLocaleString(), String(count));
  }
}
