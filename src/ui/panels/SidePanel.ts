// A docked side panel that hosts formatting, conditional formatting, sorting, charts and functions.
// Panels apply changes live — no "OK/Cancel" dialogs that hide the sheet.
import * as L from 'lucide';
import type { App } from '../app';
import { h, icon } from '../dom';
import { CFPanel } from './CFPanel';
import { ChartPanel } from './ChartPanel';
import { FormatPanel } from './FormatPanel';
import { FunctionsPanel } from './FunctionsPanel';
import { NFLPanel } from './NFLPanel';
import { ValidationPanel } from './ValidationPanel';
import { SortPanel } from './SortPanel';

export type PanelKind = 'format' | 'cf' | 'chart' | 'sort' | 'functions' | 'nfl' | 'validation';

export interface PanelContent {
  title: string;
  el: HTMLElement;
  refresh?(): void;
  /** Called when the active cell moves (cheap updates only). */
  selectionChanged?(): void;
  dispose?(): void;
}

export class SidePanel {
  readonly el: HTMLElement;
  private titleEl: HTMLHeadingElement;
  private body: HTMLDivElement;
  kind: PanelKind | null = null;
  private content: PanelContent | null = null;

  constructor(private app: App) {
    this.titleEl = h('h2', null, '');
    this.body = h('div', { style: { display: 'flex', flexDirection: 'column', flex: '1', minHeight: '0' } });
    const close = h('button', { class: 'btn', onclick: () => this.close(), 'aria-label': 'Close panel' }, icon(L.X, 16));
    this.el = h('aside', { class: 'side-panel', hidden: true }, h('div', { class: 'sp-head' }, this.titleEl, close), this.body);
    this.el.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        this.close();
      }
    });
  }

  open(kind: PanelKind, sub?: string): void {
    if (this.kind === kind && this.content) {
      if (sub && 'setTab' in this.content) (this.content as unknown as { setTab(t: string): void }).setTab(sub);
      this.content.refresh?.();
      return;
    }
    this.content?.dispose?.();
    this.kind = kind;
    const app = this.app;
    switch (kind) {
      case 'format':
        this.content = new FormatPanel(app, sub);
        break;
      case 'cf':
        this.content = new CFPanel(app);
        break;
      case 'chart':
        this.content = new ChartPanel(app);
        break;
      case 'sort':
        this.content = new SortPanel(app);
        break;
      case 'functions':
        this.content = new FunctionsPanel(app);
        break;
      case 'validation':
        this.content = new ValidationPanel(app);
        break;
      case 'nfl':
        this.content = new NFLPanel(app, sub);
        break;
    }
    this.titleEl.textContent = this.content.title;
    this.body.innerHTML = '';
    this.body.append(this.content.el);
    this.el.hidden = false;
    app.grid.resize();
  }

  close(): void {
    if (!this.kind) return;
    this.content?.dispose?.();
    this.content = null;
    this.kind = null;
    this.el.hidden = true;
    this.app.grid.resize();
    this.app.focusGrid();
  }

  selectionChanged(): void {
    this.content?.selectionChanged?.();
  }

  refresh(): void {
    if (this.content) {
      this.content.refresh?.();
      this.titleEl.textContent = this.content.title;
    }
  }
}
