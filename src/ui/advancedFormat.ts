// Advanced format: formats the selected table in one step, then asks whether to keep it,
// try another style, or put things back the way they were.
import * as L from 'lucide';
import { currentRegion } from '../core/ops/data';
import { analyze, applyPlan, describePlan, SMART_STYLE_NAMES, SMART_STYLES, targetRange, type SmartOptions, type SmartStyle } from '../core/smartFormat';
import type { Transaction } from '../core/history';
import type { App } from './app';
import { h, icon } from './dom';
import { toast } from './widgets';

let card: HTMLElement | null = null;
let cleanup: (() => void) | null = null;

function closeCard(): void {
  cleanup?.();
  cleanup = null;
  card?.remove();
  card = null;
}

function optionsFrom(app: App, style: SmartStyle): SmartOptions {
  const s = app.settings;
  const accent = s.str('smart.color') === 'accent' ? getComputedStyle(document.documentElement).getPropertyValue('--accent').trim() : s.str('smart.color');
  return {
    style,
    color: accent,
    header: s.str('smart.header') as SmartOptions['header'],
    convertText: s.bool('smart.convertText'),
    currency: s.str('smart.currency'),
    dateFormat: s.str('smart.dateFormat'),
    negatives: s.str('smart.negatives') as SmartOptions['negatives'],
    thousands: s.bool('smart.thousands'),
    decimals: s.str('smart.decimals') as SmartOptions['decimals'],
    boldTotals: s.bool('smart.boldTotals'),
    autofit: s.bool('smart.autofit'),
    freeze: s.bool('smart.freeze'),
    filter: s.bool('smart.filter'),
  };
}

export function advancedFormat(app: App): void {
  closeCard();
  const sheet = app.sheet;
  const doc = app.doc;
  const range = targetRange(doc, sheet, app.sel.range, (r, c) => currentRegion(doc, sheet, r, c));
  const empty = doc.engine.getValue(sheet, range.r1, range.c1) === null && range.r1 === range.r2 && range.c1 === range.c2;
  if (empty) {
    toast('Select a table (or any cell inside one) first, then use Advanced format.');
    return;
  }
  const first = app.settings.str('smart.style') as SmartStyle;
  const order = [first, ...SMART_STYLES.filter((s) => s !== first)];
  let attempt = 0;
  let tx: Transaction | null = null;
  let applying = false;

  const run = (): string => {
    const opts = optionsFrom(app, order[attempt % order.length]);
    const plan = analyze(doc, sheet, range, opts);
    applying = true;
    try {
      tx = app.transact('Advanced format', (t) => applyPlan(t, doc, sheet, plan, opts));
    } finally {
      applying = false;
    }
    app.select(range, { row: range.r1, col: range.c1 });
    app.grid.relayout();
    return describePlan(plan);
  };
  const undoOurs = (): boolean => {
    if (tx && doc.history.peekUndo() === tx) {
      doc.undo();
      tx = null;
      return true;
    }
    return false;
  };

  const summary = run();
  if (!tx) return;

  const title = h('div', { class: 'af-title' });
  const text = h('div', { class: 'af-text' });
  const update = (s: string) => {
    title.textContent = `Formatted with ${SMART_STYLE_NAMES[order[attempt % order.length]]} (${(attempt % order.length) + 1} of ${order.length})`;
    text.textContent = s;
  };
  update(summary);
  const keep = h('button', { class: 'btn primary', onclick: () => {
    closeCard();
    app.focusGrid();
  } }, icon(L.Check, 14), 'Keep');
  const again = h('button', { class: 'btn outline', onclick: () => {
    undoOurs();
    attempt++;
    update(run());
  } }, icon(L.RefreshCw, 14), 'Try another style');
  const revert = h('button', { class: 'btn outline', onclick: () => {
    undoOurs();
    closeCard();
    toast('Put back the way it was.');
    app.focusGrid();
  } }, icon(L.Undo2, 14), 'Revert');
  const settings = h('button', { class: 'btn', title: 'Advanced format settings', onclick: () => app.openSettings('smart') }, icon(L.Settings, 14));
  card = h(
    'div',
    { class: 'af-card', role: 'dialog', 'aria-label': 'Advanced format' },
    h('div', { class: 'af-head' }, h('span', { class: 'af-icon' }, icon(L.Sparkles, 16)), title, settings),
    text,
    h('div', { class: 'af-actions' }, keep, again, revert),
  );
  app.grid.el.append(card);

  // Any other change to the workbook means the user has moved on: keep the formatting
  const off = doc.on((e) => {
    if (e.type === 'change' && !applying && e.tx !== tx && e.direction === 'do') closeCard();
    if (e.type === 'reset') closeCard();
  });
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape' && card) {
      closeCard();
      e.stopPropagation();
    }
  };
  document.addEventListener('keydown', onKey, true);
  cleanup = () => {
    off();
    document.removeEventListener('keydown', onKey, true);
  };
}
