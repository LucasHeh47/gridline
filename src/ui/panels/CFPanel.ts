// Conditional formatting: a visual rule builder plus a list of existing rules.
import * as L from 'lucide';
import { formatRange, parseRange, rangesIntersect, type Range } from '../../core/address';
import { describeRule } from '../../core/cf';
import type { CFOperator, CFRule, CFRuleSpec, CFStyle } from '../../core/workbook';
import type { App } from '../app';
import { h, icon } from '../dom';
import { toast } from '../widgets';
import type { PanelContent } from './SidePanel';

const STYLES: { label: string; style: CFStyle }[] = [
  { label: 'Red', style: { fill: '#5c2b2b', color: '#ffb3b3' } },
  { label: 'Yellow', style: { fill: '#5a4c12', color: '#ffe08a' } },
  { label: 'Green', style: { fill: '#174a37', color: '#9ff0c9' } },
  { label: 'Blue', style: { fill: '#173e66', color: '#a8d1ff' } },
  { label: 'Red text', style: { color: '#ff6b6b' } },
  { label: 'Green text', style: { color: '#3fcf8e' } },
  { label: 'Bold', style: { bold: true } },
  { label: 'Strike', style: { strike: true, color: '#8f97a6' } },
];

type Builder = { kind: string; label: string; inputs: string[]; make: (vals: string[], style: CFStyle) => CFRuleSpec };

const BUILDERS: Builder[] = [
  { kind: 'gt', label: 'Greater than', inputs: ['Value'], make: (v, s) => cellIs('greaterThan', v, s) },
  { kind: 'lt', label: 'Less than', inputs: ['Value'], make: (v, s) => cellIs('lessThan', v, s) },
  { kind: 'between', label: 'Between', inputs: ['From', 'To'], make: (v, s) => cellIs('between', v, s) },
  { kind: 'eq', label: 'Equal to', inputs: ['Value'], make: (v, s) => cellIs('equal', v, s) },
  { kind: 'contains', label: 'Text contains', inputs: ['Text'], make: (v, s) => ({ type: 'text', operator: 'contains', text: v[0], style: s }) },
  { kind: 'dup', label: 'Duplicate values', inputs: [], make: (_v, s) => ({ type: 'duplicate', style: s }) },
  { kind: 'unique', label: 'Unique values', inputs: [], make: (_v, s) => ({ type: 'unique', style: s }) },
  { kind: 'top', label: 'Top N', inputs: ['How many'], make: (v, s) => ({ type: 'top', rank: Math.max(1, parseInt(v[0], 10) || 10), percent: false, bottom: false, style: s }) },
  { kind: 'bottom', label: 'Bottom N', inputs: ['How many'], make: (v, s) => ({ type: 'top', rank: Math.max(1, parseInt(v[0], 10) || 10), percent: false, bottom: true, style: s }) },
  { kind: 'above', label: 'Above average', inputs: [], make: (_v, s) => ({ type: 'average', below: false, style: s }) },
  { kind: 'below', label: 'Below average', inputs: [], make: (_v, s) => ({ type: 'average', below: true, style: s }) },
  { kind: 'blanks', label: 'Blank cells', inputs: [], make: (_v, s) => ({ type: 'blanks', style: s }) },
  { kind: 'errors', label: 'Errors', inputs: [], make: (_v, s) => ({ type: 'errors', style: s }) },
  { kind: 'formula', label: 'Custom formula', inputs: ['Formula (relative to the top-left cell), e.g. =$C2>100'], make: (v, s) => ({ type: 'expression', formula: v[0].replace(/^=/, ''), style: s }) },
];

function cellIs(op: CFOperator, v: string[], style: CFStyle): CFRuleSpec {
  return { type: 'cellIs', operator: op, values: v.map((x) => x.trim()), style };
}

const SCALES: { label: string; stops: string[] }[] = [
  { label: 'Green → Yellow → Red', stops: ['#3fcf8e', '#ffd43b', '#ff6b6b'] },
  { label: 'Red → Yellow → Green', stops: ['#ff6b6b', '#ffd43b', '#3fcf8e'] },
  { label: 'Blue → White → Red', stops: ['#4ea1ff', '#f1f3f5', '#ff6b6b'] },
  { label: 'Light → Green', stops: ['#1f2a24', '#3fcf8e'] },
  { label: 'Light → Blue', stops: ['#1c2533', '#4ea1ff'] },
  { label: 'Light → Red', stops: ['#2b1f22', '#ff6b6b'] },
];

const BARS = ['#4ea1ff', '#3fcf8e', '#ff9f43', '#a78bfa', '#e8618c'];
const ICONS: { set: string; label: string }[] = [
  { set: '3Arrows', label: 'Arrows' },
  { set: '3TrafficLights1', label: 'Traffic lights' },
  { set: '3Symbols', label: 'Check / ! / ✕' },
  { set: '3Flags', label: 'Flags' },
  { set: '5Rating', label: 'Rating bars' },
  { set: '3Stars', label: 'Stars' },
];

export class CFPanel implements PanelContent {
  title = 'Conditional formatting';
  readonly el: HTMLElement;
  private body: HTMLDivElement;
  private builder: Builder | null = null;
  private styleIndex = 0;
  private values: string[] = [];
  private showAll = false;

  constructor(private app: App) {
    this.body = h('div', { class: 'sp-body' });
    this.el = this.body;
    this.render();
  }

  refresh(): void {
    if (document.activeElement && this.body.contains(document.activeElement) && /INPUT|SELECT/.test((document.activeElement as HTMLElement).tagName)) return;
    this.render();
  }

  private targetRanges(): Range[] {
    const b = this.app.doc.engine.usedBounds(this.app.sheet);
    return this.app.sel.ranges.map((r) => ({
      r1: r.r1,
      c1: r.c1,
      r2: r.r2 === 1_048_575 ? Math.max(r.r1, b.maxRow) : r.r2,
      c2: r.c2 === 16_383 ? Math.max(r.c1, b.maxCol) : r.c2,
    }));
  }

  private addRule(spec: CFRuleSpec, ranges = this.targetRanges()): void {
    const rule = { ...spec, id: 'cf' + Math.random().toString(36).slice(2, 9), ranges } as CFRule;
    const s = this.app.sheet;
    this.app.transact('Conditional formatting', (tx) => tx.setProp(s, 'cf', [rule, ...s.cf]));
    this.app.statusbar.flash(`Added rule: ${describeRule(rule)}`);
  }

  private render(): void {
    const app = this.app;
    const b = this.body;
    b.innerHTML = '';
    const ranges = this.targetRanges();
    const label = ranges.map((r) => formatRange(r)).join(', ');
    b.append(h('div', { class: 'hint', style: { marginBottom: '10px' } }, `New rules apply to ${label}.`));

    // Highlight rules
    b.append(h('div', { class: 'field-label', style: { marginBottom: '6px' } }, 'Highlight cells that are…'));
    const chips = h('div', { class: 'chips', style: { marginBottom: '10px' } });
    for (const bl of BUILDERS) {
      chips.append(h('button', { class: `chip${this.builder?.kind === bl.kind ? ' on' : ''}`, onclick: () => {
        this.builder = this.builder?.kind === bl.kind ? null : bl;
        this.values = [];
        this.render();
      } }, bl.label));
    }
    b.append(chips);
    if (this.builder) {
      const bl = this.builder;
      const card = h('div', { class: 'card' });
      const inputs = bl.inputs.map((ph, i) => {
        const inp = h('input', { class: `input${bl.kind === 'formula' ? ' mono' : ''}`, placeholder: ph, value: this.values[i] ?? '', style: { width: '100%' } }) as HTMLInputElement;
        inp.addEventListener('input', () => (this.values[i] = inp.value));
        inp.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') add();
        });
        return inp;
      });
      const styles = h('div', { class: 'chips' });
      STYLES.forEach((st, i) => {
        const chip = h('button', { class: `chip${i === this.styleIndex ? ' on' : ''}`, style: { background: st.style.fill ?? '', color: st.style.color ?? '', fontWeight: st.style.bold ? '700' : '', textDecoration: st.style.strike ? 'line-through' : '' }, onclick: () => {
          this.styleIndex = i;
          this.render();
        } }, st.label);
        styles.append(chip);
      });
      const add = () => {
        const vals = inputs.map((i) => i.value.trim());
        if (vals.some((v) => !v)) {
          toast('Fill in the value first.', { kind: 'error' });
          return;
        }
        this.addRule(bl.make(vals, STYLES[this.styleIndex].style));
        this.builder = null;
        this.render();
      };
      card.append(
        h('div', { class: 'field' }, h('label', null, bl.label), ...inputs),
        h('div', { class: 'field' }, h('label', null, 'Format with'), styles),
        h('div', { class: 'row', style: { justifyContent: 'flex-end' } }, h('button', { class: 'btn primary', onclick: add }, 'Add rule')),
      );
      b.append(card);
      setTimeout(() => inputs[0]?.focus(), 0);
    }

    // Color scales
    b.append(h('div', { class: 'field-label', style: { margin: '14px 0 6px' } }, 'Color scales'));
    const scaleGrid = h('div', { class: 'preset-grid' });
    for (const sc of SCALES) {
      const demo = h('div', { class: 'demo', style: { background: `linear-gradient(90deg, ${sc.stops.join(', ')})` } });
      scaleGrid.append(h('button', { class: 'preset', title: sc.label, onclick: () => {
        const stops = sc.stops.map((color, i) => ({
          at: i === 0 ? { type: 'min' as const } : i === sc.stops.length - 1 ? { type: 'max' as const } : { type: 'percentile' as const, value: 50 },
          color,
        }));
        this.addRule({ type: 'colorScale', stops });
      } }, demo, sc.label));
    }
    b.append(scaleGrid);

    // Data bars
    b.append(h('div', { class: 'field-label', style: { margin: '14px 0 6px' } }, 'Data bars'));
    const barGrid = h('div', { class: 'preset-grid' });
    for (const color of BARS) {
      const demo = h('div', { class: 'demo', style: { background: 'var(--panel)' } }, h('div', { style: { width: '70%', background: `linear-gradient(90deg, ${color}, ${color}33)` } }));
      barGrid.append(h('button', { class: 'preset', onclick: () => this.addRule({ type: 'dataBar', color, min: { type: 'min' }, max: { type: 'max' }, gradient: true, showValue: true }) }, demo, 'Bars'));
    }
    b.append(barGrid);

    // Icon sets
    b.append(h('div', { class: 'field-label', style: { margin: '14px 0 6px' } }, 'Icon sets'));
    const iconGrid = h('div', { class: 'preset-grid' });
    for (const ic of ICONS) {
      iconGrid.append(h('button', { class: 'preset', onclick: () => {
        const n = parseInt(ic.set, 10) || 3;
        const thresholds = Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: Math.round((i * 100) / n) }));
        this.addRule({ type: 'iconSet', iconSet: ic.set, thresholds, reverse: false, showValue: true });
      } }, ic.label));
    }
    b.append(iconGrid);

    // Existing rules
    const s = app.sheet;
    const rules = this.showAll ? s.cf : s.cf.filter((r) => r.ranges.some((rr) => ranges.some((x) => rangesIntersect(rr, x))));
    const header = h('div', { class: 'row', style: { margin: '18px 0 8px', justifyContent: 'space-between' } }, h('div', { class: 'field-label' }, `Rules${this.showAll ? ' on this sheet' : ' in the selection'} (${rules.length})`));
    const toggleAll = h('button', { class: 'btn text', onclick: () => {
      this.showAll = !this.showAll;
      this.render();
    } }, this.showAll ? 'Only selection' : 'Show all');
    header.append(toggleAll);
    b.append(header);
    if (!rules.length) b.append(h('div', { class: 'hint' }, 'No rules yet. Pick one above.'));
    for (const rule of rules) {
      const idx = s.cf.indexOf(rule);
      const sample = h('div', { class: 'sample' }, 'Aa');
      if ('style' in rule) {
        sample.style.background = rule.style.fill ?? '';
        sample.style.color = rule.style.color ?? '';
        if (rule.style.bold) sample.style.fontWeight = '700';
      } else if (rule.type === 'colorScale') sample.style.background = `linear-gradient(90deg, ${rule.stops.map((x) => x.color).join(',')})`;
      else if (rule.type === 'dataBar') sample.style.background = `linear-gradient(90deg, ${rule.color} 60%, transparent 60%)`;
      const rangeInput = h('input', { class: 'input mono', value: rule.ranges.map((r) => formatRange(r)).join(', '), style: { width: '100%', height: '24px', fontSize: '12px', marginTop: '4px' } }) as HTMLInputElement;
      rangeInput.addEventListener('change', () => {
        const parsed = rangeInput.value.split(/[,\s]+/).filter(Boolean).map((x) => parseRange(x));
        if (!parsed.length || parsed.some((p) => !p)) {
          toast('Enter ranges like A1:A20, C1:C20.', { kind: 'error' });
          return;
        }
        const next = [...s.cf];
        next[idx] = { ...rule, ranges: parsed as Range[] };
        app.transact('Edit rule range', (tx) => tx.setProp(s, 'cf', next));
      });
      const up = h('button', { class: 'btn', title: 'Higher priority', onclick: () => {
        if (idx <= 0) return;
        const next = [...s.cf];
        [next[idx - 1], next[idx]] = [next[idx], next[idx - 1]];
        app.transact('Reorder rules', (tx) => tx.setProp(s, 'cf', next));
      } }, icon(L.ArrowUp, 14));
      const del = h('button', { class: 'btn danger', title: 'Delete rule', onclick: () => app.transact('Delete rule', (tx) => tx.setProp(s, 'cf', s.cf.filter((x) => x !== rule))) }, icon(L.Trash2, 14));
      b.append(h('div', { class: 'rule-card' }, sample, h('div', { class: 'info' }, h('div', { class: 't' }, describeRule(rule)), rangeInput), up, del));
    }
  }
}
