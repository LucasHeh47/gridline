// Data validation: choose what a cell may contain, an input message and an error alert.
import * as L from 'lucide';
import { formatRange, intersect, type Range } from '../../core/address';
import { describeRule, newValidationId, OPERATORS, rangesText, TYPES, validationAt, type DataValidation, type ValidationType } from '../../core/validation';
import type { App } from '../app';
import { h, icon } from '../dom';
import { toast } from '../widgets';
import type { PanelContent } from './SidePanel';

/** Parts of `a` not covered by `b` (up to four rectangles). */
function subtract(a: Range, b: Range): Range[] {
  const x = intersect(a, b);
  if (!x) return [a];
  const out: Range[] = [];
  if (a.r1 < x.r1) out.push({ r1: a.r1, c1: a.c1, r2: x.r1 - 1, c2: a.c2 });
  if (x.r2 < a.r2) out.push({ r1: x.r2 + 1, c1: a.c1, r2: a.r2, c2: a.c2 });
  if (a.c1 < x.c1) out.push({ r1: x.r1, c1: a.c1, r2: x.r2, c2: x.c1 - 1 });
  if (x.c2 < a.c2) out.push({ r1: x.r1, c1: x.c2 + 1, r2: x.r2, c2: a.c2 });
  return out;
}

export class ValidationPanel implements PanelContent {
  title = 'Data validation';
  readonly el: HTMLElement;
  private draft!: DataValidation;
  /** The rule being edited (null = a new rule for the selection). */
  private editing: DataValidation | null = null;

  constructor(private app: App) {
    this.el = h('div', { class: 'sp-body' });
    this.load();
  }

  refresh(): void {
    if (document.activeElement && this.el.contains(document.activeElement) && /INPUT|SELECT|TEXTAREA/.test((document.activeElement as HTMLElement).tagName)) return;
    this.load();
  }

  selectionChanged(): void {
    this.refresh();
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

  private load(): void {
    const a = this.app.sel.active;
    const existing = validationAt(this.app.sheet, a.row, a.col) ?? null;
    this.editing = existing;
    this.draft = existing
      ? { ...existing, ranges: existing.ranges.map((r) => ({ ...r })) }
      : { id: newValidationId(), ranges: this.targetRanges(), type: 'whole', operator: 'between', formula1: '', formula2: '', allowBlank: true, dropdown: true, showError: true, errorStyle: 'stop', showPrompt: false };
    this.render();
  }

  private render(): void {
    const d = this.draft;
    const b = this.el;
    b.innerHTML = '';
    const field = (label: string, control: HTMLElement, hint?: string) => h('div', { class: 'field' }, h('label', null, label), control, hint ? h('div', { class: 'hint' }, hint) : null);
    const select = <T extends string>(opts: [T, string][], value: T, onChange: (v: T) => void) => {
      const s = h('select', { class: 'input' }) as HTMLSelectElement;
      for (const [v, l] of opts) s.append(h('option', { value: v }, l));
      s.value = value;
      s.addEventListener('change', () => onChange(s.value as T));
      return s;
    };
    const input = (value: string, onChange: (v: string) => void, placeholder = '') => {
      const i = h('input', { class: 'input mono', value, placeholder, spellcheck: false }) as HTMLInputElement;
      i.addEventListener('input', () => onChange(i.value));
      return i;
    };
    const check = (label: string, value: boolean, onChange: (v: boolean) => void) => {
      const c = h('input', { type: 'checkbox', checked: value }) as HTMLInputElement;
      c.addEventListener('change', () => onChange(c.checked));
      return h('label', { class: 'check' }, c, label);
    };

    b.append(
      h('div', { class: 'hint', style: { marginBottom: '10px' } }, this.editing ? `Editing the rule for ${rangesText(d)}.` : `New rule for ${d.ranges.map(formatRange).join(', ')}.`),
      field('Allow', select(TYPES, d.type, (v: ValidationType) => {
        d.type = v;
        if (v === 'list' && d.formula1 && !d.formula1.startsWith('"') && !/^[$A-Za-z]/.test(d.formula1)) d.formula1 = '';
        this.render();
      })),
    );

    if (d.type === 'list') {
      const source = d.formula1.startsWith('"') ? d.formula1.slice(1, -1).replace(/""/g, '"') : d.formula1 ? '=' + d.formula1.replace(/^=/, '') : '';
      b.append(
        field(
          'Choices',
          input(source, (v) => {
            const t = v.trim();
            d.formula1 = t.startsWith('=') ? t.slice(1) : t ? `"${t.split(',').map((s) => s.trim()).filter(Boolean).join(',').replace(/"/g, '""')}"` : '';
          }, 'Yes, No, Maybe   or   =$E$2:$E$20'),
          'Separate choices with commas, or type = and a range of cells that holds them.',
        ),
        check('Show a dropdown arrow in the cell', d.dropdown, (v) => (d.dropdown = v)),
      );
    } else if (d.type === 'custom') {
      b.append(field('Formula', input(d.formula1 ? '=' + d.formula1 : '', (v) => (d.formula1 = v.trim().replace(/^=/, '')), '=AND(B2>0, B2<=C2)'), 'The value is allowed when the formula is TRUE. Write it for the first cell of the range.'));
    } else if (d.type !== 'any') {
      const two = d.operator === 'between' || d.operator === 'notBetween';
      const unit = d.type === 'date' ? 'date, like 1/1/2026' : d.type === 'time' ? 'time, like 9:00' : d.type === 'textLength' ? 'number of characters' : 'number';
      b.append(
        field('Data', select(OPERATORS, d.operator, (v) => {
          d.operator = v;
          this.render();
        })),
        field(two ? 'Minimum' : 'Value', input(this.shown(d.formula1), (v) => (d.formula1 = this.stored(v)), unit)),
      );
      if (two) b.append(field('Maximum', input(this.shown(d.formula2 ?? ''), (v) => (d.formula2 = this.stored(v)), unit)));
      b.append(h('div', { class: 'hint' }, 'You can also type a cell reference or formula, like =B1.'));
    }
    if (d.type !== 'any') b.append(check('Ignore blank cells', d.allowBlank, (v) => (d.allowBlank = v)));

    // Input message
    b.append(h('div', { class: 'field-label', style: { margin: '14px 0 6px' } }, 'Message when the cell is selected'));
    b.append(check('Show a message', d.showPrompt, (v) => {
      d.showPrompt = v;
      this.render();
    }));
    if (d.showPrompt) {
      b.append(field('Title', input(d.promptTitle ?? '', (v) => (d.promptTitle = v || undefined))), field('Message', this.textarea(d.prompt ?? '', (v) => (d.prompt = v || undefined))));
    }

    // Error alert
    b.append(h('div', { class: 'field-label', style: { margin: '14px 0 6px' } }, 'When someone types something else'));
    b.append(check('Show an alert', d.showError, (v) => {
      d.showError = v;
      this.render();
    }));
    if (d.showError) {
      b.append(
        field('Style', select([['stop', 'Stop — don\'t allow it'], ['warning', 'Warning — ask before keeping it'], ['information', 'Information — just tell them']], d.errorStyle, (v) => (d.errorStyle = v))),
        field('Title', input(d.errorTitle ?? '', (v) => (d.errorTitle = v || undefined))),
        field('Message', this.textarea(d.error ?? '', (v) => (d.error = v || undefined))),
      );
    }

    b.append(
      h(
        'div',
        { class: 'row', style: { justifyContent: 'flex-end', gap: '6px', marginTop: '14px' } },
        this.editing ? h('button', { class: 'btn outline danger', onclick: () => this.remove() }, icon(L.Trash2, 14), 'Remove rule') : null,
        h('button', { class: 'btn primary', onclick: () => this.apply() }, this.editing ? 'Save rule' : 'Add rule'),
      ),
    );

    // Checking existing data
    b.append(
      h('div', { class: 'field-label', style: { margin: '18px 0 6px' } }, 'Check the data'),
      h(
        'div',
        { class: 'row', style: { gap: '6px', flexWrap: 'wrap' } },
        h('button', { class: 'btn outline', onclick: () => this.app.commands.run('data.circleInvalid') }, icon(L.CircleDashed, 14), 'Circle invalid data'),
        h('button', { class: 'btn outline', onclick: () => this.app.commands.run('data.clearCircles') }, 'Clear circles'),
      ),
    );

    // Rules on this sheet
    const rules = this.app.sheet.validations;
    if (rules.length) {
      b.append(h('div', { class: 'field-label', style: { margin: '18px 0 6px' } }, `Rules on this sheet (${rules.length})`));
      for (const r of rules) {
        const row = h('button', { class: `card dv-rule${r === this.editing ? ' on' : ''}`, onclick: () => {
          this.app.select(r.ranges[0], { row: r.ranges[0].r1, col: r.ranges[0].c1 });
          this.load();
        } }, h('div', { class: 'mono' }, rangesText(r)), h('div', { class: 'hint' }, describeRule(r)));
        b.append(row);
      }
    }
  }

  private textarea(value: string, onChange: (v: string) => void): HTMLTextAreaElement {
    const t = h('textarea', { class: 'input', rows: 3 }) as HTMLTextAreaElement;
    t.value = value;
    t.addEventListener('input', () => onChange(t.value));
    return t;
  }

  /** Formula text as shown in a box: numbers plain, anything else with "=". */
  private shown(f: string): string {
    if (!f) return '';
    return /^[+-]?(\d+\.?\d*|\.\d+)$/.test(f) ? f : '=' + f.replace(/^=/, '');
  }

  /** What was typed in a box, as stored: numbers and dates as values, "=..." as a formula. */
  private stored(v: string): string {
    const t = v.trim();
    if (!t) return '';
    if (t.startsWith('=')) return t.slice(1);
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(t)) return t;
    // A typed date or time (1/1/2026, 9:00 AM) becomes its serial number
    const n = this.app.doc.engine.evalText(`VALUE("${t.replace(/"/g, '""')}")`, this.app.sheet, 0, 0);
    return typeof n === 'number' ? String(n) : `"${t.replace(/"/g, '""')}"`;
  }

  private apply(): void {
    const d = this.draft;
    if (d.type === 'list' && !d.formula1) {
      toast('Type the choices, separated by commas.', { kind: 'error' });
      return;
    }
    if (d.type === 'custom' && !d.formula1) {
      toast('Type a formula, like =B2>0.', { kind: 'error' });
      return;
    }
    if (d.type !== 'list' && d.type !== 'custom' && d.type !== 'any' && !d.formula1) {
      toast('Enter a value to compare with.', { kind: 'error' });
      return;
    }
    const sheet = this.app.sheet;
    const ranges = d.ranges;
    // A new rule replaces any rule on the same cells (like Excel)
    const others = sheet.validations
      .filter((r) => r !== this.editing)
      .map((r) => ({ ...r, ranges: this.editing ? r.ranges : r.ranges.flatMap((x) => ranges.reduce<Range[]>((acc, cut) => acc.flatMap((a) => subtract(a, cut)), [x])) }))
      .filter((r) => r.ranges.length);
    const next = [...others, { ...d }];
    this.app.transact(this.editing ? 'Edit data validation' : 'Add data validation', (tx) => tx.setProp(sheet, 'validations', next));
    toast(this.editing ? 'Rule saved.' : `Rule added: ${describeRule(d)}.`, { kind: 'success' });
    this.load();
  }

  private remove(): void {
    const sheet = this.app.sheet;
    const target = this.editing;
    if (!target) return;
    this.app.transact('Remove data validation', (tx) => tx.setProp(sheet, 'validations', sheet.validations.filter((r) => r !== target)));
    this.load();
  }
}
