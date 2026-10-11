// The no-code macro builder (Beta, Developer branch): a list of plain-English steps, each a sentence
// with the blanks filled in place ("Color [B2:B20] [Yellow]"). Blocks (If, Repeat, For each row)
// hold their own steps. "Show code" displays the VBA the steps turn into.
import * as L from 'lucide';
import { formatRange, quoteSheetName } from '../../core/address';
import type { StepMacro } from '../../core/macros/project';
import { compileStepMacro, describeStep, macroProcName, newStep, STEP_BY_TYPE, STEP_CATEGORIES, STEP_EXAMPLES, STEPS, type Step, type StepDef, type StepField } from '../../core/macros/steps';
import type { App } from '../app';
import { h, icon } from '../dom';
import { Popover, promptDialog, tooltip } from '../widgets';
import { highlightCode } from './highlight';

const ICONS = L as unknown as Record<string, L.IconNode>;

/** lucide icon by kebab-case name ("list-plus" → ListPlus). */
function stepIcon(name: string): L.IconNode {
  const pascal = name.replace(/(^|-)([a-z0-9])/g, (_m, _d, c: string) => c.toUpperCase());
  return ICONS[pascal] ?? L.Circle;
}

export class StepBuilder {
  readonly el: HTMLElement;
  private listEl: HTMLElement;
  private codeEl: HTMLElement;
  private showCode = false;
  private titleEl: HTMLElement;

  constructor(
    private app: App,
    private macro: StepMacro,
    private changed: () => void,
  ) {
    this.titleEl = h('h3', { class: 'sb-title' });
    this.listEl = h('div', { class: 'sb-list' });
    this.codeEl = h('div', { class: 'sb-code', hidden: true });
    const codeBtn = h('button', { class: 'btn outline', onclick: () => this.toggleCode(codeBtn) }, icon(L.Code, 14), 'Show code') as HTMLButtonElement;
    tooltip(codeBtn, 'See the VBA code these steps turn into');
    this.el = h(
      'div',
      { class: 'step-builder' },
      h(
        'div',
        { class: 'sb-head' },
        this.titleEl,
        h('button', { class: 'btn', 'aria-label': 'Rename', onclick: () => void this.rename() }, icon(L.Pencil, 14)),
        h('span', { class: 'badge' }, 'Beta'),
        h('span', { class: 'spacer' }),
        codeBtn,
        h('button', { class: 'btn primary', onclick: () => void app.macros.run(this.macro.name) }, icon(L.Play, 14), 'Run'),
      ),
      h('div', { class: 'sb-body' }, this.listEl, this.codeEl),
    );
    this.render();
  }

  dispose(): void {
    // nothing to clean up
  }

  async rename(): Promise<void> {
    const project = this.app.macros.ensureProject();
    const next = await promptDialog('Rename step macro', 'Name', this.macro.name, (v) => {
      const t = v.trim();
      if (!t) return 'Give the macro a name.';
      if (t.length > 60) return 'Keep the name under 60 characters.';
      if (project.stepMacros.some((m) => m !== this.macro && macroProcName(m.name).toLowerCase() === macroProcName(t).toLowerCase())) return 'Another step macro has (almost) the same name.';
      return null;
    });
    if (!next || next.trim() === this.macro.name) return;
    const old = this.macro.name;
    this.macro.name = next.trim();
    // Keep buttons, shortcuts and triggers pointing at it
    for (const s of project.shortcuts) if (s.macro === old) s.macro = this.macro.name;
    for (const t of project.triggers) if (t.macro === old) t.macro = this.macro.name;
    for (const sh of this.app.doc.wb.sheets) for (const b of sh.buttons) if (b.macro === old) b.macro = this.macro.name;
    this.render();
    this.changed();
  }

  private toggleCode(btn: HTMLButtonElement): void {
    this.showCode = !this.showCode;
    btn.classList.toggle('on', this.showCode);
    this.codeEl.hidden = !this.showCode;
    this.renderCode();
  }

  private renderCode(): void {
    if (!this.showCode) return;
    const { code, problems } = compileStepMacro(macroProcName(this.macro.name), this.macro.steps, this.macro.description);
    this.codeEl.innerHTML = '';
    this.codeEl.append(
      h('div', { class: 'hint' }, 'This is the VBA your steps turn into. It works in Excel too.'),
      ...problems.map((p) => h('div', { class: 'out-line err' }, icon(L.CircleAlert, 13), p)),
      Object.assign(h('pre', { class: 'code-hl static' }), { innerHTML: highlightCode(code.replace(/\r\n/g, '\n')) }),
    );
  }

  private update(): void {
    this.changed();
    this.renderCode();
  }

  // ---------- list ----------

  private render(): void {
    this.titleEl.textContent = this.macro.name;
    this.listEl.innerHTML = '';
    const desc = h('input', { class: 'input sb-desc', placeholder: 'What does this macro do? (optional)', value: this.macro.description ?? '' }) as HTMLInputElement;
    desc.addEventListener('change', () => {
      this.macro.description = desc.value.trim() || undefined;
      this.update();
    });
    this.listEl.append(desc);
    if (!this.macro.steps.length) this.listEl.append(this.examples());
    this.listEl.append(this.stepList(this.macro.steps, []));
    this.renderCode();
  }

  private examples(): HTMLElement {
    return h(
      'div',
      { class: 'sb-examples' },
      h('div', { class: 'field-label' }, 'Start from an example, or add steps below'),
      h(
        'div',
        { class: 'sb-example-grid' },
        ...STEP_EXAMPLES.map((ex) =>
          h(
            'button',
            {
              class: 'sb-example',
              onclick: () => {
                this.macro.steps = ex.steps();
                if (!this.macro.description) this.macro.description = ex.description;
                this.render();
                this.update();
              },
            },
            h('strong', null, ex.name),
            h('span', null, ex.description),
          ),
        ),
      ),
    );
  }

  /** A list of steps with an "Add step" button. `vars` are loop variables available inside. */
  private stepList(steps: Step[], vars: string[]): HTMLElement {
    const box = h('div', { class: 'sb-steps' });
    steps.forEach((step, i) => box.append(this.card(steps, i, vars)));
    const add = h('button', { class: 'btn outline sb-add' }, icon(L.Plus, 14), steps.length ? 'Add step' : 'Add a step') as HTMLButtonElement;
    add.addEventListener('click', () => this.pickStep(add, (type) => {
      steps.push(newStep(type));
      this.render();
      this.update();
    }));
    box.append(add);
    return box;
  }

  private card(list: Step[], index: number, vars: string[]): HTMLElement {
    const step = list[index];
    const def = STEP_BY_TYPE.get(step.type);
    if (!def) return h('div', { class: 'sb-card bad' }, `Unknown step "${step.type}"`);
    const tools = h(
      'div',
      { class: 'sb-tools' },
      this.toolBtn(L.ArrowUp, 'Move up', index === 0, () => this.move(list, index, -1)),
      this.toolBtn(L.ArrowDown, 'Move down', index === list.length - 1, () => this.move(list, index, 1)),
      this.toolBtn(L.Copy, 'Duplicate', false, () => {
        list.splice(index + 1, 0, structuredClone({ ...step, id: 's' + Math.random().toString(36).slice(2, 9) }));
        this.render();
        this.update();
      }),
      this.toolBtn(step.off ? L.EyeOff : L.Eye, step.off ? 'Turn this step back on' : 'Skip this step (keep it)', false, () => {
        step.off = !step.off || undefined;
        this.render();
        this.update();
      }),
      this.toolBtn(L.Trash2, 'Delete step', false, () => {
        list.splice(index, 1);
        this.render();
        this.update();
      }),
    );
    const card = h(
      'div',
      { class: `sb-card cat-${def.category.toLowerCase()}${step.off ? ' off' : ''}${def.block ? ' block' : ''}` },
      h('div', { class: 'sb-line' }, h('span', { class: 'sb-icon' }, icon(stepIcon(def.icon), 15)), this.sentence(def, step, vars), tools),
    );
    if (def.block) {
      const inner = def.provides ? [...vars, def.provides] : vars;
      if (def.provides) card.append(h('div', { class: 'sb-provides' }, 'Inside this block, use ', h('code', null, `{${def.provides}}`), def.provides === 'Row' ? ' for the row number, e.g. D{Row}' : ' for the repeat number'));
      step.body ??= [];
      card.append(h('div', { class: 'sb-inner' }, this.stepList(step.body, inner)));
      if (def.hasElse) {
        if (step.elseBody) card.append(h('div', { class: 'sb-else-label' }, 'Otherwise'), h('div', { class: 'sb-inner' }, this.stepList(step.elseBody, inner)));
        else
          card.append(
            h('button', { class: 'btn text sb-add-else', onclick: () => {
              step.elseBody = [];
              this.render();
            } }, icon(L.Plus, 12), 'Otherwise…'),
          );
      }
    }
    return card;
  }

  private toolBtn(ic: L.IconNode, label: string, disabled: boolean, run: () => void): HTMLButtonElement {
    const b = h('button', { class: 'btn', 'aria-label': label, disabled, onclick: run }, icon(ic, 13)) as HTMLButtonElement;
    tooltip(b, label);
    return b;
  }

  private move(list: Step[], i: number, d: number): void {
    const j = i + d;
    if (j < 0 || j >= list.length) return;
    [list[i], list[j]] = [list[j], list[i]];
    this.render();
    this.update();
  }

  /** The step's sentence with its fields as inline inputs. */
  private sentence(def: StepDef, step: Step, vars: string[]): HTMLElement {
    const out = h('div', { class: 'sb-sentence' });
    const parts = def.sentence.split(/(\{\w+\})/);
    for (const part of parts) {
      const m = /^\{(\w+)\}$/.exec(part);
      if (!m) {
        if (part) out.append(h('span', null, part));
        continue;
      }
      const field = def.fields.find((f) => f.key === m[1]);
      if (field) out.append(this.field(field, step, vars));
    }
    // Fields not mentioned in the sentence (like "First row is a header")
    for (const f of def.fields) if (!def.sentence.includes(`{${f.key}}`)) out.append(h('span', { class: 'sb-extra' }, `${f.label}:`), this.field(f, step, vars));
    return out;
  }

  private field(f: StepField, step: Step, vars: string[]): HTMLElement {
    const value = step.args[f.key] ?? f.dflt ?? '';
    const set = (v: string) => {
      step.args[f.key] = v;
      this.update();
    };
    if (f.kind === 'choice' || f.kind === 'color') {
      const sel = h('select', { class: `input sb-field${f.kind === 'color' ? ' color' : ''}`, 'aria-label': f.label }) as HTMLSelectElement;
      for (const [v, label] of f.options ?? []) sel.append(h('option', { value: v }, label));
      sel.value = value;
      const paint = () => {
        if (f.kind === 'color') sel.style.setProperty('--swatch', sel.value === 'none' ? 'transparent' : sel.value);
      };
      paint();
      sel.addEventListener('change', () => {
        paint();
        set(sel.value);
      });
      return sel;
    }
    if (f.kind === 'macro') {
      const sel = h('select', { class: 'input sb-field', 'aria-label': f.label }) as HTMLSelectElement;
      const list = this.app.macros.runner.list().filter((m) => m.label !== this.macro.name);
      sel.append(h('option', { value: '' }, 'Choose a macro…'));
      for (const m of list) sel.append(h('option', { value: m.kind === 'steps' ? m.label : m.id }, m.label));
      sel.value = value;
      sel.addEventListener('change', () => set(sel.value));
      return sel;
    }
    const wide = f.kind === 'text';
    const input = h('input', {
      class: `input sb-field${f.kind === 'range' || f.kind === 'cell' || f.kind === 'column' ? ' mono' : ''}${wide ? ' wide' : ''}`,
      value,
      placeholder: f.placeholder ?? f.label,
      'aria-label': f.label,
      spellcheck: false,
    }) as HTMLInputElement;
    input.size = Math.max(f.kind === 'column' ? 3 : 6, Math.min(wide ? 40 : 18, (value || f.placeholder || '').length + 1));
    input.addEventListener('input', () => {
      input.size = Math.max(f.kind === 'column' ? 3 : 6, Math.min(wide ? 40 : 18, input.value.length + 1));
    });
    input.addEventListener('change', () => set(input.value));
    if (f.kind === 'sheet') {
      const listId = 'sb-sheets';
      let dl = document.getElementById(listId) as HTMLDataListElement | null;
      if (!dl) document.body.append((dl = h('datalist', { id: listId }) as HTMLDataListElement));
      dl.innerHTML = '';
      for (const s of this.app.doc.wb.sheets) dl.append(h('option', { value: s.name }));
      input.setAttribute('list', listId);
    }
    if (f.kind === 'range' || f.kind === 'cell') {
      const pick = h('button', { class: 'btn sb-pick', 'aria-label': 'Use the selected cells' }, icon(L.SquareDashedMousePointer, 13)) as HTMLButtonElement;
      tooltip(pick, 'Use the cells selected on the sheet');
      pick.addEventListener('click', () => {
        const app = this.app;
        const r = formatRange(app.sel.range);
        const first = app.doc.wb.sheets.find((s) => !s.hidden);
        input.value = app.sheet === first ? r : `${quoteSheetName(app.sheet.name)}!${r}`;
        set(input.value);
      });
      const hint = vars.length ? h('span', { class: 'sb-var-hint' }, ...vars.map((v) => h('button', { class: 'chip', onclick: () => {
        input.value += `{${v}}`;
        set(input.value);
      } }, `{${v}}`))) : null;
      return h('span', { class: 'sb-range' }, input, pick, hint);
    }
    return input;
  }

  // ---------- adding steps ----------

  private pickStep(anchor: HTMLElement, onPick: (type: string) => void): void {
    const search = h('input', { class: 'input', placeholder: 'Search steps…', spellcheck: false }) as HTMLInputElement;
    const results = h('div', { class: 'sb-pick-list' });
    const pop = new Popover(h('div', { class: 'sb-picker' }, search, results), anchor);
    const render = () => {
      const q = search.value.trim().toLowerCase();
      results.innerHTML = '';
      for (const cat of STEP_CATEGORIES) {
        const defs = STEPS.filter((d) => d.category === cat && (!q || `${d.title} ${d.sentence} ${d.category}`.toLowerCase().includes(q)));
        if (!defs.length) continue;
        results.append(h('div', { class: 'me-section' }, CATEGORY_LABELS[cat]));
        for (const d of defs) {
          results.append(
            h(
              'button',
              { class: `sb-pick-item cat-${d.category.toLowerCase()}`, onclick: () => {
                pop.close();
                onPick(d.type);
              } },
              h('span', { class: 'sb-icon' }, icon(stepIcon(d.icon), 15)),
              h('span', { class: 'sb-pick-text' }, h('strong', null, d.title), h('span', null, describeStep(newStep(d.type)))),
            ),
          );
        }
      }
      if (!results.childElementCount) results.append(h('div', { class: 'hint' }, 'No steps match.'));
    };
    search.addEventListener('input', render);
    search.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') (results.querySelector('.sb-pick-item') as HTMLButtonElement | null)?.click();
    });
    render();
    pop.open(anchor);
    search.focus();
  }
}

const CATEGORY_LABELS: Record<string, string> = {
  Cells: 'Cells & data',
  Format: 'Formatting',
  Sheets: 'Sheets',
  Logic: 'Logic & loops',
  Talk: 'Messages & questions',
  Workbook: 'Workbook',
};
