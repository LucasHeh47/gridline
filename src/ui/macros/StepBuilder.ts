// The no-code macro builder (Beta, Developer branch): a list of plain-English steps, each a sentence
// with the blanks filled in place ("Color [B2:B20] [Yellow]"). Blocks (If, Repeat, For each…) hold
// their own steps; drag a step by its icon to move it, into or out of a block. Variables are set by
// steps and used in any box as {Name}. "Show code" displays the VBA the steps turn into.
import * as L from 'lucide';
import { formatRange, quoteSheetName } from '../../core/address';
import type { StepMacro } from '../../core/macros/project';
import {
  compileStepMacro,
  countStepsUsing,
  describeStart,
  describeStep,
  fieldShown,
  looksLikeCell,
  macroProcName,
  newStep,
  renameVariable,
  sameVariable,
  STEP_BY_TYPE,
  STEP_CATEGORIES,
  STEP_EXAMPLES,
  stepHelperCode,
  STEPS,
  stepUsesVariable,
  unusedVariableName,
  variableNameProblem,
  withStepVariables,
  type Provided,
  type Step,
  type StepDef,
  type StepField,
  type StepVariable,
} from '../../core/macros/steps';
import type { App } from '../app';
import { h, icon } from '../dom';
import { confirmDialog, dialog, Popover, promptDialog, tooltip } from '../widgets';
import { highlightCode } from './highlight';

const ICONS = L as unknown as Record<string, L.IconNode>;

/** lucide icon by kebab-case name ("list-plus" → ListPlus). */
function stepIcon(name: string): L.IconNode {
  const pascal = name.replace(/(^|-)([a-z0-9])/g, (_m, _d, c: string) => c.toUpperCase());
  return ICONS[pascal] ?? L.Circle;
}

/** Whether `list` is inside `step` (a block can't be dropped into itself). */
function holds(step: Step, list: Step[]): boolean {
  for (const inner of [step.body, step.elseBody]) {
    if (!inner) continue;
    if (inner === list || inner.some((s) => holds(s, list))) return true;
  }
  return false;
}

const ALWAYS: Provided[] = [
  { name: 'Today', hint: 'today’s date' },
  { name: 'Now', hint: 'the date and time' },
];

export class StepBuilder {
  readonly el: HTMLElement;
  private listEl: HTMLElement;
  private codeEl: HTMLElement;
  private problemsEl: HTMLElement;
  private varsEl: HTMLElement;
  private showCode = false;
  private titleEl: HTMLElement;
  private cards = new Map<HTMLElement, Step>();
  private drag: { list: Step[]; step: Step; card: HTMLElement } | null = null;
  private dropAt: { list: Step[]; index: number } | null = null;
  private dropMark = h('div', { class: 'sb-drop' });

  constructor(
    private app: App,
    private macro: StepMacro,
    private changed: () => void,
  ) {
    this.titleEl = h('h3', { class: 'sb-title' });
    this.listEl = h('div', { class: 'sb-list' });
    this.codeEl = h('div', { class: 'sb-code', hidden: true });
    this.problemsEl = h('div', { class: 'sb-problems', hidden: true });
    this.varsEl = h('div', { class: 'sb-vars' });
    // Macros made before variables had a list get one with the names their steps use
    this.macro.variables ??= withStepVariables(this.macro.steps);
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
    this.endDrag();
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

  private compiled() {
    return compileStepMacro(macroProcName(this.macro.name), this.macro.steps, this.macro.description, this.macro.variables);
  }

  private get vars(): StepVariable[] {
    return (this.macro.variables ??= []);
  }

  private renderCode(): void {
    if (!this.showCode) return;
    const { code, helpers } = this.compiled();
    const all = helpers.length ? `${code}\r\n\r\n${stepHelperCode(helpers)}` : code;
    this.codeEl.innerHTML = '';
    this.codeEl.append(
      h('div', { class: 'hint' }, 'This is the VBA your steps turn into. It works in Excel too.'),
      Object.assign(h('pre', { class: 'code-hl static' }), { innerHTML: highlightCode(all.replace(/\r\n/g, '\n')) }),
    );
  }

  /** Mistakes the steps have (a missing cell, a variable that's never set), above the steps. */
  private renderProblems(): void {
    const { problems } = this.compiled();
    this.problemsEl.hidden = !problems.length;
    this.problemsEl.replaceChildren(...problems.map((p) => h('div', { class: 'sb-problem' }, icon(L.TriangleAlert, 13), h('span', null, p))));
  }

  /** The macro's Variables: each with what it starts as, plus "New variable". */
  private renderVars(): void {
    const add = h('button', { class: 'btn outline sb-var-new', onclick: () => void this.createVariable().then((v) => v && this.render()) }, icon(L.Plus, 13), 'New variable');
    const pills = this.vars.map((v) => {
      const start = describeStart(v);
      const pill = h(
        'span',
        { class: 'sb-var-pill' },
        h('button', { class: 'sb-var-edit', onclick: () => void this.editVariable(v) }, h('code', null, `{${v.name}}`), start ? h('span', { class: 'sb-var-start' }, `= ${start}`) : null),
        h('button', { class: 'sb-var-del', 'aria-label': `Delete {${v.name}}`, onclick: () => void this.deleteVariable(v) }, icon(L.X, 12)),
      );
      tooltip(pill.firstElementChild as HTMLElement, `Rename {${v.name}} or change what it starts as`);
      // Hovering a variable points out the steps that use it
      pill.addEventListener('mouseenter', () => this.pointOut(v.name, true));
      pill.addEventListener('mouseleave', () => this.pointOut(v.name, false));
      return pill;
    });
    this.varsEl.replaceChildren(
      h('div', { class: 'sb-vars-head' }, h('span', { class: 'sb-vars-label' }, icon(L.Variable, 14), 'Variables'), h('span', { class: 'sb-vars-hint' }, pills.length ? 'Use them in any box as {Name}.' : 'A variable keeps a value, like a total or a name, that steps can use and change.'), h('span', { class: 'spacer' }), add),
      ...(pills.length ? [h('div', { class: 'sb-var-list' }, ...pills)] : []),
    );
  }

  private pointOut(name: string, on: boolean): void {
    for (const [card, step] of this.cards) card.classList.toggle('hl', on && stepUsesVariable(step, name));
  }

  /** Asks for a new variable (name and starting value) and adds it to the list. */
  private async createVariable(suggestion?: string): Promise<StepVariable | null> {
    const v = await this.variableDialog(null, suggestion);
    if (!v) return null;
    this.vars.push(v);
    this.update();
    return v;
  }

  private async editVariable(v: StepVariable): Promise<void> {
    const next = await this.variableDialog(v);
    if (!next) return;
    if (next.name !== v.name) renameVariable(this.macro.steps, v.name, next.name, this.vars);
    Object.assign(v, next);
    this.render();
    this.update();
  }

  private async deleteVariable(v: StepVariable): Promise<void> {
    const n = countStepsUsing(this.macro.steps, v.name);
    if (n && !(await confirmDialog('Delete variable', `${n} step${n === 1 ? ' uses' : 's use'} {${v.name}}. Delete it anyway? ${n === 1 ? 'That step' : 'Those steps'} will need another variable.`, 'Delete', true))) return;
    this.macro.variables = this.vars.filter((x) => x !== v);
    this.render();
    this.update();
  }

  /** The New variable / edit dialog: a unique name, and what the variable starts as. */
  private variableDialog(current: StepVariable | null, suggestion = 'Total'): Promise<StepVariable | null> {
    const others = this.vars.filter((x) => x !== current).map((x) => x.name);
    const name = h('input', { class: 'input', value: current?.name ?? unusedVariableName(suggestion, others), spellcheck: false, 'aria-label': 'Name' }) as HTMLInputElement;
    const from = h('select', { class: 'input', 'aria-label': 'Starts as' }) as HTMLSelectElement;
    for (const [v, label] of [['value', 'Text or a number'], ['cell', 'The value in a cell'], ['formula', 'A formula’s result']]) from.append(h('option', { value: v }, label));
    from.value = current?.from ?? 'value';
    const value = h('input', { class: 'input', value: current?.value ?? '', spellcheck: false, 'aria-label': 'Starting value' }) as HTMLInputElement;
    const selected = (() => {
      const app = this.app;
      const r = formatRange(app.sel.range);
      const first = app.doc.wb.sheets.find((x) => !x.hidden);
      return app.sheet === first ? r : `${quoteSheetName(app.sheet.name)}!${r}`;
    })();
    const useSel = h('button', { class: 'btn outline', onclick: () => ((value.value = selected), check()) }, icon(L.SquareDashedMousePointer, 13), `Use ${selected}`) as HTMLButtonElement;
    const hint = h('div', { class: 'hint' });
    const err = h('div', { class: 'hint sb-var-err' });
    const shape = () => {
      const k = from.value;
      value.classList.toggle('mono', k === 'cell');
      value.placeholder = k === 'cell' ? 'B2 or Sheet2!B2' : k === 'formula' ? '=SUM(B2:B10)' : 'Leave empty to start with nothing';
      if (k === 'cell' && !value.value.trim()) value.value = selected;
      if (k === 'formula' && !value.value.trim()) value.value = '=';
      useSel.hidden = k !== 'cell';
      hint.textContent =
        k === 'cell'
          ? 'Reads the cell when the macro starts. Steps can change the variable later; the cell stays as it is.'
          : k === 'formula'
            ? 'Calculated when the macro starts. It can use other variables, like ={Price}*2.'
            : 'For example 0, Paid, or {Today}.';
    };
    const body = h(
      'div',
      { class: 'sb-var-dialog' },
      h('div', { class: 'field' }, h('label', null, 'Name'), name),
      h('div', { class: 'field' }, h('label', null, 'Starts as'), h('div', { class: 'sb-var-dialog-row' }, from, value, useSel)),
      hint,
      err,
    );
    const p = dialog(current ? `Variable {${current.name}}` : 'New variable', body, [
      { label: current ? 'Save' : 'Create', value: 'ok', primary: true },
      { label: 'Cancel', value: 'cancel' },
    ]);
    const ok = body.parentElement!.querySelector('.btn.primary') as HTMLButtonElement;
    const check = () => {
      let msg = variableNameProblem(name.value, others);
      if (!msg && from.value === 'cell' && !looksLikeCell(value.value)) msg = 'Type a cell like B2 or Sheet2!B2.';
      err.textContent = msg ?? '';
      ok.disabled = !!msg;
      return !msg;
    };
    for (const el of [name, value]) {
      el.addEventListener('input', check);
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !check()) {
          e.stopPropagation();
          e.preventDefault();
        }
      });
    }
    from.addEventListener('change', () => {
      shape();
      check();
    });
    shape();
    setTimeout(() => {
      name.focus();
      name.select();
      check();
    }, 0);
    return p.then((r) => {
      if (r !== 'ok' || !check()) return null;
      const v = value.value.trim();
      return { name: name.value.trim(), from: from.value as StepVariable['from'], value: from.value === 'formula' && v === '=' ? '' : v };
    });
  }

  private update(): void {
    this.changed();
    this.renderCode();
    this.renderProblems();
    this.renderVars();
  }

  // ---------- list ----------

  private render(): void {
    this.titleEl.textContent = this.macro.name;
    this.listEl.innerHTML = '';
    this.cards.clear();
    const desc = h('input', { class: 'input sb-desc', placeholder: 'What does this macro do? (optional)', value: this.macro.description ?? '' }) as HTMLInputElement;
    desc.addEventListener('change', () => {
      this.macro.description = desc.value.trim() || undefined;
      this.update();
    });
    this.listEl.append(desc, this.varsEl, this.problemsEl);
    if (!this.macro.steps.length) this.listEl.append(this.examples());
    this.listEl.append(this.stepList(this.macro.steps, []));
    this.renderCode();
    this.renderProblems();
    this.renderVars();
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
                const vars = this.vars;
                for (const v of ex.variables?.() ?? []) if (!vars.some((x) => sameVariable(x.name, v.name))) vars.push(v);
                this.macro.variables = withStepVariables(this.macro.steps, vars);
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

  /** A list of steps with an "Add step" button. `scope` is what the blocks around it provide. */
  private stepList(steps: Step[], scope: Provided[]): HTMLElement {
    const box = h('div', { class: 'sb-steps' });
    steps.forEach((step, i) => box.append(this.card(steps, i, scope)));
    const add = h('button', { class: 'btn outline sb-add' }, icon(L.Plus, 14), steps.length ? 'Add step' : 'Add a step') as HTMLButtonElement;
    add.addEventListener('click', () =>
      this.pickStep(add, (type) => void this.addStep(steps, type)),
    );
    box.append(add);
    // Dropping a dragged step here (the innermost list under the pointer wins)
    box.addEventListener('dragover', (e) => {
      if (!this.drag || holds(this.drag.step, steps)) return;
      e.preventDefault();
      e.stopPropagation();
      if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
      const cards = [...box.children].filter((el) => el.classList.contains('sb-card')) as HTMLElement[];
      let index = cards.findIndex((c) => e.clientY < c.getBoundingClientRect().top + c.getBoundingClientRect().height / 2);
      if (index < 0) index = cards.length;
      box.insertBefore(this.dropMark, cards[index] ?? add);
      this.dropAt = { list: steps, index };
    });
    box.addEventListener('drop', (e) => {
      if (!this.drag || this.dropAt?.list !== steps) return;
      e.preventDefault();
      e.stopPropagation();
      const { list: from, step } = this.drag;
      let at = this.dropAt.index;
      const old = from.indexOf(step);
      if (old >= 0) from.splice(old, 1);
      if (from === steps && old >= 0 && old < at) at--;
      steps.splice(at, 0, step);
      this.endDrag();
      this.render();
      this.update();
    });
    return box;
  }

  /** Adds a step. One that stores a value uses the first variable, or asks for a new one. */
  private async addStep(list: Step[], type: string): Promise<void> {
    const step = newStep(type);
    const vf = STEP_BY_TYPE.get(type)?.fields.find((f) => f.kind === 'variable');
    if (vf) {
      const v = this.vars[0] ?? (await this.createVariable(vf.dflt));
      if (!v) return;
      step.args[vf.key] = v.name;
    }
    list.push(step);
    this.render();
    this.update();
  }

  private endDrag(): void {
    this.drag?.card.classList.remove('dragging');
    this.drag = null;
    this.dropAt = null;
    this.dropMark.remove();
  }

  private card(list: Step[], index: number, scope: Provided[]): HTMLElement {
    const step = list[index];
    const def = STEP_BY_TYPE.get(step.type);
    if (!def) return h('div', { class: 'sb-card bad' }, `Unknown step "${step.type}"`);
    const tools = h(
      'div',
      { class: 'sb-tools' },
      this.toolBtn(L.ArrowUp, 'Move up', index === 0, () => this.move(list, index, -1)),
      this.toolBtn(L.ArrowDown, 'Move down', index === list.length - 1, () => this.move(list, index, 1)),
      this.toolBtn(L.Copy, 'Duplicate', false, () => {
        list.splice(index + 1, 0, withNewIds(structuredClone(step)));
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
    // The icon is the drag handle; the card is only draggable while it's held, so text in the
    // boxes can still be selected with the mouse
    const handle = h('span', { class: 'sb-icon sb-handle' }, icon(stepIcon(def.icon), 15));
    tooltip(handle, 'Drag to move (into or out of a block)');
    const card = h(
      'div',
      { class: `sb-card cat-${def.category.toLowerCase()}${step.off ? ' off' : ''}${def.block ? ' block' : ''}` },
      h('div', { class: 'sb-line' }, handle, this.sentence(def, step, scope), tools),
    );
    this.cards.set(card, step);
    handle.addEventListener('pointerdown', () => (card.draggable = true));
    handle.addEventListener('pointerup', () => (card.draggable = false));
    card.addEventListener('dragstart', (e) => {
      if (!card.draggable) return;
      e.stopPropagation();
      this.drag = { list, step, card };
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', describeStep(step));
      }
      requestAnimationFrame(() => card.classList.add('dragging'));
    });
    card.addEventListener('dragend', () => {
      card.draggable = false;
      this.endDrag();
    });
    if (def.block) {
      const inner = def.provides ? [...def.provides, ...scope.filter((v) => !def.provides!.some((p) => p.name === v.name))] : scope;
      if (def.provides?.length)
        card.append(
          h(
            'div',
            { class: 'sb-provides' },
            'Inside this block, use ',
            ...def.provides.flatMap((p, i) => [i ? ', ' : '', h('code', null, `{${p.name}}`), ` for ${p.hint}`]),
          ),
        );
      step.body ??= [];
      card.append(h('div', { class: 'sb-inner' }, this.stepList(step.body, inner)));
      if (def.hasElse) {
        if (step.elseBody) card.append(h('div', { class: 'sb-else-label' }, 'Otherwise'), h('div', { class: 'sb-inner' }, this.stepList(step.elseBody, inner)));
        else
          card.append(
            h(
              'button',
              {
                class: 'btn text sb-add-else',
                onclick: () => {
                  step.elseBody = [];
                  this.render();
                },
              },
              icon(L.Plus, 12),
              'Otherwise…',
            ),
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

  /** The step's sentence with its fields as inline inputs (fields that don't apply are left out). */
  private sentence(def: StepDef, step: Step, scope: Provided[]): HTMLElement {
    const out = h('div', { class: 'sb-sentence' });
    const parts = def.sentence.split(/(\{\w+\})/);
    for (const part of parts) {
      const m = /^\{(\w+)\}$/.exec(part);
      if (!m) {
        // Punctuation sticks to the box before it ("Change [Total]: add")
        if (part.trim()) out.append(h('span', { class: /^[:,.)]/.test(part) ? 'sb-punct' : undefined }, part));
        continue;
      }
      const field = def.fields.find((f) => f.key === m[1]);
      if (field && fieldShown(field, step.args, def)) out.append(this.field(def, field, step, scope));
    }
    // Fields not mentioned in the sentence (like "First row is a header")
    for (const f of def.fields)
      if (!def.sentence.includes(`{${f.key}}`) && fieldShown(f, step.args, def)) out.append(h('span', { class: 'sb-extra' }, `${f.label}:`), this.field(def, f, step, scope));
    return out;
  }

  private field(def: StepDef, f: StepField, step: Step, scope: Provided[]): HTMLElement {
    const value = step.args[f.key] ?? f.dflt ?? '';
    const set = (v: string) => {
      step.args[f.key] = v;
      this.update();
    };
    if (f.kind === 'choice' || f.kind === 'color') {
      const sel = h('select', { class: `input sb-field${f.kind === 'color' ? ' color' : ''}${f.key === 'join' && !value ? ' quiet' : ''}`, 'aria-label': f.label }) as HTMLSelectElement;
      for (const [v, label] of f.options ?? []) sel.append(h('option', { value: v }, label));
      sel.value = value;
      const paint = () => {
        if (f.kind === 'color') sel.style.setProperty('--swatch', sel.value === 'none' ? 'transparent' : sel.value);
      };
      paint();
      // Other fields that only apply for some choices appear or disappear
      const shapes = def.fields.some((x) => x.when?.key === f.key);
      sel.addEventListener('change', () => {
        paint();
        set(sel.value);
        if (shapes) this.render();
      });
      return sel;
    }
    if (f.kind === 'variable') return this.variablePicker(f, step, value, set);
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
    const mono = f.kind === 'range' || f.kind === 'cell' || f.kind === 'column';
    const input = h('input', {
      class: `input sb-field${mono ? ' mono' : ''}${wide ? ' wide' : ''}`,
      value,
      placeholder: f.placeholder ?? f.label,
      'aria-label': f.label,
      spellcheck: false,
    }) as HTMLInputElement;
    const fit = () => (input.size = Math.max(f.kind === 'column' ? 3 : 6, Math.min(wide ? 40 : 18, (input.value || f.placeholder || '').length + 1)));
    fit();
    input.addEventListener('input', fit);
    input.addEventListener('change', () => set(input.value));
    if (f.kind === 'column') return input;
    if (f.kind === 'sheet') {
      const listId = 'sb-sheets';
      let dl = document.getElementById(listId) as HTMLDataListElement | null;
      if (!dl) document.body.append((dl = h('datalist', { id: listId }) as HTMLDataListElement));
      dl.innerHTML = '';
      for (const s of this.app.doc.wb.sheets) dl.append(h('option', { value: s.name }));
      input.setAttribute('list', listId);
    }
    const vars = this.varButton(input, scope, set);
    if (f.kind === 'range' || f.kind === 'cell') {
      const pick = h('button', { class: 'btn sb-pick', 'aria-label': 'Use the selected cells' }, icon(L.SquareDashedMousePointer, 13)) as HTMLButtonElement;
      tooltip(pick, 'Use the cells selected on the sheet');
      pick.addEventListener('click', () => {
        const app = this.app;
        const r = formatRange(app.sel.range);
        const first = app.doc.wb.sheets.find((s) => !s.hidden);
        input.value = app.sheet === first ? r : `${quoteSheetName(app.sheet.name)}!${r}`;
        fit();
        set(input.value);
      });
      // Quick chips for the names most used in addresses (D{Row}, {Cell})
      const quick = scope.filter((v) => v.name === 'Row' || v.name === 'Cell');
      const chips = quick.length
        ? h(
            'span',
            { class: 'sb-var-hint' },
            ...quick.map((v) =>
              h(
                'button',
                {
                  class: 'chip',
                  onclick: () => {
                    insertAtCaret(input, `{${v.name}}`);
                    fit();
                    set(input.value);
                  },
                },
                `{${v.name}}`,
              ),
            ),
          )
        : null;
      return h('span', { class: 'sb-range' }, input, pick, vars, chips);
    }
    return h('span', { class: 'sb-range' }, input, vars);
  }

  /** A step's variable: one of the macro's variables, or a new one. */
  private variablePicker(f: StepField, step: Step, value: string, set: (v: string) => void): HTMLSelectElement {
    const NEW = '\u0001new';
    const sel = h('select', { class: 'input sb-field var', 'aria-label': f.label }) as HTMLSelectElement;
    const match = this.vars.find((v) => sameVariable(v.name, value));
    for (const v of this.vars) sel.append(h('option', { value: v.name }, v.name));
    if (value && !match) sel.append(h('option', { value }, `${value} (not a variable)`));
    sel.append(h('option', { value: NEW }, '+ New variable…'));
    sel.value = match?.name ?? value;
    let before = sel.value;
    sel.addEventListener('change', async () => {
      if (sel.value !== NEW) {
        before = sel.value;
        set(sel.value);
        return;
      }
      sel.value = before;
      const v = await this.createVariable(f.dflt);
      if (!v) return;
      step.args[f.key] = v.name;
      this.render();
      this.update();
    });
    return sel;
  }

  /** The {…} button next to a box: inserts a variable where the cursor is. */
  private varButton(input: HTMLInputElement, scope: Provided[], set: (v: string) => void): HTMLButtonElement {
    const btn = h('button', { class: 'btn sb-pick sb-var-btn', 'aria-label': 'Insert a variable' }, icon(L.Braces, 13)) as HTMLButtonElement;
    tooltip(btn, 'Insert a variable');
    btn.addEventListener('click', () => {
      const pick = (name: string) => {
        pop.close();
        insertAtCaret(input, `{${name}}`);
        input.dispatchEvent(new Event('input'));
        set(input.value);
      };
      const item = (v: Provided) => h('button', { class: 'sb-var-item', onclick: () => pick(v.name) }, h('code', null, `{${v.name}}`), h('span', null, v.hint));
      const own = this.vars
        .filter((x) => !scope.some((v) => sameVariable(v.name, x.name)))
        .map((x) => ({ name: x.name, hint: describeStart(x) ? `starts as ${describeStart(x)}` : 'your variable' }));
      const menu = h(
        'div',
        { class: 'sb-var-menu' },
        ...(scope.length ? [h('div', { class: 'me-section' }, 'From the blocks around this step'), ...scope.map(item)] : []),
        h('div', { class: 'me-section' }, 'Your variables'),
        ...(own.length ? own.map(item) : [h('div', { class: 'hint' }, 'None yet. Click “New variable” above the steps to make one.')]),
        h('div', { class: 'me-section' }, 'Always available'),
        ...ALWAYS.map(item),
      );
      const pop = new Popover(menu, btn);
      pop.open(btn);
    });
    return btn;
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
              {
                class: `sb-pick-item cat-${d.category.toLowerCase()}`,
                onclick: () => {
                  pop.close();
                  onPick(d.type);
                },
              },
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

/** Puts text where the cursor is (or was, before the box lost focus). */
function insertAtCaret(input: HTMLInputElement, text: string): void {
  const start = input.selectionStart ?? input.value.length;
  const end = input.selectionEnd ?? start;
  input.value = input.value.slice(0, start) + text + input.value.slice(end);
  input.focus();
  input.setSelectionRange(start + text.length, start + text.length);
}

/** A copied step (and the steps inside it) with new ids. */
function withNewIds(step: Step): Step {
  step.id = 's' + Math.random().toString(36).slice(2, 9);
  step.body?.forEach(withNewIds);
  step.elseBody?.forEach(withNewIds);
  return step;
}

const CATEGORY_LABELS: Record<string, string> = {
  Cells: 'Cells & data',
  Format: 'Formatting',
  Sheets: 'Sheets',
  Variables: 'Variables',
  Logic: 'Logic & loops',
  Talk: 'Messages & questions',
  Workbook: 'Workbook',
};
