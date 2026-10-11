// The macro editor, docked under the grid so the sheet stays visible while you write and run code.
// Left: VBA modules and (on the Developer branch) step macros. Right: the code or step editor.
// Bottom: output from Debug.Print and messages about runs and errors.
import * as L from 'lucide';
import type { StepMacro, VbaModule } from '../../core/macros/project';
import type { App } from '../app';
import { h, icon } from '../dom';
import { confirmDialog, promptDialog, showMenu, toast, tooltip, type MenuItem } from '../widgets';
import { highlightCode } from './highlight';
import { StepBuilder } from './StepBuilder';

type Selection = { kind: 'module'; name: string } | { kind: 'steps'; id: string };

const HEIGHT_KEY = 'gridline.macroEditorHeight';
const NEW_MODULE = `' A macro is a list of instructions. Run it from Automate → Run macro,
' a button on the sheet, or a keyboard shortcut.
Sub HelloWorld()
    Range("A1").Value = "Hello from Gridline!"
    MsgBox "Done!"
End Sub
`;

/** Lines that start a block, so Enter indents the next line. */
const OPENS_BLOCK = /^\s*((public|private|friend)\s+)?(static\s+)?(sub|function|property)\b|^\s*(if\b.*\bthen\s*(('|rem\b).*)?$|else\b|elseif\b|for\b|do\b|while\b|with\b|select\s+case\b|case\b)/i;

export class MacroEditor {
  readonly el: HTMLElement;
  private list: HTMLElement;
  private main: HTMLElement;
  private out: HTMLElement;
  private sel: Selection | null = null;
  private textarea: HTMLTextAreaElement | null = null;
  private marker: HTMLElement | null = null;
  private steps: StepBuilder | null = null;
  private touchTimer: ReturnType<typeof setTimeout> | undefined;
  private errorLine = 0;

  constructor(private app: App) {
    this.list = h('nav', { class: 'me-list' });
    this.main = h('div', { class: 'me-main' });
    this.out = h('div', { class: 'me-output', 'aria-live': 'polite' });
    const resize = h('div', { class: 'me-resize', title: 'Drag to resize' });
    resize.addEventListener('pointerdown', (e) => this.startResize(e));
    const head = h(
      'header',
      { class: 'me-head' },
      icon(L.SquareCode, 16),
      h('h2', null, 'Macros'),
      this.tool(L.Play, 'Run', () => void this.runCurrent(), 'Run the macro the cursor is in (F5)', true),
      this.tool(L.Square, 'Stop', () => app.macros.stop(), 'Stop the running macro (Esc)'),
      h('span', { class: 'me-sep' }),
      this.toolMenu(L.Plus, 'New', () => this.newMenu()),
      this.tool(L.Keyboard, 'Shortcut', () => void app.macros.assignShortcutFor(this.currentMacroName()), 'Give the macro a keyboard shortcut'),
      this.tool(L.MousePointerClick, 'Button', () => void app.macros.addButtonFor(this.currentMacroName()), 'Put a button on the sheet that runs the macro'),
      h('span', { class: 'spacer' }),
      this.tool(L.CircleHelp, 'Help', () => this.help(), 'What can macros do?'),
      h('button', { class: 'btn', 'aria-label': 'Close the macro editor', onclick: () => app.closeMacroEditor() }, icon(L.X, 16)),
    );
    this.el = h('section', { class: 'macro-editor', hidden: true }, resize, head, h('div', { class: 'me-body' }, this.list, h('div', { class: 'me-right' }, this.main, this.out)));
    this.el.style.height = `${Number(localStorage.getItem(HEIGHT_KEY)) || 300}px`;
    this.el.addEventListener('keydown', (e) => {
      if (e.key === 'F5') {
        e.preventDefault();
        void this.runCurrent();
      } else if (e.key === 'Escape' && app.macros.running) {
        e.preventDefault();
        app.macros.stop();
      }
    });
    app.macros.onOutput = () => this.renderOutput();
  }

  private tool(ic: L.IconNode, label: string, run: () => void, tip?: string, primary = false): HTMLButtonElement {
    const b = h('button', { class: `btn${primary ? ' primary' : ''}`, onclick: run }, icon(ic, 14), h('span', { class: 'lbl' }, label)) as HTMLButtonElement;
    if (tip) tooltip(b, tip);
    return b;
  }

  private toolMenu(ic: L.IconNode, label: string, items: () => MenuItem[]): HTMLButtonElement {
    const b = h('button', { class: 'btn' }, icon(ic, 14), h('span', { class: 'lbl' }, label), icon(L.ChevronDown, 12, 'caret')) as HTMLButtonElement;
    b.addEventListener('click', () => showMenu(items(), b));
    return b;
  }

  get isOpen(): boolean {
    return !this.el.hidden;
  }

  /** Shows the editor, optionally at a module/line or a step macro. */
  open(target?: string, line?: number): void {
    const project = this.app.macros.ensureProject();
    this.el.hidden = false;
    if (target) {
      const step = project.stepMacros.find((m) => m.id === target || m.name === target);
      if (step) this.sel = { kind: 'steps', id: step.id };
      else if (project.module(target)) this.sel = { kind: 'module', name: project.module(target)!.name };
    }
    if (!this.sel || !this.exists(this.sel)) {
      const first = project.modules.find((m) => m.kind === 'standard') ?? project.modules[0];
      this.sel = first ? { kind: 'module', name: first.name } : null;
    }
    this.errorLine = line ?? 0;
    this.render();
    if (line) this.goToLine(line);
    this.app.grid.resize();
  }

  close(): void {
    this.flushTouch();
    this.el.hidden = true;
    this.app.grid.resize();
  }

  private exists(s: Selection): boolean {
    const p = this.app.doc.wb.macros;
    if (!p) return false;
    return s.kind === 'module' ? !!p.module(s.name) : p.stepMacros.some((m) => m.id === s.id);
  }

  /** Redraws after the workbook changed (another file opened, sheets renamed…). */
  refresh(): void {
    if (!this.isOpen) return;
    if (!this.app.doc.wb.macros) this.app.macros.ensureProject();
    if (this.sel && !this.exists(this.sel)) this.sel = null;
    this.renderList();
    if (!this.sel) this.render();
  }

  // ---------- rendering ----------

  private render(): void {
    this.renderList();
    this.renderMain();
    this.renderOutput();
  }

  private renderList(): void {
    const project = this.app.macros.ensureProject();
    const wb = this.app.doc.wb;
    this.list.innerHTML = '';
    const item = (sel: Selection, ic: L.IconNode, label: string, sub?: string) => {
      const on = this.sel && this.sel.kind === sel.kind && (sel.kind === 'module' ? this.sel.kind === 'module' && this.sel.name === sel.name : this.sel.kind === 'steps' && this.sel.id === sel.id);
      const el = h('button', { class: `me-item${on ? ' on' : ''}`, onclick: () => this.select(sel) }, icon(ic, 14), h('span', { class: 'mi-name' }, label), sub ? h('span', { class: 'mi-sub' }, sub) : null);
      el.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        this.select(sel);
        showMenu(this.itemMenu(sel), { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY });
      });
      return el;
    };
    const errors = new Set(this.app.macros.runner.errors().map((e) => e.module.toLowerCase()));
    const mods = project.modules.filter((m) => m.kind !== 'document');
    this.list.append(h('div', { class: 'me-section' }, 'Code'));
    for (const m of mods) this.list.append(item({ kind: 'module', name: m.name }, m.kind === 'standard' ? L.FileCode : L.Box, m.name, errors.has(m.name.toLowerCase()) ? 'has an error' : m.kind === 'class' ? 'class' : m.kind === 'form' ? 'form' : undefined));
    if (!mods.length) this.list.append(h('div', { class: 'me-empty' }, 'No code yet.'));
    this.list.append(h('div', { class: 'me-section' }, 'Workbook & sheets'));
    const wbMod = project.module(project.workbookModule);
    if (wbMod) this.list.append(item({ kind: 'module', name: wbMod.name }, L.BookOpen, 'This workbook', events(wbMod)));
    for (const s of wb.sheets) {
      const mod = s.codeName ? project.module(s.codeName) : undefined;
      if (mod) this.list.append(item({ kind: 'module', name: mod.name }, L.Sheet, s.name, events(mod)));
    }
    if (this.app.macros.stepsAvailable) {
      this.list.append(h('div', { class: 'me-section' }, 'Step macros ', h('span', { class: 'badge' }, 'Beta')));
      for (const sm of project.stepMacros) this.list.append(item({ kind: 'steps', id: sm.id }, L.ListChecks, sm.name, `${countSteps(sm.steps)} step${countSteps(sm.steps) === 1 ? '' : 's'}`));
      this.list.append(h('button', { class: 'me-item add', onclick: () => this.newStepMacro() }, icon(L.Plus, 14), h('span', { class: 'mi-name' }, 'New step macro')));
    }
  }

  private select(sel: Selection): void {
    this.flushTouch();
    this.sel = sel;
    this.errorLine = 0;
    this.render();
  }

  private renderMain(): void {
    this.steps?.dispose();
    this.steps = null;
    this.textarea = null;
    this.main.innerHTML = '';
    const project = this.app.macros.ensureProject();
    const sel = this.sel;
    if (!sel) {
      this.main.append(this.emptyState());
      return;
    }
    if (sel.kind === 'steps') {
      const sm = project.stepMacros.find((m) => m.id === sel.id)!;
      this.steps = new StepBuilder(this.app, sm, () => {
        project.changed();
        this.touchSoon();
        this.renderList();
      });
      this.main.append(this.steps.el);
      return;
    }
    const mod = project.module(sel.name)!;
    this.main.append(this.codeView(mod));
  }

  private emptyState(): HTMLElement {
    return h(
      'div',
      { class: 'me-welcome' },
      h('h3', null, 'Automate your workbook'),
      h('p', null, 'Macros repeat work for you: clean up data, build reports, color rows, or anything you would otherwise do by hand.'),
      h(
        'div',
        { class: 'row' },
        h('button', { class: 'btn primary', onclick: () => this.newModule() }, icon(L.FileCode, 14), 'Write a macro (VBA code)'),
        this.app.macros.stepsAvailable ? h('button', { class: 'btn outline', onclick: () => this.newStepMacro() }, icon(L.ListChecks, 14), 'Build one from steps (no code)') : null,
      ),
    );
  }

  private codeView(mod: VbaModule): HTMLElement {
    const code = mod.code.replace(/\r\n?/g, '\n');
    const ta = h('textarea', { class: 'code-input', spellcheck: false, autocapitalize: 'off', autocomplete: 'off', wrap: 'off', 'aria-label': `Code in ${mod.name}` }) as HTMLTextAreaElement;
    ta.value = code;
    const pre = h('pre', { class: 'code-hl', 'aria-hidden': 'true' });
    const gutter = h('div', { class: 'code-gutter', 'aria-hidden': 'true' });
    const marker = h('div', { class: 'code-error', hidden: true });
    this.textarea = ta;
    this.marker = marker;
    const paint = () => {
      pre.innerHTML = highlightCode(ta.value) + '\n';
      const lines = ta.value.split('\n').length;
      gutter.innerHTML = Array.from({ length: lines }, (_, i) => `<div${i + 1 === this.errorLine ? ' class="err"' : ''}>${i + 1}</div>`).join('');
      this.paintError(mod.name);
    };
    ta.addEventListener('input', () => {
      this.errorLine = 0;
      paint();
      this.app.macros.ensureProject().setCode(mod.name, ta.value.replace(/\n/g, '\r\n'));
      this.touchSoon();
    });
    ta.addEventListener('scroll', () => {
      pre.scrollTop = ta.scrollTop;
      pre.scrollLeft = ta.scrollLeft;
      gutter.scrollTop = ta.scrollTop;
      marker.style.transform = `translateY(${-ta.scrollTop}px)`;
    });
    ta.addEventListener('keydown', (e) => this.onCodeKey(e, ta));
    paint();
    const readOnlyNote = mod.kind === 'form' ? h('div', { class: 'me-note' }, icon(L.Info, 14), "This is a UserForm. Gridline keeps it in the file, but can't show or run forms yet.") : null;
    const docNote =
      mod.kind === 'document' && !mod.code.trim()
        ? h('div', { class: 'me-note' }, icon(L.Info, 14), 'Code here can run automatically: ', h('button', { class: 'btn text', onclick: () => this.insertEvent(mod, ta) }, 'add an event…'))
        : null;
    return h('div', { class: 'code-view' }, readOnlyNote, docNote, h('div', { class: 'code-wrap' }, gutter, h('div', { class: 'code-area' }, pre, marker, ta)));
  }

  /** Underlines the line with a syntax or runtime error. */
  private paintError(module: string): void {
    const marker = this.marker;
    if (!marker || !this.textarea) return;
    const syntax = this.app.macros.runner.errors().find((e) => e.module.toLowerCase() === module.toLowerCase());
    const line = this.errorLine || syntax?.line || 0;
    marker.hidden = !line;
    if (line) {
      const lh = parseFloat(getComputedStyle(this.textarea).lineHeight) || 19;
      marker.style.top = `${8 + (line - 1) * lh}px`;
      marker.style.height = `${lh}px`;
      marker.title = syntax && !this.errorLine ? syntax.message : '';
    }
  }

  private onCodeKey(e: KeyboardEvent, ta: HTMLTextAreaElement): void {
    const { selectionStart: s, selectionEnd: end, value } = ta;
    if (e.key === 'Tab') {
      e.preventDefault();
      const lineStart = value.lastIndexOf('\n', s - 1) + 1;
      if (e.shiftKey) {
        const m = /^( {1,4}|\t)/.exec(value.slice(lineStart));
        if (m) {
          ta.setRangeText('', lineStart, lineStart + m[0].length, 'preserve');
          ta.selectionStart = ta.selectionEnd = Math.max(lineStart, s - m[0].length);
        }
      } else if (s !== end && value.slice(s, end).includes('\n')) {
        // Indent every selected line
        const block = value.slice(lineStart, end);
        ta.setRangeText(block.replace(/^/gm, '    '), lineStart, end, 'select');
      } else ta.setRangeText('    ', s, end, 'end');
      ta.dispatchEvent(new Event('input'));
      return;
    }
    if (e.key === 'Enter' && !e.ctrlKey && !e.altKey) {
      e.preventDefault();
      const lineStart = value.lastIndexOf('\n', s - 1) + 1;
      const line = value.slice(lineStart, s);
      let indent = /^\s*/.exec(line)![0];
      if (OPENS_BLOCK.test(line) && !/^\s*end\b/i.test(line)) indent += '    ';
      let insert = '\n' + indent;
      // Typing "Sub Name()" adds the matching End Sub
      const opener = /^\s*(?:(?:public|private)\s+)?(sub|function)\s+\w+.*$/i.exec(line);
      if (opener && !new RegExp(`^\\s*end\\s+${opener[1]}`, 'im').test(value.slice(s).split(/\n\s*(?:(?:public|private)\s+)?(?:sub|function)\b/i)[0])) {
        insert += `\n${indent.slice(4)}End ${opener[1][0].toUpperCase()}${opener[1].slice(1).toLowerCase()}`;
        ta.setRangeText(insert, s, end, 'end');
        ta.selectionStart = ta.selectionEnd = s + 1 + indent.length;
      } else ta.setRangeText(insert, s, end, 'end');
      ta.dispatchEvent(new Event('input'));
    }
  }

  private goToLine(line: number): void {
    const ta = this.textarea;
    if (!ta) return;
    const lines = ta.value.split('\n');
    const pos = lines.slice(0, Math.max(0, line - 1)).reduce((n, l) => n + l.length + 1, 0);
    ta.focus();
    ta.setSelectionRange(pos, pos + (lines[line - 1]?.length ?? 0));
    const lh = parseFloat(getComputedStyle(ta).lineHeight) || 19;
    ta.scrollTop = Math.max(0, (line - 4) * lh);
  }

  renderOutput(): void {
    const lines = this.app.macros.output;
    this.out.innerHTML = '';
    const errors = this.app.macros.runner.errors();
    for (const e of errors) {
      const row = h('div', { class: 'out-line err' }, icon(L.CircleAlert, 13), `${e.module}${e.line ? `, line ${e.line}` : ''}: ${e.message}`);
      if (e.line && this.app.doc.wb.macros?.module(e.module)) row.addEventListener('click', () => this.open(e.module, e.line));
      this.out.append(row);
    }
    if (!lines.length && !errors.length) this.out.append(h('div', { class: 'out-line hint' }, 'Output from Debug.Print and messages about macro runs show up here.'));
    for (const l of lines) this.out.append(h('div', { class: `out-line${l.startsWith('✗') ? ' err' : l.startsWith('✓') ? ' ok' : ''}` }, l));
    this.out.scrollTop = this.out.scrollHeight;
  }

  // ---------- actions ----------

  /** The macro to run for "Run": the Sub the cursor is in, else the first one in the module. */
  private currentMacroName(): string | null {
    const sel = this.sel;
    const project = this.app.doc.wb.macros;
    if (!sel || !project) return null;
    if (sel.kind === 'steps') {
      const sm = project.stepMacros.find((m) => m.id === sel.id);
      return sm ? sm.name : null;
    }
    const mod = project.module(sel.name);
    if (!mod) return null;
    const text = this.textarea?.value ?? mod.code;
    const upto = this.textarea ? text.slice(0, this.textarea.selectionStart) : '';
    const sub = /^\s*(?:public\s+)?sub\s+(\w+)\s*\(\s*\)/gim;
    let current: string | null = null;
    let first: string | null = null;
    let m: RegExpExecArray | null;
    while ((m = sub.exec(text))) {
      first ??= m[1];
      if (m.index <= upto.length) current = m[1];
    }
    const name = current ?? first;
    return name ? `${mod.name}.${name}` : null;
  }

  private async runCurrent(): Promise<void> {
    this.flushTouch();
    const name = this.currentMacroName();
    if (!name) {
      const list = this.app.macros.runner.list();
      if (!list.length) {
        toast('There are no macros to run yet. Write a Sub, like: Sub Hello() … End Sub');
        return;
      }
      showMenu(
        list.map((m) => ({ label: m.label, icon: m.kind === 'steps' ? L.ListChecks : L.Play, run: () => void this.app.macros.run(m.id) })),
        this.el.querySelector('.me-head .btn.primary') as HTMLElement,
      );
      return;
    }
    await this.app.macros.run(name);
    this.renderOutput();
  }

  private newMenu(): MenuItem[] {
    const items: MenuItem[] = [
      { label: 'Code module', icon: L.FileCode, run: () => this.newModule() },
      { label: 'Class module', icon: L.Box, run: () => this.newModule('class') },
    ];
    if (this.app.macros.stepsAvailable) items.push({ separator: true }, { label: 'Step macro (no code) — Beta', icon: L.ListChecks, run: () => this.newStepMacro() });
    return items;
  }

  newModule(kind: 'standard' | 'class' = 'standard'): void {
    const project = this.app.macros.ensureProject();
    const mod = project.addModule(kind, undefined, kind === 'standard' ? NEW_MODULE.replace(/\n/g, '\r\n') : "' A class describes your own kind of object.\r\nPublic Name As String\r\n");
    this.app.doc.touch();
    this.open(mod.name);
    this.textarea?.focus();
  }

  async newStepMacro(example?: { name: string; description: string; steps: () => StepMacro['steps'] }): Promise<void> {
    const project = this.app.macros.ensureProject();
    let name = example?.name ?? 'New macro';
    let i = 2;
    const taken = (n: string) => project.stepMacros.some((m) => m.name.toLowerCase() === n.toLowerCase()) || this.app.macros.runner.list().some((m) => m.label.toLowerCase() === n.toLowerCase());
    const base = name;
    while (taken(name)) name = `${base} ${i++}`;
    const sm: StepMacro = { id: 'sm' + Math.random().toString(36).slice(2, 9), name, description: example?.description, steps: example?.steps() ?? [] };
    project.stepMacros.push(sm);
    project.changed();
    this.app.doc.touch();
    this.open(sm.id);
  }

  private itemMenu(sel: Selection): MenuItem[] {
    const project = this.app.macros.ensureProject();
    if (sel.kind === 'steps') {
      const sm = project.stepMacros.find((m) => m.id === sel.id)!;
      return [
        { label: 'Run', icon: L.Play, run: () => void this.app.macros.run(sm.name) },
        { label: 'Rename…', icon: L.TextCursorInput, run: () => this.steps?.rename() },
        { label: 'Delete', icon: L.Trash2, run: () => void this.deleteStepMacro(sm) },
      ];
    }
    const mod = project.module(sel.name)!;
    if (mod.kind === 'document') return [{ label: 'Document modules belong to the workbook and its sheets.', disabled: true }];
    return [
      { label: 'Rename…', icon: L.TextCursorInput, run: () => void this.renameModule(mod) },
      { label: 'Delete', icon: L.Trash2, run: () => void this.deleteModule(mod) },
    ];
  }

  private async renameModule(mod: VbaModule): Promise<void> {
    const project = this.app.macros.ensureProject();
    const next = await promptDialog('Rename module', 'New name', mod.name, (v) => {
      const t = v.trim();
      if (!/^[A-Za-z][A-Za-z0-9_]{0,30}$/.test(t)) return 'Use letters, numbers and _ (start with a letter).';
      if (t.toLowerCase() !== mod.name.toLowerCase() && project.module(t)) return 'Another module already has that name.';
      return null;
    });
    if (!next || next === mod.name) return;
    const old = mod.name;
    mod.name = next.trim();
    for (const s of project.shortcuts) if (s.macro.toLowerCase().startsWith(old.toLowerCase() + '.')) s.macro = mod.name + s.macro.slice(old.length);
    project.changed(true);
    this.app.doc.touch();
    this.sel = { kind: 'module', name: mod.name };
    this.render();
  }

  private async deleteModule(mod: VbaModule): Promise<void> {
    if (!(await confirmDialog(`Delete ${mod.name}?`, 'Its code will be removed from the workbook. You can still close without saving to get it back.', 'Delete', true))) return;
    this.app.macros.ensureProject().removeModule(mod.name);
    this.app.doc.touch();
    this.sel = null;
    this.open();
  }

  private async deleteStepMacro(sm: StepMacro): Promise<void> {
    if (!(await confirmDialog(`Delete "${sm.name}"?`, 'Buttons, shortcuts and triggers that run it will stop working.', 'Delete', true))) return;
    const project = this.app.macros.ensureProject();
    project.stepMacros = project.stepMacros.filter((m) => m !== sm);
    project.triggers = project.triggers.filter((t) => t.macro !== sm.name);
    project.changed();
    this.app.doc.touch();
    this.sel = null;
    this.open();
  }

  private insertEvent(mod: VbaModule, ta: HTMLTextAreaElement): void {
    const isWb = mod.name.toLowerCase() === this.app.macros.ensureProject().workbookModule.toLowerCase();
    const events: [string, string][] = isWb
      ? [
          ['When the workbook opens', 'Private Sub Workbook_Open()\n    \nEnd Sub'],
          ['Before the workbook is saved', 'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\n    \nEnd Sub'],
          ['When cells change on any sheet', 'Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)\n    \nEnd Sub'],
        ]
      : [
          ['When cells on this sheet change', 'Private Sub Worksheet_Change(ByVal Target As Range)\n    \nEnd Sub'],
          ['When this sheet is shown', 'Private Sub Worksheet_Activate()\n    \nEnd Sub'],
          ['When the selection moves', 'Private Sub Worksheet_SelectionChange(ByVal Target As Range)\n    \nEnd Sub'],
        ];
    showMenu(
      events.map(([label, code]) => ({
        label,
        run: () => {
          ta.value = (ta.value.trim() ? ta.value.trimEnd() + '\n\n' : '') + code;
          ta.dispatchEvent(new Event('input'));
          this.renderMain();
          const t = this.textarea;
          if (t) {
            const pos = t.value.lastIndexOf('    \n');
            t.focus();
            t.setSelectionRange(pos + 4, pos + 4);
          }
        },
      })),
      this.main.querySelector('.me-note .btn') as HTMLElement,
    );
  }

  private help(): void {
    showMenu(
      [
        { title: 'Macros in Gridline' },
        { label: 'Write VBA code like in Excel: Sub, Range, Cells, MsgBox…', disabled: true },
        { label: 'Press F5 to run the macro your cursor is in', disabled: true },
        { label: 'Debug.Print writes to the output below', disabled: true },
        { label: 'A whole macro run undoes with one Ctrl+Z', disabled: true },
        { separator: true },
        { label: 'Open the macros guide', icon: L.BookOpen, run: () => this.app.openGuide('macros') },
      ],
      this.el.querySelector('.me-head .spacer + .btn') as HTMLElement,
    );
  }

  // ---------- saving & sizing ----------

  private touchSoon(): void {
    clearTimeout(this.touchTimer);
    this.touchTimer = setTimeout(() => this.flushTouch(), 400);
  }

  private flushTouch(): void {
    if (this.touchTimer === undefined) return;
    clearTimeout(this.touchTimer);
    this.touchTimer = undefined;
    this.app.doc.touch();
    this.renderListKeepFocus();
    if (this.sel?.kind === 'module') this.paintError(this.sel.name);
    this.renderOutput();
  }

  private renderListKeepFocus(): void {
    const focused = document.activeElement;
    this.renderList();
    if (focused instanceof HTMLElement && this.el.contains(focused)) focused.focus();
  }

  private startResize(e: PointerEvent): void {
    e.preventDefault();
    const startY = e.clientY;
    const startH = this.el.offsetHeight;
    const move = (ev: PointerEvent) => {
      const hgt = Math.max(160, Math.min(window.innerHeight - 220, startH + (startY - ev.clientY)));
      this.el.style.height = `${hgt}px`;
      this.app.grid.resize();
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      try {
        localStorage.setItem(HEIGHT_KEY, String(this.el.offsetHeight));
      } catch {
        // ignore
      }
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }
}

function events(mod: VbaModule): string | undefined {
  const n = (mod.code.match(/^\s*(?:private\s+|public\s+)?sub\s+(workbook|worksheet)_\w+/gim) ?? []).length;
  return n ? `${n} event${n === 1 ? '' : 's'}` : undefined;
}

function countSteps(steps: StepMacro['steps']): number {
  return steps.reduce((n, s) => n + 1 + countSteps(s.body ?? []) + countSteps(s.elseBody ?? []), 0);
}

