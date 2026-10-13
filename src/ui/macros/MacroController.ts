// Macros in the app: running them (with a Stop button), the "this file has macros" banner,
// keyboard shortcuts, automatic events, and the dialogs VBA code opens (MsgBox, InputBox).
import * as L from 'lucide';
import { keyCol, keyRow, type Range } from '../../core/address';
import type { Transaction } from '../../core/history';
import type { MacroUI } from '../../core/macros/host';
import { userFunctions } from '../../core/macros/cellFunctions';
import { MacroProject } from '../../core/macros/project';
import { ensureDocumentModules, MacroRunner, type RunResult } from '../../core/macros/runner';
import type { Sheet } from '../../core/workbook';
import type { App } from '../app';
import { eventKey } from '../commands';
import { h, icon } from '../dom';
import { dialog, promptDialog, toast } from '../widgets';

const TRUST_KEY = 'gridline.trustedMacroFiles';

/** VBA MsgBox button sets → [label, return value]. */
const MSGBOX_BUTTONS: [string, number][][] = [
  [['OK', 1]],
  [['OK', 1], ['Cancel', 2]],
  [['Abort', 3], ['Retry', 4], ['Ignore', 5]],
  [['Yes', 6], ['No', 7], ['Cancel', 2]],
  [['Yes', 6], ['No', 7]],
  [['Retry', 4], ['Cancel', 2]],
];

/** What a workbook tab remembers about its macros while another tab is shown. */
export interface MacroTabState {
  enabled: boolean;
  banner: [string, BannerAction[]] | null;
}

type BannerAction = { label: string; primary?: boolean; run: () => void };

export class MacroController {
  readonly runner: MacroRunner;
  /** Macros may run in this workbook (security). */
  enabled = true;
  readonly banner: HTMLElement;
  private pill: HTMLElement | null = null;
  private bannerArgs: [string, BannerAction[]] | null = null;
  private pillTimer: ReturnType<typeof setTimeout> | undefined;
  /** Debug.Print output and run messages, shown in the macro editor. */
  readonly output: string[] = [];
  onOutput: (() => void) | null = null;

  constructor(private app: App) {
    this.runner = new MacroRunner(
      () => app.doc,
      () => this.ui(),
    );
    this.runner.onPrint = (t) => this.log(t);
    this.banner = h('div', { class: 'macro-banner', hidden: true });
  }

  get project(): MacroProject | null {
    return this.app.doc.wb.macros;
  }

  get running(): boolean {
    return this.runner.running;
  }

  /** The workbook's macro project, created on first use. */
  ensureProject(): MacroProject {
    const p = ensureDocumentModules(this.app.doc);
    return p;
  }

  /** The no-code step builder is still in development: only on the Developer branch (or a dev build). */
  get stepsAvailable(): boolean {
    return this.app.settings.bool('updates.devBranch') || __APP_VERSION__.includes('-') || !!this.project?.stepMacros.length;
  }

  log(text: string): void {
    this.output.push(text);
    if (this.output.length > 500) this.output.splice(0, this.output.length - 500);
    this.onOutput?.();
  }

  // ---------- opening files ----------

  /** Formulas can use the workbook's VBA Functions only while its macros may run. */
  private syncFunctions(): void {
    const doc = this.app.doc;
    doc.setUserFunctions(this.enabled ? userFunctions(doc) : null);
  }

  /** After a workbook is opened: decide whether its macros may run, and run Workbook_Open. */
  async workbookOpened(path: string | null, fromFile: boolean): Promise<void> {
    this.output.length = 0;
    this.hideBanner();
    const p = this.project;
    const hasMacros = !!p && (p.hasCode || p.stepMacros.length > 0);
    if (!hasMacros || !fromFile) {
      this.enabled = true;
      this.syncFunctions();
      return;
    }
    const mode = this.app.settings.str('macros.security');
    if (mode === 'enable' || (path && this.trusted().includes(path))) {
      this.enabled = true;
      this.syncFunctions();
      await this.fireOpen();
      return;
    }
    this.enabled = false;
    this.syncFunctions();
    if (mode === 'disable') {
      this.showBanner('Macros in this workbook are turned off (Settings → Macros).', [{ label: 'Settings', run: () => this.app.openSettings('macros') }]);
      return;
    }
    this.showBanner('This workbook contains macros. Macros can change your workbook — only turn them on if you trust where the file came from.', [
      { label: 'Enable macros', primary: true, run: () => void this.enable(false) },
      ...(path ? [{ label: 'Always trust this file', run: () => void this.enable(true) }] : []),
      { label: 'View code', run: () => this.app.openMacroEditor() },
    ]);
  }

  async enable(trustFile: boolean): Promise<void> {
    const path = this.app.doc.filePath;
    if (trustFile && path) {
      const list = this.trusted().filter((p) => p !== path);
      list.unshift(path);
      try {
        localStorage.setItem(TRUST_KEY, JSON.stringify(list.slice(0, 200)));
      } catch {
        // ignore
      }
    }
    this.enabled = true;
    this.hideBanner();
    this.syncFunctions();
    await this.fireOpen();
  }

  private trusted(): string[] {
    try {
      const v = JSON.parse(localStorage.getItem(TRUST_KEY) ?? '[]');
      return Array.isArray(v) ? v : [];
    } catch {
      return [];
    }
  }

  saveState(): MacroTabState {
    return { enabled: this.enabled, banner: this.banner.hidden ? null : this.bannerArgs };
  }

  restoreState(s: MacroTabState | undefined): void {
    this.enabled = s?.enabled ?? true;
    if (s?.banner) this.showBanner(...s.banner);
    else this.hideBanner();
  }

  private showBanner(text: string, actions: BannerAction[]): void {
    this.bannerArgs = [text, actions];
    this.banner.innerHTML = '';
    this.banner.append(
      icon(L.ShieldAlert, 16),
      h('span', { class: 'mb-text' }, text),
      ...actions.map((a) => h('button', { class: `btn ${a.primary ? 'primary' : 'outline'}`, onclick: a.run }, a.label)),
      h('button', { class: 'btn', 'aria-label': 'Close', onclick: () => this.hideBanner() }, icon(L.X, 14)),
    );
    this.banner.hidden = false;
    this.app.grid.resize();
  }

  private hideBanner(): void {
    if (this.banner.hidden) return;
    this.banner.hidden = true;
    this.app.grid.resize();
  }

  /** Asks to enable macros when someone runs one in a workbook that hasn't been trusted yet. */
  private async ensureEnabled(): Promise<boolean> {
    if (this.enabled) return true;
    if (this.app.settings.str('macros.security') === 'disable') {
      toast('Macros are turned off. Change this in Settings → Macros.', { kind: 'error' });
      return false;
    }
    const ok = await dialog('Enable macros?', 'Macros can change your workbook. Only enable them if you trust where this file came from.', [
      { label: 'Enable macros', value: true, primary: true },
      { label: 'Cancel', value: false },
    ]);
    if (ok) {
      this.enabled = true;
      this.hideBanner();
    }
    return ok;
  }

  // ---------- running ----------

  /** Runs a macro by name ("Module1.Main", "Main" or a step macro's name). */
  async run(name: string, opts: { quiet?: boolean } = {}): Promise<RunResult | null> {
    if (this.running) {
      toast('A macro is already running.', { kind: 'error' });
      return null;
    }
    if (!(await this.ensureEnabled())) return null;
    if (this.app.isEditing && !this.app.commitEdit()) return null;
    const label = this.runner.list().find((m) => m.id === this.runner.resolve(name))?.label ?? name;
    this.startPill(label);
    let result: RunResult;
    try {
      result = await this.runner.run(name);
    } finally {
      this.stopPill();
      this.app.statusbar.hold(null);
    }
    this.report(label, result, opts.quiet);
    return result;
  }

  stop(): void {
    this.runner.stop();
  }

  private report(label: string, r: RunResult, quiet?: boolean): void {
    if (r.ok) {
      this.log(`✓ ${label} finished${r.ms > 1000 ? ` in ${(r.ms / 1000).toFixed(1)} s` : ''}.`);
      if (!quiet) this.app.statusbar.flash(`Ran ${label}`);
      return;
    }
    if (r.stopped) {
      this.log(`■ ${label} was stopped.`);
      toast(`Stopped "${label}". Changes it already made can be undone with Ctrl+Z.`);
      return;
    }
    const e = r.error!;
    const where = e.module ? ` (${e.module}${e.line ? `, line ${e.line}` : ''})` : '';
    this.log(`✗ ${label}: ${e.message}${where}`);
    toast(`Macro "${label}" stopped: ${e.message}${where}`, {
      kind: 'error',
      ms: 10000,
      action: e.module && this.project?.module(e.module) ? { label: 'Show code', run: () => this.app.openMacroEditor(e.module, e.line) } : undefined,
    });
  }

  private startPill(label: string): void {
    clearTimeout(this.pillTimer);
    this.pillTimer = setTimeout(() => {
      this.pill = h(
        'div',
        { class: 'macro-running', role: 'status' },
        h('span', { class: 'spinner' }),
        h('span', null, `Running "${label}"…`),
        h('button', { class: 'btn outline', onclick: () => this.stop() }, icon(L.Square, 12), 'Stop'),
        h('span', { class: 'hint' }, 'Esc'),
      );
      document.body.append(this.pill);
    }, 400);
  }

  private stopPill(): void {
    clearTimeout(this.pillTimer);
    this.pill?.remove();
    this.pill = null;
  }

  // ---------- keyboard ----------

  /** Runs the macro assigned to this key combination. Returns true when one ran. */
  handleKey(e: KeyboardEvent): boolean {
    if (this.running) {
      if (e.key === 'Escape' || (e.key === 'Pause' && e.ctrlKey)) {
        e.preventDefault();
        this.stop();
      }
      return true; // the workbook is busy
    }
    const p = this.project;
    if (!p?.shortcuts.length || ['Control', 'Shift', 'Alt', 'Meta'].includes(e.key)) return false;
    const key = eventKey(e);
    const hit = p.shortcuts.find((s) => s.keys.toLowerCase() === key.toLowerCase());
    if (!hit) return false;
    e.preventDefault();
    void this.run(hit.macro);
    return true;
  }

  // ---------- events ----------

  private eventsOn(): boolean {
    return this.enabled && this.app.settings.bool('macros.events') && !this.running && this.runner.hasHandlers();
  }

  private async fireOpen(): Promise<void> {
    if (!this.eventsOn()) return;
    for (const r of await this.runner.fireOpen()) this.report('Workbook_Open', r, true);
  }

  /** After people change cells (not macros): Worksheet_Change and "when cells change" triggers. */
  cellsChanged(tx: Transaction): void {
    if (tx.source === 'macro' || !this.eventsOn()) return;
    const bySheet = new Map<Sheet, Range>();
    for (const ch of tx.changes) {
      if (ch.k !== 'cell') continue;
      const b = ch.before;
      const a = ch.after;
      if (b?.f === a?.f && b?.v === a?.v) continue;
      const r = keyRow(ch.key);
      const c = keyCol(ch.key);
      const cur = bySheet.get(ch.sheet);
      bySheet.set(ch.sheet, cur ? { r1: Math.min(cur.r1, r), c1: Math.min(cur.c1, c), r2: Math.max(cur.r2, r), c2: Math.max(cur.c2, c) } : { r1: r, c1: c, r2: r, c2: c });
    }
    if (!bySheet.size) return;
    void (async () => {
      for (const [sheet, range] of bySheet) for (const r of await this.runner.fireChange(sheet, [range])) this.report('Worksheet_Change', r, true);
    })();
  }

  sheetActivated(sheet: Sheet): void {
    if (!this.eventsOn()) return;
    void (async () => {
      for (const r of await this.runner.fireActivate(sheet)) this.report('Worksheet_Activate', r, true);
    })();
  }

  /** Workbook_BeforeSave. Returns false when the code cancelled saving. */
  async beforeSave(saveAs: boolean): Promise<boolean> {
    if (!this.eventsOn()) return true;
    const { cancel, results } = await this.runner.fireBeforeSave(saveAs);
    for (const r of results) this.report('Workbook_BeforeSave', r, true);
    if (cancel) toast('A macro in this workbook cancelled saving.');
    return !cancel;
  }

  // ---------- dialogs: pick a macro, shortcuts, buttons, triggers ----------

  /** Asks which macro to use. Returns its id (or step macro name), or null. */
  async pickMacro(title: string, current?: string | null): Promise<string | null> {
    const list = this.runner.list();
    if (!list.length) {
      toast('There are no macros yet. Open Automate → Macro editor to write one.');
      return null;
    }
    const sel = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const m of list) sel.append(h('option', { value: m.kind === 'steps' ? m.label : m.id }, m.kind === 'steps' ? `${m.label} (steps)` : m.module ? `${m.label}  —  ${m.module}` : m.label));
    const cur = current ? this.runner.resolve(current) : null;
    const hit = list.find((m) => m.id === cur);
    if (hit) sel.value = hit.kind === 'steps' ? hit.label : hit.id;
    const ok = await dialog(title, h('div', { class: 'field' }, h('label', null, 'Macro'), sel), [
      { label: 'OK', value: true, primary: true },
      { label: 'Cancel', value: false },
    ]);
    return ok ? sel.value : null;
  }

  /** Lets people press a key combination for a macro. */
  async assignShortcutFor(name: string | null): Promise<void> {
    const macro = name ?? (await this.pickMacro('Shortcut for which macro?'));
    if (!macro) return;
    const project = this.ensureProject();
    const existing = project.shortcuts.find((s) => s.macro === macro);
    let keys = existing?.keys ?? '';
    const box = h('input', { class: 'input mono shortcut-box', readOnly: true, value: keys || 'Press keys…' }) as HTMLInputElement;
    const warn = h('div', { class: 'hint' }, 'Use Ctrl with a letter, like Ctrl+Shift+M. Ctrl+Letter shortcuts also work in Excel.');
    box.addEventListener('keydown', (e) => {
      if (['Control', 'Shift', 'Alt', 'Meta', 'Tab'].includes(e.key)) return;
      e.preventDefault();
      e.stopPropagation();
      if (e.key === 'Escape') return;
      if (e.key === 'Backspace' || e.key === 'Delete') {
        keys = '';
        box.value = 'No shortcut';
        warn.textContent = 'The shortcut will be removed.';
        return;
      }
      const k = eventKey(e);
      if (!/^(Ctrl|Alt)\+/.test(k)) {
        warn.textContent = 'Include Ctrl or Alt so typing in cells still works.';
        return;
      }
      keys = k;
      box.value = k;
      const taken = this.app.commands.match(e);
      const other = project.shortcuts.find((s) => s.keys.toLowerCase() === k.toLowerCase() && s.macro !== macro);
      warn.textContent = other ? `${k} already runs "${other.macro}" — it will be moved to this macro.` : taken ? `${k} is normally "${taken.title}". In this workbook it will run the macro instead.` : `${k} will run the macro.`;
    });
    setTimeout(() => box.focus(), 0);
    const body = h('div', null, h('p', null, `Click the box and press the keys that should run "${macro}".`), box, warn);
    const ok = await dialog('Keyboard shortcut', body, [
      { label: 'Save', value: true, primary: true },
      { label: 'Cancel', value: false },
    ]);
    if (!ok) return;
    project.shortcuts = project.shortcuts.filter((s) => s.macro !== macro && (!keys || s.keys.toLowerCase() !== keys.toLowerCase()));
    if (keys) project.shortcuts.push({ macro, keys });
    project.changed(/^Ctrl\+(Shift\+)?[A-Z]$/i.test(keys) || /^Ctrl\+(Shift\+)?[A-Z]$/i.test(existing?.keys ?? ''));
    this.app.doc.touch();
    toast(keys ? `${keys} now runs "${macro}".` : `Removed the shortcut for "${macro}".`, { kind: 'success' });
  }

  /** Puts a button that runs a macro next to the selected cell. */
  async addButtonFor(name: string | null): Promise<void> {
    const macro = name ?? (await this.pickMacro('Which macro should the button run?'));
    if (!macro) return;
    const label = await promptDialog('Button text', 'Text on the button', macro.replace(/^.*\./, '').replace(/_/g, ' '), (v) => (v.trim() ? null : 'The button needs some text.'));
    if (!label) return;
    const app = this.app;
    const sheet = app.sheet;
    const a = app.sel.active;
    const button = { id: 'btn-' + Math.random().toString(36).slice(2, 9), label: label.trim(), macro, anchor: { row: a.row, col: a.col, dx: 4, dy: 3, w: Math.max(96, label.trim().length * 8 + 44), h: 30 } };
    app.transact('Add button', (tx) => tx.setProp(sheet, 'buttons', [...sheet.buttons, button]));
    this.ensureProject();
    toast('Button added. Drag it to move it; right-click to change it.', { kind: 'success' });
  }

  /** Lists and adds "run automatically" triggers. */
  async triggersDialog(): Promise<void> {
    const project = this.ensureProject();
    const EVENTS: [string, string][] = [
      ['open', 'When the workbook opens'],
      ['change', 'When cells change'],
      ['activate', 'When a sheet is shown'],
      ['beforeSave', 'Before the workbook is saved'],
    ];
    const listEl = h('div', { class: 'trigger-list' });
    const renderList = () => {
      listEl.innerHTML = '';
      if (!project.triggers.length) listEl.append(h('div', { class: 'hint' }, 'Nothing runs automatically yet.'));
      for (const t of project.triggers) {
        const what = EVENTS.find(([e]) => e === t.event)?.[1] ?? t.event;
        const where = [t.sheet ? `on ${t.sheet}` : '', t.range ? `in ${t.range}` : ''].filter(Boolean).join(' ');
        listEl.append(
          h(
            'div',
            { class: 'trigger-row' },
            icon(L.Zap, 14),
            h('span', null, `${what}${where ? ' ' + where : ''}: run "${t.macro}"`),
            h('button', { class: 'btn', 'aria-label': 'Remove', onclick: () => {
              project.triggers = project.triggers.filter((x) => x !== t);
              project.changed();
              this.app.doc.touch();
              renderList();
            } }, icon(L.Trash2, 14)),
          ),
        );
      }
    };
    renderList();
    const macroSel = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const m of this.runner.list()) macroSel.append(h('option', { value: m.kind === 'steps' ? m.label : m.id }, m.label));
    const eventSel = h('select', { class: 'input' }) as HTMLSelectElement;
    for (const [v, l] of EVENTS) eventSel.append(h('option', { value: v }, l));
    const sheetSel = h('select', { class: 'input' }) as HTMLSelectElement;
    sheetSel.append(h('option', { value: '' }, 'Any sheet'));
    for (const s of this.app.doc.wb.sheets) sheetSel.append(h('option', { value: s.name }, s.name));
    const rangeIn = h('input', { class: 'input mono', placeholder: 'Any cell (or e.g. B2:B50)' }) as HTMLInputElement;
    const whereRow = h('div', { class: 'row' }, sheetSel, rangeIn);
    const syncWhere = () => {
      whereRow.hidden = eventSel.value === 'open' || eventSel.value === 'beforeSave';
      rangeIn.hidden = eventSel.value !== 'change';
    };
    eventSel.addEventListener('change', syncWhere);
    syncWhere();
    const add = h('button', { class: 'btn outline', onclick: () => {
      if (!macroSel.value) return;
      project.triggers.push({ id: 't' + Math.random().toString(36).slice(2, 9), macro: macroSel.value, event: eventSel.value as never, sheet: whereRow.hidden ? undefined : sheetSel.value || undefined, range: rangeIn.hidden ? undefined : rangeIn.value.trim() || undefined });
      project.changed();
      this.app.doc.touch();
      renderList();
    } }, icon(L.Plus, 14), 'Add');
    const body = h(
      'div',
      { class: 'triggers' },
      h('p', null, 'Run a macro by itself when something happens. (VBA code can also use events like Workbook_Open.)'),
      listEl,
      h('div', { class: 'field-label', style: { marginTop: '12px' } }, 'Add'),
      h('div', { class: 'row' }, eventSel, macroSel),
      whereRow,
      h('div', { class: 'row', style: { justifyContent: 'flex-end' } }, add),
    );
    if (!this.runner.list().length) {
      add.disabled = true;
      macroSel.append(h('option', { value: '' }, 'No macros yet'));
    }
    await dialog('Run automatically', body, [{ label: 'Done', value: true, primary: true }], { wide: true });
  }

  /** The Run macro menu for the ribbon. */
  runMenu(): { label: string; icon?: L.IconNode; hint?: string; run?: () => void; disabled?: boolean }[] {
    const list = this.runner.list();
    const keys = new Map((this.project?.shortcuts ?? []).map((s) => [this.runner.resolve(s.macro) ?? s.macro, s.keys]));
    if (!list.length) return [{ label: 'No macros yet', disabled: true }];
    return list.map((m) => ({ label: m.label, icon: m.kind === 'steps' ? L.ListChecks : L.Play, hint: keys.get(m.id), run: () => void this.run(m.kind === 'steps' ? m.label : m.id) }));
  }

  // ---------- what VBA code can ask of the app ----------

  private ui(): MacroUI {
    const app = this.app;
    return {
      msgBox: (prompt, buttons, title) => {
        const set = MSGBOX_BUTTONS[buttons & 7] ?? MSGBOX_BUTTONS[0];
        const def = Math.min(set.length - 1, (buttons >> 8) & 3);
        const cancel = set.find(([, v]) => v === 2)?.[1] ?? (set.length === 1 ? set[0][1] : set[set.length - 1][1]);
        const kind = buttons & 0x70;
        const body = h('div', { class: 'msgbox' }, kind ? icon(kind === 16 ? L.CircleX : kind === 32 ? L.CircleHelp : kind === 48 ? L.TriangleAlert : L.Info, 28, `mb-icon k${kind}`) : null, h('p', null, prompt));
        return dialog<number>(title || 'Gridline', body, set.map(([label, value], i) => ({ label, value, primary: i === def })), { cancelValue: cancel });
      },
      inputBox: (prompt, title, dflt) => {
        return promptDialog(title || 'Gridline', prompt, dflt);
      },
      print: (t) => this.log(t),
      status: (t) => app.statusbar.hold(t),
      selection: () => ({ sheet: app.sheet, ranges: app.sel.ranges.map((r) => ({ ...r })), active: { ...app.sel.active } }),
      select: (sheet, ranges, active) => {
        if (sheet !== app.sheet) app.setActiveSheet(sheet.id, false);
        app.sel.restore({ sheetId: sheet.id, ranges, active });
        app.onSelectionChanged();
      },
      activate: (sheet) => {
        if (sheet !== app.sheet || app.doc.wb.activeSheetId !== sheet.id) app.setActiveSheet(sheet.id);
      },
      save: () => app.save({ fromMacro: true }),
      measure: (sheet, axis, index) => (sheet === app.sheet ? (axis === 'col' ? app.grid.measureColumn(index) : app.grid.measureRow(index)) : (undefined as unknown as number)),
      viewChanged: () => {
        app.grid.relayout();
        app.statusbar.refresh();
      },
    };
  }
}
