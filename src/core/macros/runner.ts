// Runs a workbook's macros: compiles VBA modules and step macros into one program, runs a macro as a
// single undo step, and fires event handlers (Workbook_Open, Worksheet_Change…) and step triggers.
import type { Range } from '../address';
import type { Doc } from '../doc';
import type { Sheet } from '../workbook';
import { MacroSession, uniqueCodeName, type MacroUI } from './host';
import { MacroProject, type MacroTrigger, type StepMacro, type TriggerEvent } from './project';
import { RangeObj } from './range';
import { compileStepMacro, macroProcName } from './steps';
import { VbaProgram, type CompileError, type MacroInfo, type ModuleSource } from './vba/interpreter';
import { StopSignal, VbError, type Value } from './vba/values';

/** Name of the hidden module holding compiled step macros (never saved to the file). */
export const STEPS_MODULE = 'GridlineSteps';

export interface RunResult {
  ok: boolean;
  stopped?: boolean;
  error?: { message: string; number: number; module: string; proc: string; line: number };
  ms: number;
}

export interface MacroEntry {
  /** What to pass to run(): "Module1.Main" or a step macro's procedure name. */
  id: string;
  /** Shown to people: "Main" or "Monthly report". */
  label: string;
  kind: 'vba' | 'steps';
  module?: string;
  stepMacro?: StepMacro;
}

/** Gives every sheet a VBA code name (Sheet1, Sheet2…) and the project a document module per sheet. */
export function ensureDocumentModules(doc: Doc): MacroProject {
  const wb = doc.wb;
  const project = (wb.macros ??= new MacroProject());
  for (const sheet of wb.sheets) if (!sheet.codeName) sheet.codeName = uniqueCodeName(wb.sheets);
  if (!project.module(project.workbookModule)) project.modules.unshift({ name: project.workbookModule, kind: 'document', code: '' });
  for (const sheet of wb.sheets) {
    if (!project.module(sheet.codeName!)) project.modules.push({ name: sheet.codeName!, kind: 'document', code: '' });
  }
  return project;
}

export class MacroRunner {
  private cache: { version: number; program: VbaProgram; stepErrors: CompileError[] } | null = null;
  private project: MacroProject | null = null;
  private current: MacroSession | null = null;
  /** Debug.Print output and run messages for the editor's output pane. */
  onPrint: (text: string) => void = () => undefined;

  constructor(
    private doc: () => Doc,
    private ui: () => MacroUI,
  ) {}

  get running(): boolean {
    return !!this.current;
  }

  /** The compiled program for the workbook's current macros. */
  program(): VbaProgram {
    const project = this.doc().wb.macros ?? new MacroProject();
    if (this.cache && this.project === project && this.cache.version === project.version) return this.cache.program;
    const sources: ModuleSource[] = project.modules.map((m) => ({ name: m.name, code: m.code, kind: m.kind === 'form' ? 'class' : m.kind }));
    const stepErrors: CompileError[] = [];
    if (project.stepMacros.length) {
      const code = project.stepMacros
        .map((m) => {
          const out = compileStepMacro(macroProcName(m.name), m.steps, m.description);
          for (const p of out.problems) stepErrors.push({ module: m.name, line: 0, message: p });
          return out.code;
        })
        .join('\r\n\r\n');
      sources.push({ name: STEPS_MODULE, code, kind: 'standard' });
    }
    const program = new VbaProgram(sources);
    this.cache = { version: project.version, program, stepErrors };
    this.project = project;
    return program;
  }

  /** Syntax errors in the VBA modules (and problems in step macros). */
  errors(): CompileError[] {
    const program = this.program();
    return [...program.errors, ...(this.cache?.stepErrors ?? [])];
  }

  /** Macros people can run: public Subs without arguments, plus step macros. */
  list(): MacroEntry[] {
    const project = this.doc().wb.macros;
    if (!project) return [];
    const out: MacroEntry[] = [];
    const vba: MacroInfo[] = this.program().macros().filter((m) => m.runnable && m.module !== STEPS_MODULE);
    for (const m of vba) {
      const mod = project.module(m.module);
      if (mod?.kind === 'document' && /^(workbook|worksheet)_/i.test(m.name)) continue; // event handlers
      out.push({ id: `${m.module}.${m.name}`, label: m.name, kind: 'vba', module: m.module });
    }
    for (const sm of project.stepMacros) out.push({ id: `${STEPS_MODULE}.${macroProcName(sm.name)}`, label: sm.name, kind: 'steps', stepMacro: sm });
    return out;
  }

  /** Finds a macro by what a button or shortcut stores ("Main", "Module1.Main" or a step macro name). */
  resolve(name: string): string | null {
    const list = this.list();
    const n = name.trim().toLowerCase();
    const hit =
      list.find((m) => m.id.toLowerCase() === n) ??
      list.find((m) => m.label.toLowerCase() === n) ??
      list.find((m) => m.id.toLowerCase().endsWith('.' + macroProcName(name).toLowerCase()));
    return hit?.id ?? null;
  }

  /** Runs a macro. Everything it changes is one undo step. */
  async run(name: string, label?: string, args: Value[] = []): Promise<RunResult> {
    const t0 = performance.now();
    if (this.current) return { ok: false, ms: 0, error: { message: 'Another macro is still running.', number: 0, module: '', proc: '', line: 0 } };
    const program = this.program();
    const errs = program.errors;
    if (errs.length) {
      const e = errs[0];
      return { ok: false, ms: 0, error: { message: e.message, number: 0, module: e.module, proc: '', line: e.line } };
    }
    const id = this.resolve(name) ?? name;
    const shown = label ?? this.list().find((m) => m.id === id)?.label ?? name;
    return this.session(`Run macro: ${shown}`, t0, (s) => program.run(s, id, args));
  }

  private async session(label: string, t0: number, fn: (s: MacroSession) => Promise<unknown>): Promise<RunResult> {
    const doc = this.doc();
    ensureDocumentModules(doc);
    const ui = this.ui();
    const s = new MacroSession(doc, { ...ui, print: (t) => this.onPrint(t) }, this.program(), label);
    this.current = s;
    try {
      await fn(s);
      return { ok: true, ms: performance.now() - t0 };
    } catch (e) {
      if (e instanceof StopSignal) return { ok: false, stopped: true, ms: performance.now() - t0 };
      if (e instanceof VbError) return { ok: false, ms: performance.now() - t0, error: { message: e.message, number: e.number, module: e.module, proc: e.proc, line: e.line } };
      return { ok: false, ms: performance.now() - t0, error: { message: (e as Error)?.message ?? String(e), number: 0, module: '', proc: '', line: 0 } };
    } finally {
      this.current = null;
      s.finish();
    }
  }

  stop(): void {
    this.current?.stop();
  }

  private wbModule(): string {
    return this.doc().wb.macros?.workbookModule ?? 'ThisWorkbook';
  }

  // ---------- events ----------

  /** Whether the workbook has any code for this event (to avoid starting sessions needlessly). */
  hasHandlers(): boolean {
    const p = this.doc().wb.macros;
    return !!p && (p.triggers.length > 0 || p.modules.some((m) => m.kind === 'document' && /\b(workbook|worksheet)_\w+/i.test(m.code)));
  }

  /** Workbook_Open plus step macros set to run when the file opens. */
  async fireOpen(): Promise<RunResult[]> {
    const out: RunResult[] = [];
    const r = await this.event(this.wbModule(), 'Workbook_Open', () => []);
    if (r) out.push(r);
    out.push(...(await this.fireTriggers('open', null)));
    return out;
  }

  async fireActivate(sheet: Sheet): Promise<RunResult[]> {
    const out: RunResult[] = [];
    if (sheet.codeName) {
      const r = await this.event(sheet.codeName, 'Worksheet_Activate', () => []);
      if (r) out.push(r);
    }
    const w = await this.event(this.wbModule(), 'Workbook_SheetActivate', (s) => [s.sheetObj(sheet)]);
    if (w) out.push(w);
    out.push(...(await this.fireTriggers('activate', sheet)));
    return out;
  }

  /** After people edit cells (not after macros). */
  async fireChange(sheet: Sheet, ranges: Range[]): Promise<RunResult[]> {
    const out: RunResult[] = [];
    const target = (s: MacroSession) => new RangeObj(s, sheet, ranges);
    if (sheet.codeName) {
      const r = await this.event(sheet.codeName, 'Worksheet_Change', (s) => [target(s)]);
      if (r) out.push(r);
    }
    const w = await this.event(this.wbModule(), 'Workbook_SheetChange', (s) => [s.sheetObj(sheet), target(s)]);
    if (w) out.push(w);
    out.push(...(await this.fireTriggers('change', sheet, ranges)));
    return out;
  }

  async fireSelectionChange(sheet: Sheet, ranges: Range[]): Promise<RunResult | null> {
    if (!sheet.codeName) return null;
    return this.event(sheet.codeName, 'Worksheet_SelectionChange', (s) => [new RangeObj(s, sheet, ranges)]);
  }

  /** Workbook_BeforeSave; returns false when the code set Cancel = True. */
  async fireBeforeSave(saveAs: boolean): Promise<{ cancel: boolean; results: RunResult[] }> {
    const results: RunResult[] = [];
    const cancel = { v: false as Value };
    const r = await this.event(this.wbModule(), 'Workbook_BeforeSave', () => [saveAs, cancel.v], (refs) => (cancel.v = refs[1]?.v ?? false));
    if (r) results.push(r);
    results.push(...(await this.fireTriggers('beforeSave', null)));
    return { cancel: cancel.v === true || (typeof cancel.v === 'number' && cancel.v !== 0), results };
  }

  /** Step macros whose trigger matches. */
  async fireTriggers(event: TriggerEvent, sheet: Sheet | null, ranges?: Range[]): Promise<RunResult[]> {
    const project = this.doc().wb.macros;
    if (!project?.triggers.length) return [];
    const out: RunResult[] = [];
    for (const t of project.triggers.filter((x) => x.event === event)) {
      if (!this.triggerMatches(t, sheet, ranges)) continue;
      out.push(await this.run(t.macro));
    }
    return out;
  }

  private triggerMatches(t: MacroTrigger, sheet: Sheet | null, ranges?: Range[]): boolean {
    if (t.sheet && sheet && t.sheet.toLowerCase() !== sheet.name.toLowerCase()) return false;
    if (t.range && ranges) {
      const doc = this.doc();
      try {
        const s = new MacroSession(doc, this.ui(), this.program(), '');
        const r = s.parseAddress(t.range, sheet ?? doc.wb.activeSheet);
        s.live.finish();
        if (r.sheet !== sheet) return false;
        return ranges.some((a) => r.areas.some((b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2));
      } catch {
        return false;
      }
    }
    return true;
  }

  /** Runs an event handler in a document module, if it exists. */
  private async event(module: string, proc: string, args: (s: MacroSession) => Value[], after?: (refs: { v: Value }[]) => void): Promise<RunResult | null> {
    if (this.current) return null; // macros don't trigger events
    const project = this.doc().wb.macros;
    const mod = project?.module(module);
    if (!mod || !new RegExp(`\\bsub\\s+${proc}\\b`, 'i').test(mod.code)) return null;
    const program = this.program();
    if (program.errors.length) return null;
    return this.session(`Macro: ${proc}`, performance.now(), async (s) => {
      const refs = args(s).map((v): { v: Value } => ({ v }));
      // (event arguments are variables the handler may change, like Cancel)
      await program.runEvent(s, module, proc, refs);
      after?.(refs);
    });
  }
}

export type { MacroUI };
