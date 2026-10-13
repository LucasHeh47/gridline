// The macros stored in a workbook: VBA modules (kept in the file's vbaProject.bin), plus Gridline's own
// step macros, shortcuts and automatic triggers (kept in a hidden part of the same file).
import type { Step } from './steps';

/** 'form' modules (UserForms) are kept in the file but their designs can't be shown or run. */
export type ModuleKind = 'standard' | 'document' | 'class' | 'form';

export interface VbaModule {
  /** Module name (Module1, ThisWorkbook, Sheet1, MyClass). */
  name: string;
  kind: ModuleKind;
  /** Code as shown in the editor (without the hidden "Attribute VB_…" header lines). */
  code: string;
  /** Hidden attribute lines from the file, written back when saving. */
  attrs?: { header: string[]; procAttrs: Record<string, string[]> };
}

/** A macro built from steps instead of code. Runs by compiling to VBA. */
export interface StepMacro {
  id: string;
  name: string;
  description?: string;
  steps: Step[];
}

export interface MacroShortcut {
  /** Macro name ("Module1.Main", "Main" or a step macro name). */
  macro: string;
  /** "Ctrl+Shift+M" style. */
  keys: string;
}

export type TriggerEvent = 'open' | 'change' | 'activate' | 'beforeSave' | 'timer';

/** Runs a macro automatically (for step macros; VBA uses Workbook_Open and friends). */
export interface MacroTrigger {
  id: string;
  macro: string;
  event: TriggerEvent;
  /** Sheet name for change/activate ("" = any sheet). */
  sheet?: string;
  /** Only changes inside this range (A1 style) fire the trigger. */
  range?: string;
  /** Minutes between runs for 'timer'. */
  every?: number;
}

/** A clickable button drawn over the grid that runs a macro. */
export interface SheetButton {
  id: string;
  label: string;
  macro: string;
  color?: string;
  /** Top-left cell plus pixel offset, and size in px (at 100% zoom), like charts. */
  anchor: { row: number; col: number; dx: number; dy: number; w: number; h: number };
}

export class MacroProject {
  modules: VbaModule[] = [];
  stepMacros: StepMacro[] = [];
  shortcuts: MacroShortcut[] = [];
  triggers: MacroTrigger[] = [];
  /** VBA project name (shown in Excel's editor). */
  projectName = 'VBAProject';
  /** The workbook's document module (ThisWorkbook; other languages of Excel use other names). */
  workbookModule = 'ThisWorkbook';
  /** The vbaProject.bin the workbook was opened with; saved back as-is while the code is unchanged. */
  original: Uint8Array | null = null;
  /** True once any VBA module was added, removed or edited (the binary must be rebuilt). */
  codeChanged = false;
  /** Bumped on every change, so caches (compiled programs, menus) know to refresh. */
  version = 0;

  get isEmpty(): boolean {
    return !this.hasCode && !this.stepMacros.length;
  }

  /** Whether any module contains code (document modules usually start empty). */
  get hasCode(): boolean {
    return this.modules.some((m) => /\S/.test(stripOptions(m.code)));
  }

  /** Whether a standard module declares a public Function with this name (one formulas can call). */
  definesFunction(name: string): boolean {
    const re = new RegExp(`^[ \\t]*(?:Public[ \\t]+)?Function[ \\t]+${name.replace(/[^A-Za-z0-9_]/g, '')}\\b`, 'im');
    return this.modules.some((m) => m.kind === 'standard' && re.test(m.code));
  }

  module(name: string): VbaModule | undefined {
    const n = name.toLowerCase();
    return this.modules.find((m) => m.name.toLowerCase() === n);
  }

  /** A name like Module1, Module2… that isn't taken. */
  uniqueModuleName(base = 'Module'): string {
    let i = 1;
    while (this.module(`${base}${i}`)) i++;
    return `${base}${i}`;
  }

  setCode(name: string, code: string): void {
    const m = this.module(name);
    if (!m || m.code === code) return;
    m.code = code;
    this.changed(true);
  }

  addModule(kind: ModuleKind, name?: string, code = ''): VbaModule {
    const m: VbaModule = { name: name ?? this.uniqueModuleName(kind === 'class' ? 'Class' : 'Module'), kind, code };
    this.modules.push(m);
    this.changed(true);
    return m;
  }

  removeModule(name: string): void {
    const i = this.modules.findIndex((m) => m.name.toLowerCase() === name.toLowerCase());
    if (i < 0) return;
    this.modules.splice(i, 1);
    this.changed(true);
  }

  changed(code = false): void {
    if (code) this.codeChanged = true;
    this.version++;
  }
}

/** Code with Option lines and blank lines removed (to tell if a module is really empty). */
function stripOptions(code: string): string {
  return code
    .split(/\r?\n/)
    .filter((l) => !/^\s*(option\s+\w+.*|'.*)?$/i.test(l))
    .join('\n');
}

/** A valid VBA identifier for a macro name typed by the user. */
export function toIdentifier(name: string): string {
  let id = name.replace(/[^A-Za-z0-9_]+/g, '_').replace(/^_+|_+$/g, '');
  if (!id) id = 'Macro';
  if (!/^[A-Za-z]/.test(id)) id = 'M_' + id;
  return id.slice(0, 60);
}
