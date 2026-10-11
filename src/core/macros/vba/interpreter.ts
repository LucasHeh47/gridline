// The VBA interpreter. It runs code asynchronously so message boxes can wait for the user and a
// long-running loop can be stopped without freezing the app.
import type { Arg, DimVar, Expr, ModuleAst, Param, Proc, Stmt } from './ast';
import { BUILTINS, CONSTANTS, MISSING, type BuiltinCtx } from './builtins';
import { VbaSyntaxError } from './lexer';
import { parseModule } from './parser';
import { compare, EMPTY, isObject, like, NULL, StopSignal, toBool, toInt, toNum, toStr, typeName, VbArray, VbDate, VbError, type Value, type VbObject } from './values';

export interface ModuleSource {
  name: string;
  code: string;
  /** standard modules, document modules (ThisWorkbook, sheets) and class modules. */
  kind: 'standard' | 'document' | 'class';
}

/** What the program needs from its surroundings: the workbook objects and the UI. */
export interface VbaHost {
  /** The Application object. */
  application: VbObject;
  /** Members usable without "Application." (Range, Cells, ActiveSheet…). For sheet modules, `me` is that sheet. */
  global(name: string, args: Value[], named: Record<string, Value> | undefined, me: VbObject | null): Value | Promise<Value>;
  /** Whether `name` is a host global (so Range(…) is a call, not an undeclared variable). */
  isGlobal(name: string): boolean;
  /** The sheet object for a document module (Sheet1, ThisWorkbook). */
  documentObject(moduleName: string): VbObject | null;
  msgBox(prompt: string, buttons: number, title: string): Promise<number>;
  inputBox(prompt: string, title: string, dflt: string): Promise<string | null>;
  print(text: string): void;
  /** Called every few thousand steps; may throw StopSignal to cancel. */
  tick(): Promise<void>;
  /** Evaluates a spreadsheet formula (for [A1] and Evaluate). */
  evaluate(text: string): Value;
}

interface Ref {
  v: Value;
  type?: string;
  /** A value computed for a call (not a variable), so the callee can convert it freely. */
  temp?: boolean;
}

interface Module {
  name: string;
  kind: ModuleSource['kind'];
  ast: ModuleAst;
  vars: Map<string, Ref>;
  consts: Map<string, Value>;
  procs: Map<string, Proc>;
  initialized: boolean;
}

interface Frame {
  module: Module;
  proc: Proc;
  locals: Map<string, Ref>;
  statics?: Map<string, Ref>;
  ret: Ref;
  withStack: Value[];
  onError: 'none' | 'resumenext' | 'label';
  errLabel?: string;
  inHandler: boolean;
  me: VbObject | null;
}

type Signal = undefined | 'exitfor' | 'exitdo' | 'exitproc';

class GotoSignal {
  constructor(
    readonly label: string,
    readonly line: number,
  ) {}
}

class ResumeSignal {
  /** 'next', 'retry' (plain Resume) or a label. */
  constructor(readonly target: string) {}
}

const MAX_DEPTH = 500;

/** Default values for declared types. */
function initial(type: string | undefined): Value {
  switch (type) {
    case 'integer':
    case 'long':
    case 'longlong':
    case 'longptr':
    case 'double':
    case 'single':
    case 'currency':
    case 'byte':
    case 'decimal':
      return 0;
    case 'string':
      return '';
    case 'boolean':
      return false;
    case 'date':
      return new VbDate(0);
    case 'variant':
    case undefined:
      return EMPTY;
    default:
      return null; // object types start as Nothing
  }
}

/** Converts a value for a typed variable (Dim x As Long: x = "5" → 5). */
function coerce(type: string | undefined, v: Value): Value {
  switch (type) {
    case 'integer':
      return toInt(v, 32767);
    case 'long':
    case 'longlong':
    case 'longptr':
      return toInt(v, type === 'long' ? 2147483647 : Number.MAX_SAFE_INTEGER);
    case 'byte':
      return toInt(v, 255);
    case 'double':
    case 'single':
    case 'decimal':
      return toNum(v);
    case 'currency':
      return Math.round(toNum(v) * 10000) / 10000;
    case 'string':
      return toStr(v);
    case 'boolean':
      return toBool(v);
    case 'date':
      return v instanceof VbDate ? v : new VbDate(toNum(v));
    default:
      return v;
  }
}

// ---------- Collection & Dictionary ----------

export class VbCollection implements VbObject {
  readonly vbType = 'Collection';
  private list: { key?: string; v: Value }[] = [];

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    switch (name) {
      case 'count':
        return this.list.length;
      case 'item':
        return this.item(args[0]);
      case 'add': {
        const v = args[0] ?? named?.item;
        const key = args[1] !== undefined && args[1] !== (MISSING as unknown) ? toStr(args[1]) : named?.key !== undefined ? toStr(named.key) : undefined;
        if (key !== undefined && this.list.some((x) => x.key?.toLowerCase() === key.toLowerCase())) throw new VbError(457, `The key "${key}" is already in the collection.`);
        const entry: { key?: string; v: Value } = { key, v };
        const before = args[2] ?? named?.before;
        const after = args[3] ?? named?.after;
        if (before !== undefined && before !== (MISSING as unknown)) this.list.splice(this.indexOf(before), 0, entry);
        else if (after !== undefined && after !== (MISSING as unknown)) this.list.splice(this.indexOf(after) + 1, 0, entry);
        else this.list.push(entry);
        return EMPTY;
      }
      case 'remove':
        this.list.splice(this.indexOf(args[0]), 1);
        return EMPTY;
    }
    throw new VbError(438, `A Collection doesn't have "${name}".`);
  }

  private indexOf(k: Value): number {
    if (typeof k === 'string') {
      const i = this.list.findIndex((x) => x.key?.toLowerCase() === k.toLowerCase());
      if (i < 0) throw new VbError(5, `There's no item with the key "${k}".`);
      return i;
    }
    const i = Math.round(toNum(k)) - 1;
    if (i < 0 || i >= this.list.length) throw new VbError(9, `Subscript out of range: the collection has ${this.list.length} items.`);
    return i;
  }

  private item(k: Value): Value {
    return this.list[this.indexOf(k)].v;
  }

  set(name: string): void {
    throw new VbError(438, `Can't set "${name}" on a Collection.`);
  }

  dflt(args: Value[]): Value {
    return this.item(args[0]);
  }

  items(): Value[] {
    return this.list.map((x) => x.v);
  }
}

export class VbDictionary implements VbObject {
  readonly vbType = 'Dictionary';
  private map = new Map<string, { key: Value; v: Value }>();
  textCompare = false;

  private k(key: Value): string {
    const s = typeof key === 'string' ? 's' + key : typeof key === 'number' ? 'n' + key : 'x' + toStr(key);
    return this.textCompare ? s.toLowerCase() : s;
  }

  get(name: string, args: Value[]): Value {
    switch (name) {
      case 'count':
        return this.map.size;
      case 'add': {
        const key = this.k(args[0]);
        if (this.map.has(key)) throw new VbError(457, `The key "${toStr(args[0])}" is already in the dictionary.`);
        this.map.set(key, { key: args[0], v: args[1] });
        return EMPTY;
      }
      case 'exists':
        return this.map.has(this.k(args[0]));
      case 'item':
        return this.dflt(args);
      case 'keys':
      case 'items': {
        const list = [...this.map.values()].map((e) => (name === 'keys' ? e.key : e.v));
        const arr = new VbArray([{ lb: 0, ub: list.length - 1 }]);
        arr.data = list;
        return arr;
      }
      case 'remove':
        if (!this.map.delete(this.k(args[0]))) throw new VbError(32811, `The key "${toStr(args[0])}" isn't in the dictionary.`);
        return EMPTY;
      case 'removeall':
        this.map.clear();
        return EMPTY;
      case 'comparemode':
        return this.textCompare ? 1 : 0;
    }
    throw new VbError(438, `A Dictionary doesn't have "${name}".`);
  }

  set(name: string, args: Value[], v: Value): void {
    if (name === 'item') return this.setDflt(args, v);
    if (name === 'comparemode') {
      this.textCompare = toNum(v) === 1;
      return;
    }
    if (name === 'key') {
      const e = this.map.get(this.k(args[0]));
      if (e) {
        this.map.delete(this.k(args[0]));
        this.map.set(this.k(v), { key: v, v: e.v });
      }
      return;
    }
    throw new VbError(438, `Can't set "${name}" on a Dictionary.`);
  }

  dflt(args: Value[]): Value {
    // Reading a missing key adds it as Empty, like Scripting.Dictionary
    const key = this.k(args[0]);
    if (!this.map.has(key)) this.map.set(key, { key: args[0], v: EMPTY });
    return this.map.get(key)!.v;
  }

  setDflt(args: Value[], v: Value): void {
    this.map.set(this.k(args[0]), { key: args[0], v });
  }

  items(): Value[] {
    return [...this.map.values()].map((e) => e.key);
  }
}

class ErrObject implements VbObject {
  readonly vbType = 'ErrObject';
  number = 0;
  description = '';
  source = '';
  line = 0;

  clear(): void {
    this.number = 0;
    this.description = '';
    this.source = '';
    this.line = 0;
  }

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    switch (name) {
      case 'number':
        return this.number;
      case 'description':
        return this.description;
      case 'source':
        return this.source;
      case 'clear':
        this.clear();
        return EMPTY;
      case 'raise': {
        const n = toNum(args[0] ?? named?.number ?? 0);
        const desc = args[2] ?? named?.description;
        throw new VbError(n, desc !== undefined && desc !== (MISSING as unknown) ? toStr(desc) : `Error ${n}`);
      }
    }
    throw new VbError(438, `Err doesn't have "${name}".`);
  }

  set(name: string, _args: Value[], v: Value): void {
    if (name === 'number') this.number = toNum(v);
    else if (name === 'description') this.description = toStr(v);
    else if (name === 'source') this.source = toStr(v);
  }

  dflt(): Value {
    return this.number;
  }
}

/** A module used as an object: Module1.DoThing */
class ModuleObject implements VbObject {
  readonly vbType = 'Module';
  constructor(
    private program: VbaProgram,
    private mod: Module,
  ) {}

  async get(name: string, args: Value[]): Promise<Value> {
    const proc = this.mod.procs.get(name);
    if (proc) return this.program.callProc(this.mod, proc, args.map((v) => ({ v })), null);
    const ref = this.mod.vars.get(name);
    if (ref) return ref.v;
    if (this.mod.consts.has(name)) return this.mod.consts.get(name);
    throw new VbError(438, `${this.mod.name} doesn't have "${name}".`);
  }

  set(name: string, _args: Value[], v: Value): void {
    const ref = this.mod.vars.get(name);
    if (!ref) throw new VbError(438, `${this.mod.name} doesn't have a variable "${name}".`);
    ref.v = v;
  }
}

export interface MacroInfo {
  module: string;
  name: string;
  /** Shown in the Run list: public Subs without required arguments. */
  runnable: boolean;
  line: number;
}

/** Errors found while compiling, with module and line. */
export interface CompileError {
  module: string;
  line: number;
  message: string;
}

export class VbaProgram {
  private modules: Module[] = [];
  readonly errors: CompileError[] = [];
  readonly err = new ErrObject();
  private depth = 0;
  private steps = 0;
  /** Set while the program runs; holds the host. */
  private host: VbaHost | null = null;
  private builtinCtx: BuiltinCtx | null = null;
  private stopped = false;

  constructor(sources: ModuleSource[]) {
    for (const src of sources) {
      try {
        const ast = parseModule(src.code);
        const procs = new Map<string, Proc>();
        for (const p of ast.procs) {
          const key = p.name.toLowerCase() + (p.kind === 'let' || p.kind === 'set' ? ':' + p.kind : '');
          procs.set(key, p);
        }
        this.modules.push({ name: src.name, kind: src.kind, ast, vars: new Map(), consts: new Map(), procs, initialized: false });
      } catch (e) {
        if (e instanceof VbaSyntaxError) this.errors.push({ module: src.name, line: e.line, message: e.message });
        else throw e;
      }
    }
  }

  /** All Subs and Functions, for the macro list. */
  macros(): MacroInfo[] {
    const out: MacroInfo[] = [];
    for (const m of this.modules) {
      if (m.kind === 'class') continue;
      for (const p of m.ast.procs) {
        out.push({ module: m.name, name: p.name, runnable: p.kind === 'sub' && !p.isPrivate && p.params.every((x) => x.optional || x.paramArray), line: p.line });
      }
    }
    return out;
  }

  hasProc(module: string | null, name: string): boolean {
    return !!this.find(module, name);
  }

  private find(module: string | null, name: string): { mod: Module; proc: Proc } | null {
    const key = name.toLowerCase();
    for (const m of this.modules) {
      if (module && m.name.toLowerCase() !== module.toLowerCase()) continue;
      const proc = m.procs.get(key);
      if (proc && (module || !proc.isPrivate || this.modules.length === 1)) return { mod: m, proc };
    }
    return null;
  }

  /** Runs a Sub (or Function) by name, e.g. "Module1.Main" or "Main". */
  async run(host: VbaHost, name: string, args: Value[] = []): Promise<Value> {
    const dot = name.lastIndexOf('.');
    const module = dot > 0 ? name.slice(0, dot) : null;
    const proc = dot > 0 ? name.slice(dot + 1) : name;
    const hit = this.find(module, proc) ?? (module ? null : this.findAny(proc));
    if (!hit) throw new VbError(1004, `There's no macro called "${name}".`);
    return this.withHost(host, () => this.callProc(hit.mod, hit.proc, args.map((v) => ({ v })), host.documentObject(hit.mod.name)));
  }

  /** Runs an event handler (Workbook_Open, Worksheet_Change…) if the module has one. Arguments are
   * passed by reference, so a handler can set Cancel = True. */
  async runEvent(host: VbaHost, module: string, name: string, args: { v: Value }[]): Promise<boolean> {
    const mod = this.modules.find((m) => m.name.toLowerCase() === module.toLowerCase());
    const proc = mod?.procs.get(name.toLowerCase());
    if (!mod || !proc) return false;
    // Handlers declared with fewer parameters (Sub Workbook_BeforeSave()) still run
    const refs = args.slice(0, proc.params.length);
    await this.withHost(host, () => this.callProc(mod, proc, refs, host.documentObject(mod.name)));
    return true;
  }

  private findAny(name: string): { mod: Module; proc: Proc } | null {
    const key = name.toLowerCase();
    for (const m of this.modules) {
      const proc = m.procs.get(key);
      if (proc) return { mod: m, proc };
    }
    return null;
  }

  private async withHost<T>(host: VbaHost, fn: () => Promise<T>): Promise<T> {
    const outer = this.host;
    this.host = host;
    this.stopped = false;
    this.builtinCtx = {
      compareText: false,
      optionBase: 0,
      msgBox: (p, b, t) => host.msgBox(p, b, t),
      inputBox: (p, t, d) => host.inputBox(p, t, d),
      createObject: (id) => this.createObject(id),
    };
    try {
      return await fn();
    } catch (e) {
      // End finishes the program normally; Stop (cancel) is passed on so the app can say so
      if (e instanceof StopSignal && e.reason === 'end') return EMPTY as T;
      throw e;
    } finally {
      this.host = outer;
    }
  }

  /** Stops the running code at the next statement. */
  stop(): void {
    this.stopped = true;
  }

  createObject(progId: string): Value {
    const id = progId.toLowerCase().trim();
    if (id === 'scripting.dictionary') return new VbDictionary();
    if (id === 'vba.collection' || id === 'collection') return new VbCollection();
    throw new VbError(429, `Can't create "${progId}": Gridline macros can only work with the workbook, not other programs or files.`);
  }

  // ---------- modules ----------

  private async initModule(mod: Module): Promise<void> {
    if (mod.initialized) return;
    mod.initialized = true;
    for (const c of mod.ast.consts) mod.consts.set(c.name.toLowerCase(), EMPTY);
    // Constants can refer to earlier constants
    const frame = this.frame(mod, { kind: 'sub', name: '(module)', params: [], body: [], isPrivate: true, line: 0 }, null);
    for (const c of mod.ast.consts) mod.consts.set(c.name.toLowerCase(), await this.eval(c.e, frame));
    for (const v of mod.ast.vars) mod.vars.set(v.dim.name.toLowerCase(), await this.declare(v.dim, frame));
  }

  private frame(module: Module, proc: Proc, me: VbObject | null): Frame {
    return { module, proc, locals: new Map(), ret: { v: proc.kind === 'function' || proc.kind === 'get' ? initial(proc.returnType) : EMPTY, type: proc.returnType }, withStack: [], onError: 'none', inHandler: false, me };
  }

  private async declare(d: DimVar, f: Frame): Promise<Ref> {
    const type = d.type;
    if (d.bounds) {
      if (!d.bounds.length) return { v: new VbArray([]), type: 'array' };
      const dims = [];
      for (const [lo, hi] of d.bounds) dims.push({ lb: lo ? toNum(await this.eval(lo, f)) : f.module.ast.optionBase, ub: toNum(await this.eval(hi, f)) });
      return { v: new VbArray(dims, () => initial(type)), type: 'array' };
    }
    if (d.isNew) return { v: this.newObject(type ?? ''), type };
    return { v: initial(type), type };
  }

  private newObject(cls: string): Value {
    const c = cls.toLowerCase();
    if (c === 'collection' || c === 'vba.collection') return new VbCollection();
    if (c === 'scripting.dictionary' || c === 'dictionary') return new VbDictionary();
    const classMod = this.modules.find((m) => m.kind === 'class' && m.name.toLowerCase() === c);
    if (classMod) return new ClassInstance(this, classMod);
    throw new VbError(429, `Can't create a "${cls}" object.`);
  }

  // ---------- calls ----------

  async callProc(mod: Module, proc: Proc, args: Ref[], me: VbObject | null, named?: Map<string, Ref>, instance?: ClassInstance): Promise<Value> {
    await this.initModule(mod);
    if (this.depth >= MAX_DEPTH) throw new VbError(28, 'Out of stack space: a procedure keeps calling itself.');
    this.depth++;
    const f = this.frame(mod, proc, me);
    // Class fields are visible as variables; the Ref objects are shared, so changes stick
    const locals = instance ? new Map(instance.fields) : f.locals;
    f.locals = locals;
    try {
      // Bind parameters
      let ai = 0;
      for (const p of proc.params) {
        if (p.paramArray) {
          const rest = args.slice(ai).map((r) => r.v);
          const arr = new VbArray([{ lb: 0, ub: rest.length - 1 }]);
          arr.data = rest;
          locals.set(p.name.toLowerCase(), { v: arr });
          ai = args.length;
          continue;
        }
        let ref = named?.get(p.name.toLowerCase()) ?? args[ai];
        ai++;
        if (!ref || ref.v === (MISSING as unknown)) {
          if (!p.optional) throw new VbError(450, `${proc.name} needs a value for "${p.name}".`);
          const v = p.dflt ? await this.eval(p.dflt, f) : p.type && p.type !== 'variant' ? initial(p.type) : (MISSING as unknown as Value);
          ref = { v, type: p.type };
        } else if (p.byVal || ref.temp) {
          // ByVal (or a computed value): the procedure gets its own copy, converted to the declared type
          ref = { v: p.type ? coerceParam(p, ref.v) : ref.v, type: p.type };
        }
        // ByRef variables are shared, so changes are seen by the caller
        locals.set(p.name.toLowerCase(), ref);
      }
      if (args.length > ai && !proc.params.some((p) => p.paramArray)) throw new VbError(450, `${proc.name} was given too many values.`);
      await this.execProcBody(f);
      return f.ret.v;
    } catch (e) {
      if (e instanceof VbError && !e.line) {
        e.proc = proc.name;
        e.module = mod.name;
      }
      throw e;
    } finally {
      this.depth--;
    }
  }

  /** Runs a procedure body, handling On Error GoTo labels and GoTo jumps at the top level. */
  private async execProcBody(f: Frame): Promise<void> {
    const body = f.proc.body;
    const labelIndex = (label: string, line: number) => {
      const i = body.findIndex((s) => s.k === 'label' && s.name === label);
      if (i < 0) throw Object.assign(new VbError(1004, `There's no label "${label}" in ${f.proc.name} (labels inside loops or Ifs aren't supported).`), { line });
      return i;
    };
    // Errors and jumps inside loops and Ifs resume at the enclosing top-level statement
    let pc = 0;
    let errorAt = 0;
    while (pc < body.length) {
      try {
        const sig = await this.execStmtGuarded(body[pc], f);
        if (sig === 'exitproc') return;
        pc++;
      } catch (e) {
        if (e instanceof GotoSignal) {
          pc = labelIndex(e.label, e.line);
          continue;
        }
        if (e instanceof ResumeSignal) {
          if (!f.inHandler) throw new VbError(20, 'Resume without an error.');
          f.inHandler = false;
          this.err.clear();
          pc = e.target === 'next' ? errorAt + 1 : e.target === 'retry' ? errorAt : labelIndex(e.target, 0);
          continue;
        }
        if (e instanceof VbError && f.onError === 'label' && !f.inHandler) {
          this.setErr(e);
          f.inHandler = true;
          errorAt = pc;
          pc = labelIndex(f.errLabel!, e.line);
          continue;
        }
        throw e;
      }
    }
  }

  private setErr(e: VbError): void {
    this.err.number = e.number;
    this.err.description = e.message;
    this.err.line = e.line;
    this.err.source = e.module;
  }

  // ---------- statements ----------

  private async execBlock(body: Stmt[], f: Frame): Promise<Signal> {
    for (const s of body) {
      const sig = await this.execStmtGuarded(s, f);
      if (sig) return sig;
    }
    return undefined;
  }

  /** Runs a statement; with On Error Resume Next, errors skip to the next statement. */
  private async execStmtGuarded(s: Stmt, f: Frame): Promise<Signal> {
    if (f.onError !== 'resumenext') return this.execStmt(s, f);
    try {
      return await this.execStmt(s, f);
    } catch (e) {
      if (e instanceof VbError) {
        this.setErr(e);
        return undefined;
      }
      throw e;
    }
  }

  /** Counts work so long runs can yield to the app and be stopped. Called per statement and loop pass. */
  private async step(): Promise<void> {
    if (++this.steps % 2000 === 0) await this.host?.tick();
    if (this.stopped) throw new StopSignal('cancel');
  }

  private async execStmt(s: Stmt, f: Frame): Promise<Signal> {
    await this.step();
    try {
      switch (s.k) {
        case 'dim':
          for (const d of s.vars) {
            const key = d.name.toLowerCase();
            if (s.isStatic) {
              const statics = (f.proc as Proc & { statics?: Map<string, Ref> }).statics ?? ((f.proc as Proc & { statics?: Map<string, Ref> }).statics = new Map());
              if (!statics.has(key)) statics.set(key, await this.declare(d, f));
              f.locals.set(key, statics.get(key)!);
            } else f.locals.set(key, await this.declare(d, f));
          }
          return;
        case 'const':
          for (const c of s.items) f.locals.set(c.name.toLowerCase(), { v: await this.eval(c.e, f), type: 'const' });
          return;
        case 'redim':
          for (const d of s.vars) {
            const dims = [];
            for (const [lo, hi] of d.bounds ?? []) dims.push({ lb: lo ? toNum(await this.eval(lo, f)) : f.module.ast.optionBase, ub: toNum(await this.eval(hi, f)) });
            const ref = this.lookupVar(d.name, f) ?? this.implicit(d.name, f);
            if (ref.v instanceof VbArray) ref.v.resize(dims, s.preserve);
            else ref.v = new VbArray(dims, () => initial(d.type));
            ref.type = 'array';
          }
          return;
        case 'assign': {
          const v = await this.eval(s.e, f);
          if (s.set) {
            if (v !== null && !isObject(v)) throw new VbError(424, 'Object required: Set needs an object (or Nothing).');
          }
          await this.assign(s.target, s.set ? v : await this.unwrapDefault(v), f, s.set);
          return;
        }
        case 'call':
          await this.evalCall(s.e as Expr & { k: 'call' }, f, true);
          return;
        case 'print': {
          const parts: string[] = [];
          for (const a of s.args) parts.push(toStr(await this.unwrapDefault(await this.eval(a, f))));
          this.host?.print(parts.join(' '));
          return;
        }
        case 'if':
          for (const b of s.branches) if (toBool(await this.value(b.cond, f))) return this.execBlock(b.body, f);
          if (s.elseBody) return this.execBlock(s.elseBody, f);
          return;
        case 'for': {
          const from = toNum(await this.value(s.from, f));
          const to = toNum(await this.value(s.to, f));
          const step = s.step ? toNum(await this.value(s.step, f)) : 1;
          if (step === 0) throw new VbError(5, 'A For loop with Step 0 would never end.');
          let i = from;
          await this.assign(s.v, i, f, false);
          while (step > 0 ? i <= to + 1e-9 : i >= to - 1e-9) {
            await this.step();
            const sig = await this.execBlock(s.body, f);
            if (sig === 'exitfor') break;
            if (sig) return sig;
            i = toNum(await this.value(s.v, f)) + step;
            await this.assign(s.v, i, f, false);
          }
          return;
        }
        case 'foreach': {
          const coll = await this.eval(s.coll, f);
          let list: Value[];
          if (coll instanceof VbArray) {
            if (coll.dims.length === 2) {
              list = [];
              for (let c = coll.dims[1].lb; c <= coll.dims[1].ub; c++) for (let r = coll.dims[0].lb; r <= coll.dims[0].ub; r++) list.push(coll.get([r, c]));
            } else list = [...coll.data];
          } else if (isObject(coll) && coll.items) list = coll.items();
          else throw new VbError(451, "For Each needs a collection, range or array.");
          for (const item of list) {
            await this.step();
            await this.assign(s.v, item, f, isObject(item) || item === null);
            const sig = await this.execBlock(s.body, f);
            if (sig === 'exitfor') break;
            if (sig) return sig;
          }
          return;
        }
        case 'do':
          for (;;) {
            await this.step();
            if (s.pre && toBool(await this.value(s.pre.cond, f)) === s.pre.until) break;
            const sig = await this.execBlock(s.body, f);
            if (sig === 'exitdo') break;
            if (sig) return sig;
            if (s.post && toBool(await this.value(s.post.cond, f)) === s.post.until) break;
          }
          return;
        case 'select': {
          const v = await this.value(s.e, f);
          const text = f.module.ast.compareText;
          for (const c of s.cases) {
            for (const t of c.tests) {
              let hit = false;
              if (t.k === 'value') hit = compare(v, await this.value(t.e, f), text) === 0;
              else if (t.k === 'range') hit = compare(v, await this.value(t.from, f), text) >= 0 && compare(v, await this.value(t.to, f), text) <= 0;
              else hit = cmpOp(t.op, compare(v, await this.value(t.e, f), text));
              if (hit) return this.execBlock(c.body, f);
            }
          }
          if (s.elseBody) return this.execBlock(s.elseBody, f);
          return;
        }
        case 'with': {
          const obj = await this.eval(s.obj, f);
          f.withStack.push(obj);
          try {
            return await this.execBlock(s.body, f);
          } finally {
            f.withStack.pop();
          }
        }
        case 'exit':
          return s.what === 'for' ? 'exitfor' : s.what === 'do' ? 'exitdo' : 'exitproc';
        case 'onerror':
          if (s.mode === 'resumenext') f.onError = 'resumenext';
          else if (s.mode === 'goto0') f.onError = 'none';
          else {
            f.onError = 'label';
            f.errLabel = s.label;
          }
          this.err.clear();
          return;
        case 'resume':
          throw new ResumeSignal(s.next ? 'next' : (s.label ?? 'retry'));
        case 'goto':
          throw new GotoSignal(s.label, s.line);
        case 'label':
          return;
        case 'end':
          throw new StopSignal('end');
      }
    } catch (e) {
      if (e instanceof VbError) {
        if (!e.line) e.line = s.line;
        if (!e.module) {
          e.module = f.module.name;
          e.proc = f.proc.name;
        }
      }
      throw e;
    }
  }

  // ---------- variables ----------

  private lookupVar(name: string, f: Frame): Ref | undefined {
    const key = name.toLowerCase();
    if (key === f.proc.name.toLowerCase() && (f.proc.kind === 'function' || f.proc.kind === 'get')) return f.ret;
    return f.locals.get(key) ?? f.module.vars.get(key) ?? this.publicVar(key, f.module);
  }

  private publicVar(key: string, from: Module): Ref | undefined {
    for (const m of this.modules) {
      if (m === from || m.kind === 'class') continue;
      const decl = m.ast.vars.find((v) => v.dim.name.toLowerCase() === key && !v.isPrivate);
      if (decl) {
        if (!m.initialized) {
          // Public variables of other modules are created on first use
          m.vars.set(key, { v: initial(decl.dim.type), type: decl.dim.type });
        }
        return m.vars.get(key);
      }
    }
    return undefined;
  }

  private implicit(name: string, f: Frame): Ref {
    if (f.module.ast.optionExplicit) throw new VbError(1004, `Variable not defined: "${name}" (add Dim ${name} — Option Explicit is on).`);
    const ref: Ref = { v: EMPTY };
    f.locals.set(name.toLowerCase(), ref);
    return ref;
  }

  private async assign(target: Expr, v: Value, f: Frame, isSet: boolean): Promise<void> {
    switch (target.k) {
      case 'name': {
        const ref = this.lookupVar(target.name, f);
        if (ref) {
          if (ref.type === 'const') throw new VbError(1004, `"${target.name}" is a constant and can't be changed.`);
          ref.v = isSet || ref.type === 'array' ? v : coerce(ref.type, v);
          return;
        }
        // Assigning to an object property without a variable: ActiveCell = 5
        const key = target.name.toLowerCase();
        if (this.host?.isGlobal(key)) {
          const obj = await this.eval(target, f);
          return this.setDefault(obj, [], v, target.name);
        }
        const mod = this.moduleByName(key);
        if (mod) throw new VbError(1004, `Can't assign to the module "${target.name}".`);
        this.implicit(target.name, f).v = v;
        return;
      }
      case 'member': {
        const obj = target.obj ? await this.eval(target.obj, f) : this.withObject(f, target.line);
        if (!isObject(obj)) throw new VbError(obj === null ? 91 : 424, obj === null ? `Object variable not set: can't set .${target.name} on Nothing.` : `Object required: can't set .${target.name} on a ${typeName(obj)}.`);
        await obj.set(target.name.toLowerCase(), [], v);
        return;
      }
      case 'call': {
        const head = target.target;
        const args = await this.argValues(target.args, f);
        if (head.k === 'name') {
          const ref = this.lookupVar(head.name, f);
          if (ref) {
            if (ref.v instanceof VbArray) {
              ref.v.set(args.values.map(toNum), v);
              return;
            }
            if (isObject(ref.v)) return this.setDefault(ref.v, args.values, v, head.name);
            throw new VbError(13, `"${head.name}" isn't an array.`);
          }
        }
        if (head.k === 'member') {
          const obj = head.obj ? await this.eval(head.obj, f) : this.withObject(f, head.line);
          if (!isObject(obj)) throw new VbError(obj === null ? 91 : 424, `Object required before .${head.name}.`);
          const name = head.name.toLowerCase();
          // obj.Prop(args) = v is either a property with arguments (dict.Item(k) = v) or the default
          // member of what Prop returns (ws.Range("A1") = 5, ws.Cells(1, 1) = 5)
          try {
            await obj.set(name, args.values, v);
            return;
          } catch (e) {
            if (!(e instanceof VbError && e.number === 438) || !args.values.length) throw e;
          }
          return this.setDefault(await obj.get(name, args.values, args.named), [], v, head.name);
        }
        // Range("A1") = 5, Cells(1, 1) = 5
        const obj = await this.evalCall(target, f, false);
        return this.setDefault(obj, [], v, head.k === 'name' ? head.name : 'value');
      }
      case 'eval': {
        // [A1] = 5
        if (!this.host) throw new VbError(1004, "This can't be assigned to.");
        return this.setDefault(this.host.evaluate(target.text), [], v, target.text);
      }
      default:
        throw new VbError(1004, "This can't be assigned to.");
    }
  }

  private async setDefault(obj: Value, args: Value[], v: Value, name: string): Promise<void> {
    if (!isObject(obj)) throw new VbError(424, `Object required: "${name}" isn't an object.`);
    if (obj.setDflt) return obj.setDflt(args, v);
    throw new VbError(438, `A ${obj.vbType} can't be assigned a value directly.`);
  }

  private withObject(f: Frame, line: number): Value {
    if (!f.withStack.length) throw Object.assign(new VbError(1004, 'A name starting with "." must be inside a With block.'), { line });
    return f.withStack[f.withStack.length - 1];
  }

  private moduleByName(key: string): Module | undefined {
    return this.modules.find((m) => m.name.toLowerCase() === key);
  }

  // ---------- expressions ----------

  /** Evaluates to a plain value: objects with a default member (like Range) give their value. */
  async value(e: Expr, f: Frame): Promise<Value> {
    return this.unwrapDefault(await this.eval(e, f));
  }

  private async unwrapDefault(v: Value): Promise<Value> {
    if (isObject(v) && v.dflt && v.vbType !== 'Collection' && v.vbType !== 'Dictionary') return v.dflt([]);
    return v;
  }

  async eval(e: Expr, f: Frame): Promise<Value> {
    switch (e.k) {
      case 'lit':
        return e.date ? new VbDate(e.v as number) : e.v;
      case 'special':
        return e.v === 'nothing' ? null : e.v === 'empty' ? EMPTY : e.v === 'null' ? NULL : f.me ?? null;
      case 'name':
        return this.evalName(e, f);
      case 'member': {
        const obj = e.obj ? await this.eval(e.obj, f) : this.withObject(f, e.line);
        return this.member(obj, e.name, [], undefined, e.line);
      }
      case 'call':
        return this.evalCall(e, f, false);
      case 'new':
        return this.newObject(e.cls);
      case 'eval':
        return this.host ? this.host.evaluate(e.text) : EMPTY;
      case 'unary': {
        const v = await this.value(e.e, f);
        if (e.op === '-') {
          if (v === NULL) return NULL;
          if (v instanceof VbDate) return new VbDate(-v.serial);
          return -toNum(v);
        }
        if (v === NULL) return NULL;
        if (typeof v === 'boolean') return !v;
        return ~toInt(v);
      }
      case 'bin':
        return this.binary(e, f);
    }
  }

  private async binary(e: Expr & { k: 'bin' }, f: Frame): Promise<Value> {
    const op = e.op;
    if (op === 'is') {
      const a = await this.eval(e.a, f);
      const b = await this.eval(e.b, f);
      return a === b || (a === null && b === null);
    }
    // And/Or short-circuit on booleans is not VBA behavior, but evaluating both sides is
    const a = await this.value(e.a, f);
    const b = await this.value(e.b, f);
    const text = f.module.ast.compareText;
    try {
      switch (op) {
        case '&':
          return (a === NULL ? '' : toStr(a)) + (b === NULL ? '' : toStr(b));
        case '+':
          if (typeof a === 'string' && typeof b === 'string') return a + b;
          if (a === NULL || b === NULL) return NULL;
          if (a instanceof VbDate || b instanceof VbDate) return new VbDate(toNum(a) + toNum(b));
          if (a === EMPTY && typeof b === 'string') return b;
          if (b === EMPTY && typeof a === 'string') return a;
          return toNum(a) + toNum(b);
        case '-':
          if (a === NULL || b === NULL) return NULL;
          if (a instanceof VbDate && b instanceof VbDate) return a.serial - b.serial;
          if (a instanceof VbDate) return new VbDate(a.serial - toNum(b));
          return toNum(a) - toNum(b);
        case '*':
          return toNum(a) * toNum(b);
        case '/': {
          const d = toNum(b);
          if (d === 0) throw new VbError(11, 'Division by zero.');
          return toNum(a) / d;
        }
        case '\\': {
          const d = toInt(b);
          if (d === 0) throw new VbError(11, 'Division by zero.');
          return Math.trunc(toInt(a) / d);
        }
        case 'mod': {
          const d = toInt(b);
          if (d === 0) throw new VbError(11, 'Division by zero.');
          return toInt(a) % d;
        }
        case '^':
          return Math.pow(toNum(a), toNum(b));
        case '=':
        case '<>':
        case '<':
        case '>':
        case '<=':
        case '>=':
          if (a === NULL || b === NULL) return NULL;
          return cmpOp(op, compare(a, b, text));
        case 'like':
          return like(toStr(a), toStr(b), text);
        case 'and':
          if (typeof a === 'boolean' && typeof b === 'boolean') return a && b;
          return toInt(a) & toInt(b);
        case 'or':
          if (typeof a === 'boolean' && typeof b === 'boolean') return a || b;
          return toInt(a) | toInt(b);
        case 'xor':
          if (typeof a === 'boolean' && typeof b === 'boolean') return a !== b;
          return toInt(a) ^ toInt(b);
        case 'eqv':
          return toBool(a) === toBool(b);
        case 'imp':
          return !toBool(a) || toBool(b);
      }
    } catch (err) {
      if (err instanceof VbError && !err.line) err.line = e.line;
      throw err;
    }
    throw new VbError(5, `Unknown operator ${op}.`);
  }

  private async evalName(e: Expr & { k: 'name' }, f: Frame): Promise<Value> {
    const key = e.name.toLowerCase();
    const ref = this.lookupVar(key, f);
    if (ref) return ref.v;
    if (f.module.consts.has(key)) return f.module.consts.get(key);
    for (const m of this.modules) {
      const c = m.ast.consts.find((x) => x.name.toLowerCase() === key && !x.isPrivate);
      if (c) {
        await this.initModule(m);
        return m.consts.get(key);
      }
    }
    // A Sub/Function called without parentheses
    const proc = f.module.procs.get(key) ?? this.findAny(key)?.proc;
    if (proc) {
      const mod = f.module.procs.has(key) ? f.module : this.findAny(key)!.mod;
      return this.callProc(mod, proc, [], this.host?.documentObject(mod.name) ?? null);
    }
    if (key in CONSTANTS) return CONSTANTS[key];
    if (key === 'err') return this.err;
    if (key === 'debug') return { vbType: 'Debug', get: (n: string, a: Value[]) => (n === 'print' ? (this.host?.print(a.map((x) => toStr(x)).join(' ')), EMPTY) : EMPTY), set: () => undefined } as VbObject;
    if (key in BUILTINS) return BUILTINS[key]([], this.builtinCtx!);
    if (key === 'vba') return { vbType: 'VBA', get: (n: string, a: Value[]) => (BUILTINS[n] ? BUILTINS[n](a, this.builtinCtx!) : EMPTY), set: () => undefined } as VbObject;
    const mod = this.moduleByName(key);
    if (mod) return mod.kind === 'document' ? this.host?.documentObject(mod.name) ?? new ModuleObject(this, mod) : new ModuleObject(this, mod);
    if (this.host?.isGlobal(key)) return this.host.global(key, [], undefined, f.me);
    // An undeclared variable is Empty (unless Option Explicit)
    if (f.module.ast.optionExplicit) throw Object.assign(new VbError(1004, `Variable not defined: "${e.name}".`), { line: e.line });
    return EMPTY;
  }

  private async argValues(args: Arg[], f: Frame): Promise<{ values: Value[]; named?: Record<string, Value> }> {
    const values: Value[] = [];
    let named: Record<string, Value> | undefined;
    for (const a of args) {
      if (!a.expr) {
        values.push(MISSING as unknown as Value);
        continue;
      }
      const v = await this.eval(a.expr, f);
      if (a.name) (named ??= {})[a.name.toLowerCase()] = v;
      else values.push(v);
    }
    return { values, named };
  }

  /** Arguments for a user procedure: plain variables are passed by reference. */
  private async argRefs(args: Arg[], f: Frame): Promise<{ refs: Ref[]; named?: Map<string, Ref> }> {
    const refs: Ref[] = [];
    let named: Map<string, Ref> | undefined;
    for (const a of args) {
      let ref: Ref;
      if (!a.expr) ref = { v: MISSING as unknown as Value };
      else if (a.expr.k === 'name') ref = this.lookupVar(a.expr.name, f) ?? { v: await this.eval(a.expr, f), temp: true };
      else ref = { v: await this.eval(a.expr, f), temp: true };
      if (a.name) (named ??= new Map()).set(a.name.toLowerCase(), ref);
      else refs.push(ref);
    }
    return { refs, named };
  }

  private async evalCall(e: Expr & { k: 'call' }, f: Frame, isStatement: boolean): Promise<Value> {
    const head = e.target;
    if (head.k === 'name') {
      const key = head.name.toLowerCase();
      const ref = this.lookupVar(key, f);
      if (ref && !(key === f.proc.name.toLowerCase() && e.args.length)) {
        const { values } = await this.argValues(e.args, f);
        if (ref.v instanceof VbArray) return values.length ? ref.v.get(values.map(toNum)) : ref.v;
        if (isObject(ref.v)) return values.length || !isStatement ? this.defaultCall(ref.v, values, head.name) : ref.v;
        if (!values.length) return ref.v;
        throw Object.assign(new VbError(13, `"${head.name}" isn't an array or a procedure.`), { line: e.line });
      }
      // User procedure
      const local = f.module.procs.get(key);
      const hit = local ? { mod: f.module, proc: local } : this.findAny(key);
      if (hit) {
        const { refs, named } = await this.argRefs(e.args, f);
        return this.callProc(hit.mod, hit.proc, refs, this.host?.documentObject(hit.mod.name) ?? null, named);
      }
      if (key in BUILTINS) {
        const { values } = await this.argValues(e.args, f);
        const ctx = { ...this.builtinCtx!, compareText: f.module.ast.compareText, optionBase: f.module.ast.optionBase };
        try {
          // Objects pass their value (Len(Range("A1"))), except to functions that look at the object itself
          const keepObjects = key === 'isobject' || key === 'typename' || key === 'ismissing';
          const plain: Value[] = [];
          for (const v of values) plain.push(keepObjects ? v : await this.unwrapDefault(v));
          return await BUILTINS[key](plain, ctx);
        } catch (err) {
          if (err instanceof VbError && !err.line) err.line = e.line;
          throw err;
        }
      }
      if (this.host?.isGlobal(key)) {
        const { values, named } = await this.argValues(e.args, f);
        try {
          return await this.host.global(key, values, named, f.me);
        } catch (err) {
          if (err instanceof VbError && !err.line) err.line = e.line;
          throw err;
        }
      }
      if (this.moduleByName(key)) return this.defaultCall(await this.evalName(head, f), (await this.argValues(e.args, f)).values, head.name);
      throw Object.assign(new VbError(1004, `There's no procedure or array called "${head.name}".`), { line: e.line });
    }
    if (head.k === 'member') {
      const obj = head.obj ? await this.eval(head.obj, f) : this.withObject(f, head.line);
      // Calling a procedure in a class instance passes arguments by reference
      if (obj instanceof ClassInstance) {
        const { refs } = await this.argRefs(e.args, f);
        return obj.callMethod(head.name.toLowerCase(), refs);
      }
      const { values, named } = await this.argValues(e.args, f);
      return this.member(obj, head.name, values, named, head.line);
    }
    // (expr)(args): default member of the result
    const obj = await this.eval(head, f);
    const { values } = await this.argValues(e.args, f);
    if (obj instanceof VbArray) return obj.get(values.map(toNum));
    return this.defaultCall(obj, values, 'value');
  }

  private async defaultCall(obj: Value, args: Value[], name: string): Promise<Value> {
    if (obj instanceof VbArray) return obj.get(args.map(toNum));
    if (!isObject(obj)) {
      if (!args.length) return obj;
      throw new VbError(obj === null ? 91 : 13, obj === null ? `Object variable not set: "${name}" is Nothing.` : `"${name}" can't take arguments.`);
    }
    if (!obj.dflt) throw new VbError(438, `A ${obj.vbType} can't be used like that.`);
    return obj.dflt(args);
  }

  private async member(obj: Value, name: string, args: Value[], named: Record<string, Value> | undefined, line: number): Promise<Value> {
    if (!isObject(obj)) {
      if (obj === null) throw Object.assign(new VbError(91, `Object variable not set: can't use .${name} on Nothing.`), { line });
      throw Object.assign(new VbError(424, `Object required: .${name} needs an object, not a ${typeName(obj)}.`), { line });
    }
    try {
      return await obj.get(name.toLowerCase(), args, named);
    } catch (err) {
      if (err instanceof VbError && !err.line) err.line = line;
      throw err;
    }
  }
}

function coerceParam(p: Param, v: Value): Value {
  return coerce(p.type, v);
}

function cmpOp(op: string, c: number): boolean {
  switch (op) {
    case '=':
      return c === 0;
    case '<>':
      return c !== 0;
    case '<':
      return c < 0;
    case '>':
      return c > 0;
    case '<=':
      return c <= 0;
    default:
      return c >= 0;
  }
}

/** An instance of a class module (Dim x As New MyClass). */
class ClassInstance implements VbObject {
  readonly vbType: string;
  fields = new Map<string, Ref>();
  private ready = false;

  constructor(
    private program: VbaProgram,
    private mod: Module,
  ) {
    this.vbType = mod.name;
  }

  private async init(): Promise<void> {
    if (this.ready) return;
    this.ready = true;
    for (const v of (this.mod as Module).ast.vars) this.fields.set(v.dim.name.toLowerCase(), { v: v.dim.isNew ? null : initial(v.dim.type), type: v.dim.type });
    const ctor = this.mod.procs.get('class_initialize');
    if (ctor) await this.program.callProc(this.mod, ctor, [], this, undefined, this);
  }

  async callMethod(name: string, args: Ref[]): Promise<Value> {
    await this.init();
    const proc = this.mod.procs.get(name);
    if (!proc) throw new VbError(438, `${this.vbType} doesn't have "${name}".`);
    return this.program.callProc(this.mod, proc, args, this, undefined, this);
  }

  async get(name: string, args: Value[]): Promise<Value> {
    await this.init();
    const proc = this.mod.procs.get(name);
    if (proc) return this.program.callProc(this.mod, proc, args.map((v) => ({ v })), this, undefined, this);
    const field = this.fields.get(name);
    if (field) return field.v;
    throw new VbError(438, `${this.vbType} doesn't have "${name}".`);
  }

  async set(name: string, args: Value[], v: Value): Promise<void> {
    await this.init();
    const setter = this.mod.procs.get(`${name}:let`) ?? this.mod.procs.get(`${name}:set`);
    if (setter) {
      const all: Value[] = [...args, v];
      await this.program.callProc(this.mod, setter, all.map((x): Ref => ({ v: x })), this, undefined, this);
      return;
    }
    const field = this.fields.get(name);
    if (field) {
      field.v = v;
      return;
    }
    throw new VbError(438, `${this.vbType} doesn't have "${name}".`);
  }
}
