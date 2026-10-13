// Connects running VBA code to the workbook through Excel's object model
// (Application, Workbook, Worksheets, Range…). One MacroSession = one run = one undo step.
import { MAX_COLS, MAX_ROWS, parseRange, quoteSheetName, type CellAddr, type Range } from '../address';
import type { Doc, LiveEdit } from '../doc';
import { FUNCTIONS } from '../formula/functions/registry';
import { translateFormula } from '../formula/refs';
import { Matrix } from '../formula/types';
import type { Transaction } from '../history';
import { isDateFormat } from '../numfmt';
import { setCellInput, setLink, forEachCellIn } from '../ops/cells';
import { clampToUsed, pasteClip, type ClipData, type PasteMode } from '../ops/clipboard';
import { applyFilter, displayText } from '../ops/data';
import { addSheet, deleteSheet, duplicateSheet, moveSheet, renameSheet, setSheetHidden } from '../ops/sheets';
import { sortRange, type SortKey } from '../ops/data';
import { ErrorValue, type ErrorCode, type Scalar } from '../values';
import { validateSheetName, type Sheet } from '../workbook';
import { hexToVb, opt, RangeObj, unsupported, vbToHex, wildcard } from './range';
import { MISSING } from './vba/builtins';
import type { VbaHost, VbaProgram } from './vba/interpreter';
import { isGen, type Pending } from './vba/flow';
import { EMPTY, isObject, NULL, StopSignal, toBool, toNum, toStr, VbArray, VbDate, VbError, type Value, type VbObject } from './vba/values';

export interface MacroUI {
  msgBox(prompt: string, buttons: number, title: string): Promise<number>;
  inputBox(prompt: string, title: string, dflt: string): Promise<string | null>;
  /** Debug.Print output. */
  print(text: string): void;
  /** Application.StatusBar text (null = clear). */
  status(text: string | null): void;
  selection(): { sheet: Sheet; ranges: Range[]; active: CellAddr };
  select(sheet: Sheet, ranges: Range[], active: CellAddr): void;
  activate(sheet: Sheet): void;
  save(): Promise<boolean>;
  /** Width (col) or height (row) in px that fits the contents. */
  measure?(sheet: Sheet, axis: 'col' | 'row', index: number): number;
  /** Called while a macro runs, with the time it has been running. */
  progress?(ms: number): void;
  /** Zoom or other view-only changes were made. */
  viewChanged?(): void;
}

export interface FindArgs {
  what: string;
  after?: Value;
  lookIn: number;
  lookAt: number;
  byColumns: boolean;
  previous: boolean;
  matchCase: boolean;
}

const ERR_NUM: Partial<Record<ErrorCode, number>> = {
  '#NULL!': 2000,
  '#DIV/0!': 2007,
  '#VALUE!': 2015,
  '#REF!': 2023,
  '#NAME?': 2029,
  '#NUM!': 2036,
  '#N/A': 2042,
};

/** A cell error as a VBA Error value (what CVErr returns). */
export function cvErr(n: number): VbObject {
  return { vbType: 'Error', get: () => n, set: () => undefined, dflt: () => n };
}

function errFromVb(n: number): ErrorValue {
  const code = (Object.entries(ERR_NUM).find(([, v]) => v === n)?.[0] ?? '#VALUE!') as ErrorCode;
  return new ErrorValue(code);
}

/** Names usable without "Application." in front. */
const GLOBALS = new Set([
  'application',
  'activesheet',
  'activeworkbook',
  'thisworkbook',
  'workbooks',
  'activecell',
  'selection',
  'sheets',
  'worksheets',
  'range',
  'cells',
  'rows',
  'columns',
  'worksheetfunction',
  'union',
  'intersect',
  'evaluate',
  'calculate',
  'names',
  'activewindow',
]);

export class MacroSession implements VbaHost {
  readonly application: VbObject;
  readonly live: LiveEdit;
  readonly workbook: WorkbookObj;
  clipboard: ClipData | null = null;
  lastFind: (FindArgs & { range: RangeObj; hit?: RangeObj }) | null = null;
  screenUpdating = true;
  readonly started = performance.now();
  private lastYield = performance.now();
  private lastPaint = performance.now();
  private stopRequested = false;
  private sheetObjs = new Map<Sheet, SheetObj>();
  private sorts = new Map<Sheet, SortObj>();
  /** For a function used in a cell: the cell (Application.Caller). */
  caller: { sheet: Sheet; row: number; col: number } | null = null;

  constructor(
    readonly doc: Doc,
    readonly ui: MacroUI,
    readonly program: VbaProgram,
    label: string,
  ) {
    this.live = doc.beginLive(label);
    this.workbook = new WorkbookObj(this);
    this.application = new ApplicationObj(this);
  }

  // ---------- VbaHost ----------

  isGlobal(name: string): boolean {
    return GLOBALS.has(name) || !!this.sheetByCodeName(name);
  }

  global(name: string, args: Value[], named: Record<string, Value> | undefined, me: VbObject | null): Pending<Value> {
    if (!GLOBALS.has(name)) {
      const sheet = this.sheetByCodeName(name);
      if (sheet) return this.sheetObj(sheet);
    }
    return this.appMember(name, args, named, me instanceof SheetObj ? me.sheet : null);
  }

  documentObject(module: string): VbObject | null {
    if (module.toLowerCase() === (this.doc.wb.macros?.workbookModule ?? 'ThisWorkbook').toLowerCase()) return this.workbook;
    const sheet = this.sheetByCodeName(module.toLowerCase());
    return sheet ? this.sheetObj(sheet) : null;
  }

  msgBox(prompt: string, buttons: number, title: string): Promise<number> {
    this.live.repaint();
    return this.ui.msgBox(prompt, buttons, title);
  }

  inputBox(prompt: string, title: string, dflt: string): Promise<string | null> {
    this.live.repaint();
    return this.ui.inputBox(prompt, title, dflt);
  }

  print(text: string): void {
    this.ui.print(text);
  }

  tick(): void | Promise<void> {
    return this.pause();
  }

  /** Lets the app repaint and respond while long-running code works. */
  private async pause(): Promise<void> {
    const now = performance.now();
    if (now - this.lastYield < 30) return;
    if (this.screenUpdating && now - this.lastPaint > 500) {
      this.live.repaint();
      this.lastPaint = now;
    }
    this.ui.progress?.(now - this.started);
    await new Promise((r) => setTimeout(r, 0));
    this.lastYield = performance.now();
    if (this.stopRequested) throw new StopSignal('cancel');
  }

  evaluate(text: string, sheet?: Sheet): Value {
    this.settle();
    const t = text.trim().replace(/^=/, '');
    const base = sheet ?? this.activeSheet();
    // A plain address or name gives a Range, like Excel
    try {
      return this.parseAddress(t, base);
    } catch {
      // not an address
    }
    return this.fromFormula(this.doc.engine.evalText(t, base, 0, 0), false, 'Evaluate');
  }

  // ---------- running ----------

  /** Asks the running code to stop at the next statement. */
  stop(): void {
    this.stopRequested = true;
    this.program.stop();
  }

  /** Records the run's changes as one undo step. */
  finish(): void {
    this.live.finish();
    this.ui.status(null);
  }

  // ---------- helpers for the object model ----------

  activeSheet(): Sheet {
    return this.doc.wb.activeSheet;
  }

  checkSheet(sheet: Sheet): void {
    if (!this.doc.wb.sheets.includes(sheet)) throw new VbError(424, `The sheet "${sheet.name}" has been deleted.`);
  }

  sheetByCodeName(name: string): Sheet | undefined {
    const n = name.toLowerCase();
    return this.doc.wb.sheets.find((s) => s.codeName?.toLowerCase() === n);
  }

  sheetObj(sheet: Sheet): SheetObj {
    let o = this.sheetObjs.get(sheet);
    if (!o) this.sheetObjs.set(sheet, (o = new SheetObj(this, sheet)));
    return o;
  }

  sortObj(sheet: Sheet): SortObj {
    let o = this.sorts.get(sheet);
    if (!o) this.sorts.set(sheet, (o = new SortObj(this, sheet)));
    return o;
  }

  whole(sheet: Sheet): RangeObj {
    return new RangeObj(this, sheet, [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }]);
  }

  /** Applies a change as part of this run's undo step. */
  edit(fn: (tx: Transaction) => void): void {
    try {
      this.live.apply(fn);
    } catch (e) {
      if (e instanceof VbError || e instanceof StopSignal) throw e;
      throw new VbError(1004, (e as Error).message || String(e));
    }
  }

  settle(): void {
    this.live.settle();
  }

  clamp(sheet: Sheet, r: Range): Range {
    if (r.r2 < MAX_ROWS - 1 && r.c2 < MAX_COLS - 1) return r;
    return clampToUsed(this.doc, sheet, r);
  }

  /** Like clamp, but refuses ranges too large to fill one cell at a time. */
  clampWrite(sheet: Sheet, r: Range): Range {
    const a = this.clamp(sheet, r);
    if ((a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1) > 2_000_000) throw new VbError(7, 'That range is too big to fill at once (over 2 million cells).');
    return a;
  }

  translate(formula: string, dr: number, dc: number): string {
    return translateFormula(formula, dr, dc);
  }

  measure(sheet: Sheet, axis: 'col' | 'row', index: number): number {
    const m = this.ui.measure?.(sheet, axis, index);
    if (m !== undefined) return m;
    if (axis === 'row') return sheet.defaultRowHeight;
    let w = 0;
    forEachCellIn(sheet, clampToUsed(this.doc, sheet, { r1: 0, c1: index, r2: MAX_ROWS - 1, c2: index }), (r, c) => {
      w = Math.max(w, displayText(this.doc, sheet, r, c).length * 7 + 12);
    });
    return w;
  }

  cellValue(sheet: Sheet, row: number, col: number, value2: boolean): Value {
    const v = this.doc.engine.getValue(sheet, row, col);
    if (v === null) return EMPTY;
    if (v instanceof ErrorValue) return cvErr(ERR_NUM[v.code] ?? 2015);
    if (typeof v === 'number' && !value2 && isDateFormat(this.doc.numFmtAt(sheet, row, col))) return new VbDate(v);
    return v;
  }

  writeCell(tx: Transaction, sheet: Sheet, row: number, col: number, input: Value): void {
    let v = input;
    if (v instanceof VbArray) v = v.data[0];
    if (isObject(v)) {
      if (v.vbType === 'Error') {
        tx.patchCell(sheet, row, col, { v: errFromVb(toNum(v.dflt?.([]) as Value)), f: undefined, cached: undefined });
        return;
      }
      const d = v.dflt?.([]);
      if (d === undefined || d instanceof Promise || isGen(d)) throw new VbError(13, `A ${v.vbType} can't be put in a cell.`);
      v = d;
    }
    if (v === EMPTY || v === null || v === NULL || v === undefined || v === '') {
      const cell = sheet.getCell(row, col);
      if (cell && (cell.v !== undefined || cell.f !== undefined)) tx.patchCell(sheet, row, col, { v: undefined, f: undefined, cached: undefined });
      return;
    }
    if (typeof v === 'string') return setCellInput(tx, this.doc, sheet, row, col, v);
    if (typeof v === 'number') {
      if (!Number.isFinite(v)) throw new VbError(6, 'Overflow: that number is too big for a cell.');
      tx.patchCell(sheet, row, col, { v, f: undefined, cached: undefined });
      return;
    }
    if (typeof v === 'boolean') {
      tx.patchCell(sheet, row, col, { v, f: undefined, cached: undefined });
      return;
    }
    if (v instanceof VbDate) {
      const patch: { v: number; f: undefined; cached: undefined; s?: number } = { v: v.serial, f: undefined, cached: undefined };
      const id = sheet.styleIdAt(row, col);
      const fmt = this.doc.wb.styles.get(id).numFmt;
      if (!fmt || fmt === 'General') {
        const whole = Number.isInteger(v.serial);
        patch.s = this.doc.wb.styles.apply(id, { numFmt: whole ? 'm/d/yyyy' : v.serial < 1 ? 'h:mm:ss AM/PM' : 'm/d/yyyy h:mm' });
      }
      tx.patchCell(sheet, row, col, patch);
      return;
    }
    throw new VbError(13, "That value can't be put in a cell.");
  }

  /** A Range from a Range object or an address/name string. `cell2` makes the box spanning both. */
  rangeArg(v: Value | undefined, sheet: Sheet, cell2?: Value): RangeObj {
    let a: RangeObj;
    if (v instanceof RangeObj) a = v;
    else if (typeof v === 'string') a = this.parseAddress(v, sheet);
    else if (v === undefined) throw new VbError(449, 'A range is required here.');
    else throw new VbError(13, 'Expected a range, or an address like "A1:B5".');
    if (cell2 === undefined) return a;
    const b = this.rangeArg(cell2, a.sheet);
    const r = { r1: Math.min(a.r.r1, b.r.r1), c1: Math.min(a.r.c1, b.r.c1), r2: Math.max(a.r.r2, b.r.r2), c2: Math.max(a.r.c2, b.r.c2) };
    return new RangeObj(this, a.sheet, [r]);
  }

  /** Parses "A1", "B2:C9", "A:A", "Sheet2!A1", "'My Sheet'!A1,C3" or a defined name. */
  parseAddress(text: string, sheet: Sheet, depth = 0): RangeObj {
    const parts = splitTopLevel(text.trim());
    let target: Sheet | null = null;
    const areas: Range[] = [];
    for (const part of parts) {
      let s = sheet;
      let addr = part.trim();
      const m = /^(?:'((?:[^']|'')+)'|([^!'"]+))!(.+)$/.exec(addr);
      if (m) {
        const name = (m[1] ?? m[2]).replace(/''/g, "'");
        const found = this.doc.wb.sheetByName(name);
        if (!found) throw new VbError(9, `There's no sheet called "${name}".`);
        s = found;
        addr = m[3];
      }
      const r = parseRange(addr);
      if (r) {
        if (target && target !== s) throw new VbError(1004, 'A range can only cover one sheet.');
        target = s;
        areas.push(r);
        continue;
      }
      const dn = this.doc.wb.getName(addr);
      if (dn && depth < 5) {
        const inner = this.parseAddress(dn.ref.replace(/^=/, ''), s, depth + 1);
        if (target && target !== inner.sheet) throw new VbError(1004, 'A range can only cover one sheet.');
        target = inner.sheet;
        areas.push(...inner.areas);
        continue;
      }
      throw new VbError(1004, `"${part.trim()}" isn't a cell address or a named range.`);
    }
    if (!target || !areas.length) throw new VbError(1004, `"${text}" isn't a cell address or a named range.`);
    return new RangeObj(this, target, areas);
  }

  defineName(name: string, ref: string): void {
    const n = name.trim();
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(n) || parseRange(n)) throw new VbError(1004, `"${name}" isn't a valid name (use letters, numbers and _ and don't make it look like a cell address).`);
    this.edit((tx) => {
      const names = new Map(this.doc.wb.names);
      names.set(n.toUpperCase(), { name: n, ref: ref.replace(/^=/, '') });
      tx.setProp(this.doc.wb, 'names', names, true);
    });
  }

  paste(sheet: Sheet, at: Range, mode: PasteMode): void {
    const clip = this.clipboard;
    if (!clip) throw new VbError(1004, "There's nothing to paste — use Range.Copy first.");
    this.edit((tx) => pasteClip(tx, this.doc, clip, sheet, at, mode));
    if (clip.cut) this.clipboard = null;
  }

  // ---------- Find ----------

  find(rng: RangeObj, a: FindArgs): Value {
    this.lastFind = { ...a, range: rng };
    return this.findFrom(rng, a, a.after);
  }

  findAgain(rng: RangeObj, after: Value | undefined, previous: boolean): Value {
    const f = this.lastFind;
    if (!f) throw new VbError(91, 'FindNext needs a Find first.');
    return this.findFrom(f.range, { ...f, previous }, after ?? f.hit ?? rng);
  }

  private findFrom(rng: RangeObj, a: FindArgs, after: Value | undefined): Value {
    if (!a.what) return null;
    this.settle();
    const re = wildcard(a.what, a.lookAt === 1, a.matchCase);
    const sheet = rng.sheet;
    const hits: [number, number][] = [];
    for (const area of rng.areas) {
      forEachCellIn(sheet, this.clamp(sheet, area), (r, c, cell) => {
        let text: string;
        if (a.lookIn === -4144) text = cell.note ?? '';
        else if (a.lookIn === -4123 && cell.f !== undefined) text = '=' + cell.f;
        else text = displayText(this.doc, sheet, r, c);
        if (text && re.test(text)) hits.push([r, c]);
      });
    }
    if (!hits.length) return null;
    const key = (r: number, c: number) => (a.byColumns ? c * MAX_ROWS + r : r * MAX_COLS + c);
    hits.sort((x, y) => key(x[0], x[1]) - key(y[0], y[1]));
    const last = rng.areas[rng.areas.length - 1];
    const from = after instanceof RangeObj ? key(after.r.r1, after.r.c1) : a.previous ? key(rng.r.r1, rng.r.c1) : key(last.r2, last.c2);
    let hit: [number, number] | undefined;
    if (a.previous) hit = [...hits].reverse().find(([r, c]) => key(r, c) < from) ?? hits[hits.length - 1];
    else hit = hits.find(([r, c]) => key(r, c) > from) ?? hits[0];
    const out = rng.sub({ r1: hit[0], c1: hit[1], r2: hit[0], c2: hit[1] });
    if (this.lastFind) this.lastFind.hit = out;
    return out;
  }

  // ---------- worksheet functions ----------

  /** Runs a spreadsheet function: WorksheetFunction.VLookup(…) → VLOOKUP(…). */
  worksheetFunction(name: string, args: Value[], strict: boolean): Value {
    const fn = name.replace(/_/g, '.').toUpperCase();
    if (!FUNCTIONS.has(fn)) throw new VbError(438, `There's no worksheet function called ${fn}.`);
    this.settle();
    const parts = args.map((v) => this.formulaArg(v));
    while (parts.length && parts[parts.length - 1] === '') parts.pop();
    return this.fromFormula(this.doc.engine.evalText(`${fn}(${parts.join(',')})`, this.activeSheet(), 0, 0), strict, fn);
  }

  private formulaArg(v: Value): string {
    if (v === (MISSING as unknown)) return '';
    if (v instanceof RangeObj) return `${quoteSheetName(v.sheet.name)}!${v.address(true, true).split(',')[0]}`;
    if (v instanceof VbArray) {
      const lit = (x: Value) => this.formulaArg(x === EMPTY ? 0 : x);
      if (v.dims.length === 2) {
        const rows: string[] = [];
        for (let r = v.dims[0].lb; r <= v.dims[0].ub; r++) {
          const cols: string[] = [];
          for (let c = v.dims[1].lb; c <= v.dims[1].ub; c++) cols.push(lit(v.get([r, c])));
          rows.push(cols.join(','));
        }
        return `{${rows.join(';')}}`;
      }
      return `{${v.data.map(lit).join(',')}}`;
    }
    if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`;
    if (typeof v === 'number') return Number.isFinite(v) ? String(v).replace('e', 'E') : '0';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (v instanceof VbDate) return String(v.serial);
    if (v === EMPTY || v === undefined) return '0';
    if (isObject(v) && v.vbType === 'Error') return errFromVb(toNum(v.dflt?.([]) as Value)).code;
    if (v === null) throw new VbError(91, 'Object variable not set (Nothing).');
    throw new VbError(13, `A ${isObject(v) ? v.vbType : 'value'} can't be passed to a worksheet function.`);
  }

  private fromFormula(v: Scalar | Matrix, strict: boolean, fn: string): Value {
    const one = (x: Scalar): Value => {
      if (x === null) return EMPTY;
      if (x instanceof ErrorValue) {
        if (strict) throw new VbError(1004, `WorksheetFunction.${fn} couldn't get a result${x.message ? ': ' + x.message : ` (${x.code})`}.`);
        return cvErr(ERR_NUM[x.code] ?? 2015);
      }
      return x;
    };
    if (v instanceof Matrix) {
      if (v.rows === 1 && v.cols === 1) return one(v.get(0, 0));
      const arr = new VbArray([
        { lb: 1, ub: v.rows },
        { lb: 1, ub: v.cols },
      ]);
      let i = 0;
      for (let r = 0; r < v.rows; r++) for (let c = 0; c < v.cols; c++) arr.data[i++] = one(v.get(r, c));
      return arr;
    }
    return one(v);
  }

  async sleep(ms: number): Promise<void> {
    const end = performance.now() + Math.min(Math.max(0, ms), 60_000);
    this.live.repaint();
    while (performance.now() < end) {
      if (this.stopRequested) throw new StopSignal('cancel');
      await new Promise((r) => setTimeout(r, Math.min(100, Math.max(0, end - performance.now()))));
    }
  }

  // ---------- Application members (also used for globals like Range and ActiveSheet) ----------

  appMember(name: string, args: Value[], named: Record<string, Value> | undefined, me: Sheet | null): Pending<Value> {
    const o = (i: number, n: string) => opt(args, named, i, n);
    const base = me ?? this.activeSheet();
    switch (name) {
      case 'application':
        return this.application;
      case 'activesheet':
        return this.sheetObj(this.activeSheet());
      case 'activeworkbook':
      case 'thisworkbook':
        return this.workbook;
      case 'workbooks':
        return args.length ? this.workbook : new WorkbooksObj(this);
      case 'activecell': {
        const sel = this.ui.selection();
        return new RangeObj(this, sel.sheet, [{ r1: sel.active.row, c1: sel.active.col, r2: sel.active.row, c2: sel.active.col }]);
      }
      case 'selection': {
        const sel = this.ui.selection();
        return new RangeObj(this, sel.sheet, sel.ranges.map((r) => ({ ...r })));
      }
      case 'sheets':
      case 'worksheets': {
        const sheets = new SheetsObj(this);
        return args.length ? sheets.item(args[0]) : sheets;
      }
      case 'range':
        if (!args.length) throw new VbError(449, 'Range needs an address, like Range("A1").');
        return this.rangeArg(args[0], base, o(1, 'cell2'));
      case 'cells':
      case 'rows':
      case 'columns':
        return this.whole(base).get(name, args, named);
      case 'worksheetfunction':
        return new WorksheetFunctionObj(this);
      case 'union': {
        const list = args.filter((a) => a !== (MISSING as unknown)).map((a) => this.rangeArg(a, base));
        if (list.some((r) => r.sheet !== list[0].sheet)) throw new VbError(1004, 'Union needs ranges on the same sheet.');
        return new RangeObj(this, list[0].sheet, list.flatMap((r) => r.areas));
      }
      case 'intersect': {
        const list = args.filter((a) => a !== (MISSING as unknown) && a !== null).map((a) => this.rangeArg(a, base));
        let cur: Range | null = list[0]?.r ?? null;
        for (const r of list.slice(1)) {
          if (!cur || r.sheet !== list[0].sheet) return null;
          const x: Range = { r1: Math.max(cur.r1, r.r.r1), c1: Math.max(cur.c1, r.r.c1), r2: Math.min(cur.r2, r.r.r2), c2: Math.min(cur.c2, r.r.c2) };
          cur = x.r1 <= x.r2 && x.c1 <= x.c2 ? x : null;
        }
        return cur ? new RangeObj(this, list[0].sheet, [cur]) : null;
      }
      case 'evaluate':
        return this.evaluate(toStr(args[0] ?? ''), base);
      case 'calculate':
      case 'calculatefull':
        this.settle();
        this.doc.engine.recalc();
        return EMPTY;
      case 'names':
        return args.length ? new NamesObj(this).item(args[0]) : new NamesObj(this);
      case 'activewindow':
        return new WindowObj(this);
    }
    return this.applicationOnly(name, args, named);
  }

  private settings = new Map<string, Value>([
    ['displayalerts', true],
    ['enableevents', true],
    ['calculation', -4105],
    ['interactive', true],
    ['cursor', -4143],
    ['displaystatusbar', true],
    ['enablecancelkey', 1],
    ['askToUpdateLinks', true],
    ['printcommunication', true],
    ['displayformulabar', true],
    ['windowstate', -4137],
    ['caption', 'Gridline'],
    ['statusbar', false],
    ['cutcopymode', false],
  ]);

  get eventsEnabled(): boolean {
    return toBool(this.settings.get('enableevents') ?? true);
  }

  private applicationOnly(name: string, args: Value[], named: Record<string, Value> | undefined): Pending<Value> {
    const o = (i: number, n: string) => opt(args, named, i, n);
    switch (name) {
      case 'name':
        return 'Gridline';
      case 'version':
        return '16.0';
      case 'build':
        return 0;
      case 'username':
        return '';
      case 'operatingsystem':
        return 'Windows (64-bit) NT 10.00';
      case 'pathseparator':
        return '\\';
      case 'decimalseparator':
        return '.';
      case 'thousandsseparator':
        return ',';
      case 'path':
      case 'defaultfilepath':
        return '';
      case 'screenupdating':
        return this.screenUpdating;
      case 'run': {
        if (!args.length) throw new VbError(449, 'Run needs the name of a macro.');
        const target = toStr(args[0]).replace(/^.*!/, '').replace(/^'|'$/g, '');
        return this.program.call(this, target, args.slice(1));
      }
      case 'wait': {
        const when = o(0, 'time');
        const serial = when instanceof VbDate ? when.serial : toNum(when ?? 0);
        const now = (Date.now() - new Date().getTimezoneOffset() * 60000) / 86400000 + 25569;
        return this.sleep((serial - now) * 86400000).then(() => true);
      }
      case 'inputbox':
        return this.askInput(o);
      default:
        return this.applicationRest(name, args, named);
    }
  }

  /** Application.InputBox: asks until the answer fits the requested type. */
  private async askInput(o: (i: number, n: string) => Value | undefined): Promise<Value> {
    {
      {
        const prompt = toStr(o(0, 'prompt') ?? '');
        const title = toStr(o(1, 'title') ?? 'Input');
        const dflt = o(2, 'default');
        const type = toNum(o(7, 'type') ?? 2);
        for (;;) {
          const answer = await this.inputBox(prompt, title, dflt instanceof RangeObj ? dflt.address(false, false) : dflt === undefined ? '' : toStr(dflt));
          if (answer === null) return false;
          if (type & 8) {
            try {
              return this.parseAddress(answer.replace(/^=/, ''), this.activeSheet());
            } catch {
              await this.msgBox('Please enter a cell reference, like A1 or B2:D10.', 48, title);
              continue;
            }
          }
          if (type & 1) {
            const n = Number(answer.replace(/,/g, ''));
            if (answer.trim() === '' || !Number.isFinite(n)) {
              await this.msgBox('Please enter a number.', 48, title);
              continue;
            }
            return n;
          }
          if (type & 4) return /^(true|yes|1)$/i.test(answer.trim());
          return answer;
        }
      }
    }
  }

  private applicationRest(name: string, args: Value[], named: Record<string, Value> | undefined): Value {
    const o = (i: number, n: string) => opt(args, named, i, n);
    switch (name) {
      case 'goto': {
        const ref = o(0, 'reference');
        if (ref === undefined) return EMPTY;
        const r = this.rangeArg(ref, this.activeSheet());
        if (r.sheet !== this.activeSheet()) this.ui.activate(r.sheet);
        this.ui.select(r.sheet, r.areas, { row: r.r.r1, col: r.r.c1 });
        return EMPTY;
      }
      case 'volatile':
      case 'onkey':
      case 'enableevents':
        if (name === 'enableevents') return this.eventsEnabled;
        return EMPTY;
      case 'quit':
        throw new StopSignal('end');
      case 'ontime':
        throw new VbError(1004, "Application.OnTime isn't supported. Use an automatic trigger on the Automate tab instead.");
      case 'getopenfilename':
      case 'getsaveasfilename':
      case 'filedialog':
      case 'dialogs':
        throw new VbError(70, "Macros can't open file dialogs in Gridline.");
      case 'undo':
        throw new VbError(1004, "Macros can't undo — press Ctrl+Z after the macro finishes.");
      case 'caller':
        if (this.caller) return new RangeObj(this, this.caller.sheet, [{ r1: this.caller.row, c1: this.caller.col, r2: this.caller.row, c2: this.caller.col }]);
        throw new VbError(1004, 'Application.Caller is only available to macros started by a button, or to functions used in cells.');
      case 'international':
        return toNum(args[0] ?? 0) === 32 ? '.' : '';
    }
    if (this.settings.has(name)) return this.settings.get(name)!;
    // Application.Sum(...) and friends: worksheet functions that return errors instead of raising them
    if (FUNCTIONS.has(name.replace(/_/g, '.').toUpperCase())) return this.worksheetFunction(name, args, false);
    throw unsupported('Application', name);
  }

  setApplication(name: string, v: Value): void {
    switch (name) {
      case 'screenupdating':
        this.screenUpdating = toBool(v);
        if (this.screenUpdating) this.live.repaint();
        return;
      case 'statusbar':
        this.settings.set(name, v);
        this.ui.status(v === false || v === EMPTY ? null : toStr(v));
        return;
      case 'cutcopymode':
        if (v === false || toNum(v) === 0) this.clipboard = null;
        return;
      case 'calculation':
        if (toNum(v) === -4105) this.settle();
        this.settings.set(name, v);
        return;
    }
    if (this.settings.has(name)) {
      this.settings.set(name, v);
      return;
    }
    throw unsupported('Application', name);
  }
}

/** Splits "A1,B2" on commas that aren't inside quotes. */
function splitTopLevel(s: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (const ch of s) {
    if (ch === "'") quoted = !quoted;
    if (ch === ',' && !quoted) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

// ---------- Application ----------

class ApplicationObj implements VbObject {
  readonly vbType = 'Application';
  constructor(private s: MacroSession) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Pending<Value> {
    return this.s.appMember(name, args, named, null);
  }

  set(name: string, _args: Value[], v: Value): void {
    this.s.setApplication(name, v);
  }
}

class WorksheetFunctionObj implements VbObject {
  readonly vbType = 'WorksheetFunction';
  constructor(private s: MacroSession) {}

  get(name: string, args: Value[]): Value {
    return this.s.worksheetFunction(name, args, true);
  }

  set(name: string): void {
    throw unsupported('WorksheetFunction', name);
  }
}

// ---------- Workbook ----------

class WorkbookObj implements VbObject {
  readonly vbType = 'Workbook';
  constructor(private s: MacroSession) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Value | Promise<Value> {
    const doc = this.s.doc;
    switch (name) {
      case 'name':
        return doc.fileName;
      case 'fullname':
        return doc.filePath ?? doc.fileName;
      case 'path':
        return doc.filePath ? doc.filePath.replace(/[\\/][^\\/]*$/, '') : '';
      case 'sheets':
      case 'worksheets': {
        const sheets = new SheetsObj(this.s);
        return args.length ? sheets.item(args[0]) : sheets;
      }
      case 'activesheet':
        return this.s.sheetObj(this.s.activeSheet());
      case 'names':
        return args.length ? new NamesObj(this.s).item(args[0]) : new NamesObj(this.s);
      case 'save':
        this.s.live.repaint();
        return this.s.ui.save();
      case 'saved':
        return !doc.dirty;
      case 'codename':
        return doc.wb.macros?.workbookModule ?? 'ThisWorkbook';
      case 'fileformat':
        return doc.wb.macros && !doc.wb.macros.isEmpty ? 52 : 51;
      case 'readonly':
        return false;
      case 'refreshall':
        this.s.settle();
        doc.engine.recalcExternal();
        return EMPTY;
      case 'protect':
      case 'unprotect':
        return EMPTY;
      case 'parent':
      case 'application':
        return this.s.application;
      case 'saveas':
      case 'savecopyas':
        throw new VbError(70, "Macros can't save to other files in Gridline. Use File → Save As.");
      case 'close':
        throw new VbError(1004, "Macros can't close the workbook in Gridline.");
      case 'activate':
        return EMPTY;
    }
    void named;
    throw unsupported('Workbook', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    if (name === 'saved') {
      if (toBool(v)) this.s.doc.markSaved();
      return;
    }
    throw unsupported('Workbook', name);
  }
}

class WorkbooksObj implements VbObject {
  readonly vbType = 'Workbooks';
  constructor(private s: MacroSession) {}

  get(name: string, args: Value[]): Value {
    if (name === 'count') return 1;
    if (name === 'item') return this.s.workbook;
    if (name === 'add' || name === 'open') throw new VbError(1004, 'Gridline macros work inside this workbook only — they can\'t open or create other workbooks.');
    void args;
    throw unsupported('Workbooks', name);
  }

  set(name: string): void {
    throw unsupported('Workbooks', name);
  }

  dflt(): Value {
    return this.s.workbook;
  }

  items(): Value[] {
    return [this.s.workbook];
  }
}

// ---------- Sheets ----------

class SheetsObj implements VbObject {
  readonly vbType = 'Sheets';
  constructor(private s: MacroSession) {}

  item(key: Value): SheetObj {
    const wb = this.s.doc.wb;
    let sheet: Sheet | undefined;
    if (typeof key === 'string') sheet = wb.sheetByName(key);
    else sheet = wb.sheets[Math.round(toNum(key)) - 1];
    if (!sheet) throw new VbError(9, typeof key === 'string' ? `There's no sheet called "${key}".` : `There's no sheet number ${toStr(key)} (the workbook has ${wb.sheets.length}).`);
    return this.s.sheetObj(sheet);
  }

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    const wb = this.s.doc.wb;
    switch (name) {
      case 'count':
        return wb.sheets.length;
      case 'item':
        return this.item(args[0]);
      case 'add':
      case 'add2': {
        const before = opt(args, named, 0, 'before');
        const after = opt(args, named, 1, 'after');
        const count = Math.max(1, Math.round(toNum(opt(args, named, 2, 'count') ?? 1)));
        let index = wb.sheets.indexOf(this.s.activeSheet());
        if (before instanceof SheetObj) index = wb.sheets.indexOf(before.sheet);
        else if (after instanceof SheetObj) index = wb.sheets.indexOf(after.sheet) + 1;
        let sheet: Sheet | null = null;
        this.s.edit((tx) => {
          for (let i = 0; i < count; i++) {
            sheet = addSheet(tx, wb, undefined, index + i);
            sheet.codeName = uniqueCodeName(wb.sheets);
          }
        });
        this.s.ui.activate(sheet!);
        return this.s.sheetObj(sheet!);
      }
      case 'parent':
        return this.s.workbook;
    }
    throw unsupported('Sheets', name);
  }

  set(name: string): void {
    throw unsupported('Sheets', name);
  }

  dflt(args: Value[]): Value {
    return this.item(args[0]);
  }

  items(): Value[] {
    return this.s.doc.wb.sheets.map((sh) => this.s.sheetObj(sh));
  }
}

/** The next free "SheetN" code name. */
export function uniqueCodeName(sheets: Sheet[]): string {
  let i = 1;
  const taken = new Set(sheets.map((s) => s.codeName?.toLowerCase()));
  while (taken.has(`sheet${i}`)) i++;
  return `Sheet${i}`;
}

export class SheetObj implements VbObject {
  readonly vbType = 'Worksheet';
  constructor(
    private s: MacroSession,
    readonly sheet: Sheet,
  ) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Value | Promise<Value> {
    const sheet = this.sheet;
    const wb = this.s.doc.wb;
    if (name !== 'name' && name !== 'index') this.s.checkSheet(sheet);
    const o = (i: number, n: string) => opt(args, named, i, n);
    switch (name) {
      case 'name':
        return sheet.name;
      case 'codename':
        return sheet.codeName ?? '';
      case 'index':
        return wb.sheets.indexOf(sheet) + 1;
      case 'range':
        if (!args.length) throw new VbError(449, 'Range needs an address, like Range("A1").');
        return this.s.rangeArg(args[0], sheet, o(1, 'cell2'));
      case 'cells':
      case 'rows':
      case 'columns':
        return this.s.whole(sheet).get(name, args, named);
      case 'usedrange': {
        let r1 = Infinity;
        let c1 = Infinity;
        const b = sheet.bounds();
        for (const key of sheet.cells.keys()) {
          r1 = Math.min(r1, Math.floor(key / MAX_COLS));
          c1 = Math.min(c1, key % MAX_COLS);
        }
        if (b.maxRow < 0) return new RangeObj(this.s, sheet, [{ r1: 0, c1: 0, r2: 0, c2: 0 }]);
        return new RangeObj(this.s, sheet, [{ r1: Math.min(r1, b.maxRow), c1: Math.min(c1, b.maxCol), r2: b.maxRow, c2: b.maxCol }]);
      }
      case 'activate':
      case 'select':
        this.s.ui.activate(sheet);
        return EMPTY;
      case 'delete':
        this.s.edit((tx) => deleteSheet(tx, wb, sheet));
        this.s.ui.activate(wb.activeSheet);
        return true;
      case 'copy':
      case 'move': {
        const before = o(0, 'before');
        const after = o(1, 'after');
        if (!(before instanceof SheetObj) && !(after instanceof SheetObj)) throw new VbError(1004, `${name === 'copy' ? 'Copy' : 'Move'} needs Before:= or After:= another sheet (macros can't create new workbooks).`);
        let moved: Sheet = sheet;
        this.s.edit((tx) => {
          if (name === 'copy') {
            moved = duplicateSheet(tx, wb, sheet);
            moved.codeName = uniqueCodeName(wb.sheets);
          }
          const ref = (before instanceof SheetObj ? before : (after as SheetObj)).sheet;
          const others = wb.sheets.filter((x) => x !== moved);
          const idx = others.indexOf(ref) + (before instanceof SheetObj ? 0 : 1);
          moveSheet(tx, wb, moved, idx);
        });
        if (name === 'copy') this.s.ui.activate(moved);
        return EMPTY;
      }
      case 'visible':
        return sheet.hidden ? 0 : -1;
      case 'tab':
        return new TabObj(this.s, sheet);
      case 'next':
      case 'previous': {
        const i = wb.sheets.indexOf(sheet) + (name === 'next' ? 1 : -1);
        return wb.sheets[i] ? this.s.sheetObj(wb.sheets[i]) : null;
      }
      case 'parent':
        return this.s.workbook;
      case 'calculate':
        this.s.settle();
        return EMPTY;
      case 'protect':
      case 'unprotect':
        return EMPTY;
      case 'protectcontents':
      case 'protectdrawingobjects':
      case 'protectscenarios':
        return false;
      case 'paste': {
        const dest = o(0, 'destination');
        if (dest !== undefined) {
          const d = this.s.rangeArg(dest, sheet);
          this.s.paste(d.sheet, d.r, 'all');
        } else {
          const sel = this.s.ui.selection();
          this.s.paste(sel.sheet, sel.ranges[0], 'all');
        }
        return EMPTY;
      }
      case 'sort':
        return this.s.sortObj(sheet);
      case 'autofiltermode':
        return !!sheet.filter;
      case 'showalldata':
        if (sheet.filter) {
          const f = sheet.filter;
          this.s.edit((tx) => applyFilter(tx, this.s.doc, sheet, { range: f.range, conds: {} }));
        }
        return EMPTY;
      case 'autofilter':
        return sheet.filter ? new RangeObj(this.s, sheet, [sheet.filter.range]) : null;
      case 'hyperlinks':
        return new SheetLinksObj(this.s, sheet);
      case 'evaluate':
        return this.s.evaluate(toStr(args[0] ?? ''), sheet);
      case 'standardwidth':
        return Math.round(((sheet.defaultColWidth - 5) / 7) * 100) / 100;
      case 'standardheight':
        return sheet.defaultRowHeight * 0.75;
      case 'type':
        return -4167;
      case 'pagesetup':
      case 'displaypagebreaks':
        return new IgnoredObj('PageSetup');
      case 'enablecalculation':
        return true;
    }
    throw unsupported('Worksheet', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    const sheet = this.sheet;
    const wb = this.s.doc.wb;
    this.s.checkSheet(sheet);
    switch (name) {
      case 'name': {
        const next = toStr(v);
        const problem = validateSheetName(next, wb, sheet.id);
        if (problem) throw new VbError(1004, problem);
        this.s.edit((tx) => renameSheet(tx, wb, sheet, next.trim()));
        return;
      }
      case 'visible': {
        const show = typeof v === 'boolean' ? v : toNum(v) === -1;
        this.s.edit((tx) => setSheetHidden(tx, wb, sheet, !show));
        if (!show && wb.activeSheetId !== this.s.activeSheet().id) this.s.ui.activate(wb.activeSheet);
        return;
      }
      case 'autofiltermode':
        if (!toBool(v) && sheet.filter) this.s.edit((tx) => applyFilter(tx, this.s.doc, sheet, null));
        return;
      case 'standardwidth': {
        const px = Math.round(toNum(v) * 7 + 5);
        this.s.edit((tx) => tx.setProp(sheet, 'defaultColWidth', px));
        return;
      }
      case 'enablecalculation':
      case 'displaypagebreaks':
      case 'scrollarea':
        return;
    }
    throw unsupported('Worksheet', name);
  }
}

class TabObj implements VbObject {
  readonly vbType = 'Tab';
  constructor(
    private s: MacroSession,
    private sheet: Sheet,
  ) {}

  get(name: string): Value {
    if (name === 'color') return this.sheet.tabColor ? hexToVb(this.sheet.tabColor) : false;
    if (name === 'colorindex') return this.sheet.tabColor ? 1 : -4142;
    if (name === 'themecolor' || name === 'tintandshade') return 0;
    throw unsupported('Tab', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    if (name === 'themecolor' || name === 'tintandshade') return;
    if (name !== 'color' && name !== 'colorindex') throw unsupported('Tab', name);
    const none = v === false || toNum(v) === -4142 || toNum(v) === -4105;
    this.s.edit((tx) => tx.setProp(this.sheet, 'tabColor', none ? undefined : name === 'color' ? vbToHex(v) : undefined));
  }
}

class SheetLinksObj implements VbObject {
  readonly vbType = 'Hyperlinks';
  constructor(
    private s: MacroSession,
    private sheet: Sheet,
  ) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    const o = (i: number, n: string) => opt(args, named, i, n);
    switch (name) {
      case 'add': {
        const anchor = this.s.rangeArg(o(0, 'anchor'), this.sheet);
        const address = toStr(o(1, 'address') ?? '');
        const sub = o(2, 'subaddress');
        const text = o(4, 'texttodisplay');
        this.s.edit((tx) => {
          if (text !== undefined) this.s.writeCell(tx, anchor.sheet, anchor.r.r1, anchor.r.c1, toStr(text));
          setLink(tx, anchor.sheet, anchor.r.r1, anchor.r.c1, address || (sub !== undefined ? '#' + toStr(sub) : ''));
        });
        return EMPTY;
      }
      case 'count': {
        let n = 0;
        for (const c of this.sheet.cells.values()) if (c.link) n++;
        return n;
      }
      case 'delete':
        this.s.edit((tx) => {
          for (const [key, c] of [...this.sheet.cells]) if (c.link) setLink(tx, this.sheet, Math.floor(key / MAX_COLS), key % MAX_COLS, undefined);
        });
        return EMPTY;
    }
    throw unsupported('Hyperlinks', name);
  }

  set(name: string): void {
    throw unsupported('Hyperlinks', name);
  }
}

/** Worksheet.Sort, as written by Excel's macro recorder. */
class SortObj implements VbObject {
  readonly vbType = 'Sort';
  private fields: { col: number; desc: boolean }[] = [];
  private range: Range | null = null;
  private header = 2;
  constructor(
    private s: MacroSession,
    private sheet: Sheet,
  ) {}

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    switch (name) {
      case 'sortfields':
        return {
          vbType: 'SortFields',
          get: (n: string, a: Value[], nm?: Record<string, Value>) => {
            if (n === 'clear') {
              this.fields = [];
              return EMPTY;
            }
            if (n === 'add' || n === 'add2') {
              const key = this.s.rangeArg(opt(a, nm, 0, 'key'), this.sheet);
              const order = opt(a, nm, 2, 'order');
              this.fields.push({ col: key.r.c1, desc: order !== undefined && toNum(order) === 2 });
              return EMPTY;
            }
            if (n === 'count') return this.fields.length;
            throw unsupported('SortFields', n);
          },
          set: () => undefined,
        } as VbObject;
      case 'setrange':
        this.range = this.s.rangeArg(opt(args, named, 0, 'rng'), this.sheet).r;
        return EMPTY;
      case 'header':
        return this.header;
      case 'apply': {
        if (!this.range) throw new VbError(1004, 'Use .SetRange before .Apply.');
        this.s.settle();
        const area = this.s.clamp(this.sheet, this.range);
        const keys: SortKey[] = this.fields.length ? this.fields : [{ col: area.c1, desc: false }];
        this.s.edit((tx) => sortRange(tx, this.s.doc, this.sheet, area, keys, this.header === 1));
        return EMPTY;
      }
      case 'matchcase':
      case 'orientation':
      case 'sortmethod':
        return EMPTY;
    }
    throw unsupported('Sort', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    if (name === 'header') this.header = toNum(v);
    else if (!['matchcase', 'orientation', 'sortmethod'].includes(name)) throw unsupported('Sort', name);
  }
}

// ---------- Names ----------

class NamesObj implements VbObject {
  readonly vbType = 'Names';
  constructor(private s: MacroSession) {}

  item(key: Value): VbObject {
    const names = [...this.s.doc.wb.names.values()];
    const dn = typeof key === 'string' ? this.s.doc.wb.getName(key) : names[Math.round(toNum(key)) - 1];
    if (!dn) throw new VbError(1004, `There's no name "${toStr(key)}" in this workbook.`);
    const s = this.s;
    return {
      vbType: 'Name',
      get: (n: string) => {
        switch (n) {
          case 'name':
            return dn.name;
          case 'refersto':
          case 'value':
            return '=' + dn.ref;
          case 'referstorange':
            return s.parseAddress(dn.ref, s.activeSheet());
          case 'visible':
            return true;
          case 'delete':
            s.edit((tx) => {
              const map = new Map(s.doc.wb.names);
              map.delete(dn.name.toUpperCase());
              tx.setProp(s.doc.wb, 'names', map, true);
            });
            return EMPTY;
        }
        throw unsupported('Name', n);
      },
      set: (n: string, _a: Value[], v: Value) => {
        if (n === 'refersto' || n === 'value') s.defineName(dn.name, toStr(v));
        else if (n !== 'visible') throw unsupported('Name', n);
      },
      dflt: () => '=' + dn.ref,
    };
  }

  get(name: string, args: Value[], named?: Record<string, Value>): Value {
    switch (name) {
      case 'count':
        return this.s.doc.wb.names.size;
      case 'item':
        return this.item(args[0]);
      case 'add': {
        const n = toStr(opt(args, named, 0, 'name') ?? '');
        const ref = opt(args, named, 1, 'refersto');
        const text = ref instanceof RangeObj ? `${quoteSheetName(ref.sheet.name)}!${ref.address(true, true)}` : toStr(ref ?? '');
        this.s.defineName(n, text);
        return this.item(n);
      }
    }
    throw unsupported('Names', name);
  }

  set(name: string): void {
    throw unsupported('Names', name);
  }

  dflt(args: Value[]): Value {
    return this.item(args[0]);
  }

  items(): Value[] {
    return [...this.s.doc.wb.names.values()].map((n) => this.item(n.name));
  }
}

// ---------- Window ----------

class WindowObj implements VbObject {
  readonly vbType = 'Window';
  constructor(private s: MacroSession) {}

  private get sheet(): Sheet {
    return this.s.activeSheet();
  }

  get(name: string): Pending<Value> {
    const sheet = this.sheet;
    switch (name) {
      case 'freezepanes':
        return sheet.freeze.rows > 0 || sheet.freeze.cols > 0;
      case 'splitrow':
        return sheet.freeze.rows;
      case 'splitcolumn':
        return sheet.freeze.cols;
      case 'displaygridlines':
        return sheet.showGrid;
      case 'zoom':
        return Math.round(sheet.view.zoom * 100);
      case 'scrollrow':
      case 'scrollcolumn':
        return 1;
      case 'activesheet':
        return this.s.sheetObj(sheet);
      case 'activecell':
      case 'selection':
        return this.s.appMember(name, [], undefined, null);
    }
    throw unsupported('Window', name);
  }

  set(name: string, _args: Value[], v: Value): void {
    const sheet = this.sheet;
    switch (name) {
      case 'freezepanes': {
        const a = this.s.ui.selection().active;
        const f = toBool(v) ? { rows: a.row, cols: a.col } : { rows: 0, cols: 0 };
        this.s.edit((tx) => tx.setProp(sheet, 'freeze', f));
        return;
      }
      case 'splitrow':
        this.s.edit((tx) => tx.setProp(sheet, 'freeze', { ...sheet.freeze, rows: Math.max(0, Math.round(toNum(v))) }));
        return;
      case 'splitcolumn':
        this.s.edit((tx) => tx.setProp(sheet, 'freeze', { ...sheet.freeze, cols: Math.max(0, Math.round(toNum(v))) }));
        return;
      case 'displaygridlines':
        this.s.edit((tx) => tx.setProp(sheet, 'showGrid', toBool(v)));
        return;
      case 'zoom':
        sheet.view.zoom = Math.max(0.25, Math.min(4, toNum(v) / 100));
        this.s.ui.viewChanged?.();
        return;
      case 'scrollrow':
      case 'scrollcolumn':
      case 'displayheadings':
      case 'displayworkbooktabs':
      case 'windowstate':
        return;
    }
    throw unsupported('Window', name);
  }
}

/** Settings Gridline doesn't have (page setup and the like): reads give 0, writes are ignored. */
class IgnoredObj implements VbObject {
  constructor(readonly vbType: string) {}
  get(): Value {
    return 0;
  }
  set(): void {}
  dflt(): Value {
    return 0;
  }
}

/** A cell error as the VBA value a function receives (CVErr). */
export function vbErrorValue(e: ErrorValue): VbObject {
  return cvErr(ERR_NUM[e.code] ?? 2015);
}

/** The cell error for a VBA error value a function returned. */
export function cellError(v: VbObject): ErrorValue {
  return errFromVb(toNum(v.dflt?.([]) as Value));
}

/**
 * Where a Function runs when a formula calls it (=VAULTVALUE("Rent", "$")). Like in Excel, it can
 * read anything in the workbook but not change it, and it can't stop to show a message.
 */
export class CellSession extends MacroSession {
  constructor(doc: Doc, program: VbaProgram, caller: { sheet: Sheet; row: number; col: number }) {
    const cell = { r1: caller.row, c1: caller.col, r2: caller.row, c2: caller.col };
    const refuse = (what: string) => () => Promise.reject(new VbError(1004, `A function used in a cell can't ${what}.`));
    // No screen: ActiveCell and Selection are the cell being calculated
    const ui: MacroUI = {
      msgBox: refuse('show messages'),
      inputBox: refuse('ask questions'),
      save: refuse('save the workbook'),
      print: () => undefined,
      status: () => undefined,
      selection: () => ({ sheet: caller.sheet, ranges: [cell], active: { row: caller.row, col: caller.col } }),
      select: () => undefined,
      activate: () => undefined,
    };
    super(doc, ui, program, 'Calculate');
    this.caller = caller;
  }

  override edit(): void {
    throw new VbError(1004, "A function used in a cell can't change cells, sheets or formatting. It can only return a value.");
  }

  override settle(): void {
    // nothing is ever pending: a cell function doesn't change anything
  }

  override tick(): void {
    // a recalculation never pauses
  }

  override msgBox(): Promise<number> {
    throw new VbError(1004, "A function used in a cell can't show messages.");
  }

  override inputBox(): Promise<string | null> {
    throw new VbError(1004, "A function used in a cell can't ask questions.");
  }

  override finish(): void {
    // nothing to record
  }
}
