// The calculation engine: tracks dependencies, recalculates incrementally, and spills arrays.
import { cellKey, formatCell, MAX_COLS, MAX_ROWS, type Range } from '../address';
import type { TouchedCell } from '../history';
import { expandStructured } from '../tableRefs';
import { Errors, ErrorValue, type Scalar } from '../values';
import type { Sheet, Workbook } from '../workbook';
import { evaluate, evaluateFormula } from './evaluator';
import { FUNCTIONS, VOLATILE_FUNCS } from './functions';
import { parseFormula, type Node, type ParsedFormula } from './parser';
import { Matrix, RangeRef, type Ctx, type UserFunction, type Value } from './types';

const SHEET_SHIFT = 17_179_869_184; // 2^34 > MAX_ROWS * MAX_COLS
const MAX_DEPTH = 400;

export function gkey(sheetId: number, row: number, col: number): number {
  return sheetId * SHEET_SHIFT + row * MAX_COLS + col;
}

function gSheet(g: number): number {
  return Math.floor(g / SHEET_SHIFT);
}

function gRow(g: number): number {
  return Math.floor((g % SHEET_SHIFT) / MAX_COLS);
}

function gCol(g: number): number {
  return (g % SHEET_SHIFT) % MAX_COLS;
}

interface RangeDep {
  sheetId: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  dep: number;
}

interface FormulaInfo {
  sheetId: number;
  row: number;
  col: number;
  text: string;
  parsed: ParsedFormula;
  cells: number[];
  ranges: RangeDep[];
  volatile: boolean;
  unsupported: string[];
}

interface Spill {
  sheetId: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  data: Scalar[][];
  /** Number format per column (from table functions like AMORTIZATION). */
  fmts?: (string | undefined)[];
}

class TooDeep {
  constructor(readonly g: number) {}
}

const BLOCK = 64; // columns per range-dependency bucket

export class Engine {
  private formulas = new Map<number, FormulaInfo>();
  private values = new Map<number, Scalar>();
  private dirty = new Set<number>();
  private computing = new Set<number>();
  private cellDeps = new Map<number, Set<number>>();
  /** sheetId -> column block -> range deps */
  private rangeDeps = new Map<number, Map<number, Set<RangeDep>>>();
  private spills = new Map<number, Spill>();
  private spillOwner = new Map<number, number>();
  private blockedSpills = new Map<number, Spill>();
  private volatiles = new Set<number>();
  /** Formulas that use downloaded data. */
  private externals = new Set<number>();
  private depth = 0;
  /** Bumped after each recalculation, so caches (conditional formatting, charts) know to refresh. */
  version = 0;
  /**
   * Functions from the workbook's VBA code (set while its macros may run). After changing it,
   * call rebuild() so formulas using them recalculate.
   */
  userFunctions: ((name: string) => UserFunction | null) | null = null;

  constructor(private wb: Workbook) {}

  // ---------- public API ----------

  rebuild(): void {
    this.formulas.clear();
    this.values.clear();
    this.dirty.clear();
    this.cellDeps.clear();
    this.rangeDeps.clear();
    this.spills.clear();
    this.spillOwner.clear();
    this.blockedSpills.clear();
    this.volatiles.clear();
    this.externals.clear();
    for (const sheet of this.wb.sheets) {
      for (const [key, cell] of sheet.cells) {
        if (cell.f === undefined) continue;
        const row = Math.floor(key / MAX_COLS);
        const col = key % MAX_COLS;
        this.register(sheet, row, col, cell.f);
      }
    }
    this.recalc();
  }

  /** Call after cells changed (from a transaction). */
  update(touched: TouchedCell[]): void {
    if (touched.length > 20000) {
      this.rebuild();
      return;
    }
    const changed: number[] = [];
    for (const t of touched) {
      const sheet = this.wb.sheetById(t.sheetId);
      if (!sheet) continue;
      const g = gkey(t.sheetId, t.row, t.col);
      if (this.formulas.has(g)) this.unregister(g);
      const cell = sheet.getCell(t.row, t.col);
      if (cell?.f !== undefined) this.register(sheet, t.row, t.col, cell.f);
      else this.values.delete(g);
      changed.push(g);
      // Typing into a spilled area blocks that spill.
      const owner = this.spillOwner.get(g);
      if (owner !== undefined && owner !== g) {
        this.dirty.add(owner);
        changed.push(owner);
      }
      for (const [anchor, sp] of this.blockedSpills) {
        if (sp.sheetId === t.sheetId && t.row >= sp.r1 && t.row <= sp.r2 && t.col >= sp.c1 && t.col <= sp.c2) {
          this.dirty.add(anchor);
          changed.push(anchor);
        }
      }
    }
    this.invalidate(changed);
    this.recalc();
  }

  recalc(): void {
    if (this.volatiles.size) {
      for (const g of this.volatiles) this.dirty.add(g);
      this.invalidate([...this.volatiles]);
    }
    let guard = 0;
    while (this.dirty.size && guard++ < 50) {
      const list = [...this.dirty].sort((a, b) => a - b);
      for (const g of list) {
        if (!this.dirty.has(g)) continue;
        this.computeWithRestarts(g);
      }
    }
    this.version++;
  }

  /** Recalculates formulas that use downloaded data (after a download finishes). */
  recalcExternal(): boolean {
    if (!this.externals.size) return false;
    for (const g of this.externals) this.dirty.add(g);
    this.invalidate([...this.externals]);
    this.recalc();
    return true;
  }

  /** Whether any formula calls a function that isn't built in (one from the workbook's VBA code). */
  callsUnknownFunctions(): boolean {
    for (const info of this.formulas.values()) for (const f of info.parsed.funcs) if (!FUNCTIONS.has(f)) return true;
    return false;
  }

  /** Whether any formula in the workbook uses Gridline-only functions (NFL, live prices…). */
  usesCustomFunctions(): boolean {
    for (const info of this.formulas.values()) {
      for (const f of info.parsed.funcs) if (FUNCTIONS.get(f)?.custom) return true;
    }
    return false;
  }

  /** The value shown in a cell (formula result, literal, or spilled element). */
  getValue(sheet: Sheet, row: number, col: number): Scalar {
    const cell = sheet.cells.get(cellKey(row, col));
    const g = gkey(sheet.id, row, col);
    if (cell?.f !== undefined) {
      if (this.dirty.has(g)) this.compute(g);
      const v = this.values.get(g);
      return v === undefined ? null : v;
    }
    if (cell && cell.v !== undefined && cell.v !== null) return cell.v;
    const owner = this.spillOwner.get(g);
    if (owner !== undefined) {
      if (this.dirty.has(owner)) this.compute(owner);
      const sp = this.spills.get(owner);
      if (sp && row >= sp.r1 && row <= sp.r2 && col >= sp.c1 && col <= sp.c2) {
        const v = sp.data[row - sp.r1][col - sp.c1];
        return v === null ? 0 : v;
      }
    }
    return null;
  }

  /** Range covered by a spilled array whose anchor is (row, col). */
  spillAt(sheet: Sheet, row: number, col: number): Range | null {
    const sp = this.spills.get(gkey(sheet.id, row, col));
    return sp ? { r1: sp.r1, c1: sp.c1, r2: sp.r2, c2: sp.c2 } : null;
  }

  /** Number format suggested by a table function for a spilled cell (dates, money…). */
  spillFormatAt(sheet: Sheet, row: number, col: number): string | undefined {
    const g = gkey(sheet.id, row, col);
    const sp = this.spills.get(this.spillOwner.get(g) ?? g);
    if (!sp?.fmts || row === sp.r1 && sp.data.length > 1 && typeof sp.data[0][col - sp.c1] === 'string') return undefined;
    return sp.fmts[col - sp.c1];
  }

  /** If the cell is inside a spill (not the anchor), returns the anchor. */
  spillAnchorOf(sheet: Sheet, row: number, col: number): { row: number; col: number } | null {
    const owner = this.spillOwner.get(gkey(sheet.id, row, col));
    if (owner === undefined) return null;
    return { row: gRow(owner), col: gCol(owner) };
  }

  blockedSpillAt(sheet: Sheet, row: number, col: number): Range | null {
    const sp = this.blockedSpills.get(gkey(sheet.id, row, col));
    return sp ? { r1: sp.r1, c1: sp.c1, r2: sp.r2, c2: sp.c2 } : null;
  }

  spillsOn(sheet: Sheet): Range[] {
    const out: Range[] = [];
    for (const sp of this.spills.values()) if (sp.sheetId === sheet.id) out.push({ r1: sp.r1, c1: sp.c1, r2: sp.r2, c2: sp.c2 });
    return out;
  }

  isUnsupported(sheet: Sheet, row: number, col: number): string[] {
    return this.formulas.get(gkey(sheet.id, row, col))?.unsupported ?? [];
  }

  /** Evaluates an AST at a position without tracking dependencies (conditional formatting, previews). */
  evalAst(ast: Node, sheet: Sheet, row: number, col: number, dr = 0, dc = 0): Value {
    const ctx = this.makeCtx(sheet, row, col, dr, dc);
    try {
      return evaluateFormula(ast, ctx);
    } catch (e) {
      if (e instanceof TooDeep) return Errors.calc('The calculation is too deeply nested.');
      throw e;
    }
  }

  /** Evaluates formula text (without '=') at a position. */
  evalText(text: string, sheet: Sheet, row: number, col: number): Scalar | Matrix {
    const p = parseFormula(this.withTables(text, sheet, row, col));
    if (!p.ast) return Errors.value(p.error?.message);
    const v = this.evalAst(p.ast, sheet, row, col);
    return this.materialize(v, sheet);
  }

  materialize(v: Value, sheet?: Sheet): Scalar | Matrix {
    if (v instanceof RangeRef) {
      if (v.isCell) return this.getValue(v.sheet, v.r1, v.c1);
      const ctx = this.makeCtx(sheet ?? v.sheet, 0, 0, 0, 0);
      let r2 = v.r2;
      let c2 = v.c2;
      if (r2 === MAX_ROWS - 1 || c2 === MAX_COLS - 1) {
        const b = ctx.bounds(v.sheet);
        if (r2 === MAX_ROWS - 1) r2 = Math.max(v.r1, b.maxRow);
        if (c2 === MAX_COLS - 1) c2 = Math.max(v.c1, b.maxCol);
      }
      return Matrix.fill(r2 - v.r1 + 1, c2 - v.c1 + 1, (r, c) => this.getValue(v.sheet, v.r1 + r, v.c1 + c));
    }
    return v;
  }

  /** Cells this formula cell reads (for "trace precedents" highlighting). */
  precedentsOf(sheet: Sheet, row: number, col: number): { sheetId: number; range: Range }[] {
    const info = this.formulas.get(gkey(sheet.id, row, col));
    if (!info) return [];
    const out: { sheetId: number; range: Range }[] = info.cells.map((g) => ({
      sheetId: gSheet(g),
      range: { r1: gRow(g), c1: gCol(g), r2: gRow(g), c2: gCol(g) },
    }));
    for (const r of info.ranges) out.push({ sheetId: r.sheetId, range: { r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 } });
    return out;
  }

  // ---------- registration ----------

  /** Formula text with table references (Sales[Amount], [@Price]) as ordinary ranges. */
  private withTables(text: string, sheet: Sheet, row: number, col: number): string {
    return this.wb.hasTables || text.includes('[') ? expandStructured(text, this.wb, sheet, row, col) : text;
  }

  private register(sheet: Sheet, row: number, col: number, text: string): void {
    const g = gkey(sheet.id, row, col);
    const parsed = parseFormula(this.withTables(text, sheet, row, col));
    const info: FormulaInfo = {
      sheetId: sheet.id,
      row,
      col,
      text,
      parsed,
      cells: [],
      ranges: [],
      volatile: false,
      unsupported: [],
    };
    let usesExternal = false;
    for (const f of parsed.funcs) {
      if (VOLATILE_FUNCS.has(f)) info.volatile = true;
      const def = FUNCTIONS.get(f);
      if (def) {
        if (def.external) usesExternal = true;
        continue;
      }
      // VBA functions can read any cell, so the ones marked Application.Volatile always recalculate
      const udf = this.userFunctions?.(f);
      if (udf) {
        if (udf.volatile) info.volatile = true;
      } else info.unsupported.push(f);
    }
    const addRef = (sheetName: string | undefined, r1: number, c1: number, r2: number, c2: number) => {
      const target = sheetName !== undefined ? this.wb.sheetByName(sheetName) : sheet;
      if (!target) return;
      if (r1 === r2 && c1 === c2) info.cells.push(gkey(target.id, r1, c1));
      else info.ranges.push({ sheetId: target.id, r1, c1, r2, c2, dep: g });
    };
    for (const ref of parsed.refs) addRef(ref.sheet, ref.r1, ref.c1, ref.r2, ref.c2);
    // Defined names: depend on whatever they refer to
    for (const name of parsed.names) {
      const dn = this.wb.getName(name);
      if (!dn) continue;
      const np = parseFormula(dn.ref);
      for (const ref of np.refs) addRef(ref.sheet, ref.r1, ref.c1, ref.r2, ref.c2);
      for (const f of np.funcs) if (VOLATILE_FUNCS.has(f)) info.volatile = true;
    }
    // Range operators (A1:INDEX(...)) and INDIRECT can't be tracked statically
    if (parsed.tokens.some((t) => t.type === 'op' && t.text === ':')) info.volatile = true;

    this.formulas.set(g, info);
    for (const c of info.cells) {
      let s = this.cellDeps.get(c);
      if (!s) this.cellDeps.set(c, (s = new Set()));
      s.add(g);
    }
    for (const r of info.ranges) {
      let bySheet = this.rangeDeps.get(r.sheetId);
      if (!bySheet) this.rangeDeps.set(r.sheetId, (bySheet = new Map()));
      for (let b = Math.floor(r.c1 / BLOCK); b <= Math.floor(r.c2 / BLOCK); b++) {
        let set = bySheet.get(b);
        if (!set) bySheet.set(b, (set = new Set()));
        set.add(r);
      }
    }
    if (info.volatile) this.volatiles.add(g);
    if (usesExternal) this.externals.add(g);
    this.dirty.add(g);
  }

  private unregister(g: number): void {
    const info = this.formulas.get(g);
    if (!info) return;
    for (const c of info.cells) this.cellDeps.get(c)?.delete(g);
    for (const r of info.ranges) {
      const bySheet = this.rangeDeps.get(r.sheetId);
      if (!bySheet) continue;
      for (let b = Math.floor(r.c1 / BLOCK); b <= Math.floor(r.c2 / BLOCK); b++) bySheet.get(b)?.delete(r);
    }
    this.formulas.delete(g);
    this.volatiles.delete(g);
    this.externals.delete(g);
    this.values.delete(g);
    this.dirty.delete(g);
    this.blockedSpills.delete(g);
    this.clearSpill(g);
  }

  // ---------- dirtiness ----------

  private dependentsOfCell(g: number, out: number[]): void {
    const direct = this.cellDeps.get(g);
    if (direct) for (const d of direct) out.push(d);
    const bySheet = this.rangeDeps.get(gSheet(g));
    if (!bySheet) return;
    const row = gRow(g);
    const col = gCol(g);
    const set = bySheet.get(Math.floor(col / BLOCK));
    if (!set) return;
    for (const r of set) {
      if (row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2) out.push(r.dep);
    }
  }

  private dependentsOfRange(sheetId: number, r1: number, c1: number, r2: number, c2: number, out: number[]): void {
    const size = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (size <= 4096) {
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
        const d = this.cellDeps.get(gkey(sheetId, r, c));
        if (d) for (const x of d) out.push(x);
      }
    } else {
      for (const [k, deps] of this.cellDeps) {
        if (gSheet(k) !== sheetId) continue;
        const r = gRow(k);
        const c = gCol(k);
        if (r >= r1 && r <= r2 && c >= c1 && c <= c2) for (const x of deps) out.push(x);
      }
    }
    const bySheet = this.rangeDeps.get(sheetId);
    if (!bySheet) return;
    const seen = new Set<RangeDep>();
    for (let b = Math.floor(c1 / BLOCK); b <= Math.floor(c2 / BLOCK); b++) {
      const set = bySheet.get(b);
      if (!set) continue;
      for (const r of set) {
        if (seen.has(r)) continue;
        seen.add(r);
        if (r.r1 <= r2 && r1 <= r.r2 && r.c1 <= c2 && c1 <= r.c2) out.push(r.dep);
      }
    }
  }

  /** Marks everything that (transitively) depends on the given cells as dirty. */
  private invalidate(start: number[]): void {
    const queue = [...start];
    const visited = new Set<number>();
    while (queue.length) {
      const g = queue.pop()!;
      if (visited.has(g)) continue;
      visited.add(g);
      const deps: number[] = [];
      this.dependentsOfCell(g, deps);
      const sp = this.spills.get(g);
      if (sp) this.dependentsOfRange(sp.sheetId, sp.r1, sp.c1, sp.r2, sp.c2, deps);
      for (const d of deps) {
        if (!this.dirty.has(d)) this.dirty.add(d);
        if (!visited.has(d)) queue.push(d);
      }
    }
  }

  // ---------- computation ----------

  private computeWithRestarts(start: number): void {
    const stack = [start];
    const onStack = new Set<number>([start]);
    let guard = 0;
    while (stack.length && guard++ < 100000) {
      const top = stack[stack.length - 1];
      if (!this.dirty.has(top)) {
        stack.pop();
        onStack.delete(top);
        continue;
      }
      try {
        this.depth = 0;
        this.compute(top);
        stack.pop();
        onStack.delete(top);
      } catch (e) {
        if (!(e instanceof TooDeep)) throw e;
        this.computing.clear();
        if (onStack.has(e.g)) {
          // A very long circular chain
          this.store(e.g, Errors.value(), true);
          continue;
        }
        stack.push(e.g);
        onStack.add(e.g);
      }
    }
  }

  private compute(g: number): void {
    if (!this.dirty.has(g) || this.computing.has(g)) return;
    const info = this.formulas.get(g);
    if (!info) {
      this.dirty.delete(g);
      return;
    }
    if (this.depth > MAX_DEPTH) throw new TooDeep(g);
    const sheet = this.wb.sheetById(info.sheetId);
    if (!sheet) {
      this.dirty.delete(g);
      return;
    }
    this.computing.add(g);
    this.depth++;
    let result: Value;
    try {
      if (!info.parsed.ast) {
        result = Errors.value(`This formula has a syntax error: ${info.parsed.error?.message ?? 'unknown problem'}`);
      } else {
        result = evaluateFormula(info.parsed.ast, this.makeCtx(sheet, info.row, info.col, 0, 0));
      }
    } finally {
      this.depth--;
      this.computing.delete(g);
    }
    if (info.unsupported.length) {
      const cached = sheet.getCell(info.row, info.col)?.cached;
      if (cached !== undefined) result = cached;
    }
    this.store(g, result, false);
  }

  private store(g: number, result: Value, circular: boolean): void {
    this.dirty.delete(g);
    const info = this.formulas.get(g);
    if (!info) return;
    const sheet = this.wb.sheetById(info.sheetId)!;
    const origin = { sheet: sheet.name, row: info.row, col: info.col };
    if (circular) {
      this.values.set(g, new ErrorValue('#CIRC!', 'This formula depends on its own result (a circular reference).', origin));
      this.clearSpill(g);
      return;
    }
    let v: Scalar | Matrix;
    if (result instanceof RangeRef) {
      v = result.isCell ? this.getValue(result.sheet, result.r1, result.c1) : this.materialize(result, sheet);
    } else v = result;

    const oldSpill = this.spills.get(g);
    if (v instanceof Matrix && (v.rows > 1 || v.cols > 1)) {
      const sp: Spill = { sheetId: info.sheetId, r1: info.row, c1: info.col, r2: info.row + v.rows - 1, c2: info.col + v.cols - 1, data: v.data, fmts: v.formats };
      if (v.rows === 0 || v.cols === 0) {
        this.clearSpill(g);
        this.values.set(g, Errors.calc('The result is an empty array.').withOrigin(origin));
      } else if (sp.r2 >= MAX_ROWS || sp.c2 >= MAX_COLS) {
        this.clearSpill(g);
        this.values.set(g, new ErrorValue('#SPILL!', 'The result is too big — it would run past the edge of the sheet.', origin));
      } else {
        const blocker = this.findBlocker(sheet, sp, g);
        if (blocker) {
          this.clearSpill(g);
          this.blockedSpills.set(g, sp);
          const size = `${v.rows}×${v.cols}`;
          this.values.set(
            g,
            new ErrorValue(
              '#SPILL!',
              `This formula returns ${size} values that need to fill ${formatCell(sp.r1, sp.c1)}:${formatCell(sp.r2, sp.c2)}, but ${formatCell(blocker.row, blocker.col)} isn't empty. Clear it to let the results spill.`,
              origin,
            ),
          );
        } else {
          this.blockedSpills.delete(g);
          this.setSpill(g, sp);
          const first = v.data[0][0];
          this.values.set(g, first === null ? 0 : first instanceof ErrorValue ? first.withOrigin(origin) : first);
        }
      }
    } else {
      this.blockedSpills.delete(g);
      this.clearSpill(g);
      let s = v instanceof Matrix ? v.get(0, 0) : v;
      if (s === null) s = 0;
      if (s instanceof ErrorValue && !s.origin) s = s.withOrigin(origin);
      this.values.set(g, s);
    }
    const newSpill = this.spills.get(g);
    if (oldSpill !== newSpill && (oldSpill || newSpill)) {
      const changed =
        !oldSpill || !newSpill || oldSpill.r2 !== newSpill.r2 || oldSpill.c2 !== newSpill.c2 || oldSpill.data !== newSpill.data;
      if (changed) {
        const deps: number[] = [];
        if (oldSpill) this.dependentsOfRange(oldSpill.sheetId, oldSpill.r1, oldSpill.c1, oldSpill.r2, oldSpill.c2, deps);
        if (newSpill) this.dependentsOfRange(newSpill.sheetId, newSpill.r1, newSpill.c1, newSpill.r2, newSpill.c2, deps);
        const fresh = deps.filter((d) => d !== g && !this.computing.has(d));
        for (const d of fresh) this.dirty.add(d);
        this.invalidate(fresh);
      }
    }
  }

  private findBlocker(sheet: Sheet, sp: Spill, anchor: number): { row: number; col: number } | null {
    for (let r = sp.r1; r <= sp.r2; r++) {
      for (let c = sp.c1; c <= sp.c2; c++) {
        if (r === sp.r1 && c === sp.c1) continue;
        const cell = sheet.cells.get(cellKey(r, c));
        if (cell && (cell.f !== undefined || (cell.v !== undefined && cell.v !== null && cell.v !== ''))) return { row: r, col: c };
        const owner = this.spillOwner.get(gkey(sheet.id, r, c));
        if (owner !== undefined && owner !== anchor) return { row: r, col: c };
      }
    }
    return null;
  }

  private setSpill(g: number, sp: Spill): void {
    const old = this.spills.get(g);
    if (old) {
      for (let r = old.r1; r <= old.r2; r++)
        for (let c = old.c1; c <= old.c2; c++) {
          const k = gkey(old.sheetId, r, c);
          if (this.spillOwner.get(k) === g) this.spillOwner.delete(k);
        }
    }
    this.spills.set(g, sp);
    for (let r = sp.r1; r <= sp.r2; r++)
      for (let c = sp.c1; c <= sp.c2; c++) {
        if (r === sp.r1 && c === sp.c1) continue;
        this.spillOwner.set(gkey(sp.sheetId, r, c), g);
      }
  }

  private clearSpill(g: number): void {
    const old = this.spills.get(g);
    if (!old) return;
    this.spills.delete(g);
    for (let r = old.r1; r <= old.r2; r++)
      for (let c = old.c1; c <= old.c2; c++) {
        const k = gkey(old.sheetId, r, c);
        if (this.spillOwner.get(k) === g) this.spillOwner.delete(k);
      }
  }

  // ---------- evaluation context ----------

  private makeCtx(sheet: Sheet, row: number, col: number, dr: number, dc: number): Ctx {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const engine = this;
    return {
      wb: this.wb,
      sheet,
      row,
      col,
      udf: (name: string) => engine.userFunctions?.(name) ?? null,
      dr,
      dc,
      scope: null,
      get(s: Sheet, r: number, c: number): Scalar {
        const g = gkey(s.id, r, c);
        if (engine.computing.has(g)) {
          return new ErrorValue(
            '#CIRC!',
            `Circular reference: ${s.name}!${formatCell(r, c)} depends on its own result.`,
            { sheet: s.name, row: r, col: c },
          );
        }
        return engine.getValue(s, r, c);
      },
      eval(node: Node): Value {
        return evaluate(node, this);
      },
      bounds(s: Sheet) {
        return engine.usedBounds(s);
      },
      spillRange(s: Sheet, r: number, c: number) {
        const g = gkey(s.id, r, c);
        if (engine.dirty.has(g)) engine.compute(g);
        const sp = engine.spills.get(g);
        return sp ? new RangeRef(s, sp.r1, sp.c1, sp.r2, sp.c2) : null;
      },
      isFormula(s: Sheet, r: number, c: number) {
        return s.getCell(r, c)?.f !== undefined;
      },
      formulaText(s: Sheet, r: number, c: number) {
        const f = s.getCell(r, c)?.f;
        return f === undefined ? null : f;
      },
      forEachInRange(ref: RangeRef, cb: (v: Scalar, r: number, c: number) => void | boolean) {
        engine.forEachInRange(ref, cb, (s, r, c) => this.get(s, r, c));
      },
    };
  }

  usedBounds(sheet: Sheet): { maxRow: number; maxCol: number } {
    const b = sheet.bounds();
    let maxRow = b.maxRow;
    let maxCol = b.maxCol;
    for (const sp of this.spills.values()) {
      if (sp.sheetId !== sheet.id) continue;
      if (sp.r2 > maxRow) maxRow = sp.r2;
      if (sp.c2 > maxCol) maxCol = sp.c2;
    }
    return { maxRow, maxCol };
  }

  /** Iterates non-empty values in a range; uses sparse iteration for big, mostly empty ranges. */
  forEachInRange(
    ref: RangeRef,
    cb: (v: Scalar, r: number, c: number) => void | boolean,
    get: (s: Sheet, r: number, c: number) => Scalar = (s, r, c) => this.getValue(s, r, c),
  ): void {
    const sheet = ref.sheet;
    const b = this.usedBounds(sheet);
    const r2 = Math.min(ref.r2, b.maxRow);
    const c2 = Math.min(ref.c2, b.maxCol);
    if (r2 < ref.r1 || c2 < ref.c1) return;
    const area = (r2 - ref.r1 + 1) * (c2 - ref.c1 + 1);
    let spillCells = 0;
    for (const sp of this.spills.values()) if (sp.sheetId === sheet.id) spillCells += (sp.r2 - sp.r1 + 1) * (sp.c2 - sp.c1 + 1);
    if (area <= 2 * (sheet.cells.size + spillCells) + 1024) {
      for (let r = ref.r1; r <= r2; r++) {
        for (let c = ref.c1; c <= c2; c++) {
          if (!this.occupied(sheet, r, c)) continue;
          const v = get(sheet, r, c);
          if (v === null) continue;
          if (cb(v, r, c) === false) return;
        }
      }
      return;
    }
    // Sparse: collect candidate cells
    const keys: number[] = [];
    for (const key of sheet.cells.keys()) {
      const r = Math.floor(key / MAX_COLS);
      const c = key % MAX_COLS;
      if (r >= ref.r1 && r <= r2 && c >= ref.c1 && c <= c2) keys.push(key);
    }
    for (const sp of this.spills.values()) {
      if (sp.sheetId !== sheet.id) continue;
      for (let r = Math.max(sp.r1, ref.r1); r <= Math.min(sp.r2, r2); r++)
        for (let c = Math.max(sp.c1, ref.c1); c <= Math.min(sp.c2, c2); c++) {
          if (r === sp.r1 && c === sp.c1) continue;
          keys.push(r * MAX_COLS + c);
        }
    }
    keys.sort((a, b2) => a - b2);
    let prev = -1;
    for (const key of keys) {
      if (key === prev) continue;
      prev = key;
      const r = Math.floor(key / MAX_COLS);
      const c = key % MAX_COLS;
      const v = get(sheet, r, c);
      if (v === null) continue;
      if (cb(v, r, c) === false) return;
    }
  }

  /** Whether an address has a cell or a spilled value. */
  private occupied(sheet: Sheet, r: number, c: number): boolean {
    return sheet.cells.has(cellKey(r, c)) || this.spillOwner.has(gkey(sheet.id, r, c));
  }
}
