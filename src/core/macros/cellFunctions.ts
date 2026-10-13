// Functions written in a workbook's VBA code, used in formulas like any other function:
//   Function VAULTVALUE(vaultName As String, wantType As String) … End Function
//   =VAULTVALUE("Rent", "$")
// As in Excel: public Functions in standard modules; a reference passed to a Variant or Range
// parameter arrives as a Range object (so the code can read .NumberFormat or .Text), and to a typed
// parameter (As String, As Double) as the cell's value; the code can read the whole workbook but
// can't change it. Application.Volatile functions recalculate after every change.
import type { Doc } from '../doc';
import { Matrix, RangeRef, type Arg, type Ctx, type UserFunction, type Value as CellValue } from '../formula/types';
import { Errors, ErrorValue, type Scalar } from '../values';
import { CellSession, cellError, vbErrorValue } from './host';
import { RangeObj } from './range';
import { compiledProject } from './runner';
import type { Param } from './vba/ast';
import { MISSING } from './vba/builtins';
import type { CellFunctionInfo } from './vba/interpreter';
import { EMPTY, isObject, NULL, StopSignal, VbArray, VbDate, VbError, type Value } from './vba/values';

/** Parameter types that take a plain value (a range passed to them gives its cell's value). */
const VALUE_TYPES = new Set(['string', 'double', 'single', 'long', 'integer', 'longlong', 'longptr', 'byte', 'currency', 'decimal', 'boolean', 'date']);

/** The formula engine's lookup for a workbook's VBA Functions. */
export function userFunctions(doc: Doc): (name: string) => UserFunction | null {
  return (name) => {
    const info = compiledProject(doc.wb.macros).program.cellFunctions().get(name.toUpperCase());
    if (!info) return null;
    return { volatile: info.volatile, call: (ctx, args) => callFunction(doc, info, ctx, args) };
  };
}

function callFunction(doc: Doc, info: CellFunctionInfo, ctx: Ctx, args: Arg[]): CellValue {
  const program = compiledProject(doc.wb.macros).program;
  const session = new CellSession(doc, program, { sheet: ctx.sheet, row: ctx.row, col: ctx.col });
  const vbArgs: Value[] = [];
  for (let i = 0; i < args.length; i++) {
    const v = toVba(args[i], info.params[i], session, ctx);
    if (v instanceof ErrorValue) return v;
    vbArgs.push(v);
  }
  try {
    return toCell(program.callFromCell(session, info.name, vbArgs));
  } catch (e) {
    if (e instanceof VbError) return Errors.value(`${info.name}: ${e.message}`);
    if (e instanceof StopSignal) return Errors.value(`${info.name} stopped.`);
    // The engine's own signals (very long dependency chains) pass through
    throw e;
  }
}

function wantsObject(p: Param | undefined): boolean {
  return !p || !p.type || !VALUE_TYPES.has(p.type) || !!p.paramArray;
}

/** A formula argument as the VBA code receives it. */
function toVba(a: Arg, p: Param | undefined, session: CellSession, ctx: Ctx): Value | ErrorValue {
  if (a === undefined) return MISSING as unknown as Value;
  if (a instanceof RangeRef) {
    if (wantsObject(p)) return new RangeObj(session, a.sheet, [{ r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 }]);
    if (a.r1 !== a.r2 || a.c1 !== a.c2) return Errors.value(`"${p?.name}" takes one value, not a range of cells.`);
    return scalarToVba(ctx.get(a.sheet, a.r1, a.c1), p);
  }
  if (a instanceof Matrix) {
    if (!wantsObject(p)) {
      if (a.rows === 1 && a.cols === 1) return scalarToVba(a.get(0, 0), p);
      return Errors.value(`"${p?.name}" takes one value, not a list.`);
    }
    // Arrays reach VBA 1-based, rows × columns, like Excel passes them
    const arr = new VbArray([
      { lb: 1, ub: a.rows },
      { lb: 1, ub: a.cols },
    ]);
    for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) arr.set([r + 1, c + 1], scalarToVba(a.get(r, c), undefined) as Value);
    return arr;
  }
  return scalarToVba(a, p);
}

function scalarToVba(v: Scalar, p: Param | undefined): Value | ErrorValue {
  if (v instanceof ErrorValue) return p?.type && VALUE_TYPES.has(p.type) ? v : vbErrorValue(v);
  if (v === null) return EMPTY;
  return v;
}

/** What the function returned, as a cell value (an array spills like any array formula). */
function toCell(v: Value): CellValue {
  if (v instanceof VbArray) {
    if (v.dims.length === 1) {
      const d = v.dims[0];
      const row: Scalar[] = [];
      for (let i = d.lb; i <= d.ub; i++) row.push(scalar(v.get([i])));
      return row.length ? Matrix.row(row) : Errors.value('The function returned an empty array.');
    }
    if (v.dims.length === 2) {
      const [dr, dc] = v.dims;
      return Matrix.fill(dr.ub - dr.lb + 1, dc.ub - dc.lb + 1, (r, c) => scalar(v.get([dr.lb + r, dc.lb + c])));
    }
    return Errors.value('The function returned an array with more than two dimensions.');
  }
  if (isObject(v) && v instanceof RangeObj) {
    // A Range: its values
    return toCell(v.dflt([]) as Value);
  }
  return scalar(v);
}

function scalar(v: Value): Scalar {
  if (v === EMPTY || v === (MISSING as unknown)) return 0;
  if (v === NULL) return Errors.na('The function returned Null.');
  if (v === null) return Errors.value('The function returned Nothing (an object that was never set).');
  if (v instanceof VbDate) return v.serial;
  if (typeof v === 'number') return Number.isFinite(v) ? v : Errors.num();
  if (typeof v === 'string' || typeof v === 'boolean') return v;
  if (isObject(v)) {
    if (v.vbType === 'Error') return cellError(v);
    if (v instanceof RangeObj) {
      const d = v.dflt([]);
      return d instanceof VbArray ? scalar(d.data[0] ?? EMPTY) : scalar(d as Value);
    }
    return Errors.value(`The function returned a ${v.vbType} object instead of a value.`);
  }
  return Errors.value('The function returned something that can\'t go in a cell.');
}
