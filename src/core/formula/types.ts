import type { Sheet, Workbook } from '../workbook';
import type { Scalar } from '../values';
import type { Node } from './parser';

/** A reference to a rectangle of cells (not yet read). */
export class RangeRef {
  constructor(
    readonly sheet: Sheet,
    readonly r1: number,
    readonly c1: number,
    readonly r2: number,
    readonly c2: number,
  ) {}

  get rows(): number {
    return this.r2 - this.r1 + 1;
  }

  get cols(): number {
    return this.c2 - this.c1 + 1;
  }

  get isCell(): boolean {
    return this.r1 === this.r2 && this.c1 === this.c2;
  }
}

/** An in-memory 2D array of values. */
export class Matrix {
  /** formats: optional number format per column, used to display spilled results (dates, money). */
  constructor(
    readonly data: Scalar[][],
    readonly formats?: (string | undefined)[],
  ) {}

  get rows(): number {
    return this.data.length;
  }

  get cols(): number {
    return this.data[0]?.length ?? 0;
  }

  get(r: number, c: number): Scalar {
    return this.data[r]?.[c] ?? null;
  }

  static fill(rows: number, cols: number, f: (r: number, c: number) => Scalar): Matrix {
    const data: Scalar[][] = new Array(rows);
    for (let r = 0; r < rows; r++) {
      const row: Scalar[] = new Array(cols);
      for (let c = 0; c < cols; c++) row[c] = f(r, c);
      data[r] = row;
    }
    return new Matrix(data);
  }

  static column(values: Scalar[]): Matrix {
    return new Matrix(values.map((v) => [v]));
  }

  static row(values: Scalar[]): Matrix {
    return new Matrix([values]);
  }
}

export type Value = Scalar | RangeRef | Matrix;
/** Function arguments: undefined marks an omitted argument, e.g. IF(A1,,1). */
export type Arg = Value | undefined;

/** A function written in the workbook's own VBA code, called from a formula (=VAULTVALUE(…)). */
export interface UserFunction {
  /** Recalculated after every change (the code calls Application.Volatile). */
  volatile: boolean;
  /** Arguments are evaluated: references stay references, so the code can get a Range. */
  call(ctx: Ctx, args: Arg[]): Value;
}

export interface Ctx {
  wb: Workbook;
  sheet: Sheet;
  row: number;
  col: number;
  /** Offset applied to relative references (used for conditional-formatting rules). */
  dr: number;
  dc: number;
  scope: Map<string, Value> | null;
  /** Reads a cell's current value, calculating it first if needed. */
  get(sheet: Sheet, row: number, col: number): Scalar;
  /** Evaluates an AST node in this context. */
  eval(node: Node): Value;
  /** Last used row/col on the sheet, including spilled arrays. */
  bounds(sheet: Sheet): { maxRow: number; maxCol: number };
  /** The spill range anchored at a cell, if any. */
  spillRange(sheet: Sheet, row: number, col: number): RangeRef | null;
  isFormula(sheet: Sheet, row: number, col: number): boolean;
  formulaText(sheet: Sheet, row: number, col: number): string | null;
  /** Iterates non-empty cells in a range in row-major order. */
  forEachInRange(ref: RangeRef, cb: (v: Scalar, r: number, c: number) => void | boolean): void;
  /** A function defined in the workbook's VBA code, when macros may run. */
  udf?(name: string): UserFunction | null;
}
