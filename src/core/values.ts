// Cell value types and error values.

export type ErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#SPILL!'
  | '#CALC!'
  | '#CIRC!'
  | '#BUSY!'
  | '#CONNECT!';

export const ERROR_CODES: ErrorCode[] = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#SPILL!',
  '#CALC!',
  '#CIRC!',
  '#BUSY!',
  '#CONNECT!',
];

export interface ErrorOrigin {
  sheet: string;
  row: number;
  col: number;
}

/**
 * An error value. Unlike Excel, errors carry a plain-English explanation and the cell
 * the error started in, so the UI can explain it and jump to the source.
 */
export class ErrorValue {
  constructor(
    readonly code: ErrorCode,
    readonly message?: string,
    public origin?: ErrorOrigin,
  ) {}

  toString(): string {
    return this.code;
  }

  withOrigin(origin: ErrorOrigin): ErrorValue {
    if (this.origin) return this;
    return new ErrorValue(this.code, this.message, origin);
  }
}

/** null means an empty cell. */
export type Scalar = number | string | boolean | null | ErrorValue;

export function isError(v: unknown): v is ErrorValue {
  return v instanceof ErrorValue;
}

export function err(code: ErrorCode, message?: string): ErrorValue {
  return new ErrorValue(code, message);
}

export const Errors = {
  div0: (msg = 'Something is being divided by zero (or by an empty cell, which counts as 0).') =>
    err('#DIV/0!', msg),
  value: (msg = 'A value has the wrong type — for example, text was used where a number is needed.') =>
    err('#VALUE!', msg),
  ref: (msg = 'The formula refers to a cell that no longer exists (it was deleted, or the reference goes off the edge of the sheet).') =>
    err('#REF!', msg),
  name: (msg = 'The formula uses a name Gridline doesn\'t recognize.') => err('#NAME?', msg),
  num: (msg = 'The calculation produced a number that is invalid or too large.') => err('#NUM!', msg),
  na: (msg = 'A value is not available — usually a lookup that found no match.') => err('#N/A', msg),
  nul: (msg = 'The ranges in this formula don\'t intersect.') => err('#NULL!', msg),
  calc: (msg = 'The calculation couldn\'t produce a result (for example, FILTER found nothing).') =>
    err('#CALC!', msg),
};

/** General explanations shown for an error code when no specific message exists. */
export const ERROR_HELP: Record<ErrorCode, { title: string; fix: string }> = {
  '#DIV/0!': {
    title: 'Division by zero',
    fix: 'Check the divisor. Wrap the formula in IFERROR(…, 0) to show something else when the divisor is empty.',
  },
  '#VALUE!': {
    title: 'Wrong type of value',
    fix: 'Look for text (often a stray space or a number stored as text) where a number is expected.',
  },
  '#REF!': {
    title: 'Broken reference',
    fix: 'Re-point the formula at existing cells. This usually happens after deleting rows, columns or sheets.',
  },
  '#NAME?': {
    title: 'Unknown name',
    fix: 'Check the function name for typos, and make sure any text is wrapped in "double quotes".',
  },
  '#NUM!': {
    title: 'Invalid number',
    fix: 'Check the inputs — e.g. square roots of negative numbers or results too large to represent.',
  },
  '#N/A': {
    title: 'Value not found',
    fix: 'The lookup value isn\'t in the lookup range. Check spelling and extra spaces, or use IFNA(…, "not found").',
  },
  '#NULL!': {
    title: 'Empty intersection',
    fix: 'Use a comma or colon between ranges instead of a space.',
  },
  '#SPILL!': {
    title: 'Result can\'t spill',
    fix: 'This formula returns several values, but cells it needs to fill are not empty. Clear those cells.',
  },
  '#CALC!': {
    title: 'Calculation problem',
    fix: 'The function couldn\'t return a result, such as FILTER with no matching rows. Add the if_empty argument.',
  },
  '#BUSY!': {
    title: 'Loading data',
    fix: 'The data is being downloaded. The cell fills in automatically when it arrives.',
  },
  '#CONNECT!': {
    title: "Couldn't get the data",
    fix: 'Check your internet connection, then use "Refresh live data" (Ctrl+Alt+F5). Previously downloaded data keeps working offline.',
  },
  '#CIRC!': {
    title: 'Circular reference',
    fix: 'This formula depends on its own result. Change one of the formulas in the loop so it no longer refers back.',
  },
};

export function parseErrorCode(s: string): ErrorCode | null {
  const up = s.toUpperCase();
  for (const c of ERROR_CODES) if (c === up) return c;
  if (up === '#GETTING_DATA') return '#N/A';
  return null;
}

/** Returns a short type tag, useful for sorting and comparisons. */
export function typeRank(v: Scalar): number {
  // Excel ascending sort order: numbers, text, logicals, errors, blanks
  if (v === null || v === '') return 4;
  if (typeof v === 'number') return 0;
  if (typeof v === 'string') return 1;
  if (typeof v === 'boolean') return 2;
  return 3;
}

export function scalarEquals(a: Scalar, b: Scalar): boolean {
  if (a instanceof ErrorValue || b instanceof ErrorValue) {
    return a instanceof ErrorValue && b instanceof ErrorValue && a.code === b.code;
  }
  return a === b;
}
