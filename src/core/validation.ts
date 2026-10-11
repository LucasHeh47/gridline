// Data validation: rules about what a cell may contain (a whole number between 1 and 500, one of a
// list of choices…). Works the way Excel's does, including its quirks: typed lists match exactly
// (case-sensitive) while lists taken from cells ignore case, text fails number rules, and blank cells
// pass when "Ignore blank" is on.
import { formatRange, rangeContains, type Range } from './address';
import type { Doc } from './doc';
import { translateFormula } from './formula/refs';
import { Matrix, RangeRef } from './formula/types';
import { formatGeneral, formatValue } from './numfmt';
import { ErrorValue, type Scalar } from './values';
import type { Sheet } from './workbook';

/** 'any': no restriction, the rule only shows an input message. */
export type ValidationType = 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
export type ValidationOp = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';

export interface DataValidation {
  id: string;
  ranges: Range[];
  type: ValidationType;
  operator: ValidationOp;
  /** Formula text without "=": a number, a reference, or for lists '"Yes,No"' or a range like $E$1:$E$9. */
  formula1: string;
  formula2?: string;
  /** "Ignore blank": empty cells pass. */
  allowBlank: boolean;
  /** Lists only: show the dropdown arrow in the cell. */
  dropdown: boolean;
  /** Show an alert when someone types something that breaks the rule. */
  showError: boolean;
  errorStyle: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  /** Show a message when the cell is selected. */
  showPrompt: boolean;
  promptTitle?: string;
  prompt?: string;
}

export const OPERATORS: [ValidationOp, string][] = [
  ['between', 'between'],
  ['notBetween', 'not between'],
  ['equal', 'equal to'],
  ['notEqual', 'not equal to'],
  ['greaterThan', 'greater than'],
  ['lessThan', 'less than'],
  ['greaterThanOrEqual', 'greater than or equal to'],
  ['lessThanOrEqual', 'less than or equal to'],
];

export const TYPES: [ValidationType, string][] = [
  ['any', 'Any value (message only)'],
  ['whole', 'Whole number'],
  ['decimal', 'Decimal number'],
  ['list', 'List of choices'],
  ['date', 'Date'],
  ['time', 'Time'],
  ['textLength', 'Text length'],
  ['custom', 'Custom formula'],
];

export function newValidationId(): string {
  return 'dv-' + Math.random().toString(36).slice(2, 10);
}

/** The rule covering a cell, if any. */
export function validationAt(sheet: Sheet, row: number, col: number): DataValidation | undefined {
  for (let i = sheet.validations.length - 1; i >= 0; i--) {
    const v = sheet.validations[i];
    if (v.ranges.some((r) => rangeContains(r, row, col))) return v;
  }
  return undefined;
}

/** Evaluates formula1/formula2 for a cell (relative references move with the cell, like Excel). */
function evalFor(doc: Doc, sheet: Sheet, dv: DataValidation, f: string, row: number, col: number): Scalar | Matrix | RangeRef {
  const t = f.trim().replace(/^=/, '');
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return Number(t);
  const origin = dv.ranges[0];
  const text = translateFormula(t, row - origin.r1, col - origin.c1);
  const v = doc.engine.evalText(text, sheet, row, col);
  return v;
}

function firstScalar(v: Scalar | Matrix | RangeRef): Scalar {
  if (v instanceof Matrix) return v.get(0, 0);
  if (v instanceof RangeRef) return null;
  return v;
}

function cellText(doc: Doc, sheet: Sheet, r: number, c: number, v: Scalar): string {
  if (v === null) return '';
  return formatValue(v, typeof v === 'number' ? doc.numFmtAt(sheet, r, c) : undefined, 255).text;
}

/** The choices of a list rule, and whether matching is case-sensitive (typed lists are). */
export function listItems(doc: Doc, sheet: Sheet, dv: DataValidation, row: number, col: number): { items: string[]; caseSensitive: boolean } {
  const f = dv.formula1.trim();
  if (f.startsWith('"')) {
    const inner = f.slice(1, f.endsWith('"') ? -1 : undefined).replace(/""/g, '"');
    return { items: inner.split(',').map((s) => s.trim()).filter((s) => s !== ''), caseSensitive: true };
  }
  const t = f.replace(/^=/, '');
  const origin = dv.ranges[0];
  const ref = doc.engine.evalText(translateFormula(t, row - origin.r1, col - origin.c1), sheet, row, col);
  const items: string[] = [];
  if (ref instanceof Matrix) {
    for (let r = 0; r < ref.rows; r++)
      for (let c = 0; c < ref.cols; c++) {
        const v = ref.get(r, c);
        if (v === null || v instanceof ErrorValue) continue;
        const s = typeof v === 'number' ? formatGeneral(v, 255) : String(v);
        if (s.trim() !== '') items.push(s.trim());
      }
  } else if (ref !== null && !(ref instanceof ErrorValue)) {
    const s = typeof ref === 'number' ? formatGeneral(ref, 255) : String(ref);
    if (s.trim() !== '') items.push(s.trim());
  }
  return { items, caseSensitive: false };
}

function compareOp(op: ValidationOp, x: number, a: number, b: number): boolean {
  switch (op) {
    case 'between':
      return x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween':
      return x < Math.min(a, b) || x > Math.max(a, b);
    case 'equal':
      return x === a;
    case 'notEqual':
      return x !== a;
    case 'greaterThan':
      return x > a;
    case 'lessThan':
      return x < a;
    case 'greaterThanOrEqual':
      return x >= a;
    default:
      return x <= a;
  }
}

/** Whether `value` (the cell's value at row, col) satisfies the rule. */
export function isValid(doc: Doc, sheet: Sheet, dv: DataValidation, row: number, col: number, value: Scalar): boolean {
  if (dv.type === 'any') return true;
  if (value === null || value === '') return dv.allowBlank;
  if (value instanceof ErrorValue) return false;
  const num = (f: string | undefined): number | null => {
    if (f === undefined || f.trim() === '') return null;
    const v = firstScalar(evalFor(doc, sheet, dv, f, row, col));
    return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : null;
  };
  switch (dv.type) {
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time': {
      if (typeof value !== 'number') return false;
      if (dv.type === 'whole' && !Number.isInteger(value)) return false;
      const a = num(dv.formula1);
      if (a === null) return true;
      const b = num(dv.formula2) ?? a;
      let x = value;
      if (dv.type === 'time') x = value - Math.floor(value);
      return compareOp(dv.operator, x, a, b);
    }
    case 'textLength': {
      const len = cellText(doc, sheet, row, col, value).length;
      const a = num(dv.formula1);
      if (a === null) return true;
      return compareOp(dv.operator, len, a, num(dv.formula2) ?? a);
    }
    case 'list': {
      const { items, caseSensitive } = listItems(doc, sheet, dv, row, col);
      const text = (typeof value === 'number' ? formatGeneral(value, 255) : String(value)).trim();
      return items.some((it) => (caseSensitive ? it === text : it.toLowerCase() === text.toLowerCase()));
    }
    case 'custom': {
      const v = firstScalar(evalFor(doc, sheet, dv, dv.formula1, row, col));
      return v === true || (typeof v === 'number' && v !== 0);
    }
  }
  return true;
}

/** Checks the current value of a cell against its rule (true when there's no rule). */
export function cellIsValid(doc: Doc, sheet: Sheet, row: number, col: number): boolean {
  const dv = validationAt(sheet, row, col);
  if (!dv) return true;
  return isValid(doc, sheet, dv, row, col, doc.engine.getValue(sheet, row, col));
}

/** Plain-English description of a rule: "a whole number between 1 and 500". */
export function describeRule(dv: DataValidation): string {
  const v = (f?: string) => (f ?? '').replace(/^=/, '');
  const op = OPERATORS.find(([o]) => o === dv.operator)?.[1] ?? dv.operator;
  const range = dv.operator === 'between' || dv.operator === 'notBetween' ? `${op} ${v(dv.formula1)} and ${v(dv.formula2)}` : `${op} ${v(dv.formula1)}`;
  switch (dv.type) {
    case 'any':
      return 'any value';
    case 'whole':
      return `a whole number ${range}`;
    case 'decimal':
      return `a number ${range}`;
    case 'date':
      return `a date ${range}`;
    case 'time':
      return `a time ${range}`;
    case 'textLength':
      return `text with a length ${range}`;
    case 'list': {
      const f = dv.formula1.trim();
      return f.startsWith('"') ? `one of: ${f.slice(1, -1).split(',').map((s) => s.trim()).join(', ')}` : `one of the values in ${v(f)}`;
    }
    case 'custom':
      return `a value where =${v(dv.formula1)} is TRUE`;
  }
}

/** Where a rule applies, for the panel ("E5:E1204"). */
export function rangesText(dv: DataValidation): string {
  return dv.ranges.map(formatRange).join(', ');
}

/** Cells (in the used part of the given area) that break their rule — for "Circle invalid data". */
export function invalidCells(doc: Doc, sheet: Sheet, limit = 5000): { row: number; col: number }[] {
  const out: { row: number; col: number }[] = [];
  const b = doc.engine.usedBounds(sheet);
  for (const dv of sheet.validations) {
    for (const r of dv.ranges) {
      for (let row = r.r1; row <= Math.min(r.r2, b.maxRow); row++)
        for (let col = r.c1; col <= Math.min(r.c2, b.maxCol); col++) {
          if (!isValid(doc, sheet, dv, row, col, doc.engine.getValue(sheet, row, col))) {
            out.push({ row, col });
            if (out.length >= limit) return out;
          }
        }
    }
  }
  return out;
}
