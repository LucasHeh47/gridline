import { Errors, ErrorValue, type ErrorCode, type Scalar } from '../../values';
import { scalarOf, toNumber } from '../coerce';
import { Matrix, RangeRef } from '../types';
import { define } from './registry';

const is = (name: string, desc: string, test: (v: Scalar) => boolean) =>
  define(name, 'Information', 'value', desc, (ctx, [a]) => test(scalarOf(ctx, a)), { lift: true });

is('ISBLANK', 'Returns TRUE if the cell is empty.', (v) => v === null);
is('ISNUMBER', 'Returns TRUE if the value is a number.', (v) => typeof v === 'number');
is('ISTEXT', 'Returns TRUE if the value is text.', (v) => typeof v === 'string');
is('ISNONTEXT', 'Returns TRUE if the value is not text.', (v) => typeof v !== 'string');
is('ISLOGICAL', 'Returns TRUE if the value is TRUE or FALSE.', (v) => typeof v === 'boolean');
is('ISERROR', 'Returns TRUE if the value is any error.', (v) => v instanceof ErrorValue);
is('ISERR', 'Returns TRUE if the value is any error except #N/A.', (v) => v instanceof ErrorValue && v.code !== '#N/A');
is('ISNA', 'Returns TRUE if the value is #N/A.', (v) => v instanceof ErrorValue && v.code === '#N/A');

define('ISEVEN', 'Information', 'number', 'Returns TRUE if the number is even.', (ctx, [a]) => {
  const n = toNumber(scalarOf(ctx, a));
  return n instanceof ErrorValue ? n : Math.trunc(n) % 2 === 0;
}, { lift: true });
define('ISODD', 'Information', 'number', 'Returns TRUE if the number is odd.', (ctx, [a]) => {
  const n = toNumber(scalarOf(ctx, a));
  return n instanceof ErrorValue ? n : Math.abs(Math.trunc(n) % 2) === 1;
}, { lift: true });
define('ISREF', 'Information', 'value', 'Returns TRUE if the value is a reference.', (_ctx, [a]) => a instanceof RangeRef);
define('ISFORMULA', 'Information', 'reference', 'Returns TRUE if the referenced cell contains a formula.', (ctx, [a]) => {
  if (!(a instanceof RangeRef)) return Errors.value('ISFORMULA needs a cell reference.');
  return ctx.isFormula(a.sheet, a.r1, a.c1);
});
define('FORMULATEXT', 'Information', 'reference', 'Returns the formula in a cell as text.', (ctx, [a]) => {
  if (!(a instanceof RangeRef)) return Errors.value('FORMULATEXT needs a cell reference.');
  const f = ctx.formulaText(a.sheet, a.r1, a.c1);
  return f === null ? Errors.na('That cell doesn\'t contain a formula.') : '=' + f;
});
define('NA', 'Information', '', 'Returns the #N/A error ("value not available").', () => Errors.na('NA() was used to mark this value as not available.'));

const ERROR_NUMS: Partial<Record<ErrorCode, number>> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#SPILL!': 9,
  '#CALC!': 14,
};
define('ERROR.TYPE', 'Information', 'error_val', 'Returns a number for an error type (1=#NULL!, 2=#DIV/0!, 3=#VALUE!...).', (ctx, [a]) => {
  const v = scalarOf(ctx, a);
  if (!(v instanceof ErrorValue)) return Errors.na('The value isn\'t an error.');
  return ERROR_NUMS[v.code] ?? Errors.na();
}, { lift: true });

define('TYPE', 'Information', 'value', 'Returns the type of a value: 1 number, 2 text, 4 logical, 16 error, 64 array.', (ctx, [a]) => {
  if (a instanceof Matrix || (a instanceof RangeRef && !a.isCell)) return 64;
  const v = scalarOf(ctx, a);
  if (v === null || typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 4;
  return 16;
});
define('N', 'Information', 'value', 'Converts a value to a number (text becomes 0).', (ctx, [a]) => {
  const v = scalarOf(ctx, a);
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof ErrorValue) return v;
  return 0;
}, { lift: true });

define('SHEET', 'Information', '[value]', 'Returns the sheet number of a reference.', (ctx, [a]) => {
  const sheet = a instanceof RangeRef ? a.sheet : ctx.sheet;
  if (typeof a === 'string') {
    const s = ctx.wb.sheetByName(a);
    return s ? ctx.wb.sheets.indexOf(s) + 1 : Errors.na();
  }
  return ctx.wb.sheets.indexOf(sheet) + 1;
});
define('SHEETS', 'Information', '[reference]', 'Returns the number of sheets.', (ctx) => ctx.wb.sheets.length);
