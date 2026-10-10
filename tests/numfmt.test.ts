import { describe, expect, it } from 'vitest';
import { dateToSerial } from '../src/core/dates';
import { parseInput } from '../src/core/input';
import { adjustDecimals, formatGeneral, formatValue } from '../src/core/numfmt';

const f = (v: number | string, fmt: string) => formatValue(v, fmt).text;

describe('number formats', () => {
  it('formats General numbers', () => {
    expect(formatGeneral(0)).toBe('0');
    expect(formatGeneral(1234567)).toBe('1234567');
    expect(formatGeneral(0.1 + 0.2)).toBe('0.3');
    expect(formatGeneral(1 / 3)).toBe('0.333333333');
    expect(formatGeneral(123456789012)).toBe('1.23457E+11');
    expect(formatGeneral(-2.5)).toBe('-2.5');
  });

  it('formats fixed decimals and thousands', () => {
    expect(f(1234.567, '0.00')).toBe('1234.57');
    expect(f(1234.567, '#,##0')).toBe('1,235');
    expect(f(1234.567, '#,##0.00')).toBe('1,234.57');
    expect(f(-1234.5, '#,##0.00')).toBe('-1,234.50');
    expect(f(0.5, '0%')).toBe('50%');
    expect(f(0.1234, '0.00%')).toBe('12.34%');
    expect(f(5, '000')).toBe('005');
    expect(f(0.5, '#.##')).toBe('.5');
    expect(f(1234567, '#,##0,')).toBe('1,235');
  });

  it('formats currency and sections', () => {
    expect(f(1234.5, '"$"#,##0.00')).toBe('$1,234.50');
    expect(f(-1234.5, '"$"#,##0.00;("$"#,##0.00)')).toBe('($1,234.50)');
    expect(f(0, '0;-0;"zero"')).toBe('zero');
    expect(formatValue(-5, '0;[Red]-0').color).toBeDefined();
    expect(f(1234.5, '[$€-407]#,##0.00')).toBe('€1,234.50');
  });

  it('formats scientific', () => {
    expect(f(12345, '0.00E+00')).toBe('1.23E+04');
    expect(f(0.00012, '0.0E+00')).toBe('1.2E-04');
  });

  it('formats dates and times', () => {
    const d = dateToSerial(2025, 3, 14);
    expect(f(d, 'yyyy-mm-dd')).toBe('2025-03-14');
    expect(f(d, 'm/d/yyyy')).toBe('3/14/2025');
    expect(f(d, 'mmm d, yyyy')).toBe('Mar 14, 2025');
    expect(f(d, 'dddd')).toBe('Friday');
    expect(f(d + 0.5 + 1 / 24 / 2, 'h:mm AM/PM')).toBe('12:30 PM');
    expect(f(d + 0.75, 'hh:mm:ss')).toBe('18:00:00');
    expect(f(1.5, '[h]:mm')).toBe('36:00');
    expect(f(1, 'yyyy-mm-dd')).toBe('1900-01-01');
    expect(f(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(f(61, 'yyyy-mm-dd')).toBe('1900-03-01');
  });

  it('formats fractions and phone numbers', () => {
    expect(f(1.5, '# ?/?')).toBe('1 1/2');
    expect(f(0.75, '?/?')).toBe('3/4');
    expect(f(5551234567, '(000) 000-0000')).toBe('(555) 123-4567');
  });

  it('formats text sections', () => {
    expect(f('abc', '@')).toBe('abc');
    expect(f('abc', '0;0;0;"Name: "@')).toBe('Name: abc');
  });

  it('adjusts decimals', () => {
    expect(adjustDecimals(undefined, 1, 5)).toBe('0.0');
    expect(adjustDecimals('0.00', -1, 5)).toBe('0.0');
    expect(adjustDecimals('#,##0', 2, 5)).toBe('#,##0.00');
    expect(adjustDecimals('0.00%', -2, 5)).toBe('0%');
  });
});

describe('input parsing', () => {
  it('parses numbers with formats', () => {
    expect(parseInput('42')).toEqual({ value: 42, numFmt: undefined });
    expect(parseInput('$1,200.50')).toEqual({ value: 1200.5, numFmt: '"$"#,##0.00' });
    expect(parseInput('15%').value).toBeCloseTo(0.15);
    expect(parseInput('(50)').value).toBe(-50);
    expect(parseInput('1,234').value).toBe(1234);
  });

  it('keeps IDs with leading zeros and long numbers as text', () => {
    expect(parseInput('00123').value).toBe('00123');
    expect(parseInput('4111111111111111111').value).toBe('4111111111111111111');
  });

  it('parses dates and times', () => {
    expect(parseInput('2025-03-14').value).toBe(dateToSerial(2025, 3, 14));
    expect(parseInput('3/14/2025').value).toBe(dateToSerial(2025, 3, 14));
    expect(parseInput('Mar 14, 2025').value).toBe(dateToSerial(2025, 3, 14));
    expect(parseInput('14-Mar-2025').value).toBe(dateToSerial(2025, 3, 14));
    expect(parseInput('2:30 PM').value).toBeCloseTo(14.5 / 24);
    expect(parseInput('13/45/2025').value).toBe('13/45/2025');
  });

  it('parses formulas, booleans and quoted text', () => {
    expect(parseInput('=A1+1')).toEqual({ formula: 'A1+1' });
    expect(parseInput('true').value).toBe(true);
    expect(parseInput("'123").value).toBe('123');
    expect(parseInput('hello').value).toBe('hello');
  });
});
