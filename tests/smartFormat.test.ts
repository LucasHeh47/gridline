import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/core/address';
import { formatValue } from '../src/core/numfmt';
import { currentRegion } from '../src/core/ops/data';
import { analyze, applyPlan, DEFAULT_SMART_OPTIONS, describePlan, targetRange, type SmartOptions } from '../src/core/smartFormat';
import { makeDoc, val } from './helpers';

const opts = (o: Partial<SmartOptions> = {}): SmartOptions => ({ ...DEFAULT_SMART_OPTIONS, ...o });

function budgetDoc() {
  return makeDoc({
    A1: 'Household spending',
    A3: 'Date', B3: 'Category', C3: 'Amount', D3: 'Share', E3: 'Year', F3: 'Notes', G3: 'Account ID',
    A4: '2026-01-05', B4: 'Groceries', C4: "'$132.48", D4: "'12.5%", E4: '2026', F4: 'Weekly shop at the market, including a few things for the party on Saturday', G4: '10045',
    A5: '2026-01-09', B5: 'Rent', C5: "'$1,500.00", D5: "'60%", E5: '2026', F5: 'January rent', G5: '10046',
    A6: '2026-01-12', B6: 'Gas', C6: '48.2', D6: '0.05', E6: '2026', F6: '', G6: '10047',
    A7: 'Total', C7: '=SUM(C4:C6)',
  });
}

describe('advanced format', () => {
  it('detects title, header, totals and column kinds', () => {
    const d = budgetDoc();
    const s = d.wb.activeSheet;
    const range = targetRange(d, s, parseRange('B5')!, (r, c) => currentRegion(d, s, r, c));
    expect(range).toEqual(parseRange('A1:G7'));
    const plan = analyze(d, s, range, opts());
    expect(plan.titleRow).toBe(0);
    expect(plan.headerRow).toBe(2);
    expect(plan.totalRow).toBe(6);
    const kinds = Object.fromEntries(plan.columns.map((c) => [c.col, c.kind]));
    expect(kinds).toEqual({ 0: 'date', 1: 'text', 2: 'currency', 3: 'percent', 4: 'year', 5: 'longText', 6: 'id' });
    expect(describePlan(plan)).toMatch(/Converted 4 numbers stored as text/);
  });

  it('formats, converts text numbers and styles the table', () => {
    const d = budgetDoc();
    const s = d.wb.activeSheet;
    const plan = analyze(d, s, parseRange('A1:G7')!, opts());
    d.transact('Advanced format', (tx) => applyPlan(tx, d, s, plan, opts()));
    expect(val(d, 'C5')).toBe(1500);
    expect(val(d, 'C7')).toBeCloseTo(1680.68, 2);
    const fmt = (a: string) => d.numFmtAt(s, parseRange(a)!.r1, parseRange(a)!.c1);
    expect(formatValue(val(d, 'C5') as number, fmt('C5')).text).toBe('$1,500.00');
    expect(formatValue(val(d, 'D4') as number, fmt('D4')).text).toBe('12.5%');
    expect(formatValue(val(d, 'E4') as number, fmt('E4')).text).toBe('2026');
    expect(formatValue(val(d, 'A4') as number, fmt('A4')).text).toBe('Jan 5, 2026');
    expect(d.wb.styles.get(s.styleIdAt(2, 0)).bold).toBe(true);
    expect(d.wb.styles.get(s.styleIdAt(6, 2)).borderBottom?.style).toBe('double');
    expect(s.freeze.rows).toBe(3);
    expect(s.cf.some((r) => r.type === 'expression')).toBe(true);
    // Undo restores everything
    d.undo();
    expect(val(d, 'C5')).toBe('$1,500.00');
    expect(s.freeze.rows).toBe(0);
    expect(s.cf.length).toBe(0);
  });

  it('handles all-text tables and plain number grids', () => {
    const d = makeDoc({ A1: 'Name', B1: 'City', A2: 'Ann', B2: 'Austin', A3: 'Bo', B3: 'Boston', A4: 'Cy', B4: 'Chicago' });
    const s = d.wb.activeSheet;
    expect(analyze(d, s, parseRange('A1:B4')!, opts()).headerRow).toBe(0);
    const n = makeDoc({ A1: '1.5', B1: '2', A2: '3.25', B2: '4' });
    const plan = analyze(n, n.wb.activeSheet, parseRange('A1:B2')!, opts());
    expect(plan.headerRow).toBeNull();
    expect(plan.columns.map((c) => c.kind)).toEqual(['number', 'integer']);
  });

  it('respects options: negatives, currency symbol and header mode', () => {
    const d = makeDoc({ A1: 'Profit', A2: '-1234.5', A3: '200' });
    const s = d.wb.activeSheet;
    const plan = analyze(d, s, parseRange('A1:A3')!, opts({ negatives: 'parens', currency: '€' }));
    expect(formatValue(-1234.5, plan.columns[0].numFmt).text).toBe('(€1,234.50)');
    expect(analyze(d, s, parseRange('A1:A3')!, opts({ header: 'never' })).headerRow).toBeNull();
  });
});
