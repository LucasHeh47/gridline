import { beforeEach, describe, expect, it } from 'vitest';
import { formatCell, keyCol, keyRow } from '../src/core/address';
import { external, FetchError } from '../src/core/data/external';
import { Doc } from '../src/core/doc';
import { buildTemplate, TEMPLATES } from '../src/core/templates';
import { ErrorValue } from '../src/core/values';
import { val } from './helpers';

beforeEach(() => {
  external.reset();
  // No network in tests: live functions report missing data
  external.fetcher = async () => {
    throw new FetchError('Not Found', 404);
  };
});

const settle = () => new Promise((r) => setTimeout(r, 120));

/** Every formula cell that evaluates to an error, as "Sheet!A1 #CODE message". */
function errors(doc: Doc, allowLive: boolean): string[] {
  const out: string[] = [];
  for (const s of doc.wb.sheets) {
    for (const [k, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const v = doc.engine.getValue(s, keyRow(k), keyCol(k));
      if (!(v instanceof ErrorValue)) continue;
      if (allowLive && (v.code === '#N/A' || v.code === '#BUSY!' || v.code === '#CONNECT!')) continue;
      out.push(`${s.name}!${formatCell(keyRow(k), keyCol(k))} ${v.code} ${v.message ?? ''} [${cell.f}]`);
    }
  }
  return out;
}

describe('templates', () => {
  it('builds every template without formula errors', async () => {
    for (const t of TEMPLATES) {
      const wb = buildTemplate(t.id);
      expect(wb, t.id).not.toBeNull();
      const doc = new Doc(wb!);
      await settle();
      const live = !!t.live;
      expect(errors(doc, live), t.id).toEqual([]);
      doc.dispose();
    }
  });

  it('budget totals transactions for the month', () => {
    const doc = new Doc(buildTemplate('budget')!);
    const s = doc.wb.sheets[0];
    expect(val(doc, 'C5', s)).toBe(4200); // paychecks logged on Transactions
    expect(val(doc, 'B8', s)).toBe(4600); // planned income
    const groceriesRow = [...Array(40).keys()].map((i) => i + 1).find((r) => val(doc, `A${r}`, s) === 'Groceries')!;
    expect(val(doc, `C${groceriesRow}`, s)).toBeCloseTo(219.58, 2);
    doc.dispose();
  });

  it('loan planner computes the payment and schedule', () => {
    const doc = new Doc(buildTemplate('loan')!);
    expect(val(doc, 'E4') as number).toBeCloseTo(2212.24, 2);
    expect(val(doc, 'E5')).toBe(360);
    expect(val(doc, 'A12')).toBe('Payment #');
    expect(val(doc, 'G372')).toBe(0);
    doc.dispose();
  });

  it('debt planner compares methods', () => {
    const doc = new Doc(buildTemplate('debt')!);
    const avalanche = val(doc, 'H6') as number;
    const snowball = val(doc, 'H7') as number;
    const minimums = val(doc, 'H8') as number;
    expect(avalanche).toBeGreaterThan(0);
    expect(avalanche).toBeLessThan(snowball);
    expect(minimums).toBeGreaterThan(snowball);
    expect(val(doc, 'A17')).toBe('Debt');
    expect(val(doc, 'A23')).toBe('Total');
    doc.dispose();
  });

  it('net worth and savings', () => {
    const doc = new Doc(buildTemplate('networth')!);
    const [n, s] = doc.wb.sheets;
    expect(val(doc, 'B15', n)).toBe(102000 - 32600);
    expect(val(doc, 'B8', s) as number).toBeCloseTo(590.43, 2);
    expect(val(doc, 'E8', s) as number).toBeGreaterThan(600000);
    doc.dispose();
  });
});
