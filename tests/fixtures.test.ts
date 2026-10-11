import { readFileSync } from 'fs';
import { describe, expect, it } from 'vitest';
import { Doc } from '../src/core/doc';
import { readXlsx, writeXlsx } from '../src/core/io/xlsx';
import { val } from './helpers';

// A workbook produced by another tool (openpyxl): default XML namespaces, absolute part paths,
// comments in non-standard locations, charts in oneCellAnchors, no cached formula values.
const fixture = () => new Uint8Array(readFileSync(new URL('./fixtures/openpyxl-sample.xlsx', import.meta.url)));

describe('files from other apps', () => {
  it('opens the openpyxl sample', async () => {
    const { wb, warnings } = await readXlsx(fixture());
    const doc = new Doc(wb);
    const [sales, summary] = wb.sheets;
    expect(sales.name).toBe('Sales Data');
    expect(warnings.some((w) => /validation/i.test(w))).toBe(false);
    expect(sales.validations.map((v) => [v.type, v.formula1, v.dropdown])).toEqual([['list', '"North,South,East,West"', true]]);
    expect(val(doc, 'F2', sales)).toBe(45);
    expect(val(doc, 'G2', sales)).toBe(3.71); // uses the defined name Rate
    expect(val(doc, 'J3', sales)).toBe(9906.25);
    expect(val(doc, 'J5', sales)).toBe(720);
    expect(val(doc, 'B4', summary)).toBe(6000); // cross-sheet whole-column SUMIFS
    expect(sales.getCell(1, 2)?.note).toBe('Top performer');
    expect(sales.getCell(7, 8)?.link).toBe('https://example.com/report');
    expect(sales.freeze).toEqual({ rows: 1, cols: 0 });
    expect(sales.filter?.range).toEqual({ r1: 0, c1: 0, r2: 20, c2: 6 });
    expect(sales.cf.map((r) => r.type)).toEqual(['cellIs', 'colorScale', 'dataBar', 'expression']);
    expect(sales.charts.map((c) => c.type)).toEqual(['column']);
    expect(summary.charts.map((c) => c.type)).toEqual(['pie', 'line']);
    expect(summary.merges).toEqual([{ r1: 6, c1: 0, r2: 7, c2: 2 }]);
    expect(wb.styles.get(sales.styleIdAt(0, 0))).toMatchObject({ bold: true, fill: '#305496', color: '#ffffff' });
  }, 60000);

  it('saves it back without losing content', async () => {
    const { wb } = await readXlsx(fixture());
    const bytes = await writeXlsx(new Doc(wb));
    const again = await readXlsx(bytes);
    const doc = new Doc(again.wb);
    const [sales, summary] = again.wb.sheets;
    expect(val(doc, 'G2', sales)).toBe(3.71);
    expect(sales.getCell(1, 2)?.note).toBe('Top performer');
    expect(sales.charts).toHaveLength(1);
    expect(summary.charts).toHaveLength(2);
    expect(sales.cf).toHaveLength(4);
    expect(sales.validations).toHaveLength(1);
  }, 60000);
});
