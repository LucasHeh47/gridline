import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/core/address';
import { createChart } from '../src/core/charts';
import { dateToSerial } from '../src/core/dates';
import { Doc } from '../src/core/doc';
import { csvToWorkbook, decodeText, detectDelimiter, encodeCSV, parseCSV, sheetToCSV } from '../src/core/io/csv';
import { readXlsx, toFileFormula, writeXlsx } from '../src/core/io/xlsx';
import { applyStyle, mergeCells } from '../src/core/ops/cells';
import { makeDoc, val } from './helpers';

describe('csv', () => {
  it('parses quoted fields and detects delimiters', () => {
    expect(parseCSV('a,"b,c",d\n1,"say ""hi""",3')).toEqual([
      ['a', 'b,c', 'd'],
      ['1', 'say "hi"', '3'],
    ]);
    expect(detectDelimiter('a;b;c\n1;2;3\n')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2\n')).toBe('\t');
  });

  it('imports values with types', () => {
    const wb = csvToWorkbook('id,name,amount,date\n007,Ann,"$1,200.50",2025-03-14\n', 'people.csv');
    const doc = new Doc(wb);
    expect(wb.sheets[0].name).toBe('people');
    expect(val(doc, 'A2')).toBe('007');
    expect(val(doc, 'C2')).toBe(1200.5);
    expect(val(doc, 'D2')).toBe(dateToSerial(2025, 3, 14));
  });

  it('exports full precision numbers and formatted dates', () => {
    const doc = makeDoc({ A1: 'x', B1: '=1/3', C1: '2025-01-02', D1: 'a,b' });
    const csv = sheetToCSV(doc, doc.wb.activeSheet);
    expect(csv).toBe('x,0.333333333333333,2025-01-02,"a,b"\r\n');
    expect(decodeText(encodeCSV('é'))).toBe('é');
  });
});

describe('xlsx', () => {
  it('round-trips values, formulas, styles, merges and charts', async () => {
    const doc = makeDoc({
      A1: 'Month',
      B1: 'Sales',
      A2: 'Jan',
      B2: '100',
      A3: 'Feb',
      B3: '250',
      B4: '=SUM(B2:B3)',
      C1: '=XLOOKUP("Feb",A2:A3,B2:B3)',
      D1: '=SEQUENCE(3)',
      E1: '2025-03-14',
      F1: '$5.25',
    });
    const s = doc.wb.activeSheet;
    doc.transact('style', (tx) => {
      applyStyle(tx, doc.wb, s, [parseRange('A1:B1')!], { bold: true, fill: '#336699', color: '#ffffff' });
      mergeCells(tx, doc, s, parseRange('G1:H2')!, 'merge');
      tx.setProp(s, 'freeze', { rows: 1, cols: 0 });
      tx.setProp(s, 'charts', [createChart(doc, s, parseRange('A1:B3')!, 'column', 5, 1)]);
      tx.setMap(s.colWidths, 0, 150);
    });
    const bytes = await writeXlsx(doc);
    expect(bytes.length).toBeGreaterThan(1000);

    const { wb } = await readXlsx(bytes);
    const back = new Doc(wb);
    const s2 = wb.sheets[0];
    expect(val(back, 'B4')).toBe(350);
    expect(s2.getCell(3, 1)?.f).toBe('SUM(B2:B3)');
    expect(s2.getCell(0, 2)?.f).toBe('XLOOKUP("Feb",A2:A3,B2:B3)');
    expect(val(back, 'C1')).toBe(250);
    expect(val(back, 'D3')).toBe(3);
    expect(val(back, 'E1')).toBe(dateToSerial(2025, 3, 14));
    expect(wb.styles.get(s2.styleIdAt(0, 0))).toMatchObject({ bold: true, fill: '#336699', color: '#ffffff' });
    expect(wb.styles.get(s2.styleIdAt(0, 5)).numFmt).toBe('"$"#,##0.00');
    expect(s2.merges).toEqual([{ r1: 0, c1: 6, r2: 1, c2: 7 }]);
    expect(s2.freeze).toEqual({ rows: 1, cols: 0 });
    expect(s2.colWidth(0)).toBe(150);
    expect(s2.charts).toHaveLength(1);
    expect(s2.charts[0].type).toBe('column');
    expect(s2.charts[0].series[0].values).toBe('Sheet1!$B$2:$B$3');
  }, 60000);

  it('prefixes new functions for Excel', () => {
    expect(toFileFormula('XLOOKUP(A1,B:B,C:C)+SUM(1)')).toBe('_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(1)');
    expect(toFileFormula('FILTER(A1:A3,B1:B3>1)')).toBe('_xlfn._xlws.FILTER(A1:A3,B1:B3>1)');
  });
});
