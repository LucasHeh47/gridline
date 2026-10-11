// Results measured in Microsoft Excel (365, Windows). Gridline must match them exactly.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { Doc } from '../src/core/doc';
import { parseExcelNumber, xMul } from '../src/core/formula/excelMath';
import { readXlsx, toFileFormula, writeXlsx } from '../src/core/io/xlsx';
import { formatValue } from '../src/core/numfmt';
import { calc, makeDoc, val } from './helpers';

describe('arithmetic matches Excel bit for bit', () => {
  // Value2 of each formula in Excel, written with .NET "R" (round-trip) formatting
  const cases: [string, number][] = [
    ['19*107.85*(1-0.05)', 1946.6925],
    ['19*107.85*0.95', 1946.6925],
    ['2049.15*(1-0.05)', 1946.6925],
    ['19*(107.85*(1-0.05))', 1946.6924999999999],
    ['0.1+0.2', 0.30000000000000004],
    ['0.5-0.4-0.1', 0],
    ['1*(0.5-0.4-0.1)', -2.7755575615628914e-17],
    ['(0.5-0.4-0.1)*1', -2.7755575615628914e-17],
    ['0.1+0.2-0.3', 0],
    ['1-0.9', 0.09999999999999998],
    ['1.1*3', 3.3000000000000003],
    ['3.3-1.1*3', 0],
    ['SUM(0.1,0.2,-0.3)', 0],
    ['SUM(1E+16,1,-1E+16)', 0],
    ['SUM(0.1,0.2)', 0.30000000000000004],
    ['1+2^-50-1', 0],
    ['1+2^-48-1', 3.552713678800501e-15],
    ['100.1-100', 0.09999999999999432],
    ['4.35*100', 434.99999999999994],
    ['0.07*100', 7.000000000000001],
    ['1/3*3', 1],
    ['10/3', 3.3333333333333335],
    ['SUM(1,2^-53,2^-53)', 1],
    ['SUM(2^-53,2^-53,1)', 1.0000000000000002],
    ['=0.66666666666666666', 0.666666666666666],
    ['=2.9999999999999999', 2.99999999999999],
  ];
  for (const [f, want] of cases) {
    it(f, () => {
      expect(calc(makeDoc(), f)).toBe(want);
    });
  }

  it('applies the near-zero rule only to a final + or - outside parentheses', () => {
    const doc = makeDoc({ A1: '1.0000000000000009', B1: '1' });
    // 1+4u in A1 (typed values are cut to 15 digits, so set it directly)
    doc.transact('t', (tx) => tx.setCell(doc.wb.activeSheet, 0, 0, { v: 1 + 4 * 2 ** -52 }));
    expect(calc(doc, '=A1-B1')).toBe(0);
    expect(calc(doc, '=(A1-B1)')).toBe(8.881784197001252e-16);
    expect(calc(doc, '=IF(TRUE,A1-B1)')).toBe(8.881784197001252e-16);
  });

  it('rounds products twice like the x87 FPU', () => {
    expect(xMul(2049.15, 0.95)).toBe(1946.6925);
    expect(2049.15 * 0.95).toBe(1946.6924999999999);
  });

  it('keeps only 15 significant digits of typed numbers (truncating)', () => {
    expect(parseExcelNumber('1234567890123456789')).toBe(1.23456789012345e18);
    expect(parseExcelNumber('12345.678901234567')).toBe(12345.6789012345);
    expect(parseExcelNumber('0.1')).toBe(0.1);
    const doc = makeDoc({ A1: '0.66666666666666666' });
    expect(val(doc, 'A1')).toBe(0.666666666666666);
  });
});

describe('displayed values match Excel', () => {
  it('rounds half-cents up like Excel (15 significant digits first)', () => {
    expect(formatValue(1087.005, '$#,##0.00').text).toBe('$1,087.01');
    expect(formatValue(633418.315, '$#,##0.00').text).toBe('$633,418.32');
    expect(formatValue(1.005, '0.00').text).toBe('1.01');
    expect(formatValue(0.145, '0.0%').text).toBe('14.5%');
    expect(formatValue(2.5, '0').text).toBe('3');
    expect(formatValue(-2.5, '0').text).toBe('-3');
    expect(formatValue(0.1 + 0.2, 'General').text).toBe('0.3');
    expect(formatValue(1087.005, 'General').text).toBe('1087.005');
  });
});

describe('formulas in files', () => {
  it('writes LET/LAMBDA names, spills and @ the way Excel requires', () => {
    expect(toFileFormula('LET(x,10,y,20,x+y)')).toBe('_xlfn.LET(_xlpm.x,10,_xlpm.y,20,_xlpm.x+_xlpm.y)');
    expect(toFileFormula('LET(f,LAMBDA(n,n*2),f(21))')).toBe('_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*2),_xlpm.f(21))');
    expect(toFileFormula('SUM(A1#)')).toBe('SUM(_xlfn.ANCHORARRAY(A1))');
    expect(toFileFormula('SUM(Data!B2#)*2')).toBe('SUM(_xlfn.ANCHORARRAY(Data!B2))*2');
    expect(toFileFormula('@A1:A5')).toBe('_xlfn.SINGLE(A1:A5)');
  });

  it('reads them back', async () => {
    const doc = makeDoc({ A1: '=SEQUENCE(3)', B1: '=SUM(A1#)', C1: '=LET(x,10,y,20,x+y)' });
    const { wb } = await readXlsx(await writeXlsx(doc));
    const s = wb.sheets[0];
    expect(s.getCell(0, 1)?.f).toBe('SUM(A1#)');
    expect(s.getCell(0, 2)?.f).toBe('LET(x,10,y,20,x+y)');
    const d = new Doc(wb);
    expect(val(d, 'B1')).toBe(6);
    expect(val(d, 'C1')).toBe(30);
  });

  it('treats built-in functions marked as add-ins (_xludf.) as the real thing', async () => {
    const doc = makeDoc({ A1: '1' });
    const zip = await JSZip.loadAsync(await writeXlsx(doc));
    let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    xml = xml.replace(/<c r="A1"[^>]*>.*?<\/c>/, '<c r="A1" t="e"><f>_xludf.CONCAT("Grid","line")</f><v>#NAME?</v></c>');
    zip.file('xl/worksheets/sheet1.xml', xml);
    const { wb } = await readXlsx(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })));
    expect(val(new Doc(wb), 'A1')).toBe('Gridline');
  });

  it('puts a chart drawing after a data bar rule, not inside it', async () => {
    const doc = makeDoc({ A1: '1', A2: '2', A3: '3' });
    const sheet = doc.wb.activeSheet;
    sheet.cf = [{ id: 'd', type: 'dataBar', ranges: [{ r1: 0, c1: 0, r2: 2, c2: 0 }], color: '#5b9bd5', min: { type: 'min' }, max: { type: 'max' }, gradient: true, showValue: true }];
    sheet.charts = [{ id: 'c', type: 'column', series: [{ values: 'Sheet1!$A$1:$A$3' }], legend: 'none', anchor: { row: 0, col: 2, dx: 0, dy: 0, w: 300, h: 200 } }];
    const xml = await (await JSZip.loadAsync(await writeXlsx(doc))).file('xl/worksheets/sheet1.xml')!.async('string');
    const cfEnd = xml.lastIndexOf('</conditionalFormatting>');
    expect(xml.indexOf('<drawing ')).toBeGreaterThan(cfEnd);
  });
});
