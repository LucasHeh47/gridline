import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/core/address';
import { CFEvaluator } from '../src/core/cf';
import { createChart, chartData } from '../src/core/charts';
import { adjustForStructure, cycleAbsolute, translateFormula } from '../src/core/formula/refs';
import { applyStyle, mergeCells } from '../src/core/ops/cells';
import { clipForWorkbook, copyRange, pasteClip, pasteText } from '../src/core/ops/clipboard';
import { applyFilter, removeDuplicates, sortRange } from '../src/core/ops/data';
import { autofill } from '../src/core/ops/fill';
import { addSheet, deleteSheet, renameSheet } from '../src/core/ops/sheets';
import { restructure } from '../src/core/ops/structure';
import { makeDoc, set, val } from './helpers';

const R = (s: string) => parseRange(s)!;

describe('reference rewriting', () => {
  it('translates relative references', () => {
    expect(translateFormula('A1+$B$2+C$3+$D4', 1, 1)).toBe('B2+$B$2+D$3+$D5');
    expect(translateFormula('SUM(A1:A10)', 2, 0)).toBe('SUM(A3:A12)');
    expect(translateFormula('A1', -1, 0)).toBe('#REF!');
    expect(translateFormula('"A1"&A1', 1, 0)).toBe('"A1"&A2');
  });

  it('adjusts for inserted and deleted rows', () => {
    const ins = { sheetName: 'Sheet1', axis: 'row' as const, at: 2, count: 2 };
    expect(adjustForStructure('A1+A3+SUM(A1:A5)', 'Sheet1', ins)).toBe('A1+A5+SUM(A1:A7)');
    const del = { sheetName: 'Sheet1', axis: 'row' as const, at: 2, count: -1 };
    expect(adjustForStructure('A3+A4+SUM(A1:A5)', 'Sheet1', del)).toBe('#REF!+A3+SUM(A1:A4)');
    expect(adjustForStructure('Other!A3', 'Sheet1', del)).toBe('Other!A3');
  });

  it('toggles absolute references with F4', () => {
    expect(cycleAbsolute('=A1+B2', 2)?.text).toBe('=$A$1+B2');
    expect(cycleAbsolute('=$A$1', 3)?.text).toBe('=A$1');
  });
});

describe('structure', () => {
  it('inserts rows and fixes formulas, then undoes', () => {
    const d = makeDoc({ A1: '1', A2: '2', A3: '=SUM(A1:A2)' });
    d.transact('insert', (tx) => restructure(tx, d.wb, d.wb.activeSheet, 'row', 1, 1));
    expect(val(d, 'A4')).toBe(3);
    expect(d.wb.activeSheet.getCell(3, 0)?.f).toBe('SUM(A1:A3)');
    set(d, { A2: '10' });
    expect(val(d, 'A4')).toBe(13);
    d.undo();
    d.undo();
    expect(val(d, 'A3')).toBe(3);
    expect(d.wb.activeSheet.getCell(2, 0)?.f).toBe('SUM(A1:A2)');
    d.redo();
    expect(val(d, 'A4')).toBe(3);
  });

  it('deleting a referenced row produces #REF!', () => {
    const d = makeDoc({ A1: '1', A2: '2', B1: '=A2*2' });
    d.transact('delete', (tx) => restructure(tx, d.wb, d.wb.activeSheet, 'row', 1, -1));
    expect(val(d, 'B1')).toBe('#REF!');
  });

  it('renaming a sheet updates references', () => {
    const d = makeDoc();
    let s2!: ReturnType<typeof addSheet>;
    d.transact('add', (tx) => {
      s2 = addSheet(tx, d.wb, 'Data');
    });
    set(d, { A1: '5' }, s2);
    const s1 = d.wb.sheets[0];
    set(d, { A1: '=Data!A1*2' }, s1);
    d.transact('rename', (tx) => renameSheet(tx, d.wb, s2, 'My Data'));
    expect(s1.getCell(0, 0)?.f).toBe("'My Data'!A1*2");
    expect(val(d, 'A1', s1)).toBe(10);
    d.transact('delete', (tx) => deleteSheet(tx, d.wb, s2));
    expect(val(d, 'A1', s1)).toBe('#REF!');
  });
});

describe('editing operations', () => {
  it('sorts with headers and keeps formulas aligned', () => {
    const d = makeDoc({ A1: 'Name', B1: 'Score', A2: 'b', B2: '2', A3: 'a', B3: '3', A4: 'c', B4: '1', C2: '=B2*10', C3: '=B3*10', C4: '=B4*10' });
    d.transact('sort', (tx) => sortRange(tx, d, d.wb.activeSheet, R('A1:C4'), [{ col: 1, desc: true }], true));
    expect([val(d, 'A2'), val(d, 'A3'), val(d, 'A4')]).toEqual(['a', 'b', 'c']);
    expect([val(d, 'C2'), val(d, 'C3'), val(d, 'C4')]).toEqual([30, 20, 10]);
  });

  it('autofills series', () => {
    const d = makeDoc({ A1: '1', A2: '3', B1: 'Item 1', C1: 'Jan', D1: '=A1*2', E1: '2025-01-31' });
    const s = d.wb.activeSheet;
    d.transact('fill', (tx) => autofill(tx, d, s, R('A1:A2'), R('A1:A5'), 'auto'));
    expect(val(d, 'A5')).toBe(9);
    d.transact('fill', (tx) => autofill(tx, d, s, R('B1:E1'), R('B1:E4'), 'auto'));
    expect(val(d, 'B4')).toBe('Item 4');
    expect(val(d, 'C3')).toBe('Mar');
    expect(s.getCell(3, 3)?.f).toBe('A4*2');
    expect(val(d, 'D4')).toBe(14);
    d.transact('fill', (tx) => autofill(tx, d, s, R('F1'), R('F1:F3'), 'auto'));
  });

  it('copies and pastes with relative references', () => {
    const d = makeDoc({ A1: '2', B1: '=A1*10' });
    const s = d.wb.activeSheet;
    const clip = copyRange(d, s, R('B1'), false);
    d.transact('paste', (tx) => pasteClip(tx, d, clip, s, R('B2:B3'), 'all'));
    set(d, { A2: '3', A3: '4' });
    expect([val(d, 'B2'), val(d, 'B3')]).toEqual([30, 40]);
    expect(clip.text).toBe('20');
  });

  it('cut and paste moves references', () => {
    const d = makeDoc({ A1: '2', C1: '=A1+1' });
    const s = d.wb.activeSheet;
    const clip = copyRange(d, s, R('A1'), true);
    d.transact('paste', (tx) => pasteClip(tx, d, clip, s, R('A5'), 'all'));
    expect(s.getCell(0, 2)?.f).toBe('A5+1');
    expect(val(d, 'C1')).toBe(3);
  });

  it('pastes into another workbook (tab) with the right styles, and never touches the source', () => {
    const a = makeDoc({ A1: '7', B1: '=A1*2' });
    const sa = a.wb.activeSheet;
    a.transact('style', (tx) => applyStyle(tx, a.wb, sa, [R('A1:B1')], { bold: true, fill: '#ffeb9c' }));
    const b = makeDoc();
    // b's style table differs: give it styles of its own first
    b.transact('style', (tx) => applyStyle(tx, b.wb, b.wb.activeSheet, [R('Z1')], { italic: true }));
    const clip = copyRange(a, sa, R('A1:B1'), true);
    const moved = clipForWorkbook(clip, a.wb, b.wb);
    expect(moved.cut).toBe(false);
    const sb = b.wb.activeSheet;
    b.transact('paste', (tx) => pasteClip(tx, b, moved, sb, R('C3'), 'all'));
    const st = b.wb.styles.get(sb.getCell(2, 2)?.s);
    expect([st.bold, st.fill, st.italic]).toEqual([true, '#ffeb9c', undefined]);
    expect(val(b, 'D3')).toBe(14);
    // The source workbook is unchanged (the app empties it itself after a cut)
    expect(val(a, 'A1')).toBe(7);
  });

  it('pastes plain text from other apps', () => {
    const d = makeDoc();
    const s = d.wb.activeSheet;
    d.transact('paste', (tx) => pasteText(tx, d, 'a\tb\r\n1\t"x\ny"\r\n', s, 0, 0));
    expect(val(d, 'A2')).toBe(1);
    expect(val(d, 'B2')).toBe('x\ny');
  });

  it('merges cells and applies styles', () => {
    const d = makeDoc({ A1: 'Title', B1: 'gone' });
    const s = d.wb.activeSheet;
    let lost = false;
    d.transact('merge', (tx) => {
      lost = mergeCells(tx, d, s, R('A1:C1'), 'center');
    });
    expect(lost).toBe(true);
    expect(s.merges).toHaveLength(1);
    expect(val(d, 'B1')).toBe(null);
    d.transact('bold', (tx) => applyStyle(tx, d.wb, s, [R('A:A')], { bold: true }));
    expect(d.wb.styles.get(s.styleIdAt(500, 0)).bold).toBe(true);
  });

  it('filters rows', () => {
    const d = makeDoc({ A1: 'Fruit', A2: 'apple', A3: 'pear', A4: 'apple', B1: 'n', B2: '1', B3: '5', B4: '9' });
    const s = d.wb.activeSheet;
    d.transact('filter', (tx) => applyFilter(tx, d, s, { range: R('A1:B4'), conds: { 0: { values: ['apple'] } } }));
    expect([...s.filterHidden]).toEqual([2]);
    d.transact('filter', (tx) => applyFilter(tx, d, s, { range: R('A1:B4'), conds: { 1: { op: { type: 'gt', value: '4' } } } }));
    expect([...s.filterHidden]).toEqual([1]);
  });

  it('removes duplicates', () => {
    const d = makeDoc({ A1: 'x', A2: 'a', A3: 'b', A4: 'A', A5: 'c' });
    let n = 0;
    d.transact('dedupe', (tx) => {
      n = removeDuplicates(tx, d, d.wb.activeSheet, R('A1:A5'), [0], true);
    });
    expect(n).toBe(1);
    expect([val(d, 'A2'), val(d, 'A3'), val(d, 'A4'), val(d, 'A5')]).toEqual(['a', 'b', 'c', null]);
  });
});

describe('conditional formatting and charts', () => {
  it('evaluates rules', () => {
    const d = makeDoc({ A1: '1', A2: '5', A3: '10', A4: '5' });
    const s = d.wb.activeSheet;
    s.cf = [
      { id: '1', ranges: [R('A1:A4')], type: 'cellIs', operator: 'greaterThan', values: ['4'], style: { fill: '#ff0000' } },
      { id: '2', ranges: [R('A1:A4')], type: 'duplicate', style: { bold: true } },
      { id: '3', ranges: [R('A1:A4')], type: 'colorScale', stops: [{ at: { type: 'min' }, color: '#000000' }, { at: { type: 'max' }, color: '#ffffff' }] },
      { id: '4', ranges: [R('A1:A4')], type: 'expression', formula: 'A1=MAX($A$1:$A$4)', style: { italic: true } },
    ];
    const cf = new CFEvaluator(d);
    expect(cf.evaluate(s, 0, 0)?.style).toBeUndefined();
    expect(cf.evaluate(s, 1, 0)?.style).toEqual({ fill: '#ff0000', bold: true });
    expect(cf.evaluate(s, 2, 0)?.style?.italic).toBe(true);
    expect(cf.evaluate(s, 2, 0)?.fill).toBe('#ffffff');
  });

  it('builds charts from a range', () => {
    const d = makeDoc({ A1: 'Month', B1: 'Sales', C1: 'Cost', A2: 'Jan', B2: '10', C2: '4', A3: 'Feb', B3: '20', C3: '8' });
    const s = d.wb.activeSheet;
    const chart = createChart(d, s, R('A1:C3'), 'column', 0, 4);
    expect(chart.series).toHaveLength(2);
    const data = chartData(d, chart, s);
    expect(data.categories).toEqual(['Jan', 'Feb']);
    expect(data.series[0]).toMatchObject({ name: 'Sales', values: [10, 20] });
    expect(data.series[1].name).toBe('Cost');
  });
});
