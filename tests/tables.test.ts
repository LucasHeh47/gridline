import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/core/address';
import { Doc } from '../src/core/doc';
import { readXlsx, writeXlsx } from '../src/core/io/xlsx';
import { restructure } from '../src/core/ops/structure';
import { duplicateSheet } from '../src/core/ops/sheets';
import {
  canRestoreOrder,
  columnKind,
  createTable,
  MergedCellsError,
  nextSort,
  removeTable,
  renameTable,
  setTableFilter,
  setTableLook,
  smartCompare,
  sortIsStale,
  sortTable,
  tableNameProblem,
} from '../src/core/tables';
import type { TableDef } from '../src/core/workbook';
import { makeDoc, set, val } from './helpers';

const R = (a: string) => parseRange(a)!;

function make(doc: Doc, range: string, opts: Parameters<typeof createTable>[4] = {}): TableDef {
  let t: TableDef | undefined;
  doc.transact('table', (tx) => (t = createTable(tx, doc, doc.wb.activeSheet, R(range), opts).table));
  return t!;
}

const table = (doc: Doc, i = 0) => doc.wb.activeSheet.tables[i];
const sort = (doc: Doc, keys: Parameters<typeof sortTable>[4]) => doc.transact('sort', (tx) => sortTable(tx, doc, doc.wb.activeSheet, table(doc).id, keys));
const column = (doc: Doc, col: string, from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => val(doc, `${col}${from + i}`));

function sample(): Doc {
  return makeDoc({
    A1: 'Item',
    B1: 'Amount',
    C1: 'Date',
    A2: 'Item 10',
    B2: '1,200',
    C2: '3/1/2026',
    A3: 'item 2',
    B3: '45',
    C3: '1/15/2026',
    A4: 'Item 1',
    B4: '',
    C4: '2/2/2026',
    A5: 'Apple',
    B5: '300',
    C5: '',
  });
}

describe('making tables', () => {
  it('names columns from the first row and keeps the names unique, without changing the cells', () => {
    const doc = makeDoc({ A1: 'Name', B1: 'Name', C1: '', A2: 'Ann', B2: 'Bo', C2: 'x' });
    // A repeat gets a number and an empty header becomes Column3, in the table only
    const t = make(doc, 'A1:C2');
    expect(t.range).toEqual(R('A1:C2'));
    expect(t.columns).toEqual(['Name', 'Name2', 'Column3']);
    expect([val(doc, 'A1'), val(doc, 'B1'), val(doc, 'C1'), val(doc, 'A2')]).toEqual(['Name', 'Name', null, 'Ann']);

    const doc2 = makeDoc({ A1: 'Name', B1: 'City', A2: 'Ann', B2: 'Oslo', A3: 'Bo', B3: 'Rome' });
    const t2 = make(doc2, 'A1:B3');
    expect(t2.range).toEqual(R('A1:B3'));
    expect(t2.columns).toEqual(['Name', 'City']);
    expect(t2.name).toBe('Table1');

    const doc3 = makeDoc({ A1: 'Amount', B1: '', C1: 'amount', A2: '1' });
    const t3 = make(doc3, 'A1:C2', { header: true });
    expect(t3.columns).toEqual(['Amount', 'Column2', 'amount2']);
    expect([val(doc3, 'B1'), val(doc3, 'C1')]).toEqual([null, 'amount']);
  });

  it('keeps every cell as it was: values, formulas, formats, column widths and rows', () => {
    const doc = makeDoc({ B2: '2024', C2: '=B2+1', D2: 'Total', B3: '5', C3: '6', D3: '=B3+C3', B4: '7', C4: '8', D4: '=B4+C4', F1: 'beside' });
    const sh = doc.wb.activeSheet;
    doc.transact('look', (tx) => tx.patchCell(sh, 1, 1, { s: doc.wb.styles.intern({ bold: true, fill: '#ffcc00', numFmt: '0' }) }));
    sh.colWidths.set(1, 140);
    const before = [...sh.cells.entries()].map(([k, c]) => JSON.stringify([k, c.v, c.f, c.s]));
    const widths = [0, 1, 2, 3, 4].map((c) => sh.colWidth(c));
    // Whole columns A:E, with an empty row above the data and an empty column A
    let made: ReturnType<typeof createTable> | undefined;
    doc.transact('table', (tx) => (made = createTable(tx, doc, sh, R('A1:E1048576'))));
    expect(made!.table.range).toEqual(R('B2:D4'));
    expect(made!.table.columns).toEqual(['2024', '2025', 'Total']);
    expect(made!.addedHeader).toBe(false);
    expect(made!.firstRowLooksLikeData).toBe(true);
    expect([...sh.cells.entries()].map(([k, c]) => JSON.stringify([k, c.v, c.f, c.s]))).toEqual(before);
    expect([0, 1, 2, 3, 4].map((c) => sh.colWidth(c))).toEqual(widths);
    expect(made!.table.color).toBe('none');
    expect(made!.table.banded).toBe(false);
  });

  it('leaves a title above the header row out of the table', () => {
    const doc = makeDoc({ A1: 'Scores', A2: 'Name', B2: 'Team', C2: 'Points', A3: 'Ann', B3: 'Red', C3: '12', A4: 'Bo', B4: 'Blue', C4: '7' });
    const t = make(doc, 'A1:C4');
    expect(t.range).toEqual(R('A2:C4'));
    expect(t.columns).toEqual(['Name', 'Team', 'Points']);
    expect(val(doc, 'A1')).toBe('Scores');
  });

  it('adds a header row when the data has none, and an empty row under a lone header', () => {
    const doc = makeDoc({ A1: '5', B1: '6', A2: '7', B2: '8', D1: 'beside' });
    const t = make(doc, 'A1:B2', { header: false });
    expect(t.range).toEqual(R('A1:B3'));
    expect([val(doc, 'A1'), val(doc, 'B1'), val(doc, 'A2')]).toEqual(['Column1', 'Column2', 5]);
    const doc2 = makeDoc({ A1: 'Just', B1: 'Headers' });
    expect(make(doc2, 'A1:B1').range).toEqual(R('A1:B2'));
  });

  it('refuses overlaps and bad names, and undoes in one step', () => {
    const doc = sample();
    make(doc, 'A1:C5');
    expect(() => make(doc, 'C1:D3')).toThrow(/overlaps the table Table1/);
    expect(tableNameProblem('My Sales', doc.wb)).toContain('spaces');
    expect(tableNameProblem('B12', doc.wb)).toContain('cell address');
    expect(tableNameProblem('table1', doc.wb)).toContain('already a table');
    expect(tableNameProblem('Sales_2026', doc.wb)).toBeNull();
    doc.undo();
    expect(doc.wb.activeSheet.tables).toEqual([]);
  });

  it('turns a sheet filter into the table’s own, and splits merged cells only when asked', () => {
    const doc = sample();
    doc.transact('merge', (tx) => tx.setProp(doc.wb.activeSheet, 'merges', [R('A4:B4')]));
    doc.transact('filter', (tx) => tx.setProp(doc.wb.activeSheet, 'filter', { range: R('A1:C5'), conds: { 0: { values: ['Apple'] } } }));
    expect(() => make(doc, 'A1:C5')).toThrow(MergedCellsError);
    expect(doc.wb.activeSheet.merges).toHaveLength(1);
    expect(doc.wb.activeSheet.tables).toEqual([]);
    make(doc, 'A1:C5', { unmerge: true });
    expect(doc.wb.activeSheet.merges).toHaveLength(0);
    expect(doc.wb.activeSheet.filter).toBeNull();
    expect(table(doc).filter).toEqual({ 0: { values: ['Apple'] } });
    expect([...doc.wb.activeSheet.filterHidden].sort()).toEqual([1, 2, 3]);
  });
});

describe('sorting tables', () => {
  it('sorts the way people expect: numbers in text, natural order, any capitals, blanks last', () => {
    expect(smartCompare('1,200', 45)).toBeGreaterThan(0);
    expect(smartCompare('$5', '12%')).toBeGreaterThan(0);
    expect(smartCompare('Item 2', 'Item 10')).toBeLessThan(0);
    expect(smartCompare('apple', 'Banana')).toBeLessThan(0);
    const doc = sample();
    make(doc, 'A1:C5');
    sort(doc, [{ col: 1, desc: false }]);
    expect(column(doc, 'B', 2, 5)).toEqual([45, 300, 1200, null]);
    sort(doc, [{ col: 1, desc: true }]);
    // Empty cells stay last both ways
    expect(column(doc, 'B', 2, 5)).toEqual([1200, 300, 45, null]);
    sort(doc, [{ col: 0, desc: false }]);
    expect(column(doc, 'A', 2, 5)).toEqual(['Apple', 'Item 1', 'item 2', 'Item 10']);
    expect(val(doc, 'A1')).toBe('Item');
  });

  it('cycles with one click: smallest first, largest first, original order', () => {
    const doc = sample();
    const t = make(doc, 'A1:C5');
    let keys = nextSort(t, 1);
    expect(keys).toEqual([{ col: 1, desc: false }]);
    sort(doc, keys);
    keys = nextSort(table(doc), 1);
    expect(keys).toEqual([{ col: 1, desc: true }]);
    sort(doc, keys);
    expect(canRestoreOrder(table(doc))).toBe(true);
    keys = nextSort(table(doc), 1);
    expect(keys).toBe('original');
    sort(doc, keys);
    expect(column(doc, 'A', 2, 5)).toEqual(['Item 10', 'item 2', 'Item 1', 'Apple']);
    expect(table(doc).sort).toEqual([]);
    expect(canRestoreOrder(table(doc))).toBe(false);
  });

  it('Shift+click sorts by more columns', () => {
    const doc = makeDoc({ A1: 'Team', B1: 'Score', A2: 'B', B2: '3', A3: 'A', B3: '9', A4: 'B', B4: '7', A5: 'A', B5: '1' });
    make(doc, 'A1:B5');
    let keys = nextSort(table(doc), 0);
    sort(doc, keys);
    keys = nextSort(table(doc), 1, { additive: true });
    expect(keys).toEqual([
      { col: 0, desc: false },
      { col: 1, desc: false },
    ]);
    sort(doc, keys);
    keys = nextSort(table(doc), 1, { additive: true });
    sort(doc, keys);
    expect(column(doc, 'B', 2, 5)).toEqual([9, 1, 7, 3]);
    // Shift+click a descending column again: it stops being a sort column
    expect(nextSort(table(doc), 1, { additive: true })).toEqual([{ col: 0, desc: false }]);
  });

  it('moves formulas with their rows and notices rows that changed since', () => {
    const doc = makeDoc({ A1: 'N', B1: 'Twice', A2: '3', B2: '=A2*2', A3: '1', B3: '=A3*2', A4: '2', B4: '=A4*2' });
    make(doc, 'A1:B4');
    sort(doc, [{ col: 0, desc: false }]);
    expect(column(doc, 'B', 2, 4)).toEqual([2, 4, 6]);
    expect(sortIsStale(doc, doc.wb.activeSheet, table(doc))).toBe(false);
    set(doc, { A2: '9' });
    expect(sortIsStale(doc, doc.wb.activeSheet, table(doc))).toBe(true);
    // Clicking the arrow again sorts the same way
    expect(nextSort(table(doc), 0, { stale: true })).toEqual([{ col: 0, desc: false }]);
  });

  it('keeps the filter while sorting and knows what a column holds', () => {
    const doc = sample();
    const t = make(doc, 'A1:C5');
    doc.transact('filter', (tx) => setTableFilter(tx, doc, doc.wb.activeSheet, t, { 0: { values: ['Apple', 'Item 1'] } }));
    sort(doc, [{ col: 0, desc: true }]);
    const hidden = [...doc.wb.activeSheet.filterHidden].sort();
    expect(hidden.map((r) => val(doc, `A${r + 1}`))).toEqual(['Item 10', 'item 2']);
    expect(columnKind(doc, doc.wb.activeSheet, t, 1)).toBe('number');
    expect(columnKind(doc, doc.wb.activeSheet, t, 0)).toBe('text');
    expect(columnKind(doc, doc.wb.activeSheet, t, 2)).toBe('date');
  });
});

describe('tables grow as you type', () => {
  it('takes rows typed just below, fills formula columns, and grows whole-column formulas', () => {
    const doc = makeDoc({ A1: 'Price', B1: 'Qty', C1: 'Total', A2: '2', B2: '3', C2: '=A2*B2', A3: '5', B3: '1', C3: '=A3*B3', E1: '=SUM(C2:C3)' });
    make(doc, 'A1:C3');
    set(doc, { A4: '10', B4: '2' });
    expect(table(doc).range).toEqual(R('A1:C4'));
    expect(val(doc, 'C4')).toBe(20);
    expect(doc.wb.activeSheet.getCell(0, 4)?.f).toBe('SUM(C2:C4)');
    expect(val(doc, 'E1')).toBe(31);
    doc.undo();
    expect(table(doc).range).toEqual(R('A1:C3'));
    expect(val(doc, 'C4')).toBe(null);
  });

  it('leaves a total typed under the table outside it', () => {
    const doc = makeDoc({ A1: 'Price', A2: '2', A3: '5' });
    make(doc, 'A1:A3');
    set(doc, { A4: '=SUM(A2:A3)' });
    expect(table(doc).range).toEqual(R('A1:A3'));
    sort(doc, [{ col: 0, desc: true }]);
    expect(column(doc, 'A', 2, 4)).toEqual([5, 2, 7]);
  });

  it('takes columns typed just right, naming them', () => {
    const doc = makeDoc({ A1: 'Name', A2: 'Ann', A3: 'Bo' });
    make(doc, 'A1:A3');
    set(doc, { B2: 'x' });
    expect(table(doc).range).toEqual(R('A1:B3'));
    expect(table(doc).columns).toEqual(['Name', 'Column2']);
    expect(val(doc, 'B1')).toBeNull();
    set(doc, { C1: 'Notes' });
    expect(table(doc).columns).toEqual(['Name', 'Column2', 'Notes']);
  });

  it('keeps the original order of new rows', () => {
    const doc = makeDoc({ A1: 'N', A2: '3', A3: '1' });
    make(doc, 'A1:A3');
    sort(doc, [{ col: 0, desc: false }]);
    set(doc, { A4: '2' });
    sort(doc, 'original');
    expect(column(doc, 'A', 2, 4)).toEqual([3, 1, 2]);
  });
});

describe('structured references', () => {
  function sales(): Doc {
    const doc = makeDoc({ A1: 'Item', B1: 'Unit Price', C1: 'Qty', A2: 'Pen', B2: '2', C2: '10', A3: 'Ink', B3: '5', C3: '3' });
    const t = make(doc, 'A1:C3');
    doc.transact('name', (tx) => renameTable(tx, doc.wb, doc.wb.activeSheet, t, 'Sales'));
    return doc;
  }

  it('calculates table and column references', () => {
    const doc = sales();
    set(doc, {
      F1: '=SUM(Sales[Qty])',
      F2: '=ROWS(Sales)',
      F3: '=COUNTA(Sales[#All])',
      F4: '=Sales[[#Headers],[Qty]]',
      F5: '=SUM(Sales[[Unit Price]:[Qty]])',
      D2: '=[@[Unit Price]]*[@Qty]',
      D3: '=[@[Unit Price]]*[@Qty]',
    });
    // D became part of the table (typed just right of it), so #All has 3 rows of 4; its header cell
    // stays empty (the name Column4 lives in the table), so COUNTA finds 11
    expect([val(doc, 'F1'), val(doc, 'F2'), val(doc, 'F3'), val(doc, 'F4'), val(doc, 'F5')]).toEqual([13, 2, 11, 'Qty', 20]);
    expect(table(doc).columns).toEqual(['Item', 'Unit Price', 'Qty', 'Column4']);
    expect([val(doc, 'D2'), val(doc, 'D3')]).toEqual([20, 15]);
    // A new row: the totals follow, and the calculated column fills in
    set(doc, { A4: 'Pad', B4: '1', C4: '4' });
    expect([val(doc, 'F1'), val(doc, 'D4')]).toEqual([17, 4]);
    set(doc, { G1: '=[@Qty]' });
    expect(val(doc, 'G1')).toBe('#REF!');
  });

  it('follows renamed tables and columns, and becomes A1 references when the table goes', () => {
    const doc = sales();
    set(doc, { F1: '=SUM(Sales[Qty])', D2: '=[@Qty]*2' });
    set(doc, { C1: 'Count' });
    expect(doc.wb.activeSheet.getCell(0, 5)?.f).toBe('SUM(Sales[Count])');
    expect(doc.wb.activeSheet.getCell(1, 3)?.f).toBe('[@Count]*2');
    doc.transact('name', (tx) => renameTable(tx, doc.wb, doc.wb.activeSheet, table(doc), 'Orders'));
    expect(doc.wb.activeSheet.getCell(0, 5)?.f).toBe('SUM(Orders[Count])');
    expect(val(doc, 'F1')).toBe(13);
    doc.transact('remove', (tx) => removeTable(tx, doc, doc.wb.activeSheet, table(doc)));
    expect(doc.wb.activeSheet.getCell(0, 5)?.f).toBe('SUM(Sheet1!$C$2:$C$3)');
    expect(doc.wb.activeSheet.getCell(1, 3)?.f).toBe('Sheet1!$C$2*2');
    expect(val(doc, 'F1')).toBe(13);
  });
});

describe('rows and columns inserted or deleted', () => {
  it('moves, grows and ends tables', () => {
    const doc = sample();
    make(doc, 'A1:C5');
    sort(doc, [{ col: 1, desc: false }]);
    const sheet = doc.wb.activeSheet;
    doc.transact('rows', (tx) => restructure(tx, doc.wb, sheet, 'row', 0, 2));
    expect(table(doc).range).toEqual(R('A3:C7'));
    doc.transact('rows', (tx) => restructure(tx, doc.wb, sheet, 'row', 4, 1));
    expect(table(doc).range).toEqual(R('A3:C8'));
    expect(table(doc).order).toHaveLength(5);
    doc.transact('cols', (tx) => restructure(tx, doc.wb, sheet, 'col', 1, 1));
    expect(table(doc).range).toEqual(R('A3:D8'));
    expect(table(doc).columns).toEqual(['Item', 'Column2', 'Amount', 'Date']);
    expect(val(doc, 'B3')).toBe('Column2');
    expect(table(doc).sort).toEqual([{ col: 2, desc: false }]);
    doc.transact('cols', (tx) => restructure(tx, doc.wb, sheet, 'col', 2, -1));
    expect(table(doc).sort).toEqual([]);
    expect(table(doc).columns).toEqual(['Item', 'Column2', 'Date']);
    doc.transact('rows', (tx) => restructure(tx, doc.wb, sheet, 'row', 2, -1));
    expect(sheet.tables).toEqual([]);
  });

  it('copies tables with a duplicated sheet', () => {
    const doc = makeDoc({ A1: 'N', A2: '1', A3: '2', B1: '=SUM(Table1[N])' });
    make(doc, 'A1:A3');
    doc.transact('dup', (tx) => duplicateSheet(tx, doc.wb, doc.wb.activeSheet));
    const copy = doc.wb.sheets[1];
    expect(copy.tables[0].name).toBe('Table2');
    expect(copy.getCell(0, 1)?.f).toBe('SUM(Table2[N])');
    set(doc, { A2: '10' }, copy);
    expect(val(doc, 'B1', copy)).toBe(12);
    expect(val(doc, 'B1', doc.wb.sheets[0])).toBe(3);
  });
});

describe('tables in .xlsx files', () => {
  it('saves real Excel tables and reads them back', async () => {
    const doc = makeDoc({ A1: 'Item', B1: 'Price', C1: 'Qty', A2: 'Pen', B2: '2', C2: '10', A3: 'Ink', B3: '5', C3: '3', A4: 'Pad', B4: '1', C4: '7' });
    const t = make(doc, 'A1:C4');
    const sheet = doc.wb.activeSheet;
    doc.transact('name', (tx) => renameTable(tx, doc.wb, sheet, t, 'Sales'));
    set(doc, { D1: 'Total', D2: '=[@Price]*[@Qty]', D3: '=[@Price]*[@Qty]', D4: '=[@Price]*[@Qty]', F1: '=SUM(Sales[Total])', F2: '=ROWS(Sales)' });
    sort(doc, [{ col: 2, desc: true }]);
    doc.transact('look', (tx) => setTableLook(tx, sheet, table(doc), { color: 'green', banded: false }));
    doc.transact('filter', (tx) => setTableFilter(tx, doc, sheet, table(doc), { 0: { values: ['Pen', 'Pad'] } }));

    const data = await writeXlsx(doc);
    const zip = await JSZip.loadAsync(data);
    const tableXml = await zip.file('xl/tables/table1.xml')!.async('string');
    expect(tableXml).toContain('name="Sales" displayName="Sales" ref="A1:D4"');
    expect(tableXml).toContain('<tableColumn id="4" name="Total"/>');
    expect(tableXml).toContain('<sortCondition descending="1" ref="C2:C4"/>');
    expect(tableXml).toContain('name="TableStyleMedium7"');
    expect(tableXml).toContain('showRowStripes="0"');
    expect(tableXml).toContain('<filterColumn colId="0"><filters><filter val="Pen"/><filter val="Pad"/></filters></filterColumn>');
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
    expect(sheetXml).toMatch(/<tableParts count="1"><tablePart r:id="rId\d+"\/><\/tableParts>/);
    expect(sheetXml).toContain('Sales[[#This Row],[Price]]*Sales[[#This Row],[Qty]]');
    // Excel can't read a bare table name in a file
    expect(sheetXml).toContain('<f>ROWS(Sales[])</f>');
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('/xl/tables/table1.xml');

    const { wb, warnings } = await readXlsx(data);
    expect(warnings).toEqual([]);
    const back = wb.sheets[0].tables[0];
    expect(back).toMatchObject({ name: 'Sales', range: R('A1:D4'), columns: ['Item', 'Price', 'Qty', 'Total'], color: 'green', banded: false, sort: [{ col: 2, desc: true }], filter: { 0: { values: ['Pen', 'Pad'] } } });
    expect(wb.sheets[0].getCell(1, 3)?.f).toBe('[@Price]*[@Qty]');
    expect(wb.sheets[0].getCell(1, 5)?.f).toBe('ROWS(Sales)');
    // The rows the filter hides stay filtered (not hidden by hand)
    expect([...wb.sheets[0].filterHidden]).toEqual([3]);
    expect(wb.sheets[0].hiddenRows.size).toBe(0);
    const reopened = new Doc(wb);
    expect(val(reopened, 'F1')).toBe(20 + 7 + 15);
    expect(val(reopened, 'D2')).toBe(20);
  });

  it('reads Excel’s styles, totals rows and table names', async () => {
    const doc = makeDoc({ A1: 'N', A2: '1', A3: '2', A4: '=SUBTOTAL(109,A2:A3)' });
    make(doc, 'A1:A4', { header: true, color: 'blue' });
    const data = await writeXlsx(doc);
    const zip = await JSZip.loadAsync(data);
    const xml = (await zip.file('xl/tables/table1.xml')!.async('string')).replace('ref="A1:A4"', 'ref="A1:A4" totalsRowCount="1"').replace(/TableStyleMedium\d+/, 'TableStyleLight10');
    zip.file('xl/tables/table1.xml', xml);
    const { wb } = await readXlsx(new Uint8Array(await zip.generateAsync({ type: 'uint8array' })));
    expect(wb.sheets[0].tables[0]).toMatchObject({ range: R('A1:A3'), color: 'orange', fileStyle: 'TableStyleLight10' });
  });
});
