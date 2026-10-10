// Sheet-level operations.
import { keyCol, keyRow } from '../address';
import { deleteSheetRefs, renameSheetRefs } from '../formula/refs';
import type { Transaction } from '../history';
import { Sheet, type Workbook } from '../workbook';

export function addSheet(tx: Transaction, wb: Workbook, name?: string, index?: number): Sheet {
  const sheet = new Sheet(name ?? wb.uniqueSheetName());
  const sheets = [...wb.sheets];
  sheets.splice(index ?? sheets.length, 0, sheet);
  tx.setProp(wb, 'sheets', sheets, true);
  tx.setProp(wb, 'activeSheetId', sheet.id);
  return sheet;
}

function rewriteAllFormulas(tx: Transaction, wb: Workbook, fn: (f: string) => string): void {
  for (const s of wb.sheets) {
    for (const [key, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const f = fn(cell.f);
      if (f !== cell.f) tx.setCell(s, keyRow(key), keyCol(key), { ...cell, f });
    }
    if (s.charts.length) {
      const charts = s.charts.map((c) => ({
        ...c,
        series: c.series.map((se) => ({
          ...se,
          name: se.name?.startsWith('=') ? '=' + fn(se.name.slice(1)) : se.name,
          categories: se.categories ? fn(se.categories) : se.categories,
          values: fn(se.values),
        })),
      }));
      tx.setProp(s, 'charts', charts);
    }
    if (s.cf.some((r) => r.type === 'expression')) {
      tx.setProp(
        s,
        'cf',
        s.cf.map((r) => (r.type === 'expression' ? { ...r, formula: fn(r.formula) } : r)),
      );
    }
  }
  if (wb.names.size) {
    const names = new Map(wb.names);
    for (const [k, dn] of names) names.set(k, { ...dn, ref: fn(dn.ref) });
    tx.setProp(wb, 'names', names, true);
  }
}

export function deleteSheet(tx: Transaction, wb: Workbook, sheet: Sheet): void {
  if (wb.sheets.length <= 1) throw new Error('A workbook needs at least one sheet.');
  const idx = wb.sheets.indexOf(sheet);
  const sheets = wb.sheets.filter((s) => s !== sheet);
  tx.setProp(wb, 'sheets', sheets, true);
  rewriteAllFormulas(tx, wb, (f) => deleteSheetRefs(f, sheet.name));
  if (wb.activeSheetId === sheet.id) {
    const visible = sheets.filter((s) => !s.hidden);
    const next = visible[Math.min(idx, visible.length - 1)] ?? sheets[0];
    tx.setProp(wb, 'activeSheetId', next.id);
  }
}

export function renameSheet(tx: Transaction, wb: Workbook, sheet: Sheet, name: string): void {
  const old = sheet.name;
  if (old === name) return;
  tx.setProp(sheet, 'name', name, true);
  rewriteAllFormulas(tx, wb, (f) => renameSheetRefs(f, old, name));
}

export function moveSheet(tx: Transaction, wb: Workbook, sheet: Sheet, toIndex: number): void {
  const sheets = wb.sheets.filter((s) => s !== sheet);
  sheets.splice(Math.max(0, Math.min(toIndex, sheets.length)), 0, sheet);
  tx.setProp(wb, 'sheets', sheets, true);
}

export function duplicateSheet(tx: Transaction, wb: Workbook, sheet: Sheet): Sheet {
  let base = `${sheet.name} (2)`;
  let i = 2;
  while (wb.sheetByName(base)) base = `${sheet.name} (${++i})`;
  const copy = new Sheet(base.slice(0, 31));
  copy.cells = new Map([...sheet.cells].map(([k, c]) => [k, { ...c }]));
  copy.colWidths = new Map(sheet.colWidths);
  copy.rowHeights = new Map(sheet.rowHeights);
  copy.hiddenCols = new Set(sheet.hiddenCols);
  copy.hiddenRows = new Set(sheet.hiddenRows);
  copy.colStyles = new Map(sheet.colStyles);
  copy.rowStyles = new Map(sheet.rowStyles);
  copy.merges = sheet.merges.map((m) => ({ ...m }));
  copy.freeze = { ...sheet.freeze };
  copy.cf = sheet.cf.map((r) => ({ ...r, id: r.id + '-copy', ranges: r.ranges.map((x) => ({ ...x })) }));
  copy.charts = sheet.charts.map((c) => ({
    ...c,
    id: c.id + '-' + Math.random().toString(36).slice(2, 7),
    series: c.series.map((se) => ({
      ...se,
      name: se.name?.startsWith('=') ? '=' + renameSheetRefs(se.name.slice(1), sheet.name, copy.name) : se.name,
      categories: se.categories ? renameSheetRefs(se.categories, sheet.name, copy.name) : undefined,
      values: renameSheetRefs(se.values, sheet.name, copy.name),
    })),
    anchor: { ...c.anchor },
  }));
  copy.tabColor = sheet.tabColor;
  copy.showGrid = sheet.showGrid;
  copy.defaultColWidth = sheet.defaultColWidth;
  copy.defaultRowHeight = sheet.defaultRowHeight;
  const sheets = [...wb.sheets];
  sheets.splice(wb.sheets.indexOf(sheet) + 1, 0, copy);
  tx.setProp(wb, 'sheets', sheets, true);
  tx.setProp(wb, 'activeSheetId', copy.id);
  return copy;
}

export function setSheetHidden(tx: Transaction, wb: Workbook, sheet: Sheet, hidden: boolean): void {
  if (hidden && wb.sheets.filter((s) => !s.hidden).length <= 1) throw new Error('At least one sheet must stay visible.');
  tx.setProp(sheet, 'hidden', hidden);
  if (hidden && wb.activeSheetId === sheet.id) {
    const next = wb.sheets.find((s) => !s.hidden && s !== sheet)!;
    tx.setProp(wb, 'activeSheetId', next.id);
  }
}
