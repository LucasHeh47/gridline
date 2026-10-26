// Right-click menus for cells, row headers and column headers.
import * as L from 'lucide';
import type { App } from './app';
import { activeTable } from './tableMenu';
import { showMenu, type MenuItem } from './widgets';

export function showContextMenu(app: App, e: MouseEvent, kind: 'cell' | 'row' | 'col'): void {
  const item = (id: string, label?: string): MenuItem => {
    const c = app.commands.get(id)!;
    return { label: label ?? c.title, icon: c.icon, hint: c.keys?.[0], checked: c.checked?.(), disabled: c.enabled ? !c.enabled() : false, run: () => app.commands.run(id) };
  };
  const pasteSpecial: MenuItem = {
    label: 'Paste special',
    icon: L.ClipboardList,
    submenu: () => [item('edit.pasteValues', 'Values only'), item('edit.pasteFormats', 'Formatting only'), item('edit.pasteFormulas', 'Formulas only'), item('edit.pasteTranspose', 'Transposed'), item('edit.pasteNoBorders', 'Everything except borders')],
  };
  const clipboard: MenuItem[] = [item('edit.cut'), item('edit.copy'), item('edit.paste'), pasteSpecial];
  const a = app.sel.active;
  const hasNote = !!app.sheet.getCell(a.row, a.col)?.note;
  const hasLink = !!app.sheet.getCell(a.row, a.col)?.link;
  const hasFormula = app.sheet.getCell(a.row, a.col)?.f !== undefined;
  const table = activeTable(app);
  let items: MenuItem[];
  if (kind === 'row') {
    const n = app.sel.fullRows().length || 1;
    const s = n > 1 ? `${n} rows` : 'row';
    items = [
      ...clipboard,
      { separator: true },
      item('ins.rowsAbove', `Insert ${s} above`),
      item('ins.rowsBelow', `Insert ${s} below`),
      item('del.rows', `Delete ${s}`),
      item('edit.clearContents'),
      { separator: true },
      item('rows.height', 'Row height…'),
      item('rows.autofit', 'Auto-fit row height'),
      item('rows.hide', `Hide ${s}`),
      item('rows.unhide', 'Unhide rows'),
      { separator: true },
      item('view.freezeHere', 'Freeze up to this row'),
      item('fmt.cells'),
    ];
  } else if (kind === 'col') {
    const n = app.sel.fullColumns().length || 1;
    const s = n > 1 ? `${n} columns` : 'column';
    items = [
      ...clipboard,
      { separator: true },
      item('ins.colsLeft', `Insert ${s} left`),
      item('ins.colsRight', `Insert ${s} right`),
      item('del.cols', `Delete ${s}`),
      item('edit.clearContents'),
      { separator: true },
      item('data.sortAsc', 'Sort sheet A → Z by this column'),
      item('data.sortDesc', 'Sort sheet Z → A by this column'),
      { separator: true },
      item('cols.width', 'Column width…'),
      item('cols.autofit', 'Auto-fit column width'),
      item('cols.hide', `Hide ${s}`),
      item('cols.unhide', 'Unhide columns'),
      { separator: true },
      item('fmt.cells'),
    ];
  } else {
    items = [
      ...clipboard,
      { separator: true },
      {
        label: 'Insert',
        icon: L.Plus,
        submenu: () => [item('ins.rowsAbove', 'Rows above'), item('ins.rowsBelow', 'Rows below'), item('ins.colsLeft', 'Columns left'), item('ins.colsRight', 'Columns right')],
      },
      {
        label: 'Delete',
        icon: L.Minus,
        submenu: () => [item('del.rows', 'Entire rows'), item('del.cols', 'Entire columns')],
      },
      {
        label: 'Clear',
        icon: L.Eraser,
        submenu: () => [item('edit.clearContents', 'Contents'), item('edit.clearFormats', 'Formatting'), item('edit.clearAll', 'Everything'), item('fmt.clearCf', 'Conditional formatting')],
      },
      { separator: true },
      ...(table
        ? [
            {
              label: `Table ${table.name}`,
              icon: L.Table,
              submenu: () => [
                item('table.sortHere', `Sort by ${table.columns[a.col - table.range.c1] ?? 'this column'}`),
                item('table.originalOrder', 'Original order'),
                { separator: true },
                item('table.select', 'Select table'),
                item('table.options', 'Table options…'),
                item('table.convert', 'Convert to normal cells'),
              ],
            } as MenuItem,
          ]
        : [{ label: 'Sort', icon: L.ArrowUpDown, submenu: () => [item('data.sortAsc'), item('data.sortDesc'), item('data.sortCustom')] } as MenuItem, item('data.filter', app.sheet.filter ? 'Remove filter' : 'Filter'), item('table.create', 'Make a table')]),
      { separator: true },
      item('fmt.cells'),
      item('fmt.cf'),
      item('edit.note', hasNote ? 'Edit note' : 'Add note'),
      ...(hasNote ? [item('edit.deleteNote')] : []),
      item('edit.link', hasLink ? 'Edit link' : 'Insert link'),
      ...(hasLink ? [item('edit.removeLink')] : []),
      ...(hasFormula ? [item('data.trace')] : []),
      { separator: true },
      {
        label: 'Insert chart',
        icon: L.ChartColumn,
        submenu: () => ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'].map((t) => item(`ins.chart.${t}`)),
      },
    ];
  }
  showMenu(items, { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY });
}
