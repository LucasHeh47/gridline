// Keyboard shortcut reference and About dialog.
import type { App } from './app';
import { h, kbd } from './dom';
import { platform } from './platform';
import { dialog } from './widgets';

const NAV: [string, string][] = [
  ['Move', 'Arrow keys'],
  ['Jump to the edge of the data', 'Ctrl+Arrow'],
  ['Extend the selection', 'Shift+Arrow'],
  ['Next cell / previous cell', 'Tab / Shift+Tab'],
  ['Confirm and move down / up', 'Enter / Shift+Enter'],
  ['Start editing the cell', 'F2 or double-click'],
  ['New line inside a cell', 'Alt+Enter'],
  ['Enter in every selected cell', 'Ctrl+Enter'],
  ['Cycle $ in a reference while editing', 'F4'],
  ['Go to cell A1 / last used cell', 'Ctrl+Home / Ctrl+End'],
  ['Page up / down', 'PageUp / PageDown'],
  ['Zoom', 'Ctrl+Mouse wheel'],
  ['Open a link', 'Ctrl+Click'],
];

export function showShortcuts(app: App): void {
  const table = h('div', { class: 'shortcut-table' });
  table.append(h('div', { class: 'cat' }, 'Moving & editing'));
  for (const [label, keys] of NAV) table.append(h('div', null, label), h('div', null, kbd(keys)));
  const byCat = new Map<string, { title: string; keys: string[] }[]>();
  for (const c of app.commands.all()) {
    if (!c.keys?.length) continue;
    const list = byCat.get(c.category) ?? [];
    list.push({ title: c.title, keys: c.keys });
    byCat.set(c.category, list);
  }
  byCat.set('Search', [{ title: 'Search all commands', keys: ['Ctrl+K'] }, ...(byCat.get('Search') ?? [])]);
  for (const [cat, list] of byCat) {
    table.append(h('div', { class: 'cat' }, cat));
    for (const c of list) table.append(h('div', null, c.title), h('div', null, ...c.keys.map((k) => kbd(k))));
  }
  const body = h('div', null, h('p', null, 'Every command is also searchable with Ctrl+K — no need to memorize these.'), table);
  void dialog('Keyboard shortcuts', body, [{ label: 'Close', value: true, primary: true }], { wide: true });
}

export async function showAbout(): Promise<void> {
  const version = await platform.appVersion();
  const body = h(
    'div',
    null,
    h('p', null, `Version ${version}${version.includes('-') ? ' (development build)' : ''}. A fast, simple spreadsheet that opens and saves Excel (.xlsx) and CSV files.`),
    h('p', null, 'Live data comes from nflverse (NFL), Yahoo Finance (stocks), Coinbase (crypto), Frankfurter (exchange rates), Open-Meteo (weather), Nager.Date (holidays) and the World Bank (country data).'),
    h('p', null, 'Formulas: 250+ functions including XLOOKUP, FILTER, SORT, UNIQUE and LET, with dynamic arrays. Errors explain themselves in plain English.'),
  );
  await dialog('About Gridline', body, [{ label: 'Close', value: true, primary: true }]);
}
