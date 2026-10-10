// File → New from template: a gallery of starter workbooks.
import * as L from 'lucide';
import { TEMPLATES } from '../core/templates';
import { h, icon } from './dom';
import { dialog } from './widgets';

const ICONS: Record<string, L.IconNode> = {
  Wallet: L.Wallet,
  House: L.House,
  TrendingDown: L.TrendingDown,
  PiggyBank: L.PiggyBank,
  ChartCandlestick: L.ChartCandlestick,
  Trophy: L.Trophy,
};

/** Shows the template gallery; resolves to the chosen template id, or null. */
export function pickTemplate(): Promise<string | null> {
  let choose: (id: string) => void = () => undefined;
  const body = h('div', { class: 'tpl-gallery' });
  for (const group of ['Money', 'NFL'] as const) {
    body.append(h('div', { class: 'tpl-group' }, group));
    const grid = h('div', { class: 'tpl-grid' });
    for (const t of TEMPLATES.filter((x) => x.group === group)) {
      grid.append(
        h(
          'button',
          { class: 'tpl-card', onclick: () => choose(t.id) },
          h('span', { class: 'tpl-icon' }, icon(ICONS[t.icon] ?? L.FileSpreadsheet, 20)),
          h('span', { class: 'tpl-text' }, h('span', { class: 'tpl-name' }, t.name), h('span', { class: 'tpl-desc' }, t.description)),
        ),
      );
    }
    body.append(grid);
  }
  const shown = dialog<string | null>('New from template', body, [{ label: 'Cancel', value: null }], { wide: true, cancelValue: null });
  return new Promise((resolve) => {
    choose = (id) => {
      // Close the dialog the same way Cancel does, then report the pick
      (body.closest('.modal')?.querySelector('.actions .btn') as HTMLButtonElement | null)?.click();
      resolve(id);
    };
    void shown.then(() => resolve(null));
  });
}
