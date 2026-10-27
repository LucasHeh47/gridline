// The app switcher (the dotted button in the title bar): Baseline, Gridline, Pixelline, Frameline and
// Beatline, to open one or start something new in it. Apps that aren't installed link to their download page.
import * as L from 'lucide';
import type { SuiteApp, SuiteAppId } from '../shared/api';
import { h, icon } from './dom';
import { platform } from './platform';
import { Popover, toast, tooltip } from './widgets';

/** This app. */
const SELF: SuiteAppId = 'gridline';

export const SUITE_LOOK: Record<SuiteAppId, { icon: L.IconNode; colors: [string, string]; action?: { label: string; args: string[] } }> = {
  baseline: { icon: L.LayoutDashboard, colors: ['#ffb547', '#ff6f61'] },
  gridline: { icon: L.Grid2x2, colors: ['#4ea1ff', '#3fcf8e'], action: { label: 'New spreadsheet', args: ['--new'] } },
  pixelline: { icon: L.Shapes, colors: ['#a98bff', '#ff7ab6'], action: { label: 'New image', args: ['--new=32x32'] } },
  frameline: { icon: L.Clapperboard, colors: ['#ff5c6c', '#a98bff'], action: { label: 'New video', args: ['--new=1920x1080'] } },
  beatline: { icon: L.AudioLines, colors: ['#2dd4bf', '#6366f1'], action: { label: 'New beat', args: ['--new'] } },
};

export function suiteMark(id: SuiteAppId, size = 16): HTMLElement {
  const look = SUITE_LOOK[id];
  const el = h('span', { class: 'suite-mark' }, icon(look.icon, size));
  el.style.background = `linear-gradient(135deg, ${look.colors[0]}, ${look.colors[1]})`;
  return el;
}

let pop: Popover | null = null;

export function suiteButton(): HTMLElement {
  const b = h('button', { class: 'btn suite-btn', 'aria-label': 'Apps', onclick: () => void openSuiteMenu(b) }, icon(L.Grip, 16));
  tooltip(b, 'Your apps: Baseline, Gridline, Pixelline, Frameline and Beatline');
  return b;
}

async function openSuiteMenu(anchor: HTMLElement): Promise<void> {
  if (pop?.isOpen) {
    pop.close();
    return;
  }
  const apps: SuiteApp[] = await platform.suite.apps();
  const list = h('div', { class: 'suite-list' });
  const row = (a: SuiteApp) => {
    const look = SUITE_LOOK[a.id];
    const self = a.id === SELF;
    const open = async (args: string[] = []) => {
      pop?.close();
      if (self) return;
      if (!a.installed) {
        platform.openExternal(a.download);
        return;
      }
      if (!(await platform.suite.launch(a.id, args))) toast(`Couldn't start ${a.name}.`, { kind: 'error' });
      else toast(`Opening ${a.name}…`);
    };
    const el = h(
      'div',
      { class: `suite-row${self ? ' self' : ''}`, role: 'menuitem', tabindex: '0' },
      suiteMark(a.id, 18),
      h('span', { class: 'suite-text' }, h('span', { class: 'suite-name' }, a.name), h('span', { class: 'suite-desc' }, self ? 'You’re here' : a.installed ? a.description : `${a.description} · not installed`)),
      !self && a.installed && look.action ? h('button', { class: 'btn text small', onclick: (e: MouseEvent) => {
        e.stopPropagation();
        void open(look.action!.args);
      } }, look.action.label) : null,
      !self && !a.installed ? h('span', { class: 'suite-get' }, icon(L.Download, 13), 'Get it') : null,
    );
    el.addEventListener('click', () => void open());
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') void open();
    });
    return el;
  };
  if (!apps.length) list.append(h('div', { class: 'hint suite-empty' }, 'The app switcher works in the desktop app.'));
  for (const a of apps) list.append(row(a));
  pop = new Popover(h('div', { class: 'suite-menu' }, h('div', { class: 'suite-title' }, 'Apps'), list), anchor, 'suite-pop');
  pop.open(anchor, 'below');
}
