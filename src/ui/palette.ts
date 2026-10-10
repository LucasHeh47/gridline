// Command palette (Ctrl+K): search every command, jump to cells and sheets, insert functions.
import * as L from 'lucide';
import type { IconNode } from 'lucide';
import { parseRange } from '../core/address';
import { FUNCTIONS } from '../core/formula/functions';
import type { App } from './app';
import { escapeHtml, h, icon, kbd } from './dom';

interface Entry {
  label: string;
  category: string;
  icon?: IconNode;
  keys?: string;
  keywords?: string;
  run: () => void;
  score?: number;
  marks?: number[];
}

/** Fuzzy match: all query characters in order. Rewards word starts and consecutive runs. */
function fuzzy(query: string, text: string): { score: number; marks: number[] } | null {
  const q = query.toLowerCase();
  const t = text.toLowerCase();
  if (!q) return { score: 0, marks: [] };
  const direct = t.indexOf(q);
  if (direct >= 0) {
    const marks = Array.from({ length: q.length }, (_, i) => direct + i);
    const wordStart = direct === 0 || /\W/.test(t[direct - 1]);
    return { score: 100 - direct + (wordStart ? 40 : 0) + q.length * 2, marks };
  }
  let ti = 0;
  let score = 0;
  const marks: number[] = [];
  let prev = -2;
  for (const ch of q) {
    if (ch === ' ') continue;
    const found = t.indexOf(ch, ti);
    if (found < 0) return null;
    if (found === prev + 1) score += 5;
    if (found === 0 || /\W/.test(t[found - 1])) score += 8;
    marks.push(found);
    prev = found;
    ti = found + 1;
  }
  return { score: score - t.length * 0.1, marks };
}

function highlight(text: string, marks: number[]): string {
  if (!marks.length) return escapeHtml(text);
  const set = new Set(marks);
  let out = '';
  for (let i = 0; i < text.length; i++) out += set.has(i) ? `<mark>${escapeHtml(text[i])}</mark>` : escapeHtml(text[i]);
  return out;
}

let open = false;

export function openPalette(app: App, initial = ''): void {
  if (open) return;
  open = true;
  const prevFocus = document.activeElement as HTMLElement | null;
  const input = h('input', { placeholder: 'Type a command, a cell like B12, a sheet name, or a function…', spellcheck: false, value: initial }) as HTMLInputElement;
  const list = h('div', { class: 'palette-list', role: 'listbox' });
  const box = h(
    'div',
    { class: 'palette', role: 'dialog' },
    h('div', { class: 'palette-input' }, icon(L.Search, 16), input),
    list,
    h('div', { class: 'palette-foot' }, h('span', null, kbd('↑↓'), ' move'), h('span', null, kbd('Enter'), ' run'), h('span', null, kbd('Esc'), ' close')),
  );
  const back = h('div', { class: 'modal-backdrop', style: { paddingTop: '10vh' } }, box);
  let entries: Entry[] = [];
  let index = 0;

  const close = (restore = true) => {
    open = false;
    back.remove();
    if (restore) {
      if (prevFocus && prevFocus.isConnected && prevFocus !== document.body) prevFocus.focus();
      else app.focusGrid();
    }
  };

  const commandEntries = (): Entry[] =>
    app.commands.all().map((c) => ({
      label: c.title + (c.checked ? (c.checked() ? '  ✓' : '') : ''),
      category: c.category,
      icon: c.icon,
      keys: c.keys?.[0],
      keywords: c.keywords,
      run: () => app.commands.run(c.id),
    }));

  const build = () => {
    const q = input.value.trim();
    const out: Entry[] = [];
    if (!q) {
      const recent = app.prefs.recentCommands.map((id) => app.commands.get(id)).filter((c): c is NonNullable<typeof c> => !!c && !c.hidden);
      for (const c of recent) out.push({ label: c.title, category: 'Recently used', icon: c.icon, keys: c.keys?.[0], run: () => app.commands.run(c.id) });
      const seen = new Set(recent.map((c) => c.title));
      for (const e of commandEntries()) if (!seen.has(e.label)) out.push(e);
      entries = out;
      return;
    }
    // Go to cell / range
    const addr = q.replace(/^=/, '');
    if (/^('?[^!]+'?!)?\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(addr) && parseRange(addr.split('!').pop()!)) {
      out.push({ label: `Go to ${addr.toUpperCase()}`, category: 'Navigate', icon: L.Crosshair, run: () => app.goTo(addr), score: 1000 });
    }
    // Sheets
    for (const s of app.doc.wb.sheets) {
      const m = fuzzy(q, s.name);
      if (m) out.push({ label: `Go to sheet: ${s.name}`, category: 'Sheets', icon: L.Sheet, run: () => app.setActiveSheet(s.id), score: m.score + 20, marks: m.marks.map((i) => i + 'Go to sheet: '.length) });
    }
    // Defined names
    for (const dn of app.doc.wb.names.values()) {
      const m = fuzzy(q, dn.name);
      if (m) out.push({ label: `Go to name: ${dn.name}`, category: 'Names', icon: L.Tag, run: () => app.goTo(dn.ref), score: m.score + 10 });
    }
    // Commands
    for (const e of commandEntries()) {
      const m = fuzzy(q, e.label);
      const k = e.keywords ? fuzzy(q, e.keywords) : null;
      if (m || (k && k.score > 20)) out.push({ ...e, score: Math.max(m?.score ?? 0, (k?.score ?? 0) - 15), marks: m?.marks ?? [] });
    }
    // Functions
    const fq = q.toUpperCase().replace(/[=(]/g, '');
    if (fq.length >= 2) {
      for (const def of FUNCTIONS.values()) {
        if (!def.name.includes(fq)) continue;
        const starts = def.name.startsWith(fq);
        out.push({
          label: `${def.name}(${def.args}) — ${def.desc}`,
          category: `Function · ${def.category}`,
          icon: L.SquareFunction,
          run: () => app.editor.start({ mode: 'enter', initial: `=${def.name}(`, source: 'cell' }),
          score: (starts ? 60 : 30) - def.name.length,
          marks: [...Array(fq.length)].map((_, i) => def.name.indexOf(fq) + i),
        });
      }
    }
    out.sort((a, b) => (b.score ?? 0) - (a.score ?? 0));
    entries = out.slice(0, 60);
  };

  const render = () => {
    list.innerHTML = '';
    let lastCat = '';
    const showSections = !input.value.trim();
    entries.forEach((e, i) => {
      if (showSections && e.category !== lastCat && (e.category === 'Recently used' || lastCat === 'Recently used' || lastCat === '')) {
        list.append(h('div', { class: 'pal-section' }, e.category === 'Recently used' ? 'Recently used' : 'All commands'));
      }
      lastCat = e.category;
      const label = h('span', { class: 'pi-label' });
      label.innerHTML = highlight(e.label, e.marks ?? []);
      const row = h(
        'div',
        { class: `pal-item${i === index ? ' sel' : ''}`, role: 'option' },
        h('span', { class: 'pi-icon' }, e.icon ? icon(e.icon, 16) : null),
        label,
        h('span', { class: 'pi-cat' }, e.category),
        e.keys ? kbd(e.keys) : null,
      );
      row.addEventListener('mousemove', () => {
        if (index !== i) {
          index = i;
          render();
        }
      });
      row.addEventListener('click', () => choose(i));
      list.append(row);
    });
    if (!entries.length) list.append(h('div', { class: 'hint', style: { padding: '14px' } }, 'Nothing matches. Try other words, like "sum", "freeze" or "currency".'));
    (list.querySelector('.pal-item.sel') as HTMLElement | null)?.scrollIntoView({ block: 'nearest' });
  };

  const choose = (i: number) => {
    const e = entries[i];
    if (!e) return;
    close(false);
    app.focusGrid();
    e.run();
  };

  input.addEventListener('input', () => {
    index = 0;
    build();
    render();
  });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') {
      index = Math.min(entries.length - 1, index + 1);
      render();
    } else if (e.key === 'ArrowUp') {
      index = Math.max(0, index - 1);
      render();
    } else if (e.key === 'Enter') choose(index);
    else if (e.key === 'Escape') close();
    else return;
    e.preventDefault();
    e.stopPropagation();
  });
  back.addEventListener('mousedown', (e) => {
    if (e.target === back) close();
  });
  document.body.append(back);
  build();
  render();
  input.focus();
  input.select();
}
