// Function browser: search all functions by name or description and insert them.
import { FUNCTIONS, type Category } from '../../core/formula/functions';
import type { App } from '../app';
import { h } from '../dom';
import type { PanelContent } from './SidePanel';

const CATS: (Category | 'All')[] = ['All', 'NFL', 'Personal finance', 'Math', 'Statistical', 'Lookup', 'Text', 'Logical', 'Date & time', 'Financial', 'Information'];

export class FunctionsPanel implements PanelContent {
  title = 'Functions';
  readonly el: HTMLElement;
  private search: HTMLInputElement;
  private list: HTMLDivElement;
  private cats: HTMLDivElement;
  private cat: Category | 'All' = 'All';

  constructor(private app: App) {
    this.search = h('input', { class: 'input', placeholder: `Search ${FUNCTIONS.size} functions — e.g. "lookup", "average", "date"`, spellcheck: false, style: { width: '100%' } }) as HTMLInputElement;
    this.search.addEventListener('input', () => this.render());
    this.cats = h('div', { class: 'chips', style: { margin: '10px 0' } });
    this.list = h('div', { class: 'fn-list' });
    this.el = h('div', { class: 'sp-body' }, this.search, this.cats, this.list);
    this.render();
    setTimeout(() => this.search.focus(), 0);
  }

  private render(): void {
    this.cats.innerHTML = '';
    for (const c of CATS) {
      this.cats.append(h('button', { class: `chip${c === this.cat ? ' on' : ''}`, onclick: () => {
        this.cat = c;
        this.render();
      } }, c));
    }
    const q = this.search.value.trim().toLowerCase();
    const defs = [...FUNCTIONS.values()]
      .filter((d) => this.cat === 'All' || d.category === this.cat)
      .filter((d) => !q || d.name.toLowerCase().includes(q) || d.desc.toLowerCase().includes(q))
      .sort((a, b) => {
        if (q) {
          const as = a.name.toLowerCase().startsWith(q) ? 0 : 1;
          const bs = b.name.toLowerCase().startsWith(q) ? 0 : 1;
          if (as !== bs) return as - bs;
        }
        return a.name.localeCompare(b.name);
      });
    this.list.innerHTML = '';
    for (const d of defs.slice(0, 300)) {
      const item = h('div', { class: 'fn-item', title: 'Click to insert' }, h('div', { class: 'n' }, `${d.name}(${d.args})`), h('div', { class: 'd' }, d.desc));
      // Keep focus in the cell editor while clicking
      item.addEventListener('mousedown', (e) => e.preventDefault());
      item.addEventListener('click', () => this.insert(d.name));
      this.list.append(item);
    }
    if (!defs.length) this.list.append(h('div', { class: 'hint' }, 'No functions match.'));
  }

  private insert(name: string): void {
    const app = this.app;
    if (app.isEditing) {
      app.editor.insertText(app.editor.text.startsWith('=') ? `${name}(` : `=${name}(`);
      return;
    }
    app.editor.start({ mode: 'enter', initial: `=${name}(`, source: 'cell' });
  }
}
