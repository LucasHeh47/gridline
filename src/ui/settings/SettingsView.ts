// The Settings page: a searchable list of every setting, grouped into sections.
// Changes apply immediately; each changed setting has a reset button.
import * as L from 'lucide';
import type { App } from '../app';
import { h, icon } from '../dom';
import { isDesktop, platform } from '../platform';
import { confirmDialog, toast } from '../widgets';
import { SECTIONS } from './schema';
import type { SettingDef } from './store';

export class SettingsView {
  readonly el: HTMLElement;
  private nav: HTMLElement;
  private content: HTMLElement;
  private search: HTMLInputElement;
  private section: string;
  private off: () => void;
  private onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape' && !document.querySelector('.modal-backdrop, .popover')) {
      e.preventDefault();
      e.stopPropagation();
      this.close();
    }
  };

  constructor(
    private app: App,
    section?: string,
    private onClosed?: () => void,
  ) {
    this.section = section && SECTIONS.some((s) => s.id === section) ? section : SECTIONS[0].id;
    this.search = h('input', { class: 'input st-search', type: 'search', placeholder: `Search ${app.settings.defs.length} settings`, spellcheck: false, 'aria-label': 'Search settings' }) as HTMLInputElement;
    this.search.addEventListener('input', () => this.render());
    this.nav = h('nav', { class: 'st-nav' });
    this.content = h('div', { class: 'st-content' });
    const close = h('button', { class: 'btn', 'aria-label': 'Close settings', onclick: () => this.close() }, icon(L.X, 18));
    const head = h('div', { class: 'st-head' }, h('div', { class: 'st-title' }, icon(L.Settings, 18), 'Settings'), this.search, h('div', { class: 'spacer' }), close);
    const foot = h(
      'div',
      { class: 'st-nav-foot' },
      h('button', { class: 'btn text', onclick: () => void this.exportSettings() }, icon(L.Download, 14), 'Export…'),
      h('button', { class: 'btn text', onclick: () => this.importSettings() }, icon(L.Upload, 14), 'Import…'),
      h('button', { class: 'btn text danger', onclick: () => void this.resetAll() }, icon(L.RotateCcw, 14), 'Reset all'),
    );
    this.el = h('div', { class: 'settings-view', role: 'dialog', 'aria-label': 'Settings' }, head, h('div', { class: 'st-body' }, h('div', { class: 'st-side' }, this.nav, foot), this.content));
    // Re-render when a setting changes (here or in another window), keeping focus where it was
    this.off = app.settings.onChange(() => {
      if (!this.el.contains(document.activeElement) || document.activeElement === this.search) this.render();
      else this.renderNav();
    });
    document.addEventListener('keydown', this.onKey, true);
    this.render();
    requestAnimationFrame(() => this.search.focus());
  }

  close(): void {
    this.off();
    document.removeEventListener('keydown', this.onKey, true);
    this.el.remove();
    this.onClosed?.();
  }

  show(section?: string): void {
    if (section && SECTIONS.some((s) => s.id === section)) {
      this.section = section;
      this.search.value = '';
      this.render();
    }
  }

  private matches(d: SettingDef, q: string): boolean {
    if (!q) return true;
    const opts = this.app.settings.options(d).map((o) => o.label).join(' ');
    const sec = SECTIONS.find((s) => s.id === d.section)?.label ?? '';
    const hay = `${d.label} ${d.description ?? ''} ${d.keywords ?? ''} ${d.group} ${sec} ${opts}`.toLowerCase();
    return q.split(/\s+/).every((w) => hay.includes(w));
  }

  private renderNav(): void {
    const s = this.app.settings;
    const q = this.search.value.trim().toLowerCase();
    this.nav.innerHTML = '';
    for (const sec of SECTIONS) {
      const defs = s.defs.filter((d) => d.section === sec.id && this.matches(d, q));
      const modified = defs.filter((d) => s.isModified(d.id)).length;
      const b = h(
        'button',
        { class: `st-nav-item${!q && sec.id === this.section ? ' on' : ''}${q && !defs.length ? ' dim' : ''}`, onclick: () => {
          this.search.value = '';
          this.section = sec.id;
          this.render();
        } },
        icon(sec.icon, 16),
        h('span', { class: 'n' }, sec.label),
        modified ? h('span', { class: 'st-mod-count', title: `${modified} changed` }, String(modified)) : h('span', { class: 'st-count' }, String(defs.length)),
      );
      this.nav.append(b);
    }
  }

  private render(): void {
    this.renderNav();
    const s = this.app.settings;
    const q = this.search.value.trim().toLowerCase();
    const scroll = this.content.scrollTop;
    this.content.innerHTML = '';
    const sections = q ? SECTIONS : SECTIONS.filter((x) => x.id === this.section);
    let shown = 0;
    for (const sec of sections) {
      const defs = s.defs.filter((d) => d.section === sec.id && this.matches(d, q));
      if (!defs.length) continue;
      shown += defs.length;
      this.content.append(h('h2', { class: 'st-section' }, icon(sec.icon, 18), sec.label));
      if (!q) this.content.append(h('p', { class: 'st-section-desc' }, sec.description));
      let group = '';
      for (const d of defs) {
        if (d.group !== group) {
          group = d.group;
          this.content.append(h('h3', { class: 'st-group' }, group));
        }
        this.content.append(this.row(d));
      }
    }
    if (!shown) this.content.append(h('div', { class: 'st-empty' }, icon(L.SearchX, 28), h('p', null, `No settings match "${this.search.value}".`)));
    if (!q) this.content.scrollTop = scroll;
  }

  private row(d: SettingDef): HTMLElement {
    const s = this.app.settings;
    const unavailable = d.desktopOnly && !isDesktop;
    const control = this.control(d, !!unavailable);
    const modified = s.isModified(d.id);
    const reset = modified
      ? h('button', { class: 'btn st-reset', title: 'Reset to default', 'aria-label': `Reset ${d.label}`, onclick: () => s.reset(d.id) }, icon(L.RotateCcw, 14))
      : h('span', { class: 'st-reset-space' });
    return h(
      'div',
      { class: `st-row${modified ? ' modified' : ''}${unavailable ? ' unavailable' : ''}`, 'data-id': d.id },
      h('div', { class: 'st-text' }, h('div', { class: 'st-label' }, d.label), d.description ? h('div', { class: 'st-desc' }, d.description) : null, unavailable ? h('div', { class: 'st-desc' }, 'Available in the desktop app.') : null),
      h('div', { class: 'st-control' }, control, reset),
    );
  }

  private control(d: SettingDef, disabled: boolean): HTMLElement {
    const s = this.app.settings;
    switch (d.type) {
      case 'toggle': {
        const input = h('input', { type: 'checkbox', role: 'switch', 'aria-label': d.label }) as HTMLInputElement;
        input.checked = s.bool(d.id);
        input.disabled = disabled;
        input.addEventListener('change', () => s.set(d.id, input.checked));
        return h('label', { class: 'switch' }, input, h('span', { class: 'track' }, h('span', { class: 'thumb' })));
      }
      case 'select': {
        const sel = h('select', { class: 'input', 'aria-label': d.label }) as HTMLSelectElement;
        for (const o of s.options(d)) sel.append(h('option', { value: o.value }, o.label));
        sel.value = s.str(d.id);
        sel.disabled = disabled;
        sel.addEventListener('change', () => s.set(d.id, sel.value));
        return sel;
      }
      case 'number': {
        const input = h('input', { class: 'input st-num', type: 'number', min: String(d.min ?? ''), max: String(d.max ?? ''), step: String(d.step ?? 1), 'aria-label': d.label }) as HTMLInputElement;
        input.value = String(s.num(d.id));
        input.disabled = disabled;
        const commit = () => {
          const n = Number(input.value);
          if (!Number.isFinite(n)) return;
          const clamped = Math.min(d.max ?? Infinity, Math.max(d.min ?? -Infinity, n));
          input.value = String(clamped);
          s.set(d.id, clamped);
        };
        input.addEventListener('change', commit);
        input.addEventListener('keydown', (e) => {
          if (e.key === 'Enter') commit();
        });
        return h('div', { class: 'st-num-wrap' }, input, d.unit ? h('span', { class: 'st-unit' }, d.unit) : null);
      }
      case 'text': {
        const input = h('input', { class: 'input', value: s.str(d.id), spellcheck: false, 'aria-label': d.label }) as HTMLInputElement;
        input.disabled = disabled;
        input.addEventListener('change', () => {
          const v = input.value.trim();
          if (v) s.set(d.id, v);
          else s.reset(d.id);
        });
        return input;
      }
      case 'multi': {
        const wrap = h('div', { class: 'chips st-multi' });
        const cur = new Set(s.list(d.id));
        for (const o of s.options(d)) {
          const on = cur.has(o.value);
          const b = h('button', { class: `chip${on ? ' on' : ''}`, 'aria-pressed': String(on), onclick: () => {
            const next = new Set(s.list(d.id));
            if (next.has(o.value)) next.delete(o.value);
            else next.add(o.value);
            // Keep the original order
            s.set(d.id, s.options(d).map((x) => x.value).filter((v) => next.has(v)));
          } }, on ? icon(L.Check, 12) : null, o.label);
          (b as HTMLButtonElement).disabled = disabled;
          wrap.append(b);
        }
        return wrap;
      }
      case 'action': {
        const b = h('button', { class: 'btn outline', onclick: () => void d.action?.run(this.app) }, d.action?.label ?? 'Run') as HTMLButtonElement;
        b.disabled = disabled;
        return b;
      }
      case 'info': {
        const span = h('span', { class: 'st-info' }, '…');
        if (!disabled && d.info) {
          void Promise.resolve(d.info(this.app))
            .then((t) => (span.textContent = t))
            .catch(() => (span.textContent = '—'));
        } else span.textContent = '—';
        return span;
      }
    }
  }

  private async exportSettings(): Promise<void> {
    const json = this.app.settings.export();
    const path = await platform.saveFileDialog('gridline-settings.json', 'json');
    if (!path) return;
    const r = await platform.saveFile(path, new TextEncoder().encode(json));
    if (r.ok) toast(`Exported ${this.app.settings.modifiedCount} changed setting${this.app.settings.modifiedCount === 1 ? '' : 's'}.`, { kind: 'success' });
    else toast(r.error ?? "Couldn't save the file.", { kind: 'error' });
  }

  private importSettings(): void {
    const input = h('input', { type: 'file', accept: '.json,application/json' }) as HTMLInputElement;
    input.addEventListener('change', async () => {
      const f = input.files?.[0];
      if (!f) return;
      try {
        const n = this.app.settings.import(await f.text());
        toast(`Imported ${n} setting${n === 1 ? '' : 's'}.`, { kind: 'success' });
      } catch (err) {
        toast(`Couldn't import: ${(err as Error).message}`, { kind: 'error' });
      }
    });
    input.click();
  }

  private async resetAll(): Promise<void> {
    if (!this.app.settings.modifiedCount) {
      toast('Everything is already at its default.');
      return;
    }
    if (await confirmDialog('Reset all settings?', `${this.app.settings.modifiedCount} changed settings will go back to their defaults.`, 'Reset all', true)) this.app.settings.resetAll();
  }
}
