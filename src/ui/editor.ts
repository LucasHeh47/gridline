// Cell editing: in-cell editor + formula bar, with highlighting, autocomplete and point-to-reference.
import { formatCell, formatRange, normalize, quoteSheetName, type Range } from '../core/address';
import { isDateFormat } from '../core/numfmt';
import { numberToText } from '../core/formula/coerce';
import { FUNCTIONS, argNames } from '../core/formula/functions';
import { autoFixFormula, parseFormula } from '../core/formula/parser';
import { cycleAbsolute } from '../core/formula/refs';
import { parseInput } from '../core/input';
import { formatValue } from '../core/numfmt';
import { ErrorValue } from '../core/values';
import type { Sheet } from '../core/workbook';
import type { App } from './app';
import { escapeHtml, h } from './dom';
import { argSuggestions } from './argSuggest';
import { canInsertRefAt, functionAtCaret, highlightHtml, identifierAtCaret, refSpans, stringArgAtCaret } from './formulaAssist';
import { toast } from './widgets';

/** A textarea with a syntax-highlighted layer behind it. */
export class HlEditor {
  readonly root: HTMLDivElement;
  readonly backdrop: HTMLDivElement;
  readonly ta: HTMLTextAreaElement;

  constructor(cls = '') {
    this.backdrop = h('div', { class: 'hl-backdrop', 'aria-hidden': 'true' });
    this.ta = h('textarea', { spellcheck: false, autocomplete: 'off', rows: 1 }) as HTMLTextAreaElement;
    this.ta.setAttribute('autocorrect', 'off');
    this.ta.setAttribute('autocapitalize', 'off');
    this.root = h('div', { class: `hl-editor ${cls}` }, this.backdrop, this.ta);
    this.ta.addEventListener('scroll', () => {
      this.backdrop.scrollTop = this.ta.scrollTop;
      this.backdrop.scrollLeft = this.ta.scrollLeft;
    });
  }

  get value(): string {
    return this.ta.value;
  }

  set value(v: string) {
    if (this.ta.value !== v) this.ta.value = v;
    this.render();
  }

  get caret(): number {
    return this.ta.selectionStart ?? this.ta.value.length;
  }

  setCaret(start: number, end = start): void {
    this.ta.setSelectionRange(start, end);
  }

  selectionEndOr(fallback: number): number {
    return this.ta.selectionEnd ?? fallback;
  }

  render(): void {
    const v = this.ta.value;
    const isFormula = v.startsWith('=');
    this.root.classList.toggle('plain', !isFormula);
    if (isFormula) this.backdrop.innerHTML = highlightHtml(v, this.caret);
    this.backdrop.scrollTop = this.ta.scrollTop;
  }

  focus(): void {
    this.ta.focus({ preventScroll: true });
  }
}

interface Session {
  sheetId: number;
  row: number;
  col: number;
  original: string;
  mode: 'enter' | 'edit';
  source: 'cell' | 'bar';
  point: { start: number; end: number; anchor: { row: number; col: number }; cursor: { row: number; col: number } } | null;
}

export type Move = 'down' | 'up' | 'right' | 'left' | 'none';

export class EditController {
  session: Session | null = null;
  readonly cell: HlEditor;
  readonly cellBox: HTMLDivElement;
  bar: HlEditor | null = null;
  private ac: {
    el: HTMLDivElement;
    items: { name: string; desc: string; kind: string }[];
    index: number;
    start: number;
    end: number;
    /** 'string': completing text inside quotes (stat names, teams…); 'func': function/name. */
    mode: 'func' | 'string';
    closeQuote: boolean;
  } | null = null;
  private hint: HTMLDivElement | null = null;
  private errorEl: HTMLDivElement | null = null;
  private measurer: HTMLSpanElement;

  constructor(private app: App) {
    this.cell = new HlEditor();
    this.cellBox = h('div', { class: 'cell-editor' }, this.cell.root);
    this.measurer = h('span', { style: { position: 'absolute', visibility: 'hidden', whiteSpace: 'pre', left: '-9999px', top: '0' } });
    document.body.append(this.measurer);
    this.bind(this.cell, 'cell');
  }

  attachBar(bar: HlEditor): void {
    this.bar = bar;
    this.bind(bar, 'bar');
    bar.ta.addEventListener('focus', () => {
      if (!this.session) this.start({ mode: 'edit', source: 'bar' });
      else this.session.source = 'bar';
    });
  }

  get isOpen(): boolean {
    return !!this.session;
  }

  get text(): string {
    return this.activeEditor().value;
  }

  private activeEditor(): HlEditor {
    return this.session?.source === 'bar' && this.bar ? this.bar : this.cell;
  }

  private otherEditor(): HlEditor | null {
    return this.session?.source === 'bar' ? this.cell : this.bar;
  }

  private bind(ed: HlEditor, source: 'cell' | 'bar'): void {
    ed.ta.addEventListener('input', () => {
      if (!this.session) return;
      this.session.source = source;
      this.session.point = null;
      this.onTextChanged();
    });
    ed.ta.addEventListener('keydown', (e) => this.onKeyDown(e, source));
    ed.ta.addEventListener('click', () => {
      if (!this.session) return;
      if (this.session.source === source) {
        this.session.mode = 'edit';
        this.session.point = null;
        ed.render();
        this.updateAssist();
        this.app.statusbar.setMode('Edit');
      }
    });
    ed.ta.addEventListener('blur', () => setTimeout(() => this.closeAssist(), 150));
  }

  /** The text shown when editing a cell (formulas with "=", numbers unformatted). */
  rawFor(sheet: Sheet, row: number, col: number): string {
    const cell = sheet.getCell(row, col);
    if (!cell) return '';
    if (cell.f !== undefined) return '=' + cell.f;
    const v = cell.v;
    if (v === undefined || v === null) return '';
    if (v instanceof ErrorValue) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    const fmt = this.app.doc.wb.styles.get(sheet.styleIdAt(row, col)).numFmt;
    if (typeof v === 'number') {
      if (fmt && isDateFormat(fmt)) {
        const hasDate = /[dy]/i.test(fmt) || /m{3,}/i.test(fmt);
        const hasTime = /[hs]/i.test(fmt);
        const f = hasDate && hasTime ? 'm/d/yyyy h:mm:ss' : hasDate ? 'm/d/yyyy' : 'h:mm:ss';
        return formatValue(v, f).text.replace(/:00$/, '');
      }
      if (fmt && /%/.test(fmt)) return numberToText(Math.round(v * 100 * 1e10) / 1e10) + '%';
      return numberToText(v);
    }
    // Text that would otherwise be read as a number/date/formula keeps an apostrophe
    const parsed = parseInput(v, fmt);
    if (fmt !== '@' && (parsed.formula !== undefined || parsed.value !== v)) return "'" + v;
    return v;
  }

  start(opts: { mode: 'enter' | 'edit'; initial?: string; source?: 'cell' | 'bar' }): void {
    const app = this.app;
    const sheet = app.sheet;
    const { row, col } = app.sel.active;
    const original = this.rawFor(sheet, row, col);
    this.session = { sheetId: sheet.id, row, col, original, mode: opts.mode, source: opts.source ?? 'cell', point: null };
    const text = opts.initial ?? original;
    this.cell.value = text;
    if (this.bar) this.bar.value = text;
    this.applyCellFont(sheet, row, col);
    this.cellBox.classList.add('open');
    app.grid.scrollToCell(row, col);
    this.position();
    const ed = this.activeEditor();
    ed.focus();
    ed.setCaret(text.length);
    this.onTextChanged();
    app.statusbar.setMode(opts.mode === 'enter' ? 'Enter' : 'Edit');
    app.grid.draw();
  }

  private applyCellFont(sheet: Sheet, row: number, col: number): void {
    const st = this.app.doc.wb.styles.get(sheet.styleIdAt(row, col));
    const z = this.app.grid.zoom;
    const px = (((st.fontSize ?? 11) * 96) / 72) * z;
    this.cellBox.style.setProperty('--editor-font-size', `${px.toFixed(1)}px`);
    this.cellBox.style.setProperty('--editor-line', `${Math.round(px * 1.3)}px`);
    this.cellBox.style.setProperty('--editor-font', `"${st.fontName ?? 'Calibri'}", Calibri, "Segoe UI", sans-serif`);
    this.cell.ta.style.fontWeight = st.bold ? '700' : '';
    this.cell.ta.style.fontStyle = st.italic ? 'italic' : '';
    this.cell.backdrop.style.fontWeight = this.cell.ta.style.fontWeight;
  }

  /** Positions and sizes the in-cell editor over its cell. */
  position(): void {
    const s = this.session;
    if (!s) return;
    const grid = this.app.grid;
    if (s.sheetId !== this.app.sheet.id) {
      this.cellBox.style.display = 'none';
      return;
    }
    this.cellBox.style.display = '';
    const r = grid.cellRect(s.row, s.col);
    const text = this.cell.value;
    const isFormula = text.startsWith('=');
    this.cellBox.classList.toggle('formula', isFormula);
    const cs = getComputedStyle(this.cell.ta);
    this.measurer.style.font = cs.font;
    const lines = text.split('\n');
    let maxW = 0;
    for (const line of lines) {
      this.measurer.textContent = line || ' ';
      maxW = Math.max(maxW, this.measurer.offsetWidth);
    }
    const availW = Math.max(r.w, grid.width - r.x - 4);
    const w = Math.min(availW, Math.max(r.w + 1, maxW + 18));
    const lineH = parseFloat(cs.lineHeight) || 20;
    // Wrapping lines when the text is wider than the available space
    const wraps = maxW + 18 > availW ? Math.ceil((maxW + 18) / availW) : 1;
    const hgt = Math.max(r.h + 1, (lines.length + wraps - 1) * lineH + 4);
    this.cellBox.style.left = `${r.x - 1}px`;
    this.cellBox.style.top = `${r.y - 1}px`;
    this.cellBox.style.width = `${w}px`;
    this.cellBox.style.height = `${Math.min(hgt, grid.height - r.y)}px`;
  }

  private onTextChanged(): void {
    const ed = this.activeEditor();
    const other = this.otherEditor();
    ed.render();
    if (other && other.value !== ed.value) other.value = ed.value;
    this.position();
    this.updateHighlights();
    this.updateAssist();
    this.hideError();
  }

  private updateHighlights(): void {
    const s = this.session;
    const grid = this.app.grid;
    if (!s) {
      grid.refHighlights = [];
      grid.draw();
      return;
    }
    const sheet = this.app.doc.wb.sheetById(s.sheetId)!;
    grid.refHighlights = refSpans(this.text).flatMap((sp) => {
      const target = sp.ref.sheet ? this.app.doc.wb.sheetByName(sp.ref.sheet) : sheet;
      if (!target) return [];
      return [{ sheetId: target.id, range: { r1: sp.ref.r1, c1: sp.ref.c1, r2: sp.ref.r2, c2: sp.ref.c2 }, color: sp.color }];
    });
    grid.draw();
  }

  canInsertRef(): boolean {
    const s = this.session;
    if (!s) return false;
    const ed = this.activeEditor();
    const caret = ed.caret;
    if (s.point && caret === s.point.end) return true;
    return canInsertRefAt(ed.value, caret);
  }

  /** Inserts (or replaces the last pointed) reference at the caret. */
  insertRef(range: Range, sheet: Sheet, extend = false, dragging = false): void {
    const s = this.session;
    if (!s) return;
    const ed = this.activeEditor();
    let r = range;
    if (extend && s.point) r = normalize({ r1: s.point.anchor.row, c1: s.point.anchor.col, r2: range.r2, c2: range.c2 });
    const prefix = sheet.id !== s.sheetId ? quoteSheetName(sheet.name) + '!' : '';
    const refText = prefix + formatRange(r);
    let text = ed.value;
    let start: number;
    if (s.point && (dragging || ed.caret === s.point.end)) {
      start = s.point.start;
      text = text.slice(0, s.point.start) + refText + text.slice(s.point.end);
    } else {
      start = ed.caret;
      text = text.slice(0, start) + refText + text.slice(ed.selectionEndOr(start));
    }
    const end = start + refText.length;
    ed.value = text;
    ed.setCaret(end);
    s.point = {
      start,
      end,
      anchor: extend && s.point ? s.point.anchor : dragging && s.point ? s.point.anchor : { row: r.r1, col: r.c1 },
      cursor: { row: range.r2, col: range.c2 },
    };
    if (!dragging && !extend) s.point.anchor = { row: r.r1, col: r.c1 };
    const other = this.otherEditor();
    if (other) other.value = text;
    ed.render();
    this.position();
    this.updateHighlights();
    this.updateAssist();
    ed.focus();
  }

  private movePoint(dr: number, dc: number, extend: boolean): void {
    const s = this.session!;
    const sheet = this.app.sheet;
    const grid = this.app.grid;
    const cur = s.point?.cursor ?? { row: s.row, col: s.col };
    let row = cur.row;
    let col = cur.col;
    if (dr) row = grid.rows.step(row, dr as 1 | -1);
    if (dc) col = grid.cols.step(col, dc as 1 | -1);
    const anchor = extend && s.point ? s.point.anchor : { row, col };
    const range = normalize({ r1: anchor.row, c1: anchor.col, r2: row, c2: col });
    const hadPoint = !!s.point;
    this.insertRef(range, sheet, false, hadPoint);
    s.point!.anchor = anchor;
    s.point!.cursor = { row, col };
    grid.scrollToCell(row, col);
  }

  private onKeyDown(e: KeyboardEvent, source: 'cell' | 'bar'): void {
    const s = this.session;
    if (!s) return;
    s.source = source;
    const ed = this.activeEditor();
    const key = e.key;
    // Autocomplete navigation
    if (this.ac) {
      if (key === 'ArrowDown' || key === 'ArrowUp') {
        e.preventDefault();
        this.ac.index = (this.ac.index + (key === 'ArrowDown' ? 1 : -1) + this.ac.items.length) % this.ac.items.length;
        this.renderAc();
        return;
      }
      if (key === 'Tab' || (key === 'Enter' && !e.altKey && !e.ctrlKey)) {
        e.preventDefault();
        this.acceptAc();
        return;
      }
      if (key === 'Escape') {
        e.preventDefault();
        this.closeAssist();
        return;
      }
    }
    if (key === 'Enter' && e.altKey) {
      e.preventDefault();
      const c = ed.caret;
      ed.value = ed.value.slice(0, c) + '\n' + ed.value.slice(ed.selectionEndOr(c));
      ed.setCaret(c + 1);
      this.onTextChanged();
      return;
    }
    if (key === 'Enter') {
      e.preventDefault();
      if (e.ctrlKey || e.metaKey) this.commit('none', true);
      else this.commit(e.shiftKey ? 'up' : 'down');
      return;
    }
    if (key === 'Tab') {
      e.preventDefault();
      this.commit(e.shiftKey ? 'left' : 'right');
      return;
    }
    if (key === 'Escape') {
      e.preventDefault();
      this.cancel();
      return;
    }
    if (key === 'F2') {
      e.preventDefault();
      s.mode = s.mode === 'enter' ? 'edit' : 'enter';
      this.app.statusbar.setMode(s.mode === 'enter' ? 'Enter' : 'Edit');
      return;
    }
    if (key === 'F4') {
      e.preventDefault();
      const r = cycleAbsolute(ed.value, ed.caret);
      if (r) {
        ed.value = r.text;
        ed.setCaret(r.caret);
        if (s.point) s.point = null;
        this.onTextChanged();
      }
      return;
    }
    const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[key] && source === 'cell' && s.mode === 'enter' && !e.altKey) {
      const [dr, dc] = arrows[key];
      if (this.canInsertRef()) {
        e.preventDefault();
        this.movePoint(dr, dc, e.shiftKey);
        return;
      }
      if (!ed.value.startsWith('=') || !e.shiftKey) {
        e.preventDefault();
        this.commit(dr > 0 ? 'down' : dr < 0 ? 'up' : dc > 0 ? 'right' : 'left');
        return;
      }
    }
    if ((key === 'ArrowLeft' || key === 'ArrowRight' || key === 'Home' || key === 'End') && s.point) s.point = null;
    // Re-render after caret moves so bracket matching follows
    setTimeout(() => {
      if (!this.session) return;
      ed.render();
      this.updateAssist();
    }, 0);
  }

  /** Commits the edit. Returns false if the formula has an error that must be fixed first. */
  commit(move: Move = 'down', fillSelection = false): boolean {
    const s = this.session;
    if (!s) return true;
    let text = this.text;
    const app = this.app;
    const sheet = app.doc.wb.sheetById(s.sheetId);
    if (!sheet) {
      this.close();
      return true;
    }
    if (text.startsWith('=') && text.length > 1) {
      const parsed = parseFormula(text.slice(1));
      if (parsed.error) {
        const fixed = autoFixFormula(text.slice(1));
        if (fixed) {
          text = '=' + fixed;
          toast('Fixed the formula by adding what was missing (closing parenthesis or quote).');
        } else {
          this.showError(parsed.error.message);
          return false;
        }
      }
    }
    const changed = text !== s.original || fillSelection;
    this.close();
    if (app.sheet.id !== s.sheetId) app.setActiveSheet(s.sheetId, false);
    app.sel.active = { row: s.row, col: s.col };
    if (changed) app.enterInput(sheet, s.row, s.col, text, fillSelection);
    if (move !== 'none') app.moveAfterEnter(move);
    app.focusGrid();
    return true;
  }

  cancel(): void {
    const s = this.session;
    this.close();
    if (s && this.app.sheet.id !== s.sheetId) this.app.setActiveSheet(s.sheetId, false);
    this.app.refreshFormulaBar();
    this.app.focusGrid();
  }

  private close(): void {
    this.session = null;
    this.cellBox.classList.remove('open');
    this.closeAssist();
    this.hideError();
    this.app.grid.refHighlights = [];
    this.app.grid.draw();
    this.app.statusbar.setMode('Ready');
    this.bar?.ta.blur();
  }

  focus(): void {
    this.activeEditor().focus();
  }

  /** Inserts text at the caret of whichever editor is active (used by the function browser). */
  insertText(text: string): void {
    const ed = this.activeEditor();
    const pos = ed.caret;
    const end = ed.selectionEndOr(pos);
    ed.value = ed.value.slice(0, pos) + text + ed.value.slice(end);
    ed.setCaret(pos + text.length);
    ed.focus();
    this.onTextChanged();
  }

  // ---------- assist: autocomplete & hints ----------

  private updateAssist(): void {
    const ed = this.activeEditor();
    const text = ed.value;
    const caret = ed.caret;
    if (!text.startsWith('=') || document.activeElement !== ed.ta) {
      this.closeAssist();
      return;
    }
    // Suggestions for text arguments ("Passing Yards", "Chiefs"…)
    const sa = stringArgAtCaret(text, caret);
    if (sa) {
      const fcq = functionAtCaret(text, sa.quote);
      const kind = fcq ? FUNCTIONS.get(fcq.name)?.suggest?.[fcq.argIndex] : undefined;
      const sugg = kind ? argSuggestions(kind, sa.partial) : [];
      const exact = sugg.length === 1 && sugg[0].value.toLowerCase() === sa.partial.toLowerCase();
      if (sugg.length && !exact) {
        const list = sugg.map((x) => ({ name: x.value, desc: x.desc, kind: x.kind }));
        const keep = this.ac?.items[this.ac.index]?.name;
        const idx = Math.max(0, list.findIndex((i) => i.name === keep));
        const st = { items: list, index: idx, start: sa.quote + 1, end: sa.end, mode: 'string' as const, closeQuote: !sa.closed };
        if (!this.ac) this.ac = { el: h('div', { class: 'ac-list', role: 'listbox' }), ...st };
        else Object.assign(this.ac, st);
        this.renderAc();
        return;
      }
      this.closeAc();
    }
    // Autocomplete
    const id = sa ? null : identifierAtCaret(text, caret);
    if (id && id.word.length >= 1) {
      const q = id.word.toUpperCase();
      const items: { name: string; desc: string; kind: string }[] = [];
      for (const def of FUNCTIONS.values()) {
        if (def.name.startsWith(q)) items.push({ name: def.name, desc: def.desc, kind: def.category });
      }
      items.sort((a, b) => a.name.length - b.name.length || a.name.localeCompare(b.name));
      if (q.length >= 2) {
        for (const def of FUNCTIONS.values()) {
          if (!def.name.startsWith(q) && def.name.includes(q)) items.push({ name: def.name, desc: def.desc, kind: def.category });
        }
      }
      for (const dn of this.app.doc.wb.names.values()) {
        if (dn.name.toUpperCase().startsWith(q)) items.unshift({ name: dn.name, desc: `Named range: ${dn.ref}`, kind: 'Name' });
      }
      const exact = items.length === 1 && items[0].name === q;
      if (items.length && !exact) {
        const keep = this.ac?.items[this.ac.index]?.name;
        const list = items.slice(0, 40);
        const idx = Math.max(0, list.findIndex((i) => i.name === keep));
        const st = { items: list, index: idx, start: id.start, end: id.end, mode: 'func' as const, closeQuote: false };
        if (!this.ac) this.ac = { el: h('div', { class: 'ac-list', role: 'listbox' }), ...st };
        else Object.assign(this.ac, st);
        this.renderAc();
      } else this.closeAc();
    } else this.closeAc();
    // Argument hint
    const fc = functionAtCaret(text, caret);
    const def = fc ? FUNCTIONS.get(fc.name) : undefined;
    if (def && !this.ac) {
      const names = argNames(def.args);
      const variadicIdx = names.indexOf('...');
      let current = fc!.argIndex;
      if (variadicIdx >= 0 && current >= variadicIdx) current = variadicIdx - 1;
      const sig = names
        .map((n, i) => (i === current ? `<b>${escapeHtml(n)}</b>` : escapeHtml(n)))
        .join(', ');
      if (!this.hint) {
        this.hint = h('div', { class: 'arg-hint' });
        document.body.append(this.hint);
      }
      const argName = (names[current] ?? '').replace(/[[\]]/g, '');
      const help = def.argHelp?.[argName];
      this.hint.innerHTML = `<div class="sig">${def.name}(${sig})</div><div class="desc">${escapeHtml(help ? `${argName}: ${help}` : def.desc)}</div>`;
      this.placeAssist(this.hint);
    } else this.closeHint();
  }

  private renderAc(): void {
    const ac = this.ac!;
    ac.el.innerHTML = '';
    ac.items.forEach((it, i) => {
      const row = h(
        'div',
        { class: `ac-item${i === ac.index ? ' sel' : ''}` },
        h('div', { class: 'ac-name' }, it.name, h('span', { class: 'ac-kind' }, it.kind)),
        h('div', { class: 'ac-desc' }, it.desc),
      );
      row.addEventListener('mousedown', (e) => {
        e.preventDefault();
        ac.index = i;
        this.acceptAc();
      });
      ac.el.append(row);
    });
    if (!ac.el.isConnected) document.body.append(ac.el);
    this.placeAssist(ac.el);
    (ac.el.children[ac.index] as HTMLElement | undefined)?.scrollIntoView({ block: 'nearest' });
    this.closeHint();
  }

  private placeAssist(el: HTMLElement): void {
    const ed = this.activeEditor();
    const rect = (this.session?.source === 'bar' ? ed.root.parentElement ?? ed.root : this.cellBox).getBoundingClientRect();
    el.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - el.offsetWidth - 8))}px`;
    let top = rect.bottom + 4;
    if (top + el.offsetHeight > window.innerHeight - 8) top = rect.top - el.offsetHeight - 4;
    el.style.top = `${top}px`;
  }

  private acceptAc(): void {
    const ac = this.ac;
    if (!ac) return;
    const item = ac.items[ac.index];
    const ed = this.activeEditor();
    if (ac.mode === 'string') {
      const text = ed.value;
      const close = ac.closeQuote ? '"' : '';
      ed.value = text.slice(0, ac.start) + item.name + close + text.slice(ac.end);
      // Put the caret after the closing quote
      ed.setCaret(ac.start + item.name.length + 1);
      this.closeAc();
      this.onTextChanged();
      return;
    }
    const isFunc = item.kind !== 'Name';
    const text = ed.value;
    const after = text.slice(ac.end);
    const insert = isFunc && !after.startsWith('(') ? item.name + '(' : item.name;
    ed.value = text.slice(0, ac.start) + insert + after;
    ed.setCaret(ac.start + insert.length);
    this.closeAc();
    this.onTextChanged();
  }

  private closeAc(): void {
    this.ac?.el.remove();
    this.ac = null;
  }

  private closeHint(): void {
    this.hint?.remove();
    this.hint = null;
  }

  private closeAssist(): void {
    this.closeAc();
    this.closeHint();
  }

  private showError(message: string): void {
    this.hideError();
    const ed = this.activeEditor();
    const rect = (this.session?.source === 'bar' ? ed.root.parentElement ?? ed.root : this.cellBox).getBoundingClientRect();
    this.errorEl = h('div', { class: 'edit-error', style: { position: 'fixed' } }, '⚠ ', message, h('div', { class: 'hint', style: { marginTop: '4px' } }, 'Fix the formula, or press Esc to cancel.'));
    document.body.append(this.errorEl);
    this.errorEl.style.left = `${rect.left}px`;
    this.errorEl.style.top = `${rect.bottom + 6}px`;
    ed.focus();
  }

  private hideError(): void {
    this.errorEl?.remove();
    this.errorEl = null;
  }

  /** Used by the formula bar to display a cell when not editing. */
  describeActive(): string {
    const a = this.app.sel.active;
    return this.rawFor(this.app.sheet, a.row, a.col);
  }

  cellLabel(): string {
    const s = this.session;
    return s ? formatCell(s.row, s.col) : '';
  }
}
