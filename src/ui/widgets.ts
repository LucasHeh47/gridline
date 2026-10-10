// Shared UI pieces: popovers, menus, tooltips, dialogs, toasts, color picker.
import { Check, ChevronRight, CircleAlert, CircleCheck, Info, X, type IconNode } from 'lucide';
import { clear, h, icon, kbd, placeNear } from './dom';

// ---------- popover ----------

let openPopover: Popover | null = null;

export class Popover {
  el: HTMLDivElement;
  private onDoc = (e: MouseEvent) => {
    if (!this.el.contains(e.target as Node) && !(this.anchorEl && this.anchorEl.contains(e.target as Node))) this.close();
  };
  private onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      e.preventDefault();
      this.close();
    }
  };
  onClose?: () => void;
  private closed = false;

  constructor(
    content: HTMLElement,
    private anchorEl?: HTMLElement,
    cls = '',
  ) {
    this.el = h('div', { class: `popover ${cls}` }, content);
  }

  /** Opens next to an element or at a point/rect. */
  open(at: HTMLElement | { left: number; top: number; right: number; bottom: number }, prefer: 'below' | 'above' | 'right' = 'below', keepOthers = false): this {
    if (!keepOthers && openPopover && openPopover !== this) openPopover.close();
    if (!keepOthers) openPopover = this;
    document.body.append(this.el);
    const rect = at instanceof HTMLElement ? at.getBoundingClientRect() : at;
    placeNear(this.el, rect, prefer);
    setTimeout(() => {
      document.addEventListener('mousedown', this.onDoc, true);
      document.addEventListener('keydown', this.onKey, true);
    }, 0);
    return this;
  }

  reposition(at: HTMLElement | DOMRect, prefer: 'below' | 'above' | 'right' = 'below'): void {
    const rect = at instanceof HTMLElement ? at.getBoundingClientRect() : at;
    placeNear(this.el, rect, prefer);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    document.removeEventListener('mousedown', this.onDoc, true);
    document.removeEventListener('keydown', this.onKey, true);
    this.el.remove();
    if (openPopover === this) openPopover = null;
    this.onClose?.();
  }

  get isOpen(): boolean {
    return !this.closed;
  }
}

export function closePopovers(): void {
  openPopover?.close();
}

// ---------- menu ----------

export type MenuItem =
  | {
      label: string;
      icon?: IconNode;
      hint?: string;
      checked?: boolean;
      disabled?: boolean;
      danger?: boolean;
      run?: () => void;
      submenu?: () => MenuItem[];
      swatch?: string;
    }
  | { separator: true }
  | { title: string };

export function showMenu(items: MenuItem[], at: HTMLElement | { left: number; top: number; right: number; bottom: number }, prefer: 'below' | 'right' = 'below', parent?: Popover): Popover {
  const list = h('div', { class: 'menu', role: 'menu' });
  const pop = new Popover(list, at instanceof HTMLElement ? at : undefined);
  let sub: Popover | null = null;
  const rows: HTMLElement[] = [];
  let focus = -1;
  const setFocus = (i: number) => {
    rows.forEach((r, k) => r.classList.toggle('focus', k === i));
    focus = i;
    rows[i]?.scrollIntoView({ block: 'nearest' });
  };
  for (const item of items) {
    if ('separator' in item) {
      list.append(h('div', { class: 'menu-sep' }));
      continue;
    }
    if ('title' in item) {
      list.append(h('div', { class: 'menu-title' }, item.title));
      continue;
    }
    const iconEl = h('span', { class: 'mi-icon' });
    if (item.checked) iconEl.append(icon(Check, 16));
    else if (item.swatch) iconEl.append(h('span', { style: { width: '14px', height: '14px', borderRadius: '3px', background: item.swatch, border: '1px solid rgba(127,127,127,.4)' } }));
    else if (item.icon) iconEl.append(icon(item.icon, 16));
    const row = h(
      'div',
      { class: `menu-item${item.disabled ? ' disabled' : ''}${item.danger ? ' danger' : ''}`, role: 'menuitem' },
      iconEl,
      h('span', { class: 'mi-label' }, item.label),
      item.hint ? h('span', { class: 'mi-hint' }, item.hint) : null,
      item.submenu ? icon(ChevronRight, 14) : null,
    );
    const activate = () => {
      if (item.submenu) {
        sub?.close();
        sub = showMenu(item.submenu(), row, 'right', pop);
        return;
      }
      pop.close();
      parent?.close();
      item.run?.();
    };
    row.addEventListener('click', activate);
    row.addEventListener('mouseenter', () => {
      setFocus(rows.indexOf(row));
      if (item.submenu) activate();
      else if (sub) {
        sub.close();
        sub = null;
      }
    });
    (row as unknown as { _activate: () => void })._activate = activate;
    rows.push(row);
    list.append(row);
  }
  pop.onClose = () => sub?.close();
  pop.el.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') setFocus(Math.min(rows.length - 1, focus + 1));
    else if (e.key === 'ArrowUp') setFocus(Math.max(0, focus - 1));
    else if (e.key === 'Enter') (rows[focus] as unknown as { _activate?: () => void })?._activate?.();
    else return;
    e.preventDefault();
  });
  pop.el.tabIndex = -1;
  pop.open(at, prefer, !!parent);
  pop.el.focus();
  return pop;
}

// ---------- tooltip ----------

/** Options set from Settings. */
export const uiOptions = {
  tooltips: true,
  tooltipDelay: 450,
  tooltipShortcuts: true,
  toastMs: 4500,
};

let tipEl: HTMLDivElement | null = null;
let tipTimer: ReturnType<typeof setTimeout> | undefined;

/** Adds a hover tooltip (with optional shortcut) to an element. */
export function tooltip(el: HTMLElement, text: string | (() => string), shortcut?: string): void {
  el.addEventListener('mouseenter', () => {
    clearTimeout(tipTimer);
    if (!uiOptions.tooltips) return;
    tipTimer = setTimeout(() => {
      hideTooltip();
      if (!el.isConnected) return;
      const t = typeof text === 'function' ? text() : text;
      tipEl = h('div', { class: 'tooltip' }, t, shortcut && uiOptions.tooltipShortcuts ? h('span', { class: 'tt-k' }, kbd(shortcut)) : null);
      document.body.append(tipEl);
      placeNear(tipEl, el.getBoundingClientRect(), 'below', 6);
    }, uiOptions.tooltipDelay);
  });
  el.addEventListener('mouseleave', hideTooltip);
  el.addEventListener('mousedown', hideTooltip);
}

export function hideTooltip(): void {
  clearTimeout(tipTimer);
  tipEl?.remove();
  tipEl = null;
}

// ---------- toast ----------

let toastHost: HTMLDivElement | null = null;

export function toast(message: string, opts: { kind?: 'info' | 'success' | 'error'; action?: { label: string; run: () => void }; ms?: number } = {}): void {
  if (!toastHost) {
    toastHost = h('div', { class: 'toasts' });
    document.body.append(toastHost);
  }
  const kind = opts.kind ?? 'info';
  const ic = kind === 'error' ? CircleAlert : kind === 'success' ? CircleCheck : Info;
  const el = h(
    'div',
    { class: `toast ${kind}`, role: 'status' },
    h('span', { class: 't-icon' }, icon(ic, 16)),
    h('span', { class: 't-msg' }, message),
  );
  const close = () => el.remove();
  if (opts.action) {
    el.append(
      h('button', {
        class: 'btn text',
        onclick: () => {
          opts.action!.run();
          close();
        },
      }, opts.action.label),
    );
  }
  el.append(h('button', { class: 'btn', onclick: close, 'aria-label': 'Dismiss' }, icon(X, 14)));
  toastHost.append(el);
  while (toastHost.children.length > 3) toastHost.firstChild?.remove();
  setTimeout(close, opts.ms ?? (kind === 'error' ? uiOptions.toastMs + 3000 : uiOptions.toastMs));
}

// ---------- dialogs ----------

export interface DialogButton<T> {
  label: string;
  value: T;
  primary?: boolean;
  danger?: boolean;
}

export function dialog<T>(title: string, body: string | HTMLElement, buttons: DialogButton<T>[], opts: { wide?: boolean; cancelValue?: T } = {}): Promise<T> {
  return new Promise((resolve) => {
    const prevFocus = document.activeElement as HTMLElement | null;
    const back = h('div', { class: 'modal-backdrop' });
    const done = (v: T) => {
      back.remove();
      document.removeEventListener('keydown', onKey, true);
      prevFocus?.focus?.();
      resolve(v);
    };
    const btns = buttons.map((b) =>
      h('button', { class: `btn ${b.primary ? 'primary' : b.danger ? 'outline danger' : 'outline'}`, onclick: () => done(b.value) }, b.label),
    );
    const modal = h(
      'div',
      { class: `modal${opts.wide ? ' wide' : ''}`, role: 'dialog', 'aria-modal': 'true' },
      h('h3', null, title),
      typeof body === 'string' ? h('p', null, body) : body,
      h('div', { class: 'actions' }, btns),
    );
    back.append(modal);
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        done(opts.cancelValue ?? buttons[buttons.length - 1].value);
      } else if (e.key === 'Enter' && !(e.target instanceof HTMLTextAreaElement)) {
        const p = buttons.find((b) => b.primary);
        if (p) {
          e.preventDefault();
          e.stopPropagation();
          done(p.value);
        }
      }
    };
    document.addEventListener('keydown', onKey, true);
    back.addEventListener('mousedown', (e) => {
      if (e.target === back) done(opts.cancelValue ?? buttons[buttons.length - 1].value);
    });
    document.body.append(back);
    (btns.find((_, i) => buttons[i].primary) ?? btns[0])?.focus();
  });
}

export function confirmDialog(title: string, message: string, okLabel = 'OK', danger = false): Promise<boolean> {
  return dialog(title, message, [
    { label: okLabel, value: true, primary: !danger, danger },
    { label: 'Cancel', value: false },
  ]);
}

export function promptDialog(title: string, label: string, value = '', validate?: (v: string) => string | null): Promise<string | null> {
  return new Promise((resolve) => {
    const input = h('input', { class: 'input', value, style: { width: '100%' } }) as HTMLInputElement;
    const err = h('div', { class: 'hint', style: { color: 'var(--danger)', minHeight: '18px', marginTop: '6px' } });
    const body = h('div', null, h('div', { class: 'field' }, h('label', null, label), input), err);
    let result: string | null = null;
    const p = dialog(title, body, [
      { label: 'OK', value: 'ok', primary: true },
      { label: 'Cancel', value: 'cancel' },
    ]);
    // Intercept OK to validate
    const modal = body.parentElement!;
    const okBtn = modal.querySelector('.btn.primary') as HTMLButtonElement;
    const check = () => {
      const msg = validate?.(input.value) ?? null;
      err.textContent = msg ?? '';
      okBtn.disabled = !!msg;
      return !msg;
    };
    input.addEventListener('input', check);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !check()) {
        e.stopPropagation();
        e.preventDefault();
      }
    });
    setTimeout(() => {
      input.focus();
      input.select();
      check();
    }, 0);
    p.then((v) => {
      result = v === 'ok' ? input.value : null;
      resolve(result);
    });
  });
}

// ---------- color picker ----------

export const PALETTE = [
  // greys
  '#ffffff', '#e8eaed', '#c4c8d0', '#9aa1ad', '#6b7280', '#4b5260', '#363b46', '#262a31', '#1b1e24', '#000000',
  // vivid
  '#ff6b6b', '#ff9f43', '#ffd43b', '#94d82d', '#3fcf8e', '#22b8cf', '#4ea1ff', '#748ffc', '#a78bfa', '#e8618c',
  // deep (good fills on dark)
  '#5c2b2b', '#5e3a1a', '#5a4c12', '#33491b', '#174a37', '#124650', '#173e66', '#2b3470', '#3f2d6b', '#5a2342',
  // soft (good fills on light)
  '#ffd6d6', '#ffe3c4', '#fff3b0', '#e3f5c4', '#c9f2df', '#c5eff5', '#d2e6ff', '#dfe3ff', '#ebe0ff', '#ffd9e8',
];

export function colorPicker(onPick: (color: string | null) => void, opts: { allowNone?: boolean; noneLabel?: string; current?: string } = {}): HTMLElement {
  const grid = h('div', { class: 'color-grid' });
  for (const c of PALETTE) {
    const cell = h('div', { class: 'color-cell', style: { background: c }, title: c });
    if (opts.current?.toLowerCase() === c) cell.style.boxShadow = '0 0 0 2px var(--accent)';
    cell.addEventListener('click', () => onPick(c));
    grid.append(cell);
  }
  const custom = h('input', { type: 'color', value: opts.current ?? '#4ea1ff', style: { width: '32px', height: '26px', border: '0', padding: '0', background: 'transparent', cursor: 'pointer' } }) as HTMLInputElement;
  custom.addEventListener('change', () => onPick(custom.value));
  const actions = h(
    'div',
    { class: 'color-actions' },
    opts.allowNone !== false ? h('button', { class: 'btn outline', onclick: () => onPick(null) }, opts.noneLabel ?? 'Automatic') : null,
    h('span', { class: 'hint', style: { marginLeft: 'auto' } }, 'Custom'),
    custom,
  );
  return h('div', null, grid, actions);
}

export function replaceChildren(el: HTMLElement, ...children: (Node | string)[]): void {
  clear(el);
  el.append(...children);
}
