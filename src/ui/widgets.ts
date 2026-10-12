// Shared UI pieces: popovers, menus, tooltips, dialogs, toasts, color picker.
import { Check, ChevronRight, CircleAlert, CircleCheck, Info, X, type IconNode } from 'lucide';
import { clear, h, icon, kbd, placeNear } from './dom';

// ---------- popover ----------
// Menus, dropdowns and pickers. They close when you press anywhere outside them (the grid too:
// it cancels pointer presses, so listening for mousedown wasn't enough), when the window loses
// focus or resizes, and when something outside them scrolls. A submenu counts as part of its menu.

/** Open popovers, oldest first (a submenu comes after the menu it belongs to). */
const stack: Popover[] = [];
let listening = false;
/** Where keyboard focus goes when a popover closes and there's nowhere better (the grid). */
let focusFallback: (() => void) | null = null;

export function setFocusFallback(fn: () => void): void {
  focusFallback = fn;
}

/** The dropdown a press on its own button just closed (so the click that follows doesn't reopen it). */
let toggled: { anchor: Element; at: number } | null = null;

function closeOutside(target: EventTarget | null): void {
  const t = target instanceof Node ? target : null;
  for (let i = stack.length - 1; i >= 0; i--) {
    const p = stack[i];
    if (t && p.el.contains(t)) break;
    if (t && p.anchor?.contains(t)) {
      // The button that opened it: works as a toggle (a submenu's row just keeps it open)
      if (i === 0) {
        toggled = { anchor: p.anchor, at: Date.now() };
        p.close();
      }
      break;
    }
    p.close();
  }
}

function listen(): void {
  if (listening) return;
  listening = true;
  document.addEventListener('pointerdown', (e) => closeOutside(e.target), true);
  document.addEventListener(
    'wheel',
    (e) => {
      if (stack.length && !stack.some((p) => p.el.contains(e.target as Node))) closeOutside(e.target);
    },
    { capture: true, passive: true },
  );
  document.addEventListener(
    'keydown',
    (e) => {
      if (e.key !== 'Escape' || !stack.length) return;
      // One level at a time: a submenu first, then its menu
      e.stopPropagation();
      e.preventDefault();
      stack[stack.length - 1].close();
    },
    true,
  );
  window.addEventListener('blur', () => closePopovers());
  window.addEventListener('resize', () => closePopovers());
}

export class Popover {
  el: HTMLDivElement;
  onClose?: () => void;
  private closed = false;
  private returnFocus: Element | null = null;

  constructor(
    content: HTMLElement,
    private anchorEl?: HTMLElement,
    cls = '',
  ) {
    this.el = h('div', { class: `popover ${cls}` }, content);
  }

  get anchor(): HTMLElement | undefined {
    return this.anchorEl;
  }

  /** Opens next to an element or at a point/rect. `keepOthers` keeps the menu a submenu belongs to. */
  open(at: HTMLElement | { left: number; top: number; right: number; bottom: number }, prefer: 'below' | 'above' | 'right' = 'below', keepOthers = false): this {
    listen();
    const wasToggled = !keepOthers && !!this.anchorEl && toggled?.anchor === this.anchorEl && Date.now() - toggled.at < 700;
    toggled = null;
    if (wasToggled) {
      // Pressing the button of an open dropdown closes it; don't reopen it on the click
      this.closed = true;
      return this;
    }
    if (!keepOthers) closePopovers();
    hideTooltip();
    this.returnFocus = document.activeElement;
    stack.push(this);
    document.body.classList.add('popover-open');
    document.body.append(this.el);
    const rect = at instanceof HTMLElement ? at.getBoundingClientRect() : at;
    placeNear(this.el, rect, prefer);
    return this;
  }

  reposition(at: HTMLElement | DOMRect, prefer: 'below' | 'above' | 'right' = 'below'): void {
    const rect = at instanceof HTMLElement ? at.getBoundingClientRect() : at;
    placeNear(this.el, rect, prefer);
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    // Submenus opened from this one go too
    const i = stack.indexOf(this);
    if (i >= 0) for (const p of stack.splice(i).slice(1).reverse()) p.close();
    const hadFocus = this.el.contains(document.activeElement);
    this.el.remove();
    if (!stack.length) document.body.classList.remove('popover-open');
    // Keyboard users continue where they were (unless the click moved focus somewhere else)
    if (hadFocus || document.activeElement === document.body) {
      const back = this.returnFocus;
      if (back instanceof HTMLElement && back.isConnected && back !== document.body) back.focus({ preventScroll: true });
      else if (!stack.length) focusFallback?.();
    }
    this.onClose?.();
  }

  get isOpen(): boolean {
    return !this.closed;
  }
}

export function closePopovers(): void {
  for (const p of [...stack].reverse()) p.close();
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
    if (item.submenu) row.dataset.submenu = '1';
    if (item.disabled) row.setAttribute('aria-disabled', 'true');
    rows.push(row);
    list.append(row);
  }
  pop.onClose = () => sub?.close();
  const enabled = (i: number) => !!rows[i] && !rows[i].classList.contains('disabled');
  const step = (from: number, dir: 1 | -1) => {
    for (let k = 1; k <= rows.length; k++) {
      const i = (from + dir * k + rows.length * 2) % rows.length;
      if (enabled(i)) return i;
    }
    return from;
  };
  pop.el.addEventListener('keydown', (e) => {
    const act = () => (rows[focus] as unknown as { _activate?: () => void })?._activate?.();
    if (e.key === 'ArrowDown') setFocus(step(focus < 0 ? -1 : focus, 1));
    else if (e.key === 'ArrowUp') setFocus(step(focus < 0 ? 0 : focus, -1));
    else if (e.key === 'Home') setFocus(step(-1, 1));
    else if (e.key === 'End') setFocus(step(rows.length, -1));
    else if (e.key === 'Enter' || e.key === ' ') act();
    else if (e.key === 'ArrowRight' && rows[focus]?.dataset.submenu) {
      act();
      // Start on the submenu's first item
      sub?.el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home' }));
    } else if (e.key === 'ArrowLeft' && parent) pop.close();
    else if (e.key === 'Tab') {
      pop.close();
      parent?.close();
    } else if (e.key.length === 1 && /\S/.test(e.key) && !e.ctrlKey && !e.metaKey && !e.altKey) {
      // Jump to the next item starting with the typed letter
      const ch = e.key.toLowerCase();
      for (let k = 1; k <= rows.length; k++) {
        const i = (focus + k) % rows.length;
        if (enabled(i) && rows[i].querySelector('.mi-label')?.textContent?.trim().toLowerCase().startsWith(ch)) {
          setFocus(i);
          break;
        }
      }
    } else return;
    e.preventDefault();
    e.stopPropagation();
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
  closePopovers();
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
