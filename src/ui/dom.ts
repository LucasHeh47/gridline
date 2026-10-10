// Tiny DOM helpers (no framework).
import { createElement as lucideCreate, type IconNode } from 'lucide';

type Child = Node | string | null | undefined | false;
type Props = {
  class?: string;
  style?: Partial<CSSStyleDeclaration> | string;
  title?: string;
  [key: string]: unknown;
};

/** Creates an element: h('div', { class: 'x', onclick: fn }, child1, child2). */
export function h<K extends keyof HTMLElementTagNameMap>(tag: K, props?: Props | null, ...children: (Child | Child[])[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  if (props) {
    for (const [k, v] of Object.entries(props)) {
      if (v === undefined || v === null || v === false) continue;
      if (k === 'class') el.className = String(v);
      else if (k === 'style') {
        if (typeof v === 'string') el.setAttribute('style', v);
        else Object.assign(el.style, v);
      } else if (k.startsWith('on') && typeof v === 'function') {
        el.addEventListener(k.slice(2), v as EventListener);
      } else if (k === 'dataset') {
        Object.assign(el.dataset, v);
      } else if (k in el && typeof v !== 'string') {
        (el as unknown as Record<string, unknown>)[k] = v;
      } else el.setAttribute(k, v === true ? '' : String(v));
    }
  }
  append(el, children);
  return el;
}

export function append(el: Element, children: (Child | Child[])[]): void {
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(typeof c === 'string' ? document.createTextNode(c) : c);
  }
}

export function icon(node: IconNode, size = 16, cls?: string): SVGElement {
  const el = lucideCreate(node, { width: size, height: size, 'stroke-width': 1.75 });
  if (cls) el.setAttribute('class', cls);
  return el;
}

export function clear(el: Element): void {
  while (el.firstChild) el.removeChild(el.firstChild);
}

/** Places a fixed-position element next to an anchor rect, keeping it on screen. */
export function placeNear(el: HTMLElement, rect: { left: number; top: number; right: number; bottom: number }, prefer: 'below' | 'above' | 'right' = 'below', gap = 4): void {
  el.style.left = '0px';
  el.style.top = '0px';
  const w = el.offsetWidth;
  const hgt = el.offsetHeight;
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  let x: number;
  let y: number;
  if (prefer === 'right') {
    x = rect.right + gap;
    y = rect.top;
    if (x + w > vw - 8) x = rect.left - w - gap;
  } else {
    x = rect.left;
    y = prefer === 'below' ? rect.bottom + gap : rect.top - hgt - gap;
    if (prefer === 'below' && y + hgt > vh - 8 && rect.top - hgt - gap > 8) y = rect.top - hgt - gap;
    if (prefer === 'above' && y < 8) y = rect.bottom + gap;
  }
  x = Math.max(8, Math.min(x, vw - w - 8));
  y = Math.max(8, Math.min(y, vh - hgt - 8));
  el.style.left = `${Math.round(x)}px`;
  el.style.top = `${Math.round(y)}px`;
}

export function isMac(): boolean {
  return /Mac/.test(navigator.platform);
}

/** Formats a shortcut like "Ctrl+Shift+K" for display. */
export function shortcutLabel(s: string): string {
  if (isMac()) return s.replace(/Ctrl\+/g, '⌘').replace(/Alt\+/g, '⌥').replace(/Shift\+/g, '⇧');
  return s;
}

export function kbd(s: string): HTMLElement {
  return h('span', { class: 'kbd' }, shortcutLabel(s));
}

export function debounce<T extends (...args: never[]) => void>(fn: T, ms: number): T {
  let t: ReturnType<typeof setTimeout> | undefined;
  return ((...args: Parameters<T>) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  }) as T;
}

export function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
