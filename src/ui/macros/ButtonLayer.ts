// Buttons on the sheet that run macros. Click runs; drag moves; the corner handle resizes;
// right-click edits or deletes.
import * as L from 'lucide';
import type { SheetButton } from '../../core/macros/project';
import type { App } from '../app';
import { h, icon } from '../dom';
import { promptDialog, showMenu, type MenuItem } from '../widgets';

const COLORS: [string, string][] = [
  ['', 'Accent (default)'],
  ['#16a34a', 'Green'],
  ['#dc2626', 'Red'],
  ['#ea580c', 'Orange'],
  ['#7c3aed', 'Purple'],
  ['#475569', 'Gray'],
];

export class ButtonLayer {
  readonly el: HTMLDivElement;
  private els = new Map<string, HTMLButtonElement>();

  constructor(private app: App) {
    this.el = h('div', { class: 'button-layer', style: { position: 'absolute', left: '0', top: '0', width: '0', height: '0', overflow: 'visible' } });
  }

  /** Creates/removes button elements to match the active sheet. */
  sync(): void {
    const buttons = this.app.sheet.buttons;
    const ids = new Set(buttons.map((b) => b.id));
    for (const [id, el] of this.els) {
      if (!ids.has(id)) {
        el.remove();
        this.els.delete(id);
      }
    }
    for (const b of buttons) {
      let el = this.els.get(b.id);
      if (!el) {
        el = this.create(b.id);
        this.els.set(b.id, el);
        this.el.append(el);
      }
      el.querySelector('.mb-label')!.textContent = b.label;
      el.style.setProperty('--btn-color', b.color || 'var(--accent)');
      el.title = `Runs "${b.macro}"`;
    }
    this.position();
  }

  private def(id: string): SheetButton | undefined {
    return this.app.sheet.buttons.find((b) => b.id === id);
  }

  position(): void {
    const g = this.app.grid;
    if (!g.rows) return;
    const z = g.zoom;
    const sheet = this.app.sheet;
    const clipX = g.headerW + g.frozenW;
    const clipY = g.headerH + g.frozenH;
    for (const [id, el] of this.els) {
      const b = this.def(id);
      if (!b) continue;
      const a = b.anchor;
      const x = g.colX(a.col) + a.dx * z;
      const y = g.rowY(a.row) + a.dy * z;
      const w = a.w * z;
      const hgt = a.h * z;
      Object.assign(el.style, { left: `${x}px`, top: `${y}px`, width: `${w}px`, height: `${hgt}px`, fontSize: `${13 * z}px` });
      const hidden = (a.col >= sheet.freeze.cols && x + w < clipX) || (a.row >= sheet.freeze.rows && y + hgt < clipY) || x > g.width || y > g.height;
      el.style.display = hidden ? 'none' : '';
      const cl = a.col >= sheet.freeze.cols ? Math.max(0, clipX - x) : 0;
      const ct = a.row >= sheet.freeze.rows ? Math.max(0, clipY - y) : 0;
      el.style.clipPath = cl > 0 || ct > 0 ? `inset(${ct}px 0 0 ${cl}px)` : '';
    }
  }

  private create(id: string): HTMLButtonElement {
    const grip = h('span', { class: 'mb-grip', 'aria-hidden': 'true' });
    const el = h('button', { class: 'macro-button' }, icon(L.Play, 12), h('span', { class: 'mb-label' }), grip) as HTMLButtonElement;
    el.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      this.drag(e, id, e.target === grip ? 'resize' : 'move');
    });
    el.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      e.stopPropagation();
      showMenu(this.menu(id), { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY });
    });
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        const b = this.def(id);
        if (b) void this.app.macros.run(b.macro);
      }
    });
    return el;
  }

  /** Drags past a few pixels move/resize the button; a plain click runs its macro. */
  private drag(e: PointerEvent, id: string, mode: 'move' | 'resize'): void {
    const b = this.def(id);
    if (!b) return;
    const g = this.app.grid;
    const z = g.zoom;
    const el = this.els.get(id)!;
    const start = { x: e.clientX, y: e.clientY, left: el.offsetLeft, top: el.offsetTop, w: b.anchor.w, h: b.anchor.h };
    let moved = false;
    const move = (ev: PointerEvent) => {
      const dx = ev.clientX - start.x;
      const dy = ev.clientY - start.y;
      if (!moved && Math.hypot(dx, dy) < 4) return;
      moved = true;
      if (mode === 'move') {
        el.style.left = `${start.left + dx}px`;
        el.style.top = `${start.top + dy}px`;
      } else {
        el.style.width = `${Math.max(48, start.w * z + dx)}px`;
        el.style.height = `${Math.max(22, start.h * z + dy)}px`;
      }
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (!moved) {
        void this.app.macros.run(b.macro);
        return;
      }
      const sheet = this.app.sheet;
      const next = { ...b.anchor };
      if (mode === 'move') {
        const row = g.rowAt(el.offsetTop + 1);
        const col = g.colAt(el.offsetLeft + 1);
        next.row = row;
        next.col = col;
        next.dx = Math.max(0, Math.round((el.offsetLeft - g.colX(col)) / z));
        next.dy = Math.max(0, Math.round((el.offsetTop - g.rowY(row)) / z));
      } else {
        next.w = Math.round(el.offsetWidth / z);
        next.h = Math.round(el.offsetHeight / z);
      }
      this.update(sheet.buttons.indexOf(b), { ...b, anchor: next }, mode === 'move' ? 'Move button' : 'Resize button');
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  private update(index: number, next: SheetButton | null, label: string): void {
    const sheet = this.app.sheet;
    const list = [...sheet.buttons];
    if (index < 0) return;
    if (next) list[index] = next;
    else list.splice(index, 1);
    this.app.transact(label, (tx) => tx.setProp(sheet, 'buttons', list));
  }

  private menu(id: string): MenuItem[] {
    const b = this.def(id)!;
    const sheet = this.app.sheet;
    const idx = sheet.buttons.indexOf(b);
    return [
      { label: `Run "${b.macro}"`, icon: L.Play, run: () => void this.app.macros.run(b.macro) },
      { separator: true },
      {
        label: 'Change text…',
        icon: L.Type,
        run: async () => {
          const t = await promptDialog('Button text', 'Text', b.label, (v) => (v.trim() ? null : 'The button needs some text.'));
          if (t) this.update(idx, { ...b, label: t.trim() }, 'Change button text');
        },
      },
      {
        label: 'Run a different macro…',
        icon: L.Repeat,
        run: async () => {
          const m = await this.app.macros.pickMacro('Which macro should the button run?', b.macro);
          if (m) this.update(idx, { ...b, macro: m }, 'Change button macro');
        },
      },
      { label: 'Color', icon: L.Palette, submenu: () => COLORS.map(([c, name]): MenuItem => ({ label: name, swatch: c || undefined, checked: (b.color ?? '') === c, run: () => this.update(idx, { ...b, color: c || undefined }, 'Change button color') })) },
      { separator: true },
      { label: 'Delete button', icon: L.Trash2, danger: true, run: () => this.update(idx, null, 'Delete button') },
    ];
  }
}
