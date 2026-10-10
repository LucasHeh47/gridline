// User settings: definitions live in schema.ts; values are saved in localStorage and shared by
// every window (a change in one window is picked up by the others).
import type { IconNode } from 'lucide';
import type { App } from '../app';

export type SettingValue = boolean | number | string | string[];

export interface SettingOption {
  value: string;
  label: string;
}

/** Parts of the UI to update after a setting changes. */
export type Refresh = 'theme' | 'chrome' | 'grid' | 'layout' | 'ribbon' | 'titlebar' | 'status' | 'charts' | 'timers' | 'live';

export interface SettingDef {
  id: string;
  section: string;
  group: string;
  label: string;
  description?: string;
  type: 'toggle' | 'select' | 'number' | 'text' | 'multi' | 'action' | 'info';
  default?: SettingValue;
  options?: SettingOption[] | (() => SettingOption[]);
  min?: number;
  max?: number;
  step?: number;
  unit?: string;
  /** Extra words for search. */
  keywords?: string;
  /** Only does something in the desktop app. */
  desktopOnly?: boolean;
  /** UI parts to refresh after a change. */
  refresh?: Refresh[];
  /** Applies the value. Runs once at startup and after every change. */
  apply?: (value: never, app: App) => void;
  /** 'action' rows: a button. */
  action?: { label: string; run: (app: App) => unknown };
  /** 'info' rows: text shown on the right (may load asynchronously). */
  info?: (app: App) => string | Promise<string>;
}

export interface SectionDef {
  id: string;
  label: string;
  icon: IconNode;
  description: string;
}

const KEY = 'gridline.settings';

export class Settings {
  private values: Record<string, SettingValue> = {};
  private listeners = new Set<(id: string) => void>();
  readonly byId = new Map<string, SettingDef>();

  constructor(readonly defs: SettingDef[]) {
    for (const d of defs) this.byId.set(d.id, d);
    this.values = this.read();
    // Another window changed a setting: adopt it
    window.addEventListener('storage', (e) => {
      if (e.key !== KEY) return;
      const next = this.read();
      const changed = new Set([...Object.keys(next), ...Object.keys(this.values)].filter((k) => JSON.stringify(next[k]) !== JSON.stringify(this.values[k])));
      this.values = next;
      for (const id of changed) this.emit(id);
    });
  }

  private read(): Record<string, SettingValue> {
    try {
      const raw = localStorage.getItem(KEY);
      if (raw) {
        const v = JSON.parse(raw) as unknown;
        if (v && typeof v === 'object' && !Array.isArray(v)) return v as Record<string, SettingValue>;
      }
    } catch {
      // unreadable settings fall back to defaults
    }
    return {};
  }

  private save(): void {
    try {
      localStorage.setItem(KEY, JSON.stringify(this.values));
    } catch {
      // storage may be unavailable (private mode); settings still work for this session
    }
  }

  get<T extends SettingValue>(id: string): T {
    const d = this.byId.get(id);
    const v = id in this.values ? this.values[id] : d?.default;
    return (this.valid(d, v) ? v : d?.default) as T;
  }

  bool(id: string): boolean {
    return this.get<boolean>(id) === true;
  }

  num(id: string): number {
    return Number(this.get<number>(id));
  }

  str(id: string): string {
    return String(this.get<string>(id) ?? '');
  }

  list(id: string): string[] {
    const v = this.get<string[]>(id);
    return Array.isArray(v) ? v : [];
  }

  options(d: SettingDef): SettingOption[] {
    return typeof d.options === 'function' ? d.options() : d.options ?? [];
  }

  private valid(d: SettingDef | undefined, v: unknown): boolean {
    if (!d || v === undefined) return false;
    switch (d.type) {
      case 'toggle':
        return typeof v === 'boolean';
      case 'number':
        return typeof v === 'number' && Number.isFinite(v) && (d.min === undefined || v >= d.min) && (d.max === undefined || v <= d.max);
      case 'select':
        return typeof v === 'string' && this.options(d).some((o) => o.value === v);
      case 'multi':
        return Array.isArray(v) && v.every((x) => typeof x === 'string');
      case 'text':
        return typeof v === 'string';
      default:
        return false;
    }
  }

  set(id: string, value: SettingValue): void {
    const d = this.byId.get(id);
    if (!d || !this.valid(d, value)) return;
    if (JSON.stringify(value) === JSON.stringify(d.default)) delete this.values[id];
    else this.values[id] = value;
    this.save();
    this.emit(id);
  }

  /** A stored value as saved, even for settings that no longer exist (for migrations). */
  raw(id: string): SettingValue | undefined {
    return this.values[id];
  }

  /** Forgets a stored value that no longer has a setting. */
  drop(id: string): void {
    if (!(id in this.values) || this.byId.has(id)) return;
    delete this.values[id];
    this.save();
  }

  isModified(id: string): boolean {
    return id in this.values;
  }

  get modifiedCount(): number {
    return Object.keys(this.values).filter((k) => this.byId.has(k)).length;
  }

  reset(id: string): void {
    if (!(id in this.values)) return;
    delete this.values[id];
    this.save();
    this.emit(id);
  }

  resetAll(): void {
    const ids = Object.keys(this.values);
    this.values = {};
    this.save();
    for (const id of ids) this.emit(id);
  }

  /** Settings that differ from the defaults, as JSON. */
  export(): string {
    return JSON.stringify({ gridlineSettings: 1, values: this.values }, null, 2);
  }

  /** Loads exported settings. Returns how many were applied. */
  import(json: string): number {
    const data = JSON.parse(json) as { values?: Record<string, SettingValue> };
    if (!data || typeof data !== 'object' || !data.values) throw new Error("This file doesn't contain Gridline settings.");
    let n = 0;
    for (const [id, v] of Object.entries(data.values)) {
      if (this.valid(this.byId.get(id), v)) {
        this.values[id] = v;
        n++;
        this.emit(id);
      }
    }
    this.save();
    return n;
  }

  onChange(cb: (id: string) => void): () => void {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  }

  private emit(id: string): void {
    for (const l of this.listeners) l(id);
  }
}
