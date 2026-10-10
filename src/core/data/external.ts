// Live data for formulas (NFL stats, stock prices, exchange rates).
// Formulas are synchronous, so a function asks for a resource: if it's loaded the data is
// returned; otherwise a download starts, the cell shows #BUSY!, and when the download
// finishes every formula that uses live data is recalculated.
import { ErrorValue } from '../values';

/** Downloads text. maxAgeSec lets the desktop shell serve a cached copy from disk. */
export type Fetcher = (url: string, opts: { maxAgeSec: number; headers?: Record<string, string> }) => Promise<string>;

export class FetchError extends Error {
  constructor(
    message: string,
    readonly status?: number,
  ) {
    super(message);
  }
}

interface Entry {
  status: 'loading' | 'ready' | 'error';
  value?: unknown;
  error?: string;
  notFound?: boolean;
  at: number;
}

export interface Resource<T> {
  key: string;
  url: string;
  /** How long a downloaded copy can be reused (seconds). */
  maxAgeSec: number;
  parse: (text: string) => T;
  headers?: Record<string, string>;
  /** Plain-English description used in messages: "NFL player stats for 2020". */
  label: string;
}

/** Result of a request: the data, or an error value to return from the formula. */
export type Got<T> = { ok: true; value: T } | { ok: false; error: ErrorValue };

const ERROR_RETRY_MS = 60_000;

export class ExternalData {
  private entries = new Map<string, Entry>();
  private listeners = new Set<() => void>();
  private activity = new Set<(pending: number) => void>();
  private notifyTimer: ReturnType<typeof setTimeout> | null = null;
  fetcher: Fetcher | null = null;
  /** Set to force the next requests to bypass caches (Refresh live data). */
  private forceUntil = 0;

  get pending(): number {
    let n = 0;
    for (const e of this.entries.values()) if (e.status === 'loading') n++;
    return n;
  }

  onChange(cb: () => void): () => void {
    this.listeners.add(cb);
    return () => this.listeners.delete(cb);
  }

  /** Called whenever a download starts or finishes (for progress indicators). */
  onActivity(cb: (pending: number) => void): () => void {
    this.activity.add(cb);
    return () => this.activity.delete(cb);
  }

  private emitActivity(): void {
    const n = this.pending;
    for (const a of this.activity) a(n);
  }

  private notify(): void {
    // Batch: many downloads finishing together trigger a single recalculation
    if (this.notifyTimer) return;
    this.notifyTimer = setTimeout(() => {
      this.notifyTimer = null;
      for (const l of this.listeners) l();
    }, 30);
  }

  /** Returns the resource if loaded; otherwise starts loading it and returns #BUSY!. */
  get<T>(res: Resource<T>): Got<T> {
    const e = this.entries.get(res.key);
    if (e?.status === 'ready') return { ok: true, value: e.value as T };
    if (e?.status === 'loading') return { ok: false, error: busy(res.label) };
    if (e?.status === 'error' && Date.now() - e.at < ERROR_RETRY_MS) {
      return { ok: false, error: e.notFound ? notFound(res.label, e.error) : new ErrorValue('#CONNECT!', `Couldn't download ${res.label}: ${e.error}`) };
    }
    if (!this.fetcher) return { ok: false, error: new ErrorValue('#CONNECT!', 'Live data isn\'t available in this window.') };
    this.load(res);
    return { ok: false, error: busy(res.label) };
  }

  /** Starts loading (if needed) and resolves when the resource is ready. Used by tests and panels. */
  async ensure<T>(res: Resource<T>): Promise<T> {
    const e = this.entries.get(res.key);
    if (e?.status === 'ready') return e.value as T;
    await this.load(res);
    const after = this.entries.get(res.key);
    if (after?.status === 'ready') return after.value as T;
    throw new FetchError(after?.error ?? 'download failed');
  }

  peek<T>(key: string): T | undefined {
    const e = this.entries.get(key);
    return e?.status === 'ready' ? (e.value as T) : undefined;
  }

  private inflight = new Map<string, Promise<void>>();

  private load<T>(res: Resource<T>): Promise<void> {
    const existing = this.inflight.get(res.key);
    if (existing) return existing;
    const fetcher = this.fetcher;
    if (!fetcher) return Promise.resolve();
    this.entries.set(res.key, { status: 'loading', at: Date.now() });
    queueMicrotask(() => this.emitActivity());
    const maxAgeSec = Date.now() < this.forceUntil ? 0 : res.maxAgeSec;
    const p = fetcher(res.url, { maxAgeSec, headers: res.headers })
      .then((text) => {
        const value = res.parse(text);
        this.entries.set(res.key, { status: 'ready', value, at: Date.now() });
      })
      .catch((err: unknown) => {
        const status = err instanceof FetchError ? err.status : undefined;
        const msg = err instanceof Error ? err.message : String(err);
        this.entries.set(res.key, { status: 'error', error: msg, notFound: status === 404, at: Date.now() });
      })
      .finally(() => {
        this.inflight.delete(res.key);
        this.emitActivity();
        this.notify();
      });
    this.inflight.set(res.key, p);
    return p;
  }

  /** Drops downloaded data so it's fetched again (skipping the disk cache for a moment). */
  refresh(filter?: (key: string) => boolean): void {
    for (const k of [...this.entries.keys()]) if (!filter || filter(k)) this.entries.delete(k);
    this.forceUntil = Date.now() + 15_000;
    this.notify();
  }

  /** Test helper: forget everything. */
  reset(): void {
    this.entries.clear();
    this.inflight.clear();
  }
}

function busy(label: string): ErrorValue {
  return new ErrorValue('#BUSY!', `Loading ${label}…`);
}

function notFound(label: string, detail?: string): ErrorValue {
  void detail;
  return new ErrorValue('#N/A', `There's no ${label} available (it may not have been published yet).`);
}

/** The shared instance used by all live-data functions. */
export const external = new ExternalData();
