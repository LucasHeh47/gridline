// A document = workbook + calculation engine + undo history.
import { external } from './data/external';
import { Engine } from './formula/engine';
import { History, Transaction } from './history';
import { Workbook, type Sheet } from './workbook';

export type DocEvent =
  | { type: 'change'; tx: Transaction; direction: 'do' | 'undo' | 'redo' }
  | { type: 'reset' }
  | { type: 'saved' }
  /** Downloaded data arrived and formulas were recalculated. */
  | { type: 'data' }
  /** Something outside the undo history changed (macros), or a running macro wants a repaint. */
  | { type: 'meta' };

export class Doc {
  wb: Workbook;
  engine: Engine;
  history = new History();
  filePath: string | null = null;
  fileName = 'Untitled';
  /** Format the document was loaded from / last saved as. */
  fileType: 'xlsx' | 'xlsm' | 'csv' | 'tsv' | null = null;
  private savedVersion = 0;
  private listeners = new Set<(e: DocEvent) => void>();
  /** Hook for the UI to attach selection state to transactions. */
  selectionProvider: (() => Transaction['selectionBefore']) | null = null;

  constructor(wb?: Workbook) {
    this.wb = wb ?? new Workbook();
    this.engine = new Engine(this.wb);
    this.engine.rebuild();
    this.unsubscribe = external.onChange(() => {
      if (this.engine.recalcExternal()) this.emit({ type: 'data' });
    });
  }

  private unsubscribe: () => void;

  /** Stops listening for live data (for short-lived documents such as template builders). */
  dispose(): void {
    this.unsubscribe();
    this.listeners.clear();
  }

  /** The number format a cell shows: its own, or the one a table formula (AMORTIZATION, NFL.SCHEDULE…) suggests. */
  numFmtAt(sheet: Sheet, row: number, col: number): string | undefined {
    const own = this.wb.styles.get(sheet.styleIdAt(row, col)).numFmt;
    if (own && own !== 'General') return own;
    return this.engine.spillFormatAt(sheet, row, col) ?? own;
  }

  get dirty(): boolean {
    return this.history.version !== this.savedVersion;
  }

  markSaved(): void {
    this.savedVersion = this.history.version;
    this.emit({ type: 'saved' });
  }

  /** Replaces the whole workbook (after opening a file). */
  load(wb: Workbook): void {
    this.wb = wb;
    this.engine = new Engine(wb);
    this.engine.rebuild();
    this.history.clear();
    this.history.version = 0;
    this.savedVersion = 0;
    this.emit({ type: 'reset' });
  }

  on(fn: (e: DocEvent) => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  private emit(e: DocEvent): void {
    for (const l of this.listeners) l(e);
  }

  /** Marks the document as changed for edits that aren't undoable (editing macro code). */
  touch(): void {
    this.history.version++;
    this.emit({ type: 'meta' });
  }

  /** Starts an edit made in many steps over time (a running macro) that undoes as one action. */
  beginLive(label: string): LiveEdit {
    return new LiveEdit(this, label, (e) => this.emit(e), (tx) => this.syncEngine(tx));
  }

  /**
   * Applies changes, lets `inspect` look at the result (with formulas recalculated), then reverts
   * everything. Nothing is recorded and no events fire, so the screen never shows the trial.
   */
  dryRun<T>(fn: (tx: Transaction) => void, inspect: (tx: Transaction) => T): T {
    const tx = new Transaction('preview', this.wb);
    const activeSheetId = this.wb.activeSheetId;
    try {
      fn(tx);
      this.syncEngine(tx);
      return inspect(tx);
    } finally {
      tx.replay(false);
      this.wb.activeSheetId = activeSheetId;
      this.syncEngine(tx);
    }
  }

  /** Runs `fn` as a single undoable action. */
  transact(label: string, fn: (tx: Transaction) => void): Transaction {
    const tx = new Transaction(label, this.wb);
    tx.selectionBefore = this.selectionProvider?.() ?? undefined;
    try {
      fn(tx);
    } catch (e) {
      // Roll back partial changes so the document stays consistent
      tx.replay(false);
      this.syncEngine(tx);
      throw e;
    }
    if (tx.isEmpty) return tx;
    tx.selectionAfter = this.selectionProvider?.() ?? undefined;
    this.syncEngine(tx);
    this.history.push(tx);
    this.emit({ type: 'change', tx, direction: 'do' });
    return tx;
  }

  /** Lets the UI update the "after" selection once it has moved the cursor. */
  setSelectionAfter(tx: Transaction): void {
    tx.selectionAfter = this.selectionProvider?.() ?? tx.selectionAfter;
  }

  undo(): Transaction | undefined {
    const tx = this.history.popUndo();
    if (!tx) return;
    tx.replay(false);
    this.syncEngine(tx);
    this.emit({ type: 'change', tx, direction: 'undo' });
    return tx;
  }

  redo(): Transaction | undefined {
    const tx = this.history.popRedo();
    if (!tx) return;
    tx.replay(true);
    this.syncEngine(tx);
    this.emit({ type: 'change', tx, direction: 'redo' });
    return tx;
  }

  /** Recalculates NOW(), TODAY(), RAND() and what depends on them. */
  recalcVolatile(): void {
    this.engine.recalc();
    this.emit({ type: 'data' });
  }

  /** Forces a full recalculation (F9). */
  recalcAll(): void {
    this.engine.rebuild();
    this.emit({ type: 'reset' });
  }

  private syncEngine(tx: Transaction): void {
    if (tx.structural) this.engine.rebuild();
    else {
      const touched = tx.touchedCells();
      if (touched.length) this.engine.update(touched);
    }
  }
}

/**
 * An edit applied in many small steps while code runs. Each step takes effect immediately; formulas
 * recalculate before anything is read back; finish() records everything as one undo step.
 */
export class LiveEdit {
  readonly tx: Transaction;
  private pending: Transaction[] = [];
  private closed = false;

  constructor(
    private doc: Doc,
    label: string,
    private emit: (e: DocEvent) => void,
    private sync: (tx: Transaction) => void,
  ) {
    this.tx = new Transaction(label, doc.wb);
    this.tx.source = 'macro';
    this.tx.selectionBefore = doc.selectionProvider?.() ?? undefined;
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  /** Applies a change now. Structural changes (rows, sheets) recalculate right away. */
  apply(fn: (tx: Transaction) => void): void {
    if (this.closed) throw new Error('This edit has already finished.');
    const step = new Transaction(this.tx.label, this.doc.wb);
    try {
      fn(step);
    } finally {
      if (!step.isEmpty) {
        this.tx.absorb(step);
        if (step.structural) {
          this.flushPending();
          this.sync(step);
        } else this.pending.push(step);
      }
    }
  }

  /** Recalculates formulas affected by changes so far (call before reading values). */
  settle(): void {
    this.flushPending();
  }

  private flushPending(): void {
    if (!this.pending.length) return;
    const merged = new Transaction(this.tx.label, this.doc.wb);
    for (const p of this.pending) merged.absorb(p);
    this.pending = [];
    this.sync(merged);
  }

  /** Lets the screen show progress while the code is still running. */
  repaint(): void {
    this.settle();
    this.emit({ type: 'meta' });
  }

  /** Records the whole edit as one undoable action. */
  finish(): void {
    if (this.closed) return;
    this.settle();
    this.closed = true;
    if (this.tx.isEmpty) {
      this.emit({ type: 'meta' });
      return;
    }
    this.tx.selectionAfter = this.doc.selectionProvider?.() ?? undefined;
    this.doc.history.push(this.tx);
    this.emit({ type: 'change', tx: this.tx, direction: 'do' });
  }
}
