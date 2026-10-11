// Undoable transactions. Every user action is one Transaction; undo replays its changes backwards.
import { cellKey, keyCol, keyRow } from './address';
import { config } from './config';
import type { Cell, Sheet, Workbook } from './workbook';

type Change =
  | { k: 'cell'; sheet: Sheet; key: number; before: Cell | undefined; after: Cell | undefined }
  | { k: 'map'; map: Map<unknown, unknown>; key: unknown; before: unknown; after: unknown; sheet?: Sheet }
  | { k: 'set'; set: Set<unknown>; key: unknown; before: boolean; after: boolean; sheet?: Sheet }
  | { k: 'prop'; target: object; prop: string; before: unknown; after: unknown; structural: boolean };

export interface TouchedCell {
  sheetId: number;
  row: number;
  col: number;
}

export interface SelectionSnapshot {
  sheetId: number;
  ranges: { r1: number; c1: number; r2: number; c2: number }[];
  active: { row: number; col: number };
}

export class Transaction {
  readonly changes: Change[] = [];
  /** Set when the change requires the formula engine to rebuild everything. */
  structural = false;
  selectionBefore?: SelectionSnapshot;
  selectionAfter?: SelectionSnapshot;
  /** Set for changes made by a macro, so they don't fire "cell changed" events. */
  source?: 'macro';
  private cellIndex = new Map<string, Change & { k: 'cell' }>();

  constructor(
    readonly label: string,
    readonly wb: Workbook,
  ) {}

  get isEmpty(): boolean {
    return this.changes.length === 0;
  }

  setCell(sheet: Sheet, row: number, col: number, cell: Cell | undefined): void {
    const key = cellKey(row, col);
    const before = sheet.cells.get(key);
    const normalized = cell && Object.keys(cell).length ? cell : undefined;
    const id = `${sheet.id}:${key}`;
    const existing = this.cellIndex.get(id);
    if (existing) existing.after = normalized;
    else {
      const ch = { k: 'cell' as const, sheet, key, before, after: normalized };
      this.cellIndex.set(id, ch);
      this.changes.push(ch);
    }
    sheet.rawSet(key, normalized);
  }

  /** Merges fields into a cell (pass undefined to remove a field). */
  patchCell(sheet: Sheet, row: number, col: number, patch: Partial<Cell>): void {
    const before = sheet.getCell(row, col);
    const next: Cell = { ...(before ?? {}) };
    for (const k of Object.keys(patch) as (keyof Cell)[]) {
      const v = patch[k];
      if (v === undefined) delete next[k];
      else (next as Record<string, unknown>)[k] = v;
    }
    this.setCell(sheet, row, col, next);
  }

  setMap<K, V>(map: Map<K, V>, key: K, value: V | undefined, sheet?: Sheet): void {
    const before = map.get(key);
    if (before === value) return;
    this.changes.push({ k: 'map', map: map as Map<unknown, unknown>, key, before, after: value, sheet });
    if (value === undefined) map.delete(key);
    else map.set(key, value);
  }

  setMember<K>(set: Set<K>, key: K, present: boolean, sheet?: Sheet): void {
    const before = set.has(key);
    if (before === present) return;
    this.changes.push({ k: 'set', set: set as Set<unknown>, key, before, after: present, sheet });
    if (present) set.add(key);
    else set.delete(key);
  }

  /**
   * Replaces a property. Collections must be replaced (not mutated) so the old value
   * stays intact for undo. `structural` forces a full formula-engine rebuild.
   */
  setProp<T extends object, P extends keyof T>(target: T, prop: P, value: T[P], structural = false): void {
    const before = target[prop];
    if (before === value) return;
    this.changes.push({ k: 'prop', target, prop: prop as string, before, after: value, structural });
    target[prop] = value;
    if (structural) this.structural = true;
    if ((target as unknown as Sheet).invalidateBounds) (target as unknown as Sheet).invalidateBounds();
  }

  /** Appends the changes of another (already applied) transaction, for edits made in steps. */
  absorb(other: Transaction): void {
    for (const ch of other.changes) this.changes.push(ch);
    if (other.structural) this.structural = true;
  }

  /**
   * Writes the opposite of this transaction into `into` (to undo it later, after other edits). Parts
   * that were changed again since are left alone and counted in `skipped`.
   */
  revertInto(into: Transaction): { reverted: number; skipped: number } {
    let reverted = 0;
    let skipped = 0;
    for (let i = this.changes.length - 1; i >= 0; i--) {
      const ch = this.changes[i];
      switch (ch.k) {
        case 'cell':
          if (ch.sheet.cells.get(ch.key) !== ch.after) {
            skipped++;
            continue;
          }
          into.setCell(ch.sheet, keyRow(ch.key), keyCol(ch.key), ch.before);
          break;
        case 'map':
          if (ch.map.get(ch.key) !== ch.after) {
            skipped++;
            continue;
          }
          into.setMap(ch.map, ch.key, ch.before, ch.sheet);
          break;
        case 'set':
          if (ch.set.has(ch.key) !== ch.after) {
            skipped++;
            continue;
          }
          into.setMember(ch.set, ch.key, ch.before, ch.sheet);
          break;
        case 'prop':
          if ((ch.target as Record<string, unknown>)[ch.prop] !== ch.after) {
            skipped++;
            continue;
          }
          into.setProp(ch.target as Record<string, unknown>, ch.prop, ch.before, ch.structural);
          break;
      }
      reverted++;
    }
    return { reverted, skipped };
  }

  /** Cells whose content changed, for incremental recalculation. */
  touchedCells(): TouchedCell[] {
    const out: TouchedCell[] = [];
    for (const ch of this.changes) {
      if (ch.k !== 'cell') continue;
      const b = ch.before;
      const a = ch.after;
      if (b?.f === a?.f && b?.v === a?.v) continue; // formatting only
      out.push({ sheetId: ch.sheet.id, row: keyRow(ch.key), col: keyCol(ch.key) });
    }
    return out;
  }

  /** Re-applies or reverts all changes. */
  replay(forward: boolean): void {
    const list = forward ? this.changes : [...this.changes].reverse();
    for (const ch of list) {
      switch (ch.k) {
        case 'cell':
          ch.sheet.rawSet(ch.key, forward ? ch.after : ch.before);
          break;
        case 'map': {
          const v = forward ? ch.after : ch.before;
          if (v === undefined) ch.map.delete(ch.key);
          else ch.map.set(ch.key, v);
          ch.sheet?.invalidateBounds();
          break;
        }
        case 'set': {
          const v = forward ? ch.after : ch.before;
          if (v) ch.set.add(ch.key);
          else ch.set.delete(ch.key);
          break;
        }
        case 'prop':
          (ch.target as Record<string, unknown>)[ch.prop] = forward ? ch.after : ch.before;
          if ((ch.target as unknown as Sheet).invalidateBounds) (ch.target as unknown as Sheet).invalidateBounds();
          break;
      }
    }
  }
}

export class History {
  private undoStack: Transaction[] = [];
  private redoStack: Transaction[] = [];
  /** Increments on every change; used to detect unsaved changes. */
  version = 0;

  push(tx: Transaction): void {
    if (tx.isEmpty) return;
    this.undoStack.push(tx);
    while (this.undoStack.length > Math.max(10, config.undoLimit)) this.undoStack.shift();
    this.redoStack = [];
    this.version++;
  }

  popUndo(): Transaction | undefined {
    const tx = this.undoStack.pop();
    if (tx) {
      this.redoStack.push(tx);
      this.version++;
    }
    return tx;
  }

  popRedo(): Transaction | undefined {
    const tx = this.redoStack.pop();
    if (tx) {
      this.undoStack.push(tx);
      this.version++;
    }
    return tx;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  peekUndo(): Transaction | undefined {
    return this.undoStack[this.undoStack.length - 1];
  }

  peekUndoLabel(): string | undefined {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }

  peekRedoLabel(): string | undefined {
    return this.redoStack[this.redoStack.length - 1]?.label;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
