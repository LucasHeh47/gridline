// A parsed CSV file with fast column lookup.
import { parseCSV } from '../io/csv';

export class Table {
  readonly headers: string[];
  readonly rows: string[][];
  private index = new Map<string, number>();

  constructor(text: string) {
    const all = parseCSV(text.replace(/^﻿/, ''));
    this.headers = (all[0] ?? []).map((h) => h.trim());
    this.rows = all.slice(1).filter((r) => r.length > 1);
    this.headers.forEach((h, i) => this.index.set(h, i));
  }

  col(name: string): number {
    return this.index.get(name) ?? -1;
  }

  has(name: string): boolean {
    return this.index.has(name);
  }

  str(row: string[], name: string): string {
    const i = this.index.get(name);
    if (i === undefined) return '';
    const v = row[i] ?? '';
    return v === 'NA' ? '' : v;
  }

  /** Numeric value, or null when blank / missing. */
  num(row: string[], name: string): number | null {
    const i = this.index.get(name);
    if (i === undefined) return null;
    const v = row[i];
    if (v === undefined || v === '' || v === 'NA') return null;
    const n = Number(v);
    return isNaN(n) ? null : n;
  }
}
