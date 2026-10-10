// Fast row/column position lookups with custom sizes and hidden rows/columns.

export class Axis {
  private keys: number[] = [];
  private sizes: number[] = [];
  /** prefix[j] = sum of (size - def) for the first j overrides */
  private prefix: number[] = [0];
  readonly def: number;

  constructor(
    readonly count: number,
    def: number,
    custom: Map<number, number>,
    hidden: Set<number>[],
    zoom: number,
  ) {
    this.def = Math.max(1, Math.round(def * zoom));
    const overrides = new Map<number, number>();
    for (const [i, s] of custom) overrides.set(i, Math.max(0, Math.round(s * zoom)));
    for (const set of hidden) for (const i of set) overrides.set(i, 0);
    this.keys = [...overrides.keys()].sort((a, b) => a - b);
    this.sizes = this.keys.map((k) => overrides.get(k)!);
    let acc = 0;
    this.prefix = [0];
    for (const s of this.sizes) {
      acc += s - this.def;
      this.prefix.push(acc);
    }
  }

  /** Number of overrides with index < i */
  private rank(i: number): number {
    let lo = 0;
    let hi = this.keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  size(i: number): number {
    const j = this.rank(i);
    return this.keys[j] === i ? this.sizes[j] : this.def;
  }

  /** Pixel position of the start of index i. */
  offset(i: number): number {
    return i * this.def + this.prefix[this.rank(i)];
  }

  /** Index whose span contains px (clamped to the valid range). */
  indexAt(px: number): number {
    if (px <= 0) return this.firstVisible(0);
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offset(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  firstVisible(from: number): number {
    let i = from;
    while (i < this.count - 1 && this.size(i) === 0) i++;
    return i;
  }

  isHidden(i: number): boolean {
    return this.size(i) === 0;
  }

  /** Next visible index in a direction (or the same index if none). */
  step(i: number, dir: 1 | -1): number {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.size(j) === 0) j += dir;
    if (j < 0 || j >= this.count) return i;
    return j;
  }
}
