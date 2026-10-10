// Conditional formatting evaluation.
import { MAX_COLS, MAX_ROWS, rangeContains } from './address';
import type { Doc } from './doc';
import { compareValues, toText } from './formula/coerce';
import { gkey } from './formula/engine';
import { parseFormula, type Node } from './formula/parser';
import { Matrix, RangeRef } from './formula/types';
import { parseNumberLike } from './input';
import { mixColors } from './styles';
import { ErrorValue, type Scalar } from './values';
import type { CFRule, CFStyle, CFValue, Sheet } from './workbook';

export interface CFResult {
  style?: CFStyle;
  fill?: string;
  bar?: { from: number; to: number; color: string; gradient: boolean; negative: boolean };
  icon?: { set: string; index: number; count: number };
  hideValue?: boolean;
}

interface RuleStats {
  nums: number[];
  sorted: number[];
  min: number;
  max: number;
  avg: number;
  counts: Map<string, number>;
}

function keyOf(v: Scalar): string {
  if (typeof v === 'string') return 's' + v.toLowerCase();
  if (v instanceof ErrorValue) return 'e' + v.code;
  return typeof v + String(v);
}

function truthy(v: Scalar): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') return v.toUpperCase() === 'TRUE';
  return false;
}

export class CFEvaluator {
  private version = -1;
  private cache = new Map<number, CFResult | null>();
  private stats = new Map<CFRule, RuleStats>();
  private asts = new Map<string, Node | null>();
  private cfRef = new WeakMap<Sheet, CFRule[]>();

  constructor(private doc: Doc) {}

  /** Clears caches; call when rules change. */
  invalidate(): void {
    this.cache.clear();
    this.stats.clear();
  }

  private sync(sheet: Sheet): void {
    const v = this.doc.engine.version;
    if (v !== this.version || this.cfRef.get(sheet) !== sheet.cf) {
      this.version = v;
      this.cfRef.set(sheet, sheet.cf);
      this.cache.clear();
      this.stats.clear();
    }
  }

  evaluate(sheet: Sheet, row: number, col: number): CFResult | null {
    if (!sheet.cf.length) return null;
    this.sync(sheet);
    const g = gkey(sheet.id, row, col);
    if (this.cache.has(g)) return this.cache.get(g)!;
    let result: CFResult | null = null;
    const rules = sheet.cf.filter((r) => r.ranges.some((rg) => rangeContains(rg, row, col)));
    if (rules.length) {
      result = {};
      const value = this.doc.engine.getValue(sheet, row, col);
      for (const rule of rules) {
        const applied = this.applyRule(rule, sheet, row, col, value, result);
        if (applied && rule.stopIfTrue) break;
      }
      if (!result.style && !result.fill && !result.bar && !result.icon) result = null;
    }
    this.cache.set(g, result);
    return result;
  }

  private ruleStats(rule: CFRule, sheet: Sheet): RuleStats {
    let s = this.stats.get(rule);
    if (s) return s;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    const bounds = this.doc.engine.usedBounds(sheet);
    for (const rg of rule.ranges) {
      const r2 = Math.min(rg.r2, rg.r2 === MAX_ROWS - 1 ? bounds.maxRow : rg.r2);
      const c2 = Math.min(rg.c2, rg.c2 === MAX_COLS - 1 ? bounds.maxCol : rg.c2);
      for (let r = rg.r1; r <= r2; r++)
        for (let c = rg.c1; c <= c2; c++) {
          const v = this.doc.engine.getValue(sheet, r, c);
          if (typeof v === 'number') nums.push(v);
          if (v !== null && v !== '') {
            const k = keyOf(v);
            counts.set(k, (counts.get(k) ?? 0) + 1);
          }
        }
    }
    const sorted = [...nums].sort((a, b) => a - b);
    s = {
      nums,
      sorted,
      min: sorted.length ? sorted[0] : 0,
      max: sorted.length ? sorted[sorted.length - 1] : 0,
      avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0,
      counts,
    };
    this.stats.set(rule, s);
    return s;
  }

  private ast(text: string): Node | null {
    let a = this.asts.get(text);
    if (a === undefined) {
      a = parseFormula(text.replace(/^=/, '')).ast;
      this.asts.set(text, a);
    }
    return a;
  }

  /** Evaluates a rule formula relative to the rule's top-left cell. */
  private evalFormula(rule: CFRule, text: string, sheet: Sheet, row: number, col: number): Scalar {
    const ast = this.ast(text);
    if (!ast) return null;
    const anchor = rule.ranges[0];
    const v = this.doc.engine.evalAst(ast, sheet, anchor.r1, anchor.c1, row - anchor.r1, col - anchor.c1);
    if (v instanceof RangeRef) return this.doc.engine.getValue(v.sheet, v.r1, v.c1);
    if (v instanceof Matrix) return v.get(0, 0);
    return v;
  }

  private operand(rule: CFRule, text: string, sheet: Sheet, row: number, col: number): Scalar {
    if (text.startsWith('=')) return this.evalFormula(rule, text, sheet, row, col);
    const n = parseNumberLike(text);
    if (n !== null) return n;
    return text.replace(/^"(.*)"$/, '$1');
  }

  private threshold(cv: CFValue, s: RuleStats, rule: CFRule, sheet: Sheet, row: number, col: number): number {
    switch (cv.type) {
      case 'min':
        return s.min;
      case 'max':
        return s.max;
      case 'num':
        return Number(cv.value ?? 0);
      case 'percent':
        return s.min + ((s.max - s.min) * Number(cv.value ?? 0)) / 100;
      case 'percentile': {
        if (!s.sorted.length) return 0;
        const p = Math.max(0, Math.min(100, Number(cv.value ?? 0))) / 100;
        const pos = p * (s.sorted.length - 1);
        const lo = Math.floor(pos);
        const hi = Math.min(lo + 1, s.sorted.length - 1);
        return s.sorted[lo] + (s.sorted[hi] - s.sorted[lo]) * (pos - lo);
      }
      case 'formula': {
        const v = this.evalFormula(rule, String(cv.value ?? ''), sheet, row, col);
        return typeof v === 'number' ? v : 0;
      }
    }
  }

  private applyRule(rule: CFRule, sheet: Sheet, row: number, col: number, v: Scalar, out: CFResult): boolean {
    const mergeStyle = (st: CFStyle) => {
      out.style = { ...st, ...(out.style ?? {}) };
      return true;
    };
    switch (rule.type) {
      case 'cellIs': {
        if (v === null) return false;
        const a = this.operand(rule, rule.values[0] ?? '', sheet, row, col);
        const b = this.operand(rule, rule.values[1] ?? '', sheet, row, col);
        if (v instanceof ErrorValue) return false;
        const ca = compareValues(v, a);
        let ok = false;
        switch (rule.operator) {
          case 'greaterThan':
            ok = ca > 0;
            break;
          case 'lessThan':
            ok = ca < 0;
            break;
          case 'greaterThanOrEqual':
            ok = ca >= 0;
            break;
          case 'lessThanOrEqual':
            ok = ca <= 0;
            break;
          case 'equal':
            ok = ca === 0;
            break;
          case 'notEqual':
            ok = ca !== 0;
            break;
          case 'between':
          case 'notBetween': {
            const lo = compareValues(a, b) <= 0 ? a : b;
            const hi = lo === a ? b : a;
            const inside = compareValues(v, lo) >= 0 && compareValues(v, hi) <= 0;
            ok = rule.operator === 'between' ? inside : !inside;
            break;
          }
        }
        return ok ? mergeStyle(rule.style) : false;
      }
      case 'text': {
        if (v === null) return false;
        const t = toText(v);
        if (t instanceof ErrorValue) return false;
        const s = t.toLowerCase();
        const q = rule.text.toLowerCase();
        const ok =
          rule.operator === 'contains'
            ? s.includes(q)
            : rule.operator === 'notContains'
              ? !s.includes(q)
              : rule.operator === 'beginsWith'
                ? s.startsWith(q)
                : s.endsWith(q);
        return ok ? mergeStyle(rule.style) : false;
      }
      case 'duplicate':
      case 'unique': {
        if (v === null || v === '') return false;
        const n = this.ruleStats(rule, sheet).counts.get(keyOf(v)) ?? 0;
        const ok = rule.type === 'duplicate' ? n > 1 : n === 1;
        return ok ? mergeStyle(rule.style) : false;
      }
      case 'top': {
        if (typeof v !== 'number') return false;
        const s = this.ruleStats(rule, sheet);
        if (!s.sorted.length) return false;
        const count = rule.percent ? Math.max(1, Math.floor((s.sorted.length * rule.rank) / 100)) : Math.min(rule.rank, s.sorted.length);
        const ok = rule.bottom ? v <= s.sorted[count - 1] : v >= s.sorted[s.sorted.length - count];
        return ok ? mergeStyle(rule.style) : false;
      }
      case 'average': {
        if (typeof v !== 'number') return false;
        const s = this.ruleStats(rule, sheet);
        const ok = rule.below ? v < s.avg : v > s.avg;
        return ok ? mergeStyle(rule.style) : false;
      }
      case 'blanks':
        return v === null || (typeof v === 'string' && v.trim() === '') ? mergeStyle(rule.style) : false;
      case 'noBlanks':
        return v !== null && !(typeof v === 'string' && v.trim() === '') ? mergeStyle(rule.style) : false;
      case 'errors':
        return v instanceof ErrorValue ? mergeStyle(rule.style) : false;
      case 'noErrors':
        return !(v instanceof ErrorValue) ? mergeStyle(rule.style) : false;
      case 'expression':
        return truthy(this.evalFormula(rule, rule.formula, sheet, row, col)) ? mergeStyle(rule.style) : false;
      case 'colorScale': {
        if (typeof v !== 'number' || out.fill) return false;
        const s = this.ruleStats(rule, sheet);
        const stops = rule.stops.map((st) => ({ at: this.threshold(st.at, s, rule, sheet, row, col), color: st.color }));
        if (v <= stops[0].at) out.fill = stops[0].color;
        else if (v >= stops[stops.length - 1].at) out.fill = stops[stops.length - 1].color;
        else {
          for (let i = 0; i < stops.length - 1; i++) {
            const a = stops[i];
            const b = stops[i + 1];
            if (v >= a.at && v <= b.at) {
              const t = b.at === a.at ? 0 : (v - a.at) / (b.at - a.at);
              out.fill = mixColors(a.color, b.color, t);
              break;
            }
          }
        }
        return true;
      }
      case 'dataBar': {
        if (typeof v !== 'number' || out.bar) return false;
        const s = this.ruleStats(rule, sheet);
        const lo = rule.min.type === 'min' ? Math.min(0, s.min) : this.threshold(rule.min, s, rule, sheet, row, col);
        let hi = rule.max.type === 'max' ? Math.max(0, s.max) : this.threshold(rule.max, s, rule, sheet, row, col);
        if (hi <= lo) hi = lo + 1;
        const clamp = (x: number) => Math.max(0, Math.min(1, x));
        const axis = clamp((0 - lo) / (hi - lo));
        const pos = clamp((v - lo) / (hi - lo));
        out.bar = { from: Math.min(axis, pos), to: Math.max(axis, pos), color: rule.color, gradient: rule.gradient, negative: v < 0 };
        if (!rule.showValue) out.hideValue = true;
        return true;
      }
      case 'iconSet': {
        if (typeof v !== 'number' || out.icon) return false;
        const s = this.ruleStats(rule, sheet);
        const ts = rule.thresholds.map((t) => this.threshold(t, s, rule, sheet, row, col));
        let idx = 0;
        for (let i = 0; i < ts.length; i++) if (v >= ts[i]) idx = i;
        const count = rule.thresholds.length;
        out.icon = { set: rule.iconSet, index: rule.reverse ? count - 1 - idx : idx, count };
        if (!rule.showValue) out.hideValue = true;
        return true;
      }
    }
    return false;
  }
}

/** Readable labels for the rules list in the UI. */
export function describeRule(rule: CFRule): string {
  const op: Record<string, string> = {
    greaterThan: '>',
    lessThan: '<',
    greaterThanOrEqual: '≥',
    lessThanOrEqual: '≤',
    equal: '=',
    notEqual: '≠',
  };
  switch (rule.type) {
    case 'cellIs':
      if (rule.operator === 'between') return `Between ${rule.values[0]} and ${rule.values[1]}`;
      if (rule.operator === 'notBetween') return `Not between ${rule.values[0]} and ${rule.values[1]}`;
      return `Value ${op[rule.operator]} ${rule.values[0]}`;
    case 'text': {
      const m: Record<string, string> = { contains: 'Text contains', notContains: 'Text doesn\'t contain', beginsWith: 'Text begins with', endsWith: 'Text ends with' };
      return `${m[rule.operator]} "${rule.text}"`;
    }
    case 'duplicate':
      return 'Duplicate values';
    case 'unique':
      return 'Unique values';
    case 'top':
      return `${rule.bottom ? 'Bottom' : 'Top'} ${rule.rank}${rule.percent ? '%' : ''}`;
    case 'average':
      return rule.below ? 'Below average' : 'Above average';
    case 'blanks':
      return 'Blank cells';
    case 'noBlanks':
      return 'Non-blank cells';
    case 'errors':
      return 'Cells with errors';
    case 'noErrors':
      return 'Cells without errors';
    case 'expression':
      return `Formula: =${rule.formula.replace(/^=/, '')}`;
    case 'colorScale':
      return `${rule.stops.length}-color scale`;
    case 'dataBar':
      return 'Data bars';
    case 'iconSet':
      return `Icon set (${rule.iconSet})`;
  }
}
