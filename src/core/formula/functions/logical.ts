import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import { compareValues, isArrayLike, scalarOf, toBool, toMatrix } from '../coerce';
import { evaluate } from '../evaluator';
import type { Node } from '../parser';
import { Matrix, RangeRef, type Arg, type Ctx, type Value } from '../types';
import { define } from './registry';

const nodesOf = (a: Arg[]) => a as unknown as Node[];

function evalArg(ctx: Ctx, n: Node | undefined, dflt: Value): Value {
  if (n === undefined) return dflt;
  if (n.t === 'missing') return 0;
  return evaluate(n, ctx);
}

function pick(v: Value, r: number, c: number, ctx: Ctx): Scalar {
  if (isArrayLike(v)) {
    const m = v instanceof Matrix ? v : toMatrix(ctx, v);
    const rr = m.rows === 1 ? 0 : r;
    const cc = m.cols === 1 ? 0 : c;
    if (rr >= m.rows || cc >= m.cols) return Errors.na();
    return m.get(rr, cc);
  }
  return scalarOf(ctx, v);
}

define(
  'IF',
  'Logical',
  'logical_test, value_if_true, [value_if_false]',
  'Returns one value if a condition is TRUE and another if it\'s FALSE.',
  (ctx, a) => {
    const n = nodesOf(a);
    const cond = evaluate(n[0], ctx);
    if (isArrayLike(cond)) {
      const cm = toMatrix(ctx, cond);
      const t = evalArg(ctx, n[1], true);
      const f = evalArg(ctx, n[2], false);
      const tm = isArrayLike(t) ? toMatrix(ctx, t) : null;
      const fm = isArrayLike(f) ? toMatrix(ctx, f) : null;
      const rows = Math.max(cm.rows, tm?.rows ?? 1, fm?.rows ?? 1);
      const cols = Math.max(cm.cols, tm?.cols ?? 1, fm?.cols ?? 1);
      return Matrix.fill(rows, cols, (r, c) => {
        const b = toBool(pick(cm, r, c, ctx));
        if (isError(b)) return b;
        return b ? pick(tm ?? t, r, c, ctx) : pick(fm ?? f, r, c, ctx);
      });
    }
    const b = toBool(scalarOf(ctx, cond));
    if (isError(b)) return b;
    return b ? evalArg(ctx, n[1], true) : evalArg(ctx, n[2], false);
  },
  { lazy: true },
);

define(
  'IFS',
  'Logical',
  'logical_test1, value_if_true1, [logical_test2], [value_if_true2], ...',
  'Checks conditions in order and returns the value for the first TRUE one.',
  (ctx, a) => {
    const n = nodesOf(a);
    if (n.length % 2) return Errors.value('IFS needs pairs: a condition, then its result.');
    for (let i = 0; i < n.length; i += 2) {
      const b = toBool(scalarOf(ctx, evaluate(n[i], ctx)));
      if (isError(b)) return b;
      if (b) return evalArg(ctx, n[i + 1], 0);
    }
    return Errors.na('None of the IFS conditions were TRUE. Add TRUE, "default" as the last pair to catch everything else.');
  },
  { lazy: true, min: 2 },
);

function errorCatcher(onlyNA: boolean) {
  return (ctx: Ctx, a: Arg[]): Value => {
    const n = nodesOf(a);
    const v = evaluate(n[0], ctx);
    const isCaught = (x: Scalar) => x instanceof ErrorValue && (!onlyNA || x.code === '#N/A');
    if (isArrayLike(v)) {
      const m = toMatrix(ctx, v);
      let fallback: Value | undefined;
      return Matrix.fill(m.rows, m.cols, (r, c) => {
        const x = m.get(r, c);
        if (!isCaught(x)) return x;
        if (fallback === undefined) fallback = evalArg(ctx, n[1], '');
        return pick(fallback, r, c, ctx);
      });
    }
    const s = scalarOf(ctx, v);
    if (isCaught(s)) return evalArg(ctx, n[1], '');
    return s;
  };
}

define('IFERROR', 'Logical', 'value, value_if_error', 'Returns value_if_error if the value is an error; otherwise the value itself.', errorCatcher(false), { lazy: true });
define('IFNA', 'Logical', 'value, value_if_na', 'Returns value_if_na if the value is #N/A; otherwise the value itself.', errorCatcher(true), { lazy: true });

function logicalValues(ctx: Ctx, args: Arg[]): boolean[] | ErrorValue {
  const out: boolean[] = [];
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof RangeRef || a instanceof Matrix) {
      const m = toMatrix(ctx, a);
      for (const row of m.data)
        for (const v of row) {
          if (v instanceof ErrorValue) return v;
          if (typeof v === 'boolean') out.push(v);
          else if (typeof v === 'number') out.push(v !== 0);
        }
    } else {
      const b = toBool(a);
      if (isError(b)) return b;
      out.push(b);
    }
  }
  if (!out.length) return Errors.value('There are no TRUE/FALSE values to check.');
  return out;
}

define('AND', 'Logical', 'logical1, [logical2], ...', 'Returns TRUE if all arguments are TRUE.', (ctx, args) => {
  const v = logicalValues(ctx, args);
  return isError(v) ? v : v.every(Boolean);
});
define('OR', 'Logical', 'logical1, [logical2], ...', 'Returns TRUE if any argument is TRUE.', (ctx, args) => {
  const v = logicalValues(ctx, args);
  return isError(v) ? v : v.some(Boolean);
});
define('XOR', 'Logical', 'logical1, [logical2], ...', 'Returns TRUE if an odd number of arguments are TRUE.', (ctx, args) => {
  const v = logicalValues(ctx, args);
  return isError(v) ? v : v.filter(Boolean).length % 2 === 1;
});
define(
  'NOT',
  'Logical',
  'logical',
  'Reverses TRUE and FALSE.',
  (ctx, [a]) => {
    const b = toBool(scalarOf(ctx, a));
    return isError(b) ? b : !b;
  },
  { lift: true },
);
define('TRUE', 'Logical', '', 'Returns the logical value TRUE.', () => true);
define('FALSE', 'Logical', '', 'Returns the logical value FALSE.', () => false);

define(
  'SWITCH',
  'Logical',
  'expression, value1, result1, [default_or_value2], [result2], ...',
  'Compares an expression against a list of values and returns the matching result.',
  (ctx, a) => {
    const n = nodesOf(a);
    const x = scalarOf(ctx, evaluate(n[0], ctx));
    if (x instanceof ErrorValue) return x;
    let i = 1;
    for (; i + 1 < n.length; i += 2) {
      const v = scalarOf(ctx, evaluate(n[i], ctx));
      if (v instanceof ErrorValue) return v;
      if (typeof v === typeof x && compareValues(x, v) === 0) return evalArg(ctx, n[i + 1], 0);
    }
    if (i < n.length) return evalArg(ctx, n[i], 0);
    return Errors.na('No value matched in SWITCH, and there is no default value.');
  },
  { lazy: true, min: 3 },
);

define(
  'CHOOSE',
  'Logical',
  'index_num, value1, [value2], ...',
  'Picks a value from a list by its position, e.g. CHOOSE(2, "a", "b", "c") returns "b".',
  (ctx, a) => {
    const n = nodesOf(a);
    const idxV = evaluate(n[0], ctx);
    const choose = (iv: Scalar): Value => {
      if (iv instanceof ErrorValue) return iv;
      const i = Math.trunc(typeof iv === 'number' ? iv : Number(iv));
      if (!(i >= 1 && i < n.length)) return Errors.value(`index_num must be between 1 and ${n.length - 1}.`);
      return evalArg(ctx, n[i], 0);
    };
    if (isArrayLike(idxV)) {
      const m = toMatrix(ctx, idxV);
      return Matrix.fill(m.rows, m.cols, (r, c) => scalarOf(ctx, choose(m.get(r, c))));
    }
    return choose(scalarOf(ctx, idxV));
  },
  { lazy: true, min: 2 },
);

define(
  'LET',
  'Logical',
  'name1, value1, [name2], [value2], ..., calculation',
  'Assigns names to calculation results, so you can reuse them: LET(x, A1*2, x+x).',
  (ctx, a) => {
    const n = nodesOf(a);
    if (n.length < 3 || n.length % 2 === 0) return Errors.value('LET needs name/value pairs followed by a final calculation.');
    const scope = new Map(ctx.scope ?? []);
    const inner: Ctx = { ...ctx, scope };
    for (let i = 0; i + 1 < n.length; i += 2) {
      const nameNode = n[i];
      if (nameNode.t !== 'name') return Errors.value('In LET, every odd argument must be a name (like x or total).');
      scope.set(nameNode.name.toUpperCase(), evaluate(n[i + 1], inner));
    }
    return evaluate(n[n.length - 1], inner);
  },
  { lazy: true, min: 3 },
);
