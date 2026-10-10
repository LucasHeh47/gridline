// AST evaluation.
import { formatCell, MAX_COLS, MAX_ROWS } from '../address';
import { Errors, ErrorValue, isError, type Scalar } from '../values';
import { compareValues, describeRef, isArrayLike, numResult, scalarOf, toMatrix, toNumber, toText } from './coerce';
import { FUNCTIONS, suggestFunction, type FnDef } from './functions/registry';
import { parseFormula, type Node } from './parser';
import { Matrix, RangeRef, type Arg, type Ctx, type Value } from './types';

export function evaluate(node: Node, ctx: Ctx): Value {
  switch (node.t) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return new ErrorValue(node.v, node.v === '#REF!' ? Errors.ref().message : `The formula contains the error value ${node.v}.`);
    case 'missing':
      return null;
    case 'ref':
      return resolveRef(node, ctx);
    case 'name':
      return resolveName(node.name, ctx);
    case 'func':
      return callFunction(node, ctx);
    case 'unary': {
      const v = evaluate(node.e, ctx);
      if (node.op === '@') return scalarOf(ctx, v);
      if (node.op === '+') return v;
      return liftUnary(ctx, v, (x) => {
        const n = toNumber(x);
        return isError(n) ? n : -n;
      });
    }
    case 'pct':
      return liftUnary(ctx, evaluate(node.e, ctx), (x) => {
        const n = toNumber(x);
        return isError(n) ? n : n / 100;
      });
    case 'bin':
      return binary(node, ctx);
    case 'range': {
      const l = evaluate(node.l, ctx);
      const r = evaluate(node.r, ctx);
      if (l instanceof ErrorValue) return l;
      if (r instanceof ErrorValue) return r;
      if (!(l instanceof RangeRef) || !(r instanceof RangeRef) || l.sheet !== r.sheet)
        return Errors.value('Both sides of ":" must be references on the same sheet.');
      return new RangeRef(
        l.sheet,
        Math.min(l.r1, r.r1),
        Math.min(l.c1, r.c1),
        Math.max(l.r2, r.r2),
        Math.max(l.c2, r.c2),
      );
    }
    case 'spill': {
      const v = evaluate(node.e, ctx);
      if (v instanceof RangeRef) return spillOf(v, ctx);
      return Errors.ref('The # (spill) operator must follow a cell reference.');
    }
    case 'array':
      return new Matrix(node.rows.map((row) => row.map((el) => evaluate(el, ctx) as Scalar)));
  }
}

function spillOf(ref: RangeRef, ctx: Ctx): Value {
  const sp = ctx.spillRange(ref.sheet, ref.r1, ref.c1);
  if (sp) return sp;
  if (ctx.isFormula(ref.sheet, ref.r1, ref.c1)) return new RangeRef(ref.sheet, ref.r1, ref.c1, ref.r1, ref.c1);
  return Errors.ref(`${describeRef(ref, ctx)}# refers to a spilled range, but ${formatCell(ref.r1, ref.c1)} has no formula that spills.`);
}

function resolveRef(node: Node & { t: 'ref' }, ctx: Ctx): Value {
  const ref = node.ref;
  let sheet = ctx.sheet;
  if (ref.sheet !== undefined) {
    const s = ctx.wb.sheetByName(ref.sheet);
    if (!s) return Errors.ref(`There's no sheet named "${ref.sheet}". It may have been renamed or deleted.`);
    sheet = s;
  }
  const r1 = ref.ar1 ? ref.r1 : ref.r1 + ctx.dr;
  const r2 = ref.ar2 ? ref.r2 : ref.r2 + ctx.dr;
  const c1 = ref.ac1 ? ref.c1 : ref.c1 + ctx.dc;
  const c2 = ref.ac2 ? ref.c2 : ref.c2 + ctx.dc;
  if (r1 < 0 || c1 < 0 || r2 >= MAX_ROWS || c2 >= MAX_COLS) return Errors.ref('The reference goes past the edge of the sheet.');
  const rr = new RangeRef(sheet, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
  if (node.spill) return spillOf(rr, ctx);
  return rr;
}

const nameCache = new Map<string, Node | null>();

function resolveName(name: string, ctx: Ctx): Value {
  const upper = name.toUpperCase();
  if (ctx.scope?.has(upper)) return ctx.scope.get(upper)!;
  const dn = ctx.wb.getName(upper);
  if (dn) {
    let ast = nameCache.get(dn.ref);
    if (ast === undefined) {
      ast = parseFormula(dn.ref).ast;
      nameCache.set(dn.ref, ast);
    }
    if (!ast) return Errors.name(`The defined name "${name}" has an invalid formula.`);
    return evaluate(ast, { ...ctx, dr: 0, dc: 0, scope: null });
  }
  const fn = FUNCTIONS.get(upper);
  if (fn) return Errors.name(`"${name}" is a function — add parentheses, like ${upper}(…).`);
  const sug = suggestFunction(upper);
  const hint = sug ? ` Did you mean the function ${sug}()?` : '';
  return Errors.name(`"${name}" isn't a known name. If it's meant to be text, put it in "double quotes".${hint}`);
}

function liftUnary(ctx: Ctx, v: Value, f: (x: Scalar) => Scalar): Value {
  if (isArrayLike(v)) {
    const m = toMatrix(ctx, v);
    return Matrix.fill(m.rows, m.cols, (r, c) => f(m.get(r, c)));
  }
  return f(scalarOf(ctx, v));
}

function describeOperand(node: Node, ctx: Ctx): string | null {
  if (node.t === 'ref' && node.ref.kind === 'cell') {
    const r = node.ref;
    const row = r.ar1 ? r.r1 : r.r1 + ctx.dr;
    const col = r.ac1 ? r.c1 : r.c1 + ctx.dc;
    return (r.sheet ? r.sheet + '!' : '') + formatCell(row, col);
  }
  return null;
}

function arith(op: string, a: Scalar, b: Scalar, node: Node & { t: 'bin' }, ctx: Ctx): Scalar {
  if (op === '&') {
    const x = toText(a);
    if (isError(x)) return x;
    const y = toText(b);
    if (isError(y)) return y;
    return x + y;
  }
  if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
    if (a instanceof ErrorValue) return a;
    if (b instanceof ErrorValue) return b;
    const c = compareValues(a, b);
    switch (op) {
      case '=':
        return c === 0;
      case '<>':
        return c !== 0;
      case '<':
        return c < 0;
      case '>':
        return c > 0;
      case '<=':
        return c <= 0;
      default:
        return c >= 0;
    }
  }
  const x = toNumber(a);
  if (isError(x)) return withOperand(x, node.l, ctx);
  const y = toNumber(b);
  if (isError(y)) return withOperand(y, node.r, ctx);
  switch (op) {
    case '+':
      return numResult(x + y);
    case '-':
      return numResult(x - y);
    case '*':
      return numResult(x * y);
    case '/': {
      if (y === 0) {
        const who = describeOperand(node.r, ctx);
        if (who) return Errors.div0(`You're dividing by ${who}, which is ${b === null ? 'empty' : '0'}.`);
        return Errors.div0('The divisor (the value after "/") is 0.');
      }
      return numResult(x / y);
    }
    case '^': {
      if (x === 0 && y === 0) return Errors.num('0 raised to the power 0 is undefined.');
      if (x < 0 && !Number.isInteger(y)) return Errors.num('A negative number can\'t be raised to a fractional power.');
      return numResult(Math.pow(x, y));
    }
  }
  return Errors.value();
}

function withOperand(e: ErrorValue, node: Node, ctx: Ctx): ErrorValue {
  if (e.origin || e.code !== '#VALUE!') return e;
  const who = describeOperand(node, ctx);
  if (!who || !e.message) return e;
  return new ErrorValue(e.code, `${who} contains ${e.message.charAt(0).toLowerCase()}${e.message.slice(1)}`);
}

function binary(node: Node & { t: 'bin' }, ctx: Ctx): Value {
  const lv = evaluate(node.l, ctx);
  const rv = evaluate(node.r, ctx);
  if (isArrayLike(lv) || isArrayLike(rv)) {
    const lm = toMatrix(ctx, lv);
    const rm = toMatrix(ctx, rv);
    const rows = Math.max(lm.rows, rm.rows);
    const cols = Math.max(lm.cols, rm.cols);
    const pick = (m: Matrix, r: number, c: number): Scalar => {
      const rr = m.rows === 1 ? 0 : r;
      const cc = m.cols === 1 ? 0 : c;
      if (rr >= m.rows || cc >= m.cols) return Errors.na('The arrays in this formula have different sizes.');
      return m.get(rr, cc);
    };
    return Matrix.fill(rows, cols, (r, c) => arith(node.op, pick(lm, r, c), pick(rm, r, c), node, ctx));
  }
  return arith(node.op, scalarOf(ctx, lv), scalarOf(ctx, rv), node, ctx);
}

function callFunction(node: Node & { t: 'func' }, ctx: Ctx): Value {
  const def = FUNCTIONS.get(node.name);
  if (!def) {
    const sug = suggestFunction(node.name);
    return Errors.name(
      `There's no function called ${node.name}.` + (sug ? ` Did you mean ${sug}?` : ' Check the spelling.'),
    );
  }
  const n = node.args.length;
  if (n < def.min) {
    return Errors.value(`${def.name} needs at least ${def.min} argument${def.min === 1 ? '' : 's'}: ${def.name}(${def.args}).`);
  }
  if (n > def.max) {
    return Errors.value(`${def.name} takes at most ${def.max} argument${def.max === 1 ? '' : 's'}: ${def.name}(${def.args}).`);
  }
  if (def.lazy) return def.fn(ctx, node.args as unknown as Arg[]);
  const args: Arg[] = node.args.map((a) => (a.t === 'missing' ? undefined : evaluate(a, ctx)));
  if (def.lift) return liftCall(def, ctx, args);
  return def.fn(ctx, args);
}

function liftCall(def: FnDef, ctx: Ctx, args: Arg[]): Value {
  const lifted = def.lift === true ? args.map((_, i) => i) : (def.lift as number[]);
  let rows = 1;
  let cols = 1;
  let any = false;
  const mats: (Matrix | null)[] = args.map((a, i) => {
    if (!lifted.includes(i) || !isArrayLike(a)) return null;
    const m = toMatrix(ctx, a);
    rows = Math.max(rows, m.rows);
    cols = Math.max(cols, m.cols);
    any = true;
    return m;
  });
  if (!any) return def.fn(ctx, args);
  return Matrix.fill(rows, cols, (r, c) => {
    const a2 = args.map((a, i) => {
      const m = mats[i];
      if (!m) return a;
      const rr = m.rows === 1 ? 0 : r;
      const cc = m.cols === 1 ? 0 : c;
      if (rr >= m.rows || cc >= m.cols) return Errors.na();
      return m.get(rr, cc);
    });
    const v = def.fn(ctx, a2);
    return v instanceof RangeRef || v instanceof Matrix ? scalarOf(ctx, v) : v;
  });
}
