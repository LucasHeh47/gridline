// Formula parser: tokens -> AST. Produces friendly syntax error messages.
import type { ErrorCode } from '../values';
import { tokenize, type RefInfo, type Token } from './tokenizer';

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: ErrorCode }
  | { t: 'ref'; ref: RefInfo; spill?: boolean }
  | { t: 'name'; name: string }
  | { t: 'func'; name: string; args: Node[]; pos: number }
  | { t: 'unary'; op: '-' | '+' | '@'; e: Node }
  | { t: 'pct'; e: Node }
  /** paren: written inside parentheses (Excel's near-zero rule skips those). */
  | { t: 'bin'; op: string; l: Node; r: Node; paren?: boolean }
  | { t: 'range'; l: Node; r: Node }
  | { t: 'spill'; e: Node }
  | { t: 'array'; rows: Node[][] }
  | { t: 'missing' };

export class FormulaSyntaxError extends Error {
  constructor(
    message: string,
    readonly pos: number,
  ) {
    super(message);
  }
}

const BIN_PREC: Record<string, number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
};

/** Strips Excel's file-format prefixes for newer functions. */
export function normalizeFuncName(name: string): string {
  let n = name.toUpperCase();
  if (n.startsWith('_XLFN.')) n = n.slice(6);
  if (n.startsWith('_XLWS.')) n = n.slice(6);
  if (n.startsWith('_XLFN.')) n = n.slice(6);
  return n;
}

class Parser {
  private i = 0;
  private toks: Token[];

  constructor(
    toks: Token[],
    private src: string,
  ) {
    this.toks = toks.filter((t) => t.type !== 'ws');
  }

  private peek(): Token | undefined {
    return this.toks[this.i];
  }

  private next(): Token | undefined {
    return this.toks[this.i++];
  }

  private fail(msg: string, tok?: Token): never {
    throw new FormulaSyntaxError(msg, tok ? tok.start : this.src.length);
  }

  parse(): Node {
    if (this.toks.length === 0) this.fail('The formula is empty. Type something after "=".');
    const node = this.expr(0);
    const t = this.peek();
    if (t) {
      if (t.type === 'rparen') this.fail('There\'s an extra closing parenthesis ")".', t);
      if (t.type === 'comma') this.fail('Unexpected comma. Commas only separate function arguments.', t);
      if (t.type === 'ref' || t.type === 'num' || t.type === 'name' || t.type === 'func' || t.type === 'str')
        this.fail(`Missing an operator before "${t.text}" (for example + or &).`, t);
      this.fail(`Unexpected "${t.text}".`, t);
    }
    return node;
  }

  private expr(minPrec: number): Node {
    let left = this.unary();
    for (;;) {
      const t = this.peek();
      if (!t || t.type !== 'op') break;
      const prec = BIN_PREC[t.text];
      if (prec === undefined || prec < minPrec) break;
      this.next();
      if (!this.peek()) this.fail(`Something is missing after "${t.text}".`);
      const right = this.expr(prec + 1);
      left = { t: 'bin', op: t.text, l: left, r: right };
    }
    return left;
  }

  private unary(): Node {
    const t = this.peek();
    if (t && t.type === 'op' && (t.text === '-' || t.text === '+' || t.text === '@')) {
      this.next();
      if (!this.peek()) this.fail(`Something is missing after "${t.text}".`);
      return { t: 'unary', op: t.text as '-' | '+' | '@', e: this.unary() };
    }
    return this.postfix();
  }

  private postfix(): Node {
    let node = this.primary();
    for (;;) {
      const t = this.peek();
      if (t && t.type === 'op' && t.text === '%') {
        this.next();
        node = { t: 'pct', e: node };
      } else if (t && t.type === 'op' && t.text === '#') {
        this.next();
        node = node.t === 'ref' ? { ...node, spill: true } : { t: 'spill', e: node };
      } else if (t && t.type === 'op' && t.text === ':') {
        this.next();
        const right = this.primary();
        node = { t: 'range', l: node, r: right };
      } else break;
    }
    return node;
  }

  private primary(): Node {
    const t = this.next();
    if (!t) this.fail('The formula ends too early.');
    switch (t.type) {
      case 'num':
        return { t: 'num', v: t.value as number };
      case 'str':
        if (t.unterminated) this.fail('Text is missing its closing quote (").', t);
        return { t: 'str', v: t.value as string };
      case 'bool':
        return { t: 'bool', v: t.value as boolean };
      case 'err':
        return { t: 'err', v: t.value as ErrorCode };
      case 'ref':
        return { t: 'ref', ref: t.ref! };
      case 'name':
        return { t: 'name', name: t.text };
      case 'func':
        return this.call(t);
      case 'lparen': {
        if (this.peek()?.type === 'rparen') this.fail('Empty parentheses "()" need something inside.', this.peek());
        const e = this.expr(0);
        const close = this.next();
        if (!close || close.type !== 'rparen') this.fail('Missing a closing parenthesis ")".', close);
        return e.t === 'bin' ? { ...e, paren: true } : e;
      }
      case 'lbrace':
        return this.array(t);
      case 'rparen':
        this.fail('Unexpected ")" — there may be an extra closing parenthesis or a missing value.', t);
      // eslint-disable-next-line no-fallthrough
      case 'comma':
        this.fail('Unexpected comma — a value is missing.', t);
      // eslint-disable-next-line no-fallthrough
      case 'op':
        this.fail(`"${t.text}" needs a value before it.`, t);
      // eslint-disable-next-line no-fallthrough
      default:
        this.fail(`"${t.text}" isn't valid in a formula.`, t);
    }
  }

  private call(t: Token): Node {
    const name = normalizeFuncName(t.text);
    const open = this.next();
    if (!open || open.type !== 'lparen') this.fail(`Expected "(" after ${name}.`, open);
    const args: Node[] = [];
    if (this.peek()?.type === 'rparen') {
      this.next();
      return { t: 'func', name, args, pos: t.start };
    }
    for (;;) {
      const p = this.peek();
      if (!p) this.fail(`${name}( is missing its closing parenthesis ")".`);
      if (p.type === 'comma' || p.type === 'rparen') {
        args.push({ t: 'missing' });
      } else {
        args.push(this.expr(0));
      }
      const sep = this.next();
      if (!sep) this.fail(`${name}( is missing its closing parenthesis ")".`);
      if (sep.type === 'rparen') break;
      if (sep.type !== 'comma') {
        if (sep.type === 'semicolon') this.fail('Use commas (,) to separate function arguments.', sep);
        this.fail(`Expected "," or ")" in ${name}( but found "${sep.text}".`, sep);
      }
    }
    return { t: 'func', name, args, pos: t.start };
  }

  private array(open: Token): Node {
    const rows: Node[][] = [[]];
    for (;;) {
      const t = this.peek();
      if (!t) this.fail('Array constant is missing its closing "}".', open);
      if (t.type === 'rbrace') {
        this.next();
        break;
      }
      let el: Node;
      if (t.type === 'op' && t.text === '-') {
        this.next();
        const n = this.next();
        if (!n || n.type !== 'num') this.fail('Array constants can only contain numbers, text, TRUE/FALSE or errors.', n);
        el = { t: 'num', v: -(n.value as number) };
      } else {
        this.next();
        if (t.type === 'num') el = { t: 'num', v: t.value as number };
        else if (t.type === 'str') el = { t: 'str', v: t.value as string };
        else if (t.type === 'bool') el = { t: 'bool', v: t.value as boolean };
        else if (t.type === 'err') el = { t: 'err', v: t.value as ErrorCode };
        else this.fail('Array constants can only contain numbers, text, TRUE/FALSE or errors.', t);
      }
      rows[rows.length - 1].push(el);
      const sep = this.peek();
      if (sep?.type === 'comma') this.next();
      else if (sep?.type === 'semicolon') {
        this.next();
        rows.push([]);
      } else if (sep?.type !== 'rbrace') this.fail('Separate array items with , (columns) or ; (rows).', sep);
    }
    const width = rows[0].length;
    if (rows.some((r) => r.length !== width)) this.fail('Every row of an array constant must have the same number of items.', open);
    return { t: 'array', rows };
  }
}

export interface ParsedFormula {
  ast: Node | null;
  error?: FormulaSyntaxError;
  tokens: Token[];
  refs: RefInfo[];
  funcs: Set<string>;
  names: Set<string>;
}

export function parseFormula(src: string): ParsedFormula {
  const tokens = tokenize(src);
  const refs: RefInfo[] = [];
  const funcs = new Set<string>();
  const names = new Set<string>();
  for (const t of tokens) {
    if (t.type === 'ref') refs.push(t.ref!);
    else if (t.type === 'func') funcs.add(normalizeFuncName(t.text));
    else if (t.type === 'name') names.add(t.text.toUpperCase());
  }
  try {
    const ast = new Parser(tokens, src).parse();
    return { ast, tokens, refs, funcs, names };
  } catch (e) {
    if (e instanceof FormulaSyntaxError) return { ast: null, error: e, tokens, refs, funcs, names };
    throw e;
  }
}

/** Tries to auto-repair common mistakes (missing closing parens/quotes). Returns null if no fix applies. */
export function autoFixFormula(src: string): string | null {
  let s = src;
  const toks = tokenize(s);
  const last = toks[toks.length - 1];
  if (last && last.type === 'str' && last.unterminated) s += '"';
  let depth = 0;
  for (const t of tokenize(s)) {
    if (t.type === 'lparen') depth++;
    else if (t.type === 'rparen') depth--;
  }
  if (depth > 0) s += ')'.repeat(depth);
  if (s === src) return null;
  return parseFormula(s).error ? null : s;
}
