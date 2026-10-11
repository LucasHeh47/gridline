// VBA parser: tokens -> module syntax tree.
import type { Arg, CaseTest, DimVar, Expr, ModuleAst, Param, Proc, Stmt } from './ast';
import { tokenize, VbaSyntaxError, type Token } from './lexer';

const BINARY: { ops: string[]; words?: boolean }[] = [
  // lowest precedence first
  { ops: ['imp'], words: true },
  { ops: ['eqv'], words: true },
  { ops: ['xor'], words: true },
  { ops: ['or'], words: true },
  { ops: ['and'], words: true },
];
const COMPARE = ['=', '<>', '<', '>', '<=', '>='];

/** Words that end an expression or statement and can't be variable names. */
const RESERVED = new Set(['then', 'to', 'step', 'else', 'elseif', 'and', 'or', 'not', 'xor', 'eqv', 'imp', 'mod', 'is', 'like', 'as', 'in', 'each', 'new']);

const END_OF: Record<string, string> = { If: 'End If', For: 'Next', Do: 'Loop', While: 'Wend', Select: 'End Select', With: 'End With' };

class Parser {
  private i = 0;
  /** Nesting depth of single-line Ifs being parsed. */
  private inline = 0;

  constructor(private toks: Token[]) {}

  private get t(): Token {
    return this.toks[this.i];
  }

  private peek(n = 1): Token {
    return this.toks[Math.min(this.i + n, this.toks.length - 1)];
  }

  private fail(msg: string, tok: Token = this.t): never {
    throw new VbaSyntaxError(msg, tok.line);
  }

  private isWord(w: string, tok: Token = this.t): boolean {
    return tok.type === 'word' && tok.lower === w;
  }

  private isOp(op: string, tok: Token = this.t): boolean {
    return tok.type === 'op' && tok.text === op;
  }

  private eatWord(w: string): boolean {
    if (this.isWord(w)) {
      this.i++;
      return true;
    }
    return false;
  }

  private expectWord(w: string, what?: string): void {
    if (!this.eatWord(w)) this.fail(`Expected ${what ?? w[0].toUpperCase() + w.slice(1)}${this.t.type === 'eof' ? ' before the end of the code' : ` but found "${this.t.text.trim() || 'end of line'}"`}.`);
  }

  private eatOp(op: string): boolean {
    if (this.isOp(op)) {
      this.i++;
      return true;
    }
    return false;
  }

  private expectOp(op: string): void {
    if (!this.eatOp(op)) this.fail(`Expected "${op}" but found "${this.t.text.trim() || 'end of line'}".`);
  }

  private atEnd(): boolean {
    return this.t.type === 'nl' || this.t.type === 'eof';
  }

  private endStatement(): void {
    if (this.t.type === 'eof') return;
    if (this.t.type !== 'nl') this.fail(`Expected the end of the statement but found "${this.t.text}".`);
    // Inside a single-line If, the line break ends the If itself, so leave it for ifStatement
    if (this.inline && this.t.text === '\n') return;
    this.i++;
  }

  private skipNewlines(): void {
    while (this.t.type === 'nl') this.i++;
  }

  private ident(what = 'a name'): string {
    const t = this.t;
    if (t.type !== 'word' || RESERVED.has(t.lower)) this.fail(`Expected ${what} but found "${t.text.trim() || 'end of line'}".`);
    this.i++;
    return t.text.replace(/[$%&!#@]$/, '');
  }

  // ---------- module ----------

  module(): ModuleAst {
    const mod: ModuleAst = { procs: [], vars: [], consts: [], optionExplicit: false, optionBase: 0, compareText: false };
    for (;;) {
      this.skipNewlines();
      if (this.t.type === 'eof') break;
      const w = this.t.lower;
      if (w === 'attribute') {
        while (!this.atEnd()) this.i++;
        continue;
      }
      if (w === 'option') {
        this.i++;
        if (this.eatWord('explicit')) mod.optionExplicit = true;
        else if (this.eatWord('base')) {
          mod.optionBase = this.t.value ?? 0;
          this.i++;
        } else if (this.eatWord('compare')) {
          mod.compareText = this.isWord('text');
          this.i++;
        } else if (this.eatWord('private')) this.eatWord('module');
        this.endStatement();
        continue;
      }
      let isPrivate = false;
      if (['private', 'public', 'global', 'friend', 'static'].includes(w)) {
        isPrivate = w === 'private';
        const next = this.peek().lower;
        if (['sub', 'function', 'property', 'const', 'declare', 'type', 'enum', 'dim', 'withevents'].includes(next) || this.peek().type === 'word') this.i++;
      }
      const head = this.t.lower;
      if (head === 'sub' || head === 'function' || head === 'property') {
        mod.procs.push(this.proc(isPrivate));
        continue;
      }
      if (head === 'const') {
        this.i++;
        for (const c of this.constItems()) mod.consts.push({ ...c, isPrivate });
        this.endStatement();
        continue;
      }
      if (head === 'declare') {
        // Windows API declarations can't run here; skip them (calls report a clear error)
        while (!this.atEnd()) this.i++;
        continue;
      }
      if (head === 'type') {
        this.skipUntilEnd('type');
        continue;
      }
      if (head === 'enum') {
        this.i++;
        this.ident('an Enum name');
        this.endStatement();
        let next = 0;
        for (;;) {
          this.skipNewlines();
          if (this.isWord('end') && this.isWord('enum', this.peek())) {
            this.i += 2;
            break;
          }
          if (this.toks[this.i].type === 'eof') this.fail('Missing End Enum.');
          const name = this.ident('an Enum member');
          let e: Expr = { k: 'lit', v: next };
          if (this.eatOp('=')) e = this.expr();
          if (e.k === 'lit' && typeof e.v === 'number') next = e.v + 1;
          else next++;
          mod.consts.push({ name, e, isPrivate });
          this.endStatement();
        }
        continue;
      }
      if (head === 'dim' || ['private', 'public', 'global'].includes(w)) {
        if (head === 'dim') this.i++;
        if (this.isWord('withevents')) this.i++;
        const line = this.t.line;
        for (const dim of this.dimList()) mod.vars.push({ dim, isPrivate, line });
        this.endStatement();
        continue;
      }
      this.fail(`"${this.t.text}" must be inside a Sub or Function.`);
    }
    return mod;
  }

  private skipUntilEnd(word: string): void {
    while (this.t.type !== 'eof' && !(this.isWord('end') && this.isWord(word, this.peek()))) this.i++;
    if (this.t.type !== 'eof') this.i += 2;
  }

  private proc(isPrivate: boolean): Proc {
    const line = this.t.line;
    const headWord = this.t.lower;
    this.i++;
    let kind: Proc['kind'] = headWord === 'sub' ? 'sub' : 'function';
    if (headWord === 'property') {
      const pk = this.t.lower;
      if (pk !== 'get' && pk !== 'let' && pk !== 'set') this.fail('Expected Get, Let or Set after Property.');
      kind = pk;
      this.i++;
    }
    const name = this.ident(headWord === 'sub' ? 'the Sub name' : 'the name');
    const params: Param[] = [];
    if (this.eatOp('(')) {
      if (!this.isOp(')')) {
        do params.push(this.param());
        while (this.eatOp(','));
      }
      this.expectOp(')');
    }
    let returnType: string | undefined;
    if (this.eatWord('as')) returnType = this.typeName();
    this.endStatement();
    const endWord = headWord === 'sub' ? 'sub' : headWord === 'function' ? 'function' : 'property';
    const body = this.block((a, b) => a === 'end' && b === endWord);
    this.expectWord('end');
    this.expectWord(endWord, `End ${endWord[0].toUpperCase() + endWord.slice(1)}`);
    this.endStatement();
    return { kind, name, params, body, isPrivate, line, returnType };
  }

  private param(): Param {
    const p: Param = { name: '', byVal: false, optional: false, dflt: null, paramArray: false, isArray: false };
    for (;;) {
      if (this.eatWord('optional')) p.optional = true;
      else if (this.eatWord('byval')) p.byVal = true;
      else if (this.eatWord('byref')) p.byVal = false;
      else if (this.eatWord('paramarray')) p.paramArray = true;
      else break;
    }
    p.name = this.ident('a parameter name');
    if (this.eatOp('(')) {
      this.expectOp(')');
      p.isArray = true;
    }
    if (this.eatWord('as')) p.type = this.typeName();
    if (this.eatOp('=')) p.dflt = this.expr();
    return p;
  }

  private typeName(): string {
    let name = this.ident('a type');
    while (this.eatOp('.')) name += '.' + this.ident('a type');
    if (this.isOp('*')) {
      // String * 10 (fixed length)
      this.i++;
      this.expr();
    }
    return name.toLowerCase();
  }

  private constItems(): { name: string; e: Expr }[] {
    const items: { name: string; e: Expr }[] = [];
    do {
      const name = this.ident('a constant name');
      if (this.eatWord('as')) this.typeName();
      this.expectOp('=');
      items.push({ name, e: this.expr() });
    } while (this.eatOp(','));
    return items;
  }

  private dimList(): DimVar[] {
    const vars: DimVar[] = [];
    do {
      if (this.isWord('withevents')) this.i++;
      const v: DimVar = { name: this.ident('a variable name') };
      if (this.eatOp('(')) {
        v.bounds = [];
        if (!this.isOp(')')) {
          do {
            const a = this.expr();
            if (this.eatWord('to')) v.bounds.push([a, this.expr()]);
            else v.bounds.push([null, a]);
          } while (this.eatOp(','));
        }
        this.expectOp(')');
      }
      if (this.eatWord('as')) {
        if (this.eatWord('new')) v.isNew = true;
        v.type = this.typeName();
      }
      vars.push(v);
    } while (this.eatOp(','));
    return vars;
  }

  // ---------- statements ----------

  /** Statements until `stop(firstWord, secondWord)` matches the next statement. */
  private block(stop: (a: string, b: string) => boolean): Stmt[] {
    const body: Stmt[] = [];
    for (;;) {
      this.skipNewlines();
      if (this.t.type === 'eof') return body;
      const a = this.t.type === 'word' ? this.t.lower : '';
      const b = this.peek().type === 'word' ? this.peek().lower : '';
      if (stop(a, b)) return body;
      body.push(...this.statement());
    }
  }

  /** Blocks being parsed, for clearer errors ("The If on line 3 is missing its End If"). */
  private open: { name: string; line: number }[] = [];

  private inBlock<T>(name: string, line: number, fn: () => T): T {
    this.open.push({ name, line });
    try {
      return fn();
    } finally {
      this.open.pop();
    }
  }

  private whileStatement(line: number): Stmt[] {
    this.i++;
    const cond = this.expr();
    this.endStatement();
    const body = this.block((a) => a === 'wend');
    this.expectWord('wend');
    this.endSimple();
    return [{ k: 'do', pre: { until: false, cond }, post: null, body, line }];
  }

  private withStatement(line: number): Stmt[] {
    this.i++;
    const obj = this.expr();
    this.endStatement();
    const body = this.block((a, b) => a === 'end' && b === 'with');
    this.expectWord('end');
    this.expectWord('with', 'End With');
    this.endSimple();
    return [{ k: 'with', obj, body, line }];
  }

  /** One statement (a single-line If can hold several). */
  private statement(): Stmt[] {
    const t = this.t;
    const line = t.line;
    const w = t.type === 'word' ? t.lower : '';
    // Label: "Name:" at the start of a line
    const prev = this.toks[this.i - 1];
    if (t.type === 'word' && this.peek().type === 'nl' && this.peek().text === ':' && (!prev || prev.text === '\n') && !RESERVED.has(w)) {
      this.i += 2;
      return [{ k: 'label', name: t.text.toLowerCase(), line }];
    }
    if (t.type === 'num' && prev?.text !== ':' && this.peek().type === 'word') {
      // Line numbers (10 Print …) are ignored
      this.i++;
      return this.statement();
    }
    switch (w) {
      case 'dim':
      case 'static':
      case 'private':
      case 'public': {
        this.i++;
        if (this.isWord('const')) {
          this.i++;
          const items = this.constItems();
          this.endStatement();
          return [{ k: 'const', items, line }];
        }
        const vars = this.dimList();
        this.endStatement();
        return [{ k: 'dim', vars, isStatic: w === 'static', line }];
      }
      case 'const': {
        this.i++;
        const items = this.constItems();
        this.endStatement();
        return [{ k: 'const', items, line }];
      }
      case 'redim': {
        this.i++;
        const preserve = this.eatWord('preserve');
        const vars = this.dimList();
        this.endStatement();
        return [{ k: 'redim', preserve, vars, line }];
      }
      case 'set':
      case 'let': {
        this.i++;
        const target = this.postfix();
        this.expectOp('=');
        const e = this.expr();
        this.endSimple();
        return [{ k: 'assign', target, e, set: w === 'set', line }];
      }
      case 'call': {
        this.i++;
        const e = this.postfix();
        this.endSimple();
        return [{ k: 'call', e: e.k === 'call' ? e : { k: 'call', target: e, args: [], line }, line }];
      }
      case 'if':
        return this.inBlock('If', line, () => [this.ifStatement()]);
      case 'for':
        return this.inBlock('For', line, () => [this.forStatement()]);
      case 'do':
        return this.inBlock('Do', line, () => [this.doStatement()]);
      case 'while':
        return this.inBlock('While', line, () => this.whileStatement(line));
      case 'select':
        return this.inBlock('Select', line, () => [this.selectStatement()]);
      case 'with':
        return this.inBlock('With', line, () => this.withStatement(line));
      case 'exit': {
        this.i++;
        const what = this.t.lower as 'sub' | 'function' | 'for' | 'do' | 'property';
        if (!['sub', 'function', 'for', 'do', 'property'].includes(what)) this.fail('Expected Sub, Function, For or Do after Exit.');
        this.i++;
        this.endSimple();
        return [{ k: 'exit', what, line }];
      }
      case 'on': {
        this.i++;
        if (!this.eatWord('error')) this.fail('Only "On Error" statements are supported.');
        if (this.eatWord('resume')) {
          this.expectWord('next');
          this.endSimple();
          return [{ k: 'onerror', mode: 'resumenext', line }];
        }
        this.expectWord('goto', 'GoTo');
        if (this.t.type === 'num' && this.t.value === 0) {
          this.i++;
          this.endSimple();
          return [{ k: 'onerror', mode: 'goto0', line }];
        }
        if (this.isOp('-') && this.peek().value === 1) {
          this.i += 2;
          this.endSimple();
          return [{ k: 'onerror', mode: 'goto0', line }];
        }
        const label = this.t.text.toLowerCase();
        this.i++;
        this.endSimple();
        return [{ k: 'onerror', mode: 'label', label, line }];
      }
      case 'resume': {
        this.i++;
        if (this.eatWord('next')) {
          this.endSimple();
          return [{ k: 'resume', next: true, line }];
        }
        const label = this.atEnd() || this.isWord('else') ? undefined : (this.i++, this.toks[this.i - 1].text.toLowerCase());
        this.endSimple();
        return [{ k: 'resume', next: false, label, line }];
      }
      case 'goto': {
        this.i++;
        const label = this.t.text.toLowerCase();
        this.i++;
        this.endSimple();
        return [{ k: 'goto', label, line }];
      }
      case 'end':
        if (this.peek().type === 'nl' || this.peek().type === 'eof') {
          this.i++;
          this.endSimple();
          return [{ k: 'end', line }];
        }
        {
          const open = this.open[this.open.length - 1];
          if (open) this.fail(`The ${open.name} on line ${open.line} is missing its ${END_OF[open.name]} (found "End ${this.peek().text}" first).`);
        }
        this.fail(`Unexpected "End ${this.peek().text}" — it doesn't match an open block.`);
        break;
      case 'debug':
        if (this.isOp('.', this.peek()) && this.peek(2).lower === 'print') {
          this.i += 3;
          const args: Expr[] = [];
          while (!this.atEnd() && !this.isWord('else')) {
            if (this.eatOp(';') || this.eatOp(',')) continue;
            args.push(this.expr());
          }
          this.endSimple();
          return [{ k: 'print', args, line }];
        }
        break;
      case 'else':
      case 'elseif':
      case 'next':
      case 'loop':
      case 'wend':
      case 'case':
        this.fail(`"${t.text}" without a matching ${w === 'next' ? 'For' : w === 'loop' ? 'Do' : w === 'wend' ? 'While' : w === 'case' ? 'Select Case' : 'If'}.`);
    }
    // Assignment or call
    const target = this.postfix();
    if (this.eatOp('=')) {
      const e = this.expr();
      this.endSimple();
      return [{ k: 'assign', target, e, set: false, line }];
    }
    if (this.atEnd() || this.isWord('else')) {
      this.endSimple();
      return [{ k: 'call', e: target.k === 'call' ? target : { k: 'call', target, args: [], line }, line }];
    }
    // Call without parentheses: Foo a, b, Name:=c
    const args = this.argList(false);
    this.endSimple();
    if (target.k === 'call' && target.args.length === 1 && !target.args[0].name) {
      // "Foo (1), 2": the first argument was in parentheses
      return [{ k: 'call', e: { k: 'call', target: target.target, args: [target.args[0], ...args], line }, line }];
    }
    return [{ k: 'call', e: { k: 'call', target, args, line }, line }];
  }

  /** End of a simple statement: newline, or "Else" inside a single-line If. */
  private endSimple(): void {
    if (this.isWord('else')) return;
    this.endStatement();
  }

  private ifStatement(): Stmt {
    const line = this.t.line;
    this.i++;
    const cond = this.expr();
    this.expectWord('then');
    if (!this.atEnd()) {
      // Single-line If: If x Then a: b Else c
      this.inline++;
      try {
        return this.singleLineIf(cond, line);
      } finally {
        this.inline--;
      }
    }
    this.endStatement();
    const branches: { cond: Expr; body: Stmt[] }[] = [];
    let elseBody: Stmt[] | null = null;
    let c = cond;
    for (;;) {
      const body = this.block((a, b) => a === 'elseif' || a === 'else' || (a === 'end' && b === 'if') || a === 'endif');
      branches.push({ cond: c, body });
      if (this.eatWord('elseif')) {
        c = this.expr();
        this.expectWord('then');
        this.endStatement();
        continue;
      }
      if (this.eatWord('else')) {
        if (this.isWord('if')) {
          // "Else If" on one line works like ElseIf when the If is a block
          this.i++;
          c = this.expr();
          this.expectWord('then');
          this.endStatement();
          continue;
        }
        this.endStatement();
        elseBody = this.block((a, b) => (a === 'end' && b === 'if') || a === 'endif');
      }
      break;
    }
    if (this.eatWord('endif')) {
      this.endStatement();
      return { k: 'if', branches, elseBody, line };
    }
    this.expectWord('end', 'End If');
    this.expectWord('if', 'End If');
    this.endSimple();
    return { k: 'if', branches, elseBody, line };
  }

  /** A statement inside a single-line If (no block statements). */
  private singleLineIf(cond: Expr, line: number): Stmt {
    const body: Stmt[] = [];
    while (!this.isWord('else') && !(this.t.type === 'nl' && this.t.text === '\n') && this.t.type !== 'eof') {
      if (this.t.type === 'nl') {
        this.i++;
        continue;
      }
      body.push(...this.statementInline());
    }
    let elseBody: Stmt[] | null = null;
    if (this.eatWord('else')) {
      elseBody = [];
      while (!(this.t.type === 'nl' && this.t.text === '\n') && this.t.type !== 'eof') {
        if (this.t.type === 'nl') {
          this.i++;
          continue;
        }
        elseBody.push(...this.statementInline());
      }
    }
    if (this.t.type === 'nl' && this.inline === 1) this.i++;
    return { k: 'if', branches: [{ cond, body }], elseBody, line };
  }

  private statementInline(): Stmt[] {
    return this.statement();
  }

  private forStatement(): Stmt {
    const line = this.t.line;
    this.i++;
    if (this.eatWord('each')) {
      const v: Expr = { k: 'name', name: this.ident('a variable name'), line };
      this.expectWord('in');
      const coll = this.expr();
      this.endStatement();
      const body = this.block((a) => a === 'next');
      this.expectWord('next');
      if (!this.atEnd()) this.ident();
      this.endSimple();
      return { k: 'foreach', v, coll, body, line };
    }
    const v: Expr = { k: 'name', name: this.ident('a counter variable'), line };
    this.expectOp('=');
    const from = this.expr();
    this.expectWord('to');
    const to = this.expr();
    const step = this.eatWord('step') ? this.expr() : null;
    this.endStatement();
    const body = this.block((a) => a === 'next');
    this.expectWord('next');
    // "Next i" or "Next j, i" (closes two loops — rare; only the first name is checked)
    if (!this.atEnd() && !this.isWord('else')) {
      this.ident();
      while (this.eatOp(',')) this.ident();
    }
    this.endSimple();
    return { k: 'for', v, from, to, step, body, line };
  }

  private doStatement(): Stmt {
    const line = this.t.line;
    this.i++;
    let pre: { until: boolean; cond: Expr } | null = null;
    if (this.isWord('while') || this.isWord('until')) {
      const until = this.t.lower === 'until';
      this.i++;
      pre = { until, cond: this.expr() };
    }
    this.endStatement();
    const body = this.block((a) => a === 'loop');
    this.expectWord('loop');
    let post: { until: boolean; cond: Expr } | null = null;
    if (this.isWord('while') || this.isWord('until')) {
      const until = this.t.lower === 'until';
      this.i++;
      post = { until, cond: this.expr() };
    }
    this.endSimple();
    return { k: 'do', pre, post, body, line };
  }

  private selectStatement(): Stmt {
    const line = this.t.line;
    this.i++;
    this.expectWord('case', 'Select Case');
    const e = this.expr();
    this.endStatement();
    const cases: { tests: CaseTest[]; body: Stmt[] }[] = [];
    let elseBody: Stmt[] | null = null;
    for (;;) {
      this.skipNewlines();
      if (this.isWord('end') && this.isWord('select', this.peek())) break;
      if (this.t.type === 'eof') this.fail('Missing End Select.');
      this.expectWord('case');
      if (this.eatWord('else')) {
        if (this.t.type === 'nl') this.i++;
        elseBody = this.block((a, b) => a === 'case' || (a === 'end' && b === 'select'));
        continue;
      }
      const tests: CaseTest[] = [];
      do {
        if (this.eatWord('is')) {
          const op = this.t.text;
          if (!COMPARE.includes(op)) this.fail('Expected a comparison after Case Is.');
          this.i++;
          tests.push({ k: 'is', op, e: this.expr() });
        } else {
          const a = this.expr();
          if (this.eatWord('to')) tests.push({ k: 'range', from: a, to: this.expr() });
          else tests.push({ k: 'value', e: a });
        }
      } while (this.eatOp(','));
      if (this.t.type === 'nl') this.i++;
      const body = this.block((a, b) => a === 'case' || (a === 'end' && b === 'select'));
      cases.push({ tests, body });
    }
    this.i += 2;
    this.endSimple();
    return { k: 'select', e, cases, elseBody, line };
  }

  // ---------- expressions ----------

  expr(): Expr {
    return this.binaryLevel(0);
  }

  private binaryLevel(level: number): Expr {
    if (level >= BINARY.length) return this.notLevel();
    let a = this.binaryLevel(level + 1);
    for (;;) {
      const t = this.t;
      if (t.type === 'word' && BINARY[level].ops.includes(t.lower)) {
        this.i++;
        const b = this.binaryLevel(level + 1);
        a = { k: 'bin', op: t.lower, a, b, line: t.line };
      } else return a;
    }
  }

  private notLevel(): Expr {
    if (this.eatWord('not')) return { k: 'unary', op: 'not', e: this.notLevel() };
    return this.compare();
  }

  private compare(): Expr {
    let a = this.concat();
    for (;;) {
      const t = this.t;
      if ((t.type === 'op' && COMPARE.includes(t.text)) || (t.type === 'word' && (t.lower === 'like' || t.lower === 'is'))) {
        this.i++;
        const b = this.concat();
        a = { k: 'bin', op: t.type === 'op' ? t.text : t.lower, a, b, line: t.line };
      } else return a;
    }
  }

  private concat(): Expr {
    let a = this.additive();
    while (this.isOp('&')) {
      const line = this.t.line;
      this.i++;
      a = { k: 'bin', op: '&', a, b: this.additive(), line };
    }
    return a;
  }

  private additive(): Expr {
    let a = this.modulo();
    while (this.isOp('+') || this.isOp('-')) {
      const t = this.t;
      this.i++;
      a = { k: 'bin', op: t.text, a, b: this.modulo(), line: t.line };
    }
    return a;
  }

  private modulo(): Expr {
    let a = this.intDiv();
    while (this.isWord('mod')) {
      const line = this.t.line;
      this.i++;
      a = { k: 'bin', op: 'mod', a, b: this.intDiv(), line };
    }
    return a;
  }

  private intDiv(): Expr {
    let a = this.multiplicative();
    while (this.isOp('\\')) {
      const line = this.t.line;
      this.i++;
      a = { k: 'bin', op: '\\', a, b: this.multiplicative(), line };
    }
    return a;
  }

  private multiplicative(): Expr {
    let a = this.unary();
    while (this.isOp('*') || this.isOp('/')) {
      const t = this.t;
      this.i++;
      a = { k: 'bin', op: t.text, a, b: this.unary(), line: t.line };
    }
    return a;
  }

  private unary(): Expr {
    if (this.eatOp('-')) return { k: 'unary', op: '-', e: this.unary() };
    if (this.eatOp('+')) return this.unary();
    return this.power();
  }

  private power(): Expr {
    let a = this.postfix();
    while (this.isOp('^')) {
      const line = this.t.line;
      this.i++;
      // -2^2 is -4: the exponent binds tighter than unary minus
      const b = this.isOp('-') ? (this.i++, { k: 'unary' as const, op: '-', e: this.postfix() }) : this.postfix();
      a = { k: 'bin', op: '^', a, b, line };
    }
    return a;
  }

  postfix(): Expr {
    let e = this.primary();
    for (;;) {
      const t = this.t;
      if (this.isOp('(') && e.k !== 'lit' && e.k !== 'special') {
        this.i++;
        const args = this.argList(true);
        this.expectOp(')');
        e = { k: 'call', target: e, args, line: t.line };
      } else if ((this.isOp('.') || this.isOp('!')) && !t.spaced) {
        this.i++;
        const nt = this.t;
        if (nt.type !== 'word') this.fail('Expected a name after "."');
        this.i++;
        e = { k: 'member', obj: e, name: nt.text.replace(/[$%&!#@]$/, ''), bang: t.text === '!', line: nt.line };
      } else return e;
    }
  }

  private argList(inParens: boolean): Arg[] {
    const args: Arg[] = [];
    if (inParens && this.isOp(')')) return args;
    for (;;) {
      if (this.isOp(',')) {
        args.push({});
        this.i++;
        continue;
      }
      if (inParens ? this.isOp(')') : this.atEnd() || this.isWord('else')) {
        // trailing empty argument after a comma
        if (args.length && this.toks[this.i - 1].text === ',') args.push({});
        break;
      }
      let name: string | undefined;
      if (this.t.type === 'word' && this.isOp(':=', this.peek())) {
        name = this.t.lower;
        this.i += 2;
      }
      args.push({ expr: this.expr(), name });
      if (!this.eatOp(',')) break;
      if (inParens ? this.isOp(')') : this.atEnd()) args.push({});
    }
    return args;
  }

  private primary(): Expr {
    const t = this.t;
    switch (t.type) {
      case 'num':
        this.i++;
        return { k: 'lit', v: t.value! };
      case 'str':
        this.i++;
        return { k: 'lit', v: t.text };
      case 'date':
        this.i++;
        return { k: 'lit', v: t.value!, date: true };
      case 'op':
        if (t.text === '(') {
          this.i++;
          const e = this.expr();
          this.expectOp(')');
          return e;
        }
        if (t.text === '.' || t.text === '!') {
          // Inside With: .Value
          this.i++;
          const nt = this.t;
          if (nt.type !== 'word') this.fail('Expected a name after "."');
          this.i++;
          return { k: 'member', obj: null, name: nt.text.replace(/[$%&!#@]$/, ''), bang: t.text === '!', line: nt.line };
        }
        break;
      case 'word': {
        const w = t.lower;
        if (w === 'true' || w === 'false') {
          this.i++;
          return { k: 'lit', v: w === 'true' };
        }
        if (w === 'nothing' || w === 'empty' || w === 'null' || w === 'me') {
          this.i++;
          return { k: 'special', v: w };
        }
        if (w === 'new') {
          this.i++;
          return { k: 'new', cls: this.typeName(), line: t.line };
        }
        if (t.text.startsWith('[')) {
          this.i++;
          return { k: 'eval', text: t.text.slice(1, -1), line: t.line };
        }
        if (RESERVED.has(w)) break;
        this.i++;
        return { k: 'name', name: t.text.replace(/[$%&!#@]$/, ''), line: t.line };
      }
    }
    this.fail(t.type === 'nl' || t.type === 'eof' ? 'This line ends too early — something is missing.' : `Unexpected "${t.text}".`);
  }
}

/** Parses a whole module. Throws VbaSyntaxError with a line number. */
export function parseModule(src: string): ModuleAst {
  return new Parser(tokenize(src)).module();
}

/** Parses a single expression (used for tests and the immediate window). */
export function parseExpression(src: string): Expr {
  const p = new Parser(tokenize(src));
  return p.expr();
}
