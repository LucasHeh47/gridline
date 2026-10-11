// Syntax tree for VBA modules.

export type Expr =
  | { k: 'lit'; v: number | string | boolean; date?: boolean }
  | { k: 'special'; v: 'nothing' | 'empty' | 'null' | 'me' }
  | { k: 'name'; name: string; line: number }
  /** obj.name — `obj` null means the current With object (".Value"). */
  | { k: 'member'; obj: Expr | null; name: string; bang?: boolean; line: number }
  /** target(args) — a call, an array index or a default-member access. */
  | { k: 'call'; target: Expr; args: Arg[]; line: number }
  | { k: 'unary'; op: string; e: Expr }
  | { k: 'bin'; op: string; a: Expr; b: Expr; line: number }
  | { k: 'new'; cls: string; line: number }
  /** [A1] shorthand for Evaluate("A1"). */
  | { k: 'eval'; text: string; line: number };

export interface Arg {
  expr?: Expr;
  /** Named argument (Destination:=…). */
  name?: string;
}

export interface DimVar {
  name: string;
  /** Array bounds: [lower, upper] per dimension; empty for a dynamic array `a()`. */
  bounds?: [Expr | null, Expr][];
  type?: string;
  isNew?: boolean;
}

export type CaseTest = { k: 'is'; op: string; e: Expr } | { k: 'range'; from: Expr; to: Expr } | { k: 'value'; e: Expr };

export type Stmt = { line: number } & (
  | { k: 'dim'; vars: DimVar[]; isStatic?: boolean }
  | { k: 'const'; items: { name: string; e: Expr }[] }
  | { k: 'redim'; preserve: boolean; vars: DimVar[] }
  | { k: 'assign'; target: Expr; e: Expr; set: boolean }
  | { k: 'call'; e: Expr }
  | { k: 'if'; branches: { cond: Expr; body: Stmt[] }[]; elseBody: Stmt[] | null }
  | { k: 'for'; v: Expr; from: Expr; to: Expr; step: Expr | null; body: Stmt[] }
  | { k: 'foreach'; v: Expr; coll: Expr; body: Stmt[] }
  | { k: 'do'; pre: { until: boolean; cond: Expr } | null; post: { until: boolean; cond: Expr } | null; body: Stmt[] }
  | { k: 'select'; e: Expr; cases: { tests: CaseTest[]; body: Stmt[] }[]; elseBody: Stmt[] | null }
  | { k: 'with'; obj: Expr; body: Stmt[] }
  | { k: 'exit'; what: 'sub' | 'function' | 'for' | 'do' | 'property' }
  | { k: 'onerror'; mode: 'resumenext' | 'goto0' | 'label'; label?: string }
  | { k: 'resume'; next: boolean; label?: string }
  | { k: 'goto'; label: string }
  | { k: 'label'; name: string }
  | { k: 'end' }
  | { k: 'print'; args: Expr[] }
);

export interface Param {
  name: string;
  byVal: boolean;
  optional: boolean;
  dflt: Expr | null;
  paramArray: boolean;
  isArray: boolean;
  type?: string;
}

export interface Proc {
  kind: 'sub' | 'function' | 'get' | 'let' | 'set';
  name: string;
  params: Param[];
  body: Stmt[];
  isPrivate: boolean;
  line: number;
  returnType?: string;
}

export interface ModuleAst {
  procs: Proc[];
  vars: { dim: DimVar; isPrivate: boolean; line: number }[];
  consts: { name: string; e: Expr; isPrivate: boolean }[];
  optionExplicit: boolean;
  optionBase: number;
  compareText: boolean;
}
