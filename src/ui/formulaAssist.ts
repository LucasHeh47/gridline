// Helpers for formula editing: syntax highlighting, reference colors, function context.
import { normalizeFuncName } from '../core/formula/parser';
import { tokenize, type RefInfo, type Token } from '../core/formula/tokenizer';
import { escapeHtml } from './dom';
import { REF_COLORS } from './grid/GridView';

export interface RefSpan {
  start: number;
  end: number;
  ref: RefInfo;
  color: string;
}

/** Tokens of the formula part (text after '='), with positions shifted to the full text. */
function formulaTokens(text: string): Token[] {
  if (!text.startsWith('=')) return [];
  return tokenize(text.slice(1)).map((t) => ({ ...t, start: t.start + 1, end: t.end + 1 }));
}

/** References in a formula, each with a stable color (same text = same color). */
export function refSpans(text: string): RefSpan[] {
  const out: RefSpan[] = [];
  const colorOf = new Map<string, string>();
  for (const t of formulaTokens(text)) {
    if (t.type !== 'ref' || !t.ref) continue;
    const key = t.text.replace(/\$/g, '').toUpperCase();
    let color = colorOf.get(key);
    if (!color) {
      color = REF_COLORS[colorOf.size % REF_COLORS.length];
      colorOf.set(key, color);
    }
    out.push({ start: t.start, end: t.end, ref: t.ref, color });
  }
  return out;
}

/** HTML for the highlight layer behind the editor textarea. */
export function highlightHtml(text: string, caret: number): string {
  if (!text.startsWith('=')) return escapeHtml(text) + '​';
  const toks = formulaTokens(text);
  const refs = refSpans(text);
  const refColor = new Map(refs.map((r) => [r.start, r.color]));
  // Matching parenthesis for the caret
  const parenPairs = new Map<number, number>();
  const stack: number[] = [];
  for (const t of toks) {
    if (t.type === 'lparen') stack.push(t.start);
    else if (t.type === 'rparen' && stack.length) {
      const o = stack.pop()!;
      parenPairs.set(o, t.start);
      parenPairs.set(t.start, o);
    }
  }
  let match = new Set<number>();
  for (const pos of [caret - 1, caret]) {
    if (parenPairs.has(pos)) {
      match = new Set([pos, parenPairs.get(pos)!]);
      break;
    }
  }
  let html = '<span class="hl-op">=</span>';
  for (const t of toks) {
    const txt = escapeHtml(t.text);
    let cls = '';
    let style = '';
    switch (t.type) {
      case 'ref':
        style = `color:${refColor.get(t.start)}`;
        break;
      case 'func':
        cls = 'hl-func';
        break;
      case 'str':
        cls = 'hl-str';
        break;
      case 'num':
        cls = 'hl-num';
        break;
      case 'err':
        cls = 'hl-err';
        break;
      case 'bool':
        cls = 'hl-bool';
        break;
      case 'op':
      case 'comma':
        cls = 'hl-op';
        break;
      case 'lparen':
      case 'rparen':
        if (match.has(t.start)) cls = 'hl-paren-match';
        break;
    }
    html += cls || style ? `<span${cls ? ` class="${cls}"` : ''}${style ? ` style="${style}"` : ''}>${txt}</span>` : txt;
  }
  return html + '​';
}

export interface FuncContext {
  name: string;
  argIndex: number;
  start: number;
}

/** The innermost function call surrounding the caret, and which argument the caret is in. */
export function functionAtCaret(text: string, caret: number): FuncContext | null {
  if (!text.startsWith('=')) return null;
  const toks = formulaTokens(text).filter((t) => t.end <= caret);
  const stack: { name: string | null; argIndex: number; start: number; brace: boolean }[] = [];
  let lastFunc: Token | null = null;
  for (const t of toks) {
    if (t.type === 'ws') continue;
    if (t.type === 'func') {
      lastFunc = t;
      continue;
    }
    if (t.type === 'lparen') stack.push({ name: lastFunc ? normalizeFuncName(lastFunc.text) : null, argIndex: 0, start: lastFunc?.start ?? t.start, brace: false });
    else if (t.type === 'lbrace') stack.push({ name: null, argIndex: 0, start: t.start, brace: true });
    else if (t.type === 'rparen' || t.type === 'rbrace') stack.pop();
    else if (t.type === 'comma' && stack.length) stack[stack.length - 1].argIndex++;
    lastFunc = null;
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    const s = stack[i];
    if (s.brace) return null;
    if (s.name) return { name: s.name, argIndex: s.argIndex, start: s.start };
  }
  return null;
}

/** The identifier being typed at the caret (for autocomplete). */
export function identifierAtCaret(text: string, caret: number): { start: number; end: number; word: string } | null {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  const m = /[A-Za-z_][A-Za-z0-9_.]*$/.exec(before);
  if (!m) return null;
  // Not inside a string
  const quotes = (before.match(/"/g) ?? []).length;
  if (quotes % 2) return null;
  const start = m.index;
  const prev = text[start - 1];
  if (prev && /[A-Za-z0-9_.$'!:]/.test(prev)) return null;
  const after = /^[A-Za-z0-9_.]*/.exec(text.slice(caret))![0];
  // Pure cell references like A1 don't need suggestions
  const word = m[0];
  if (/^[A-Za-z]{1,3}\d+$/.test(word + after)) return null;
  return { start, end: caret + after.length, word };
}

/** Whether a reference can be inserted at the caret (after an operator, comma or open paren). */
export function canInsertRefAt(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false;
  const before = text.slice(0, caret).replace(/\s+$/, '');
  if (before === '=') return true;
  const last = before[before.length - 1];
  if (!last) return false;
  const quotes = (before.match(/"/g) ?? []).length;
  if (quotes % 2) return false;
  return '=(,+-*/^&<>:;{'.includes(last);
}

/** If the caret is inside a quoted text argument, where it starts and what's typed so far. */
export function stringArgAtCaret(text: string, caret: number): { quote: number; partial: string; end: number; closed: boolean } | null {
  if (!text.startsWith('=')) return null;
  let open = -1;
  for (let i = 1; i < caret; i++) {
    if (text[i] !== '"') continue;
    if (open < 0) open = i;
    else if (text[i + 1] === '"' && i + 1 < caret) i++;
    else open = -1;
  }
  if (open < 0) return null;
  const close = text.indexOf('"', caret);
  return { quote: open, partial: text.slice(open + 1, caret), end: close < 0 ? text.length : close, closed: close >= 0 };
}
