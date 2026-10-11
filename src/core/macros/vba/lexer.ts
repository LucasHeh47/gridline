// VBA tokenizer: turns module source into tokens, joining " _" line continuations and dropping comments.

export type TokType = 'word' | 'num' | 'str' | 'date' | 'op' | 'nl' | 'eof';

export interface Token {
  type: TokType;
  /** Source text (for words: as written; `lower` is the normalized form). */
  text: string;
  lower: string;
  /** Numeric value for numbers and dates (dates as an Excel/VBA serial). */
  value?: number;
  line: number;
  /** True when whitespace came right before this token (matters for `Foo (x)`). */
  spaced: boolean;
}

export class VbaSyntaxError extends Error {
  constructor(
    message: string,
    readonly line: number,
  ) {
    super(message);
  }
}

const OPS = ['<>', '<=', '>=', ':=', '=', '<', '>', '+', '-', '*', '/', '\\', '^', '&', '(', ')', ',', '.', '!', ':', ';', '#'];

/** Serial number for a VBA date literal like #3/14/2026# or #2026-03-14 14:30#. */
function parseDateLiteral(text: string, line: number): number {
  const t = text.trim();
  let y = 1899;
  let m = 12;
  let d = 30;
  let rest = t;
  let match = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*(.*)$/.exec(t);
  if (match) {
    m = +match[1];
    d = +match[2];
    y = +match[3];
    if (y < 100) y += y < 30 ? 2000 : 1900;
    rest = match[4];
  } else if ((match = /^(\d{4})-(\d{1,2})-(\d{1,2})\s*(.*)$/.exec(t))) {
    y = +match[1];
    m = +match[2];
    d = +match[3];
    rest = match[4];
  }
  let time = 0;
  const tm = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(rest);
  if (tm) {
    let h = +tm[1];
    if (tm[4]) {
      if (/pm/i.test(tm[4]) && h < 12) h += 12;
      if (/am/i.test(tm[4]) && h === 12) h = 0;
    }
    time = (h * 3600 + +tm[2] * 60 + +(tm[3] ?? 0)) / 86400;
  } else if (rest) throw new VbaSyntaxError(`Invalid date #${text}#`, line);
  const ms = Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30);
  return Math.round(ms / 86400000) + time;
}

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  let line = 1;
  const n = src.length;
  let spaced = false;
  const push = (type: TokType, text: string, value?: number) => {
    out.push({ type, text, lower: text.toLowerCase(), value, line, spaced });
    spaced = false;
  };
  const atStatementStart = () => {
    const last = out[out.length - 1];
    return !last || last.type === 'nl';
  };
  while (i < n) {
    const ch = src[i];
    // Line continuation: " _" followed by end of line
    if (ch === '_' && (i === 0 || /[ \t]/.test(src[i - 1]))) {
      let j = i + 1;
      while (j < n && (src[j] === ' ' || src[j] === '\t')) j++;
      if (j >= n || src[j] === '\r' || src[j] === '\n') {
        if (src[j] === '\r' && src[j + 1] === '\n') j++;
        i = j + 1;
        line++;
        spaced = true;
        continue;
      }
    }
    if (ch === ' ' || ch === '\t' || ch === ' ') {
      i++;
      spaced = true;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      i++;
      // Keep the line break after a ":" so single-line If statements know where they end
      const last = out[out.length - 1];
      if (last && (last.type !== 'nl' || last.text === ':')) push('nl', '\n');
      line++;
      continue;
    }
    // Comment
    if (ch === "'") {
      while (i < n && src[i] !== '\r' && src[i] !== '\n') i++;
      continue;
    }
    // String
    if (ch === '"') {
      let s = '';
      i++;
      for (;;) {
        if (i >= n || src[i] === '\r' || src[i] === '\n') throw new VbaSyntaxError('A text value is missing its closing quote (").', line);
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            s += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += src[i++];
      }
      push('str', s);
      continue;
    }
    // Date literal #...# (a lone # is the file-number operator, which isn't supported anyway)
    if (ch === '#') {
      const end = src.indexOf('#', i + 1);
      const body = end > 0 ? src.slice(i + 1, end) : '';
      if (end > 0 && !/[\r\n]/.test(body) && /\d/.test(body)) {
        push('date', body, parseDateLiteral(body, line));
        i = end + 1;
        continue;
      }
    }
    // Numbers (decimal, &H hex, &O octal)
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      const m = /^(\d*\.?\d+(?:[eE][+-]?\d+)?|\d+\.)[%&@!#^]?/.exec(src.slice(i));
      const text = m![0];
      push('num', text, parseFloat(text.replace(/[%&@!#^]$/, '')));
      i += text.length;
      continue;
    }
    if (ch === '&' && /[hHoO]/.test(src[i + 1] ?? '')) {
      const m = /^&([hH][0-9a-fA-F]+|[oO][0-7]+)&?/.exec(src.slice(i));
      if (m) {
        const body = m[1];
        const v = /^[hH]/.test(body) ? parseInt(body.slice(1), 16) : parseInt(body.slice(1), 8);
        // &HFFFF is -1 as an Integer, &HFFFFFFFF as a Long (16/32-bit two's complement)
        const len = body.length - 1;
        const signed = len <= 4 && /^[hH]/.test(body) && !m[0].endsWith('&') ? (v >= 0x8000 ? v - 0x10000 : v) : v >= 0x80000000 ? v - 0x100000000 : v;
        push('num', m[0], signed);
        i += m[0].length;
        continue;
      }
    }
    // Words (identifiers and keywords), with an optional type suffix ($ % & ! # @)
    if (/[A-Za-z_À-￿]/.test(ch)) {
      const m = /^[A-Za-z_À-￿][A-Za-z0-9_À-￿]*[$%&!#@]?/.exec(src.slice(i))!;
      let text = m[0];
      // "x&y" is concatenation, not a Long suffix, when followed by a word
      if (/[&!#@]$/.test(text) && /[A-Za-z0-9_"(]/.test(src[i + text.length] ?? '')) text = text.slice(0, -1);
      // Rem comment at the start of a statement
      if (text.toLowerCase() === 'rem' && atStatementStart()) {
        while (i < n && src[i] !== '\r' && src[i] !== '\n') i++;
        continue;
      }
      push('word', text);
      i += text.length;
      continue;
    }
    // Bracketed names: [A1] evaluates, [My Sheet] names
    if (ch === '[') {
      const end = src.indexOf(']', i);
      if (end < 0) throw new VbaSyntaxError('Missing ]', line);
      push('word', src.slice(i, end + 1));
      i = end + 1;
      continue;
    }
    const op = OPS.find((o) => src.startsWith(o, i));
    if (!op) throw new VbaSyntaxError(`Unexpected character "${ch}"`, line);
    // A colon separates statements (except in := named arguments)
    if (op === ':') {
      push('nl', ':');
      i++;
      continue;
    }
    push('op', op);
    i += op.length;
  }
  if (out.length && out[out.length - 1].type !== 'nl') push('nl', '\n');
  push('eof', '');
  return out;
}
