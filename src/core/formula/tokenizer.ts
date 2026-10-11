// Formula tokenizer. Lossless: concatenating token texts reproduces the input exactly,
// which lets reference-rewriting (insert rows, copy/paste, rename sheet) preserve formatting.
import { MAX_COLS, MAX_ROWS, nameToCol, quoteSheetName, colToName } from '../address';
import { parseExcelNumber } from './excelMath';

export type TokType =
  | 'num'
  | 'str'
  | 'bool'
  | 'err'
  | 'ref'
  | 'func'
  | 'name'
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'comma'
  | 'semicolon'
  | 'lbrace'
  | 'rbrace'
  | 'ws'
  | 'unknown';

export interface RefInfo {
  sheet?: string;
  kind: 'cell' | 'range' | 'cols' | 'rows';
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  ar1: boolean;
  ac1: boolean;
  ar2: boolean;
  ac2: boolean;
}

export interface Token {
  type: TokType;
  text: string;
  start: number;
  end: number;
  ref?: RefInfo;
  /** For strings: the unescaped value. For numbers: numeric value. */
  value?: string | number | boolean;
  /** For strings: whether the closing quote is missing. */
  unterminated?: boolean;
}

const ERROR_LIT = /#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|CIRC!|GETTING_DATA)/iy;
const QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
const UNQUOTED_SHEET = /([A-Za-z_À-￿][\w.À-￿]*)!/y;
const CELL_RANGE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?/y;
const COL_RANGE = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const ROW_RANGE = /(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;
const NUMBER = /(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/y;
const IDENT = /[A-Za-z_\\À-￿][\w.\\À-￿]*/y;
const AFTER_REF_BAD = /[A-Za-z0-9_.(À-￿]/;

function matchAt(re: RegExp, s: string, pos: number): RegExpExecArray | null {
  re.lastIndex = pos;
  return re.exec(s);
}

/** Tries to read a reference (without sheet prefix) at pos. */
function readRef(s: string, pos: number): { len: number; ref: RefInfo } | null {
  // Column range A:C
  let m = matchAt(COL_RANGE, s, pos);
  if (m && !AFTER_REF_BAD.test(s[pos + m[0].length] ?? '')) {
    const c1 = nameToCol(m[2]);
    const c2 = nameToCol(m[4]);
    if (c1 >= 0 && c2 >= 0 && c1 < MAX_COLS && c2 < MAX_COLS) {
      return {
        len: m[0].length,
        ref: {
          kind: 'cols',
          r1: 0,
          r2: MAX_ROWS - 1,
          c1: Math.min(c1, c2),
          c2: Math.max(c1, c2),
          ar1: true,
          ar2: true,
          ac1: !!m[1],
          ac2: !!m[3],
        },
      };
    }
  }
  // Row range 1:3
  m = matchAt(ROW_RANGE, s, pos);
  if (m && !AFTER_REF_BAD.test(s[pos + m[0].length] ?? '')) {
    const r1 = parseInt(m[2], 10) - 1;
    const r2 = parseInt(m[4], 10) - 1;
    if (r1 >= 0 && r2 >= 0 && r1 < MAX_ROWS && r2 < MAX_ROWS) {
      return {
        len: m[0].length,
        ref: {
          kind: 'rows',
          r1: Math.min(r1, r2),
          r2: Math.max(r1, r2),
          c1: 0,
          c2: MAX_COLS - 1,
          ar1: !!m[1],
          ar2: !!m[3],
          ac1: true,
          ac2: true,
        },
      };
    }
  }
  m = matchAt(CELL_RANGE, s, pos);
  if (m) {
    let len = m[0].length;
    let hasSecond = m[6] !== undefined;
    // If the full range is followed by a bad char, fall back to just the first cell.
    if (AFTER_REF_BAD.test(s[pos + len] ?? '')) {
      if (!hasSecond) return null;
      return null;
    }
    const c1 = nameToCol(m[2]);
    const r1 = parseInt(m[4], 10) - 1;
    if (c1 < 0 || c1 >= MAX_COLS || r1 < 0 || r1 >= MAX_ROWS) return null;
    if (!hasSecond) {
      return {
        len,
        ref: { kind: 'cell', r1, c1, r2: r1, c2: c1, ar1: !!m[3], ac1: !!m[1], ar2: !!m[3], ac2: !!m[1] },
      };
    }
    const c2 = nameToCol(m[6]);
    const r2 = parseInt(m[8], 10) - 1;
    if (c2 < 0 || c2 >= MAX_COLS || r2 < 0 || r2 >= MAX_ROWS) return null;
    // Keep the written order's absolute flags but normalize coordinates.
    const swapR = r2 < r1;
    const swapC = c2 < c1;
    len = m[0].length;
    return {
      len,
      ref: {
        kind: 'range',
        r1: swapR ? r2 : r1,
        r2: swapR ? r1 : r2,
        c1: swapC ? c2 : c1,
        c2: swapC ? c1 : c2,
        ar1: swapR ? !!m[7] : !!m[3],
        ar2: swapR ? !!m[3] : !!m[7],
        ac1: swapC ? !!m[5] : !!m[1],
        ac2: swapC ? !!m[1] : !!m[5],
      },
    };
  }
  return null;
}

export function tokenize(src: string): Token[] {
  const toks: Token[] = [];
  let i = 0;
  const n = src.length;
  const push = (type: TokType, start: number, end: number, extra?: Partial<Token>) => {
    toks.push({ type, text: src.slice(start, end), start, end, ...extra });
  };
  const prevSignificant = () => {
    for (let k = toks.length - 1; k >= 0; k--) if (toks[k].type !== 'ws') return toks[k];
    return undefined;
  };
  // A reference can't directly follow a value (e.g. "1A1"), but can follow operators/open parens.
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      let j = i + 1;
      while (j < n && /\s/.test(src[j])) j++;
      push('ws', i, j);
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let val = '';
      let closed = false;
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            val += '"';
            j += 2;
            continue;
          }
          closed = true;
          j++;
          break;
        }
        val += src[j++];
      }
      push('str', i, j, { value: val, unterminated: !closed });
      i = j;
      continue;
    }
    if (ch === '#') {
      const prev = toks[toks.length - 1];
      if (prev && (prev.type === 'ref' || (prev.type === 'rparen'))) {
        push('op', i, i + 1); // spill operator
        i++;
        continue;
      }
      const m = matchAt(ERROR_LIT, src, i);
      if (m) {
        push('err', i, i + m[0].length, { value: m[0].toUpperCase() });
        i += m[0].length;
        continue;
      }
      push('unknown', i, i + 1);
      i++;
      continue;
    }
    // Sheet-qualified reference
    if (ch === "'") {
      const m = matchAt(QUOTED_SHEET, src, i);
      if (m) {
        const r = readRef(src, i + m[0].length);
        if (r) {
          r.ref.sheet = m[1].replace(/''/g, "'");
          push('ref', i, i + m[0].length + r.len, { ref: r.ref });
          i += m[0].length + r.len;
          continue;
        }
        // Sheet prefix with #REF!
        const em = matchAt(ERROR_LIT, src, i + m[0].length);
        if (em) {
          push('err', i, i + m[0].length + em[0].length, { value: '#REF!' });
          i += m[0].length + em[0].length;
          continue;
        }
      }
      push('unknown', i, i + 1);
      i++;
      continue;
    }
    const prev = prevSignificant();
    const valueBefore =
      prev &&
      (prev.type === 'num' || prev.type === 'ref' || prev.type === 'str' || prev.type === 'rparen') &&
      toks[toks.length - 1] === prev;
    if (!valueBefore) {
      const sm = matchAt(UNQUOTED_SHEET, src, i);
      if (sm) {
        const r = readRef(src, i + sm[0].length);
        if (r) {
          r.ref.sheet = sm[1];
          push('ref', i, i + sm[0].length + r.len, { ref: r.ref });
          i += sm[0].length + r.len;
          continue;
        }
        const em = matchAt(ERROR_LIT, src, i + sm[0].length);
        if (em) {
          push('err', i, i + sm[0].length + em[0].length, { value: '#REF!' });
          i += sm[0].length + em[0].length;
          continue;
        }
      }
      const r = readRef(src, i);
      if (r) {
        push('ref', i, i + r.len, { ref: r.ref });
        i += r.len;
        continue;
      }
    }
    if ((ch >= '0' && ch <= '9') || (ch === '.' && /\d/.test(src[i + 1] ?? ''))) {
      const m = matchAt(NUMBER, src, i)!;
      push('num', i, i + m[0].length, { value: parseExcelNumber(m[0]) });
      i += m[0].length;
      continue;
    }
    const im = matchAt(IDENT, src, i);
    if (im) {
      const word = im[0];
      let j = i + word.length;
      let k = j;
      while (k < n && (src[k] === ' ')) k++;
      if (src[k] === '(') {
        push('func', i, j, { value: word });
      } else if (/^(TRUE|FALSE)$/i.test(word)) {
        push('bool', i, j, { value: word.toUpperCase() === 'TRUE' });
      } else {
        push('name', i, j, { value: word });
      }
      i = j;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      push('op', i, i + 2);
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%:@'.includes(ch)) {
      push('op', i, i + 1);
      i++;
      continue;
    }
    if (ch === '(') push('lparen', i, i + 1);
    else if (ch === ')') push('rparen', i, i + 1);
    else if (ch === ',') push('comma', i, i + 1);
    else if (ch === ';') push('semicolon', i, i + 1);
    else if (ch === '{') push('lbrace', i, i + 1);
    else if (ch === '}') push('rbrace', i, i + 1);
    else push('unknown', i, i + 1);
    i++;
  }
  return toks;
}

/** Serializes a reference back to A1 text. */
export function refToString(ref: RefInfo): string {
  const prefix = ref.sheet !== undefined ? quoteSheetName(ref.sheet) + '!' : '';
  const col = (c: number, abs: boolean) => (abs ? '$' : '') + colToName(c);
  const row = (r: number, abs: boolean) => (abs ? '$' : '') + (r + 1);
  switch (ref.kind) {
    case 'cell':
      return prefix + col(ref.c1, ref.ac1) + row(ref.r1, ref.ar1);
    case 'range':
      return prefix + col(ref.c1, ref.ac1) + row(ref.r1, ref.ar1) + ':' + col(ref.c2, ref.ac2) + row(ref.r2, ref.ar2);
    case 'cols':
      return prefix + col(ref.c1, ref.ac1) + ':' + col(ref.c2, ref.ac2);
    case 'rows':
      return prefix + row(ref.r1, ref.ar1) + ':' + row(ref.r2, ref.ar2);
  }
}
