import { formatValue } from '../../numfmt';
import { parseNumberLike } from '../../input';
import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import { argBool, argInt, argNum, argText, eachValue, numberToText, scalarOf, toText } from '../coerce';
import { wildcardToRegExp } from '../criteria';
import { Matrix, RangeRef, type Arg, type Ctx, type Value } from '../types';
import { define } from './registry';

function textFn(f: (s: string) => Value) {
  return (ctx: Ctx, [a]: Arg[]): Value => {
    const s = argText(ctx, a);
    return isError(s) ? s : f(s);
  };
}

define('LEN', 'Text', 'text', 'Returns the number of characters in text.', textFn((s) => s.length), { lift: true });
define('LOWER', 'Text', 'text', 'Converts text to lowercase.', textFn((s) => s.toLowerCase()), { lift: true });
define('UPPER', 'Text', 'text', 'Converts text to uppercase.', textFn((s) => s.toUpperCase()), { lift: true });
define(
  'PROPER',
  'Text',
  'text',
  'Capitalizes the first letter of each word.',
  textFn((s) => s.toLowerCase().replace(/(^|[^\p{L}'])(\p{L})/gu, (_m, p, ch) => p + ch.toUpperCase())),
  { lift: true },
);
define('TRIM', 'Text', 'text', 'Removes extra spaces: leading, trailing and repeated spaces between words.', textFn((s) => s.replace(/ +/g, ' ').trim()), { lift: true });
// eslint-disable-next-line no-control-regex
define('CLEAN', 'Text', 'text', 'Removes non-printable characters.', textFn((s) => s.replace(/[\x00-\x1f]/g, '')), { lift: true });
define('T', 'Text', 'value', 'Returns the text if the value is text, otherwise "".', (ctx, [a]) => {
  const v = scalarOf(ctx, a);
  return typeof v === 'string' ? v : v instanceof ErrorValue ? v : '';
}, { lift: true });

define(
  'LEFT',
  'Text',
  'text, [num_chars]',
  'Returns the first characters of text.',
  (ctx, [a, b]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const n = argInt(ctx, b, 1);
    if (isError(n)) return n;
    if (n < 0) return Errors.value('num_chars can\'t be negative.');
    return s.slice(0, n);
  },
  { lift: true },
);
define(
  'RIGHT',
  'Text',
  'text, [num_chars]',
  'Returns the last characters of text.',
  (ctx, [a, b]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const n = argInt(ctx, b, 1);
    if (isError(n)) return n;
    if (n < 0) return Errors.value('num_chars can\'t be negative.');
    return n === 0 ? '' : s.slice(-n);
  },
  { lift: true },
);
define(
  'MID',
  'Text',
  'text, start_num, num_chars',
  'Returns characters from the middle of text, starting at start_num.',
  (ctx, [a, b, c]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const st = argInt(ctx, b);
    if (isError(st)) return st;
    const n = argInt(ctx, c);
    if (isError(n)) return n;
    if (st < 1) return Errors.value('start_num must be at least 1.');
    if (n < 0) return Errors.value('num_chars can\'t be negative.');
    return s.substr(st - 1, n);
  },
  { lift: true },
);

function concatValues(ctx: Ctx, args: Arg[]): string | ErrorValue {
  let out = '';
  for (const a of args) {
    let err: ErrorValue | null = null;
    eachValue(ctx, a, (v) => {
      if (err) return;
      const t = toText(v);
      if (isError(t)) err = t;
      else out += t;
    });
    if (err) return err;
  }
  return out;
}

define('CONCAT', 'Text', 'text1, [text2], ...', 'Joins text from several values or ranges.', (ctx, args) => concatValues(ctx, args));
define(
  'CONCATENATE',
  'Text',
  'text1, [text2], ...',
  'Joins several text values into one.',
  (ctx, args) => {
    let out = '';
    for (const a of args) {
      const t = toText(scalarOf(ctx, a));
      if (isError(t)) return t;
      out += t;
    }
    return out;
  },
  { lift: true },
);
define(
  'TEXTJOIN',
  'Text',
  'delimiter, ignore_empty, text1, [text2], ...',
  'Joins text with a delimiter, e.g. TEXTJOIN(", ", TRUE, A1:A5).',
  (ctx, [d, ig, ...rest]) => {
    const delim = argText(ctx, d);
    if (isError(delim)) return delim;
    const ignore = argBool(ctx, ig, true);
    if (isError(ignore)) return ignore;
    const parts: string[] = [];
    for (const a of rest) {
      let err: ErrorValue | null = null;
      eachValue(ctx, a, (v) => {
        if (err) return;
        const t = toText(v);
        if (isError(t)) err = t;
        else if (!(ignore && t === '')) parts.push(t);
      });
      if (err) return err;
    }
    return parts.join(delim);
  },
);

define('REPT', 'Text', 'text, number_times', 'Repeats text a number of times.', (ctx, [a, b]) => {
  const s = argText(ctx, a);
  if (isError(s)) return s;
  const n = argInt(ctx, b);
  if (isError(n)) return n;
  if (n < 0) return Errors.value('number_times can\'t be negative.');
  if (s.length * n > 32767) return Errors.value('The result would be longer than 32,767 characters.');
  return s.repeat(n);
}, { lift: true });

define('EXACT', 'Text', 'text1, text2', 'Checks whether two texts are exactly the same (case-sensitive).', (ctx, [a, b]) => {
  const x = argText(ctx, a);
  if (isError(x)) return x;
  const y = argText(ctx, b);
  if (isError(y)) return y;
  return x === y;
}, { lift: true });

function findFn(caseSensitive: boolean) {
  return (ctx: Ctx, [a, b, c]: Arg[]): Value => {
    const needle = argText(ctx, a);
    if (isError(needle)) return needle;
    const hay = argText(ctx, b);
    if (isError(hay)) return hay;
    const start = argInt(ctx, c, 1);
    if (isError(start)) return start;
    if (start < 1 || start > hay.length + 1) return Errors.value('start_num is outside the text.');
    let idx: number;
    if (caseSensitive) idx = hay.indexOf(needle, start - 1);
    else if (/[*?~]/.test(needle)) {
      const re = new RegExp(wildcardToRegExp(needle).source.slice(1, -1), 'i');
      const m = re.exec(hay.slice(start - 1));
      idx = m ? m.index + start - 1 : -1;
    } else idx = hay.toLowerCase().indexOf(needle.toLowerCase(), start - 1);
    if (idx < 0) return Errors.value(`"${needle}" wasn't found in "${hay.length > 40 ? hay.slice(0, 40) + '…' : hay}".`);
    return idx + 1;
  };
}
define('FIND', 'Text', 'find_text, within_text, [start_num]', 'Finds text within other text (case-sensitive) and returns its position.', findFn(true), { lift: true });
define('SEARCH', 'Text', 'find_text, within_text, [start_num]', 'Finds text within other text (not case-sensitive, allows * and ?) and returns its position.', findFn(false), { lift: true });

define(
  'SUBSTITUTE',
  'Text',
  'text, old_text, new_text, [instance_num]',
  'Replaces occurrences of old_text with new_text.',
  (ctx, [a, b, c, d]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const old = argText(ctx, b);
    if (isError(old)) return old;
    const nw = argText(ctx, c);
    if (isError(nw)) return nw;
    if (!old) return s;
    if (d === undefined) return s.split(old).join(nw);
    const inst = argInt(ctx, d);
    if (isError(inst)) return inst;
    if (inst < 1) return Errors.value('instance_num must be at least 1.');
    let idx = -1;
    for (let i = 0; i < inst; i++) {
      idx = s.indexOf(old, idx + 1);
      if (idx < 0) return s;
    }
    return s.slice(0, idx) + nw + s.slice(idx + old.length);
  },
  { lift: true },
);
define(
  'REPLACE',
  'Text',
  'old_text, start_num, num_chars, new_text',
  'Replaces part of text, by position, with different text.',
  (ctx, [a, b, c, d]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const st = argInt(ctx, b);
    if (isError(st)) return st;
    const n = argInt(ctx, c);
    if (isError(n)) return n;
    const nw = argText(ctx, d);
    if (isError(nw)) return nw;
    if (st < 1 || n < 0) return Errors.value();
    return s.slice(0, st - 1) + nw + s.slice(st - 1 + n);
  },
  { lift: true },
);

define(
  'TEXT',
  'Text',
  'value, format_text',
  'Formats a number as text using a format code, e.g. TEXT(A1, "$#,##0.00") or TEXT(A1, "yyyy-mm-dd").',
  (ctx, [a, b]) => {
    let v = scalarOf(ctx, a);
    if (v instanceof ErrorValue) return v;
    const fmt = argText(ctx, b);
    if (isError(fmt)) return fmt;
    if (typeof v === 'string') {
      const n = parseNumberLike(v);
      if (n !== null) v = n;
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return formatValue(v, fmt, 255).text;
  },
  { lift: true },
);
define('VALUE', 'Text', 'text', 'Converts text that looks like a number (or date) into a number.', (ctx, [a]) => {
  const v = scalarOf(ctx, a);
  if (typeof v === 'number') return v;
  if (v instanceof ErrorValue) return v;
  if (v === null) return 0;
  const n = parseNumberLike(String(v));
  return n === null ? Errors.value(`"${v}" doesn't look like a number.`) : n;
}, { lift: true });
define(
  'NUMBERVALUE',
  'Text',
  'text, [decimal_separator], [group_separator]',
  'Converts text to a number using the given separators (handy for European formats like 1.234,56).',
  (ctx, [a, b, c]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const dec = argText(ctx, b, '.');
    if (isError(dec)) return dec;
    const grp = argText(ctx, c, ',');
    if (isError(grp)) return grp;
    let t = s.replace(/\s/g, '');
    if (grp) t = t.split(grp[0]).join('');
    if (dec && dec[0] !== '.') t = t.replace(dec[0], '.');
    let pct = 0;
    while (t.endsWith('%')) {
      pct++;
      t = t.slice(0, -1);
    }
    if (t === '') return 0;
    const n = Number(t);
    if (isNaN(n)) return Errors.value(`"${s}" can't be converted to a number.`);
    return n / Math.pow(100, pct);
  },
  { lift: true },
);

function fixedFormat(n: number, decimals: number, commas: boolean, currency = ''): string {
  const d = Math.max(0, Math.min(decimals, 127));
  const f = Math.pow(10, decimals);
  const rounded = Math.round(Math.abs(n) * f) / f;
  let [i, frac = ''] = rounded.toFixed(d).split('.');
  if (commas) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const body = currency + i + (d > 0 ? '.' + frac : '');
  return n < 0 && rounded !== 0 ? (currency ? `(${body})` : '-' + body) : body;
}

define('FIXED', 'Text', 'number, [decimals], [no_commas]', 'Formats a number as text with a fixed number of decimals.', (ctx, [a, b, c]) => {
  const n = argNum(ctx, a);
  if (isError(n)) return n;
  const d = argInt(ctx, b, 2);
  if (isError(d)) return d;
  const nc = argBool(ctx, c, false);
  if (isError(nc)) return nc;
  return fixedFormat(n, d, !nc);
}, { lift: true });
define('DOLLAR', 'Text', 'number, [decimals]', 'Formats a number as currency text.', (ctx, [a, b]) => {
  const n = argNum(ctx, a);
  if (isError(n)) return n;
  const d = argInt(ctx, b, 2);
  if (isError(d)) return d;
  return fixedFormat(n, d, true, '$');
}, { lift: true });

define('CHAR', 'Text', 'number', 'Returns the character for a code number (1-255).', (ctx, [a]) => {
  const n = argInt(ctx, a);
  if (isError(n)) return n;
  if (n < 1 || n > 255) return Errors.value('CHAR needs a number from 1 to 255.');
  return String.fromCharCode(n);
}, { lift: true });
define('CODE', 'Text', 'text', 'Returns the code number of the first character.', (ctx, [a]) => {
  const s = argText(ctx, a);
  if (isError(s)) return s;
  if (!s) return Errors.value('CODE needs at least one character.');
  return s.charCodeAt(0);
}, { lift: true });
define('UNICHAR', 'Text', 'number', 'Returns the Unicode character for a number.', (ctx, [a]) => {
  const n = argInt(ctx, a);
  if (isError(n)) return n;
  if (n < 1 || n > 0x10ffff) return Errors.value();
  return String.fromCodePoint(n);
}, { lift: true });
define('UNICODE', 'Text', 'text', 'Returns the Unicode number of the first character.', (ctx, [a]) => {
  const s = argText(ctx, a);
  if (isError(s)) return s;
  if (!s) return Errors.value();
  return s.codePointAt(0)!;
}, { lift: true });

function splitAround(before: boolean) {
  return (ctx: Ctx, [a, b, c, d, e, f]: Arg[]): Value => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const delim = argText(ctx, b);
    if (isError(delim)) return delim;
    const inst = argInt(ctx, c, 1);
    if (isError(inst)) return inst;
    const mode = argInt(ctx, d, 0);
    if (isError(mode)) return mode;
    const matchEnd = argBool(ctx, e, false);
    if (isError(matchEnd)) return matchEnd;
    if (inst === 0) return Errors.value('instance_num can\'t be 0.');
    const hay = mode === 1 ? s.toLowerCase() : s;
    const nd = mode === 1 ? delim.toLowerCase() : delim;
    const positions: number[] = [];
    if (nd === '') positions.push(inst > 0 ? 0 : s.length);
    else {
      let i = hay.indexOf(nd);
      while (i >= 0) {
        positions.push(i);
        i = hay.indexOf(nd, i + nd.length);
      }
    }
    if (matchEnd) positions.push(s.length);
    const k = inst > 0 ? inst - 1 : positions.length + inst;
    const pos = positions[k];
    if (pos === undefined) return f !== undefined ? scalarOf(ctx, f) : Errors.na(`"${delim}" wasn't found in the text.`);
    return before ? s.slice(0, pos) : s.slice(pos + nd.length);
  };
}
define('TEXTBEFORE', 'Text', 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]', 'Returns the text before a delimiter.', splitAround(true), { lift: [0, 2] });
define('TEXTAFTER', 'Text', 'text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found]', 'Returns the text after a delimiter.', splitAround(false), { lift: [0, 2] });

define(
  'TEXTSPLIT',
  'Text',
  'text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]',
  'Splits text into columns (and rows) using delimiters.',
  (ctx, [a, b, c, d, e, f]) => {
    const s = argText(ctx, a);
    if (isError(s)) return s;
    const delims = (v: Arg): string[] => {
      if (v === undefined) return [];
      if (v instanceof Matrix || v instanceof RangeRef) {
        const out: string[] = [];
        eachValue(ctx, v, (x) => {
          const t = toText(x);
          if (!isError(t) && t) out.push(t);
        });
        return out;
      }
      const t = toText(scalarOf(ctx, v));
      return isError(t) || !t ? [] : [t];
    };
    const colD = delims(b);
    const rowD = delims(c);
    const ignore = argBool(ctx, d, false);
    if (isError(ignore)) return ignore;
    const mode = argInt(ctx, e, 0);
    if (isError(mode)) return mode;
    const pad = f === undefined ? Errors.na() : scalarOf(ctx, f);
    const esc = (x: string) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const split = (str: string, ds: string[]) => {
      if (!ds.length) return [str];
      const re = new RegExp(ds.map(esc).join('|'), mode === 1 ? 'gi' : 'g');
      let parts = str.split(re);
      if (ignore) parts = parts.filter((p) => p !== '');
      return parts;
    };
    const rows = split(s, rowD).map((r) => split(r, colD));
    const width = Math.max(...rows.map((r) => r.length));
    return Matrix.fill(rows.length, width, (r, cc) => (rows[r][cc] !== undefined ? rows[r][cc] : pad));
  },
);

define('VALUETOTEXT', 'Text', 'value, [format]', 'Converts any value to text.', (ctx, [a, b]) => {
  const v = scalarOf(ctx, a);
  const strict = argInt(ctx, b, 0);
  if (isError(strict)) return strict;
  if (v === null) return '';
  if (v instanceof ErrorValue) return v.code;
  if (typeof v === 'string') return strict === 1 ? `"${v}"` : v;
  if (typeof v === 'number') return numberToText(v);
  return v ? 'TRUE' : 'FALSE';
}, { lift: [0] });

define('HYPERLINK', 'Text', 'link_location, [friendly_name]', 'Creates a clickable link (Ctrl+click to open).', (ctx, [a, b]) => {
  const link = argText(ctx, a);
  if (isError(link)) return link;
  if (b === undefined) return link;
  const v: Scalar = scalarOf(ctx, b);
  return v;
});
