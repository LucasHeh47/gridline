// Excel number format codes: parsing and rendering.
import { DAYS, MONTHS, serialToParts } from './dates';
import { ErrorValue, type Scalar } from './values';

export interface Formatted {
  text: string;
  color?: string;
  /** For accounting-style formats: text is split around a fill point. */
  fillAt?: number;
}

type Tok =
  | { t: 'lit'; v: string }
  | { t: 'digit'; v: '0' | '#' | '?' }
  | { t: 'point' }
  | { t: 'comma' }
  | { t: 'percent' }
  | { t: 'exp'; sign: '+' | '-' }
  | { t: 'date'; v: string }
  | { t: 'ampm'; v: string }
  | { t: 'elapsed'; v: string }
  | { t: 'text' }
  | { t: 'fill'; v: string }
  | { t: 'skip'; v: string }
  | { t: 'slash' };

interface Section {
  toks: Tok[];
  color?: string;
  cond?: { op: string; val: number };
  isDate: boolean;
  isText: boolean;
  isGeneral: boolean;
  hasAmPm: boolean;
  percent: number;
  scale: number; // number of trailing scale commas
  thousands: boolean;
  intDigits: { t: 'digit'; v: '0' | '#' | '?' }[];
  fracDigits: number;
  fracMin: number;
  exp: boolean;
  fraction: boolean;
  secDecimals: number;
}

interface Parsed {
  sections: Section[];
}

const COLORS: Record<string, string> = {
  black: '#000000',
  white: '#ffffff',
  red: '#ff4d4f',
  green: '#3fcf8e',
  blue: '#4ea1ff',
  yellow: '#ffd43b',
  magenta: '#e64980',
  cyan: '#22b8cf',
};

const cache = new Map<string, Parsed>();

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQuote = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '"') inQuote = !inQuote;
    if (ch === '\\' && !inQuote && i + 1 < fmt.length) {
      cur += ch + fmt[i + 1];
      i++;
      continue;
    }
    if (ch === ';' && !inQuote) {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function parseSection(src: string): Section {
  const sec: Section = {
    toks: [],
    isDate: false,
    isText: false,
    isGeneral: false,
    hasAmPm: false,
    percent: 0,
    scale: 0,
    thousands: false,
    intDigits: [],
    fracDigits: 0,
    fracMin: 0,
    exp: false,
    fraction: false,
    secDecimals: 0,
  };
  const toks = sec.toks;
  let i = 0;
  const s = src;
  while (i < s.length) {
    const ch = s[i];
    const lower = ch.toLowerCase();
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      const lit = end < 0 ? s.slice(i + 1) : s.slice(i + 1, end);
      toks.push({ t: 'lit', v: lit });
      i = end < 0 ? s.length : end + 1;
      continue;
    }
    if (ch === '\\') {
      toks.push({ t: 'lit', v: s[i + 1] ?? '' });
      i += 2;
      continue;
    }
    if (ch === '_') {
      toks.push({ t: 'skip', v: s[i + 1] ?? ' ' });
      i += 2;
      continue;
    }
    if (ch === '*') {
      toks.push({ t: 'fill', v: s[i + 1] ?? ' ' });
      i += 2;
      continue;
    }
    if (ch === '[') {
      const end = s.indexOf(']', i);
      const inner = end < 0 ? s.slice(i + 1) : s.slice(i + 1, end);
      i = end < 0 ? s.length : end + 1;
      const low = inner.toLowerCase();
      if (COLORS[low]) {
        sec.color = COLORS[low];
      } else if (/^color\s*\d+$/.test(low)) {
        sec.color = undefined;
      } else if (/^(<=|>=|<>|<|>|=)\s*-?[\d.]+$/.test(inner)) {
        const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner)!;
        sec.cond = { op: m[1], val: parseFloat(m[2]) };
      } else if (/^(h+|m+|s+)$/i.test(inner)) {
        toks.push({ t: 'elapsed', v: low });
        sec.isDate = true;
      } else if (inner.startsWith('$')) {
        // Currency/locale: [$€-407] -> "€"
        const sym = inner.slice(1).split('-')[0];
        if (sym) toks.push({ t: 'lit', v: sym });
      }
      continue;
    }
    if (/^general/i.test(s.slice(i))) {
      sec.isGeneral = true;
      toks.push({ t: 'lit', v: '\u0000GENERAL' });
      i += 7;
      continue;
    }
    if (/^(am\/pm|a\/p)/i.test(s.slice(i))) {
      const m = /^(am\/pm|a\/p)/i.exec(s.slice(i))!;
      toks.push({ t: 'ampm', v: m[1] });
      sec.hasAmPm = true;
      sec.isDate = true;
      i += m[1].length;
      continue;
    }
    if ('ymdhse'.includes(lower) && !(lower === 'e' && /^e[+-]/i.test(s.slice(i)))) {
      if (lower === 'e' && !sec.isDate) {
        // A bare 'e' outside dates is a literal
        toks.push({ t: 'lit', v: ch });
        i++;
        continue;
      }
      let j = i;
      while (j < s.length && s[j].toLowerCase() === lower) j++;
      toks.push({ t: 'date', v: s.slice(i, j).toLowerCase() });
      sec.isDate = true;
      i = j;
      continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') {
      toks.push({ t: 'digit', v: ch });
      i++;
      continue;
    }
    if (ch === '.') {
      toks.push({ t: 'point' });
      i++;
      continue;
    }
    if (ch === ',') {
      toks.push({ t: 'comma' });
      i++;
      continue;
    }
    if (ch === '%') {
      toks.push({ t: 'percent' });
      sec.percent++;
      i++;
      continue;
    }
    if ((ch === 'E' || ch === 'e') && (s[i + 1] === '+' || s[i + 1] === '-')) {
      toks.push({ t: 'exp', sign: s[i + 1] as '+' | '-' });
      sec.exp = true;
      i += 2;
      continue;
    }
    if (ch === '@') {
      toks.push({ t: 'text' });
      sec.isText = true;
      i++;
      continue;
    }
    if (ch === '/') {
      toks.push({ t: 'slash' });
      i++;
      continue;
    }
    toks.push({ t: 'lit', v: ch });
    i++;
  }

  if (sec.isDate) {
    // Distinguish months from minutes: m after h or before s is minutes.
    const dateToks = toks.filter((t) => t.t === 'date' || t.t === 'elapsed') as { t: string; v: string }[];
    for (let k = 0; k < dateToks.length; k++) {
      const t = dateToks[k];
      if (t.t !== 'date' || t.v[0] !== 'm' || t.v.length > 2) continue;
      const prev = dateToks[k - 1];
      const next = dateToks[k + 1];
      if ((prev && (prev.v[0] === 'h' || (prev.t === 'elapsed' && prev.v[0] === 'h'))) || (next && next.v[0] === 's')) {
        t.v = t.v.length === 1 ? 'n' : 'nn';
      }
    }
    // Fractional seconds: s.000
    for (let k = 0; k < toks.length; k++) {
      const t = toks[k];
      if (t.t === 'point' && k > 0 && (toks[k - 1].t === 'date' || toks[k - 1].t === 'elapsed')) {
        let n = 0;
        while (toks[k + 1 + n]?.t === 'digit') n++;
        if (n > 0) {
          sec.secDecimals = n;
          toks.splice(k, n + 1, { t: 'date', v: 'f'.repeat(n) });
        }
      }
    }
    return sec;
  }

  // Number analysis
  let seenPoint = false;
  let seenExp = false;
  let seenSlash = false;
  let lastDigitIdx = -1;
  toks.forEach((t, idx) => {
    if (t.t === 'digit') lastDigitIdx = idx;
  });
  for (let k = 0; k < toks.length; k++) {
    const t = toks[k];
    if (t.t === 'point' && !seenExp) seenPoint = true;
    else if (t.t === 'exp') seenExp = true;
    else if (t.t === 'slash') {
      seenSlash = true;
      sec.fraction = true;
    } else if (t.t === 'digit') {
      if (seenSlash || seenExp) continue;
      if (seenPoint) {
        sec.fracDigits++;
        if (t.v === '0') sec.fracMin = sec.fracDigits;
      } else sec.intDigits.push(t);
    } else if (t.t === 'comma') {
      // Comma between digits = thousands separator; trailing commas = scaling
      const before = toks.slice(0, k).some((x) => x.t === 'digit');
      if (k > lastDigitIdx && before) {
        sec.scale++;
        toks[k] = { t: 'lit', v: '' };
      } else if (before && !seenPoint) {
        sec.thousands = true;
        toks[k] = { t: 'lit', v: '' };
      } else {
        toks[k] = { t: 'lit', v: ',' };
      }
    }
  }
  return sec;
}

export function parseFormat(fmt: string): Parsed {
  let p = cache.get(fmt);
  if (p) return p;
  const parts = splitSections(fmt || 'General');
  p = { sections: parts.slice(0, 4).map(parseSection) };
  cache.set(fmt, p);
  return p;
}

export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt) return false;
  const p = parseFormat(fmt);
  return p.sections[0]?.isDate ?? false;
}

export function isTextFormat(fmt: string | undefined): boolean {
  return fmt === '@';
}

function pad(n: number, len: number): string {
  let s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}

/** Formats a number using Excel's "General" rules, limited to ~11 characters. */
export function formatGeneral(n: number, maxChars = 11): string {
  if (!isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const abs = Math.abs(n);
  if (Number.isInteger(n) && abs < 1e11) return String(n);
  if (abs >= 1e11 || abs < 1e-9) return formatSci(n, maxChars);
  // Try decreasing precision until it fits.
  for (let digits = 10; digits >= 1; digits--) {
    let s = n.toPrecision(digits);
    if (s.includes('e')) return formatSci(n, maxChars);
    if (s.includes('.')) s = s.replace(/\.?0+$/, '');
    if (s.length <= maxChars) return s;
  }
  return formatSci(n, maxChars);
}

function formatSci(n: number, maxChars: number): string {
  for (let digits = 6; digits >= 1; digits--) {
    const [mant, ex] = n.toExponential(digits - 1).split('e');
    const m = mant.includes('.') ? mant.replace(/\.?0+$/, '') : mant;
    const e = parseInt(ex, 10);
    const s = `${m}E${e < 0 ? '-' : '+'}${pad(Math.abs(e), 2)}`;
    if (s.length <= maxChars || digits === 1) return s;
  }
  return n.toExponential();
}

function pickSection(p: Parsed, v: number): { sec: Section; abs: boolean } {
  const s = p.sections;
  const hasCond = s.some((x) => x.cond);
  if (hasCond) {
    for (let i = 0; i < Math.min(s.length, 3); i++) {
      const c = s[i].cond;
      if (c && testCond(v, c)) return { sec: s[i], abs: i > 0 && c.op.includes('<') && c.val <= 0 };
    }
    // Fallback: first section without a condition
    const free = s.find((x) => !x.cond && !x.isText);
    if (free) return { sec: free, abs: false };
    return { sec: s[0], abs: false };
  }
  if (s.length === 1 || (s.length >= 1 && s[0].isText && s.length === 1)) return { sec: s[0], abs: false };
  if (v > 0 || (v === 0 && s.length < 3)) return { sec: s[0], abs: false };
  if (v < 0) return { sec: s[1] ?? s[0], abs: !!s[1] };
  return { sec: s[2] ?? s[0], abs: false };
}

function testCond(v: number, c: { op: string; val: number }): boolean {
  switch (c.op) {
    case '<':
      return v < c.val;
    case '<=':
      return v <= c.val;
    case '>':
      return v > c.val;
    case '>=':
      return v >= c.val;
    case '=':
      return v === c.val;
    case '<>':
      return v !== c.val;
  }
  return false;
}

function fmtDate(sec: Section, v: number): string {
  if (v < 0) return '#'.repeat(10);
  const p = serialToParts(v);
  let out = '';
  for (let k = 0; k < sec.toks.length; k++) {
    const t = sec.toks[k];
    switch (t.t) {
      case 'lit':
        out += t.v;
        break;
      case 'skip':
        out += ' ';
        break;
      case 'slash':
        out += '/';
        break;
      case 'point':
        out += '.';
        break;
      case 'comma':
        out += ',';
        break;
      case 'ampm': {
        const pm = p.H >= 12;
        if (t.v.length > 3) out += pm ? 'PM' : 'AM';
        else out += pm ? 'P' : 'A';
        break;
      }
      case 'elapsed': {
        const totalSec = Math.round(v * 86400);
        const ch = t.v[0];
        const n = ch === 'h' ? Math.floor(totalSec / 3600) : ch === 'm' ? Math.floor(totalSec / 60) : totalSec;
        out += pad(n, t.v.length);
        break;
      }
      case 'date': {
        const d = t.v;
        const c0 = d[0];
        if (c0 === 'y') out += d.length <= 2 ? pad(p.y % 100, 2) : pad(p.y, 4);
        else if (c0 === 'm') {
          if (d.length === 1) out += p.m;
          else if (d.length === 2) out += pad(p.m, 2);
          else if (d.length === 3) out += MONTHS[p.m - 1].slice(0, 3);
          else if (d.length === 5) out += MONTHS[p.m - 1][0];
          else out += MONTHS[p.m - 1];
        } else if (c0 === 'd') {
          if (d.length === 1) out += p.d;
          else if (d.length === 2) out += pad(p.d, 2);
          else if (d.length === 3) out += DAYS[p.dow].slice(0, 3);
          else out += DAYS[p.dow];
        } else if (c0 === 'h') {
          let h = p.H;
          if (sec.hasAmPm) {
            h = h % 12;
            if (h === 0) h = 12;
          }
          out += d.length >= 2 ? pad(h, 2) : String(h);
        } else if (c0 === 'n') out += d.length >= 2 ? pad(p.M, 2) : String(p.M);
        else if (c0 === 's') {
          let S = p.S;
          if (sec.secDecimals === 0 && p.ms >= 500) S = Math.min(59, S + 1);
          out += d.length >= 2 ? pad(S, 2) : String(S);
        } else if (c0 === 'f') {
          out += '.' + pad(Math.floor(p.ms / Math.pow(10, 3 - d.length)), d.length);
        } else if (c0 === 'e') out += p.y;
        break;
      }
      default:
        break;
    }
  }
  return out;
}

function groupThousands(s: string): string {
  let out = '';
  let cnt = 0;
  for (let i = s.length - 1; i >= 0; i--) {
    out = s[i] + out;
    if (++cnt % 3 === 0 && i > 0) out = ',' + out;
  }
  return out;
}

function fmtNumber(sec: Section, v: number, useAbs: boolean): Formatted {
  let n = useAbs ? Math.abs(v) : v;
  n = n * Math.pow(100, sec.percent);
  n = n / Math.pow(1000, sec.scale);
  const neg = n < 0;
  n = Math.abs(n);

  if (sec.fraction) return { text: (neg ? '-' : '') + fmtFraction(sec, n) };

  let mantissaStr: string;
  let expPart = '';
  if (sec.exp) {
    const intCount = Math.max(1, sec.intDigits.length);
    let e = n === 0 ? 0 : Math.floor(Math.log10(n));
    // Engineering-style: exponent multiple of the int digit count when using ##0.0E+0
    if (sec.intDigits.length > 1 && sec.intDigits.some((d) => d.v === '#')) {
      e = Math.floor(e / intCount) * intCount;
    } else {
      e = e - (intCount - 1);
    }
    let m = n / Math.pow(10, e);
    m = roundTo(m, sec.fracDigits);
    if (m >= Math.pow(10, intCount) && n !== 0) {
      m /= 10;
      e += 1;
    }
    mantissaStr = m.toFixed(sec.fracDigits);
    const expTok = sec.toks.find((t) => t.t === 'exp') as { t: 'exp'; sign: '+' | '-' };
    let expDigits = 0;
    let afterExp = false;
    for (const t of sec.toks) {
      if (t.t === 'exp') afterExp = true;
      else if (afterExp && t.t === 'digit') expDigits++;
    }
    expPart = 'E' + (e < 0 ? '-' : expTok.sign === '+' ? '+' : '') + pad(Math.abs(e), Math.max(1, expDigits));
  } else {
    mantissaStr = roundTo(n, sec.fracDigits).toFixed(sec.fracDigits);
  }

  let [intStr, fracStr = ''] = mantissaStr.split('.');
  // Trim optional fractional digits
  if (sec.fracDigits > sec.fracMin) {
    let keep = fracStr.length;
    while (keep > sec.fracMin && fracStr[keep - 1] === '0') keep--;
    const optional = fracStr.slice(keep);
    // '?' placeholders become spaces, '#' are dropped
    const fracToks = digitToksAfterPoint(sec);
    let padded = fracStr.slice(0, keep);
    for (let k = keep; k < fracStr.length; k++) {
      if (fracToks[k]?.v === '?') padded += ' ';
    }
    void optional;
    fracStr = padded;
  }

  // Integer part: minimum digits by count of '0' placeholders
  const minInt = sec.intDigits.filter((d) => d.v === '0').length;
  if (intStr === '0' && minInt === 0) intStr = '';
  while (intStr.length < minInt) intStr = '0' + intStr;
  if (sec.thousands) intStr = groupThousands(intStr);

  // Assemble by walking tokens: place integer digits at the first int placeholder,
  // but honour literals between placeholders (e.g. phone numbers).
  const intToks = sec.intDigits.length;
  const hasLiteralInsideInt = (() => {
    let seenDigit = false;
    let inInt = true;
    for (const t of sec.toks) {
      if (t.t === 'point' || t.t === 'exp') inInt = false;
      if (!inInt) break;
      if (t.t === 'digit') seenDigit = true;
      else if (seenDigit && t.t === 'lit' && t.v !== '') {
        // literal followed later by another digit?
        return true;
      }
    }
    return false;
  })();

  let out = '';
  let fillAt: number | undefined;
  if (hasLiteralInsideInt && !sec.thousands) {
    // Distribute digits right-to-left into placeholders.
    const digits = mantissaStr.split('.')[0].replace(/^0+(?=\d)/, '');
    const placeholders: number[] = [];
    sec.toks.forEach((t, idx) => {
      if (t.t === 'digit' && !afterPointIdx(sec, idx)) placeholders.push(idx);
    });
    const assigned = new Map<number, string>();
    let di = digits.length - 1;
    for (let p = placeholders.length - 1; p >= 0; p--) {
      const tokIdx = placeholders[p];
      const ph = (sec.toks[tokIdx] as { t: 'digit'; v: string }).v;
      if (p === 0) {
        assigned.set(tokIdx, di >= 0 ? digits.slice(0, di + 1) : ph === '0' ? '0' : ph === '?' ? ' ' : '');
      } else {
        assigned.set(tokIdx, di >= 0 ? digits[di] : ph === '0' ? '0' : ph === '?' ? ' ' : '');
      }
      di--;
    }
    let fracIdx = 0;
    let pastPoint = false;
    for (let idx = 0; idx < sec.toks.length; idx++) {
      const t = sec.toks[idx];
      if (t.t === 'digit') {
        if (pastPoint) out += fracStr[fracIdx++] ?? '';
        else out += assigned.get(idx) ?? '';
      } else if (t.t === 'point') {
        pastPoint = true;
        out += '.';
      } else out += renderLit(t);
    }
  } else {
    let intPlaced = false;
    let fracPlaced = false;
    let pastPoint = false;
    let pastExp = false;
    for (const t of sec.toks) {
      if (t.t === 'digit') {
        if (pastExp) continue;
        if (!pastPoint) {
          if (!intPlaced) {
            out += intStr;
            intPlaced = true;
          }
        } else if (!fracPlaced) {
          out += fracStr;
          fracPlaced = true;
        }
      } else if (t.t === 'point') {
        if (!intPlaced && intToks === 0) {
          out += intStr;
          intPlaced = true;
        }
        pastPoint = true;
        out += '.';
      } else if (t.t === 'exp') {
        pastExp = true;
        out += expPart;
      } else if (t.t === 'percent') out += '%';
      else if (t.t === 'fill') fillAt = out.length;
      else out += renderLit(t);
    }
    if (!intPlaced && intStr) out = intStr + out;
  }
  if (neg && /[1-9]/.test(out)) out = '-' + out;
  return { text: out, color: sec.color, fillAt };
}

function afterPointIdx(sec: Section, idx: number): boolean {
  for (let k = 0; k < idx; k++) if (sec.toks[k].t === 'point' || sec.toks[k].t === 'exp') return true;
  return false;
}

function digitToksAfterPoint(sec: Section): { v: string }[] {
  const out: { v: string }[] = [];
  let past = false;
  for (const t of sec.toks) {
    if (t.t === 'point') past = true;
    else if (t.t === 'exp') break;
    else if (past && t.t === 'digit') out.push(t);
  }
  return out;
}

function renderLit(t: Tok): string {
  switch (t.t) {
    case 'lit':
      return t.v.startsWith('\u0000') ? '' : t.v;
    case 'skip':
      return ' ';
    case 'percent':
      return '%';
    case 'slash':
      return '/';
    default:
      return '';
  }
}

function roundTo(n: number, digits: number): number {
  const f = Math.pow(10, digits);
  return Math.round((n + Number.EPSILON * n) * f) / f;
}

function fmtFraction(sec: Section, n: number): string {
  // Count denominator placeholders or fixed denominator
  let denDigits = 0;
  let fixedDen = 0;
  let afterSlash = false;
  let hasWhole = false;
  let seenDigitBeforeSpace = false;
  for (const t of sec.toks) {
    if (t.t === 'slash') afterSlash = true;
    else if (afterSlash && t.t === 'digit') denDigits++;
    else if (afterSlash && t.t === 'lit' && /\d/.test(t.v)) fixedDen = fixedDen * 10 + parseInt(t.v, 10);
    else if (!afterSlash && t.t === 'digit') seenDigitBeforeSpace = true;
    else if (!afterSlash && t.t === 'lit' && t.v === ' ' && seenDigitBeforeSpace) hasWhole = true;
  }
  const whole = hasWhole ? Math.floor(n) : 0;
  const frac = n - whole;
  let bestNum = 0;
  let bestDen = 1;
  if (fixedDen) {
    bestDen = fixedDen;
    bestNum = Math.round(frac * fixedDen);
  } else {
    const maxDen = Math.pow(10, Math.max(1, denDigits)) - 1;
    let bestErr = Infinity;
    for (let d = 1; d <= maxDen; d++) {
      const num = Math.round(frac * d);
      const e = Math.abs(frac - num / d);
      if (e < bestErr - 1e-12) {
        bestErr = e;
        bestNum = num;
        bestDen = d;
      }
    }
  }
  if (hasWhole) {
    if (bestNum === 0) return String(whole);
    if (bestNum === bestDen) return String(whole + 1);
    return `${whole ? whole + ' ' : ''}${bestNum}/${bestDen}`;
  }
  return `${bestNum}/${bestDen}`;
}

/**
 * Formats a scalar value for display.
 * @param maxChars only used for General numbers — the renderer passes the cell width.
 */
export function formatValue(v: Scalar, fmt: string | undefined, maxChars = 11): Formatted {
  if (v === null) return { text: '' };
  if (v instanceof ErrorValue) return { text: v.code };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE' };
  const f = fmt || 'General';
  if (typeof v === 'string') {
    const p = parseFormat(f);
    const textSec = p.sections.length >= 4 ? p.sections[3] : p.sections.find((s) => s.isText);
    if (!textSec) return { text: v };
    let out = '';
    for (const t of textSec.toks) {
      if (t.t === 'text') out += v;
      else out += renderLit(t);
    }
    return { text: out, color: textSec.color };
  }
  if (f === 'General' || f === 'general') return { text: formatGeneral(v, maxChars) };
  if (f === '@') return { text: formatGeneral(v, maxChars) };
  const p = parseFormat(f);
  const { sec, abs } = pickSection(p, v);
  if (sec.isGeneral) {
    const t = formatGeneral(abs ? Math.abs(v) : v, maxChars);
    let out = '';
    for (const tok of sec.toks) {
      if (tok.t === 'lit' && tok.v === '\u0000GENERAL') out += t;
      else out += renderLit(tok);
    }
    return { text: out, color: sec.color };
  }
  if (sec.isDate) return { text: fmtDate(sec, v), color: sec.color };
  if (sec.isText && !sec.intDigits.length) {
    return { text: formatGeneral(v, maxChars), color: sec.color };
  }
  return fmtNumber(sec, v, abs);
}

// ---- Format presets for the UI ----

export interface FormatPreset {
  id: string;
  label: string;
  fmt: string;
}

export const FORMAT_PRESETS: FormatPreset[] = [
  { id: 'general', label: 'Automatic', fmt: 'General' },
  { id: 'number', label: 'Number', fmt: '#,##0.00' },
  { id: 'integer', label: 'Whole number', fmt: '#,##0' },
  { id: 'currency', label: 'Currency', fmt: '"$"#,##0.00' },
  { id: 'accounting', label: 'Accounting', fmt: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)' },
  { id: 'percent', label: 'Percent', fmt: '0.00%' },
  { id: 'percent0', label: 'Percent (whole)', fmt: '0%' },
  { id: 'scientific', label: 'Scientific', fmt: '0.00E+00' },
  { id: 'fraction', label: 'Fraction', fmt: '# ?/?' },
  { id: 'date', label: 'Date', fmt: 'm/d/yyyy' },
  { id: 'dateIso', label: 'Date (ISO)', fmt: 'yyyy-mm-dd' },
  { id: 'dateLong', label: 'Long date', fmt: 'dddd, mmmm d, yyyy' },
  { id: 'time', label: 'Time', fmt: 'h:mm AM/PM' },
  { id: 'time24', label: 'Time (24h)', fmt: 'hh:mm:ss' },
  { id: 'datetime', label: 'Date & time', fmt: 'm/d/yyyy h:mm' },
  { id: 'duration', label: 'Duration', fmt: '[h]:mm:ss' },
  { id: 'text', label: 'Plain text', fmt: '@' },
];

/** Changes the number of decimals in a format (for the +/- decimal buttons). */
export function adjustDecimals(fmt: string | undefined, delta: number, sample: Scalar): string {
  let f = fmt && fmt !== 'General' ? fmt : '';
  if (!f) {
    // Derive from the sample value
    let decimals = 0;
    if (typeof sample === 'number') {
      const s = formatGeneral(sample);
      const dot = s.indexOf('.');
      decimals = dot >= 0 && !s.includes('E') ? s.length - dot - 1 : 0;
    }
    const nd = Math.max(0, decimals + delta);
    return nd === 0 ? '0' : '0.' + '0'.repeat(nd);
  }
  return splitSections(f)
    .map((sec) => {
      if (/[ymdhs]/i.test(sec.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''))) return sec;
      // Find the last run of 0/# after a point, or the last digit placeholder
      const m = /([0#?])(\.([0#?]*))?(?![\s\S]*[0#?])/.exec(sec);
      if (!m) return sec;
      const idx = m.index;
      const decimals = m[3]?.length ?? 0;
      const nd = Math.max(0, decimals + delta);
      const replacement = m[1] + (nd > 0 ? '.' + '0'.repeat(nd) : '');
      return sec.slice(0, idx) + replacement + sec.slice(idx + m[0].length);
    })
    .join(';');
}
