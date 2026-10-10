// Criteria matching for COUNTIF, SUMIFS, etc. ("<5", ">=10", "app*", "<>", "").
import { parseNumberLike } from '../input';
import { ErrorValue, type Scalar } from '../values';
import { compareValues } from './coerce';

export type Matcher = (v: Scalar) => boolean;

export function wildcardToRegExp(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      re += '\\' + pattern[++i];
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp('^' + re + '$', 'i');
}

export function hasWildcard(s: string): boolean {
  return /(^|[^~])[*?]/.test(s);
}

export function makeMatcher(crit: Scalar): Matcher {
  if (crit === null) return (v) => v === null || v === '' || v === 0;
  if (typeof crit === 'number') {
    return (v) => (typeof v === 'number' && v === crit) || (typeof v === 'string' && parseNumberLike(v) === crit);
  }
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit instanceof ErrorValue) return (v) => v instanceof ErrorValue && v.code === crit.code;

  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(crit)!;
  const op = m[1] ?? '=';
  const rest = m[2];
  const num = parseNumberLike(rest);
  const bool = /^(TRUE|FALSE)$/i.test(rest) ? rest.toUpperCase() === 'TRUE' : null;

  if (rest === '' && (op === '=' || !m[1])) {
    // "" or "=" matches blank cells
    return (v) => v === null || v === '';
  }
  if (rest === '' && op === '<>') return (v) => v !== null && v !== '';

  if (op === '=' || op === '<>') {
    let test: Matcher;
    if (num !== null) test = (v) => typeof v === 'number' ? v === num : typeof v === 'string' && parseNumberLike(v) === num;
    else if (bool !== null) test = (v) => v === bool;
    else if (rest.startsWith('#')) test = (v) => v instanceof ErrorValue && v.code === rest.toUpperCase();
    else if (hasWildcard(rest) || rest.includes('~')) {
      const re = wildcardToRegExp(rest);
      test = (v) => typeof v === 'string' && re.test(v);
    } else {
      const lower = rest.toLowerCase();
      test = (v) => typeof v === 'string' && v.toLowerCase() === lower;
    }
    return op === '=' ? test : (v) => !test(v);
  }
  // Relational: numbers compare with numbers, text with text
  if (num !== null) {
    return (v) => {
      if (typeof v !== 'number') return false;
      return relational(op, v - num);
    };
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    return relational(op, compareValues(v, rest));
  };
}

function relational(op: string, cmp: number): boolean {
  switch (op) {
    case '<':
      return cmp < 0;
    case '<=':
      return cmp <= 0;
    case '>':
      return cmp > 0;
    case '>=':
      return cmp >= 0;
  }
  return false;
}
