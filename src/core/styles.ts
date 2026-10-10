// Cell styles. Styles are immutable and interned in a table; cells refer to them by id.

export type HAlign = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VAlign = 'top' | 'middle' | 'bottom';
export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface BorderSide {
  style: BorderStyle;
  color?: string; // '#rrggbb'; undefined = automatic
}

export interface Style {
  fontName?: string;
  fontSize?: number; // points
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string; // text color; undefined = automatic (contrast-aware)
  fill?: string; // background color
  numFmt?: string;
  hAlign?: HAlign;
  vAlign?: VAlign;
  wrap?: boolean;
  indent?: number;
  rotation?: number;
  borderTop?: BorderSide;
  borderRight?: BorderSide;
  borderBottom?: BorderSide;
  borderLeft?: BorderSide;
}

export type StylePatch = { [K in keyof Style]?: Style[K] | null };

const KEYS: (keyof Style)[] = [
  'fontName',
  'fontSize',
  'bold',
  'italic',
  'underline',
  'strike',
  'color',
  'fill',
  'numFmt',
  'hAlign',
  'vAlign',
  'wrap',
  'indent',
  'rotation',
  'borderTop',
  'borderRight',
  'borderBottom',
  'borderLeft',
];

export const DEFAULT_FONT = 'Calibri';
export const DEFAULT_FONT_SIZE = 11;

function keyOf(s: Style): string {
  let k = '';
  for (const key of KEYS) {
    const v = s[key];
    if (v === undefined || v === false || v === '' ) continue;
    k += key + '=' + (typeof v === 'object' ? `${(v as BorderSide).style}/${(v as BorderSide).color ?? ''}` : String(v)) + ';';
  }
  return k;
}

function clean(s: Style): Style {
  const out: Style = {};
  for (const key of KEYS) {
    const v = s[key];
    if (v === undefined || v === null || v === false || v === '') continue;
    if (key === 'numFmt' && v === 'General') continue;
    if (key === 'hAlign' && v === 'general') continue;
    if (key === 'vAlign' && v === 'bottom') continue;
    if (key === 'indent' && v === 0) continue;
    if (key === 'rotation' && v === 0) continue;
    (out as Record<string, unknown>)[key] = v;
  }
  return out;
}

export class StyleTable {
  private styles: Style[] = [{}];
  private index = new Map<string, number>([['', 0]]);

  get(id: number | undefined): Style {
    return (id !== undefined && this.styles[id]) || this.styles[0];
  }

  intern(s: Style): number {
    const c = clean(s);
    const k = keyOf(c);
    let id = this.index.get(k);
    if (id === undefined) {
      id = this.styles.length;
      this.styles.push(Object.freeze(c));
      this.index.set(k, id);
    }
    return id;
  }

  /** Returns the id of `base` with `patch` applied (null in patch removes a property). */
  apply(baseId: number | undefined, patch: StylePatch): number {
    const base = this.get(baseId);
    const next: Style = { ...base };
    for (const key of Object.keys(patch) as (keyof Style)[]) {
      const v = patch[key];
      if (v === null || v === undefined) delete next[key];
      else (next as Record<string, unknown>)[key] = v;
    }
    return this.intern(next);
  }

  get size(): number {
    return this.styles.length;
  }

  all(): readonly Style[] {
    return this.styles;
  }
}

export function isDefaultStyle(s: Style): boolean {
  return Object.keys(s).length === 0;
}

// ---- Color helpers ----

export function parseHex(hex: string): [number, number, number] {
  let h = hex.replace('#', '');
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  if (h.length === 8) h = h.slice(2); // ARGB
  const n = parseInt(h, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

export function toHex(r: number, g: number, b: number): string {
  const c = (x: number) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0');
  return `#${c(r)}${c(g)}${c(b)}`;
}

export function luminance(hex: string): number {
  const [r, g, b] = parseHex(hex).map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export function contrastRatio(a: string, b: string): number {
  const la = luminance(a);
  const lb = luminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

export function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}

export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue2rgb(p, q, h + 1 / 3) * 255, hue2rgb(p, q, h) * 255, hue2rgb(p, q, h - 1 / 3) * 255];
}

/** Applies an Excel theme tint (-1..1) to a color. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [r, g, b] = parseHex(hex);
  const [h, s, l] = rgbToHsl(r, g, b);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [nr, ng, nb] = hslToRgb(h, s, nl);
  return toHex(nr, ng, nb);
}

/** Inverts the lightness of a color, keeping its hue — used to keep dark text readable on dark backgrounds. */
export function invertLightness(hex: string): string {
  const [r, g, b] = parseHex(hex);
  const [h, s, l] = rgbToHsl(r, g, b);
  const [nr, ng, nb] = hslToRgb(h, s, Math.max(0.65, 1 - l));
  return toHex(nr, ng, nb);
}

export function mixColors(a: string, b: string, t: number): string {
  const [ar, ag, ab] = parseHex(a);
  const [br, bg, bb] = parseHex(b);
  return toHex(ar + (br - ar) * t, ag + (bg - ag) * t, ab + (bb - ab) * t);
}
