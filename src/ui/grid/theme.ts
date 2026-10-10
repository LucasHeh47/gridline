// Colors used by the canvas renderer, read from CSS variables, plus dark-mode color adaptation.
import { contrastRatio, hslToRgb, luminance, parseHex, rgbToHsl, toHex } from '../../core/styles';

export interface GridTheme {
  dark: boolean;
  cellBg: string;
  gridLine: string;
  headerBg: string;
  headerText: string;
  headerLine: string;
  headerSelBg: string;
  headerSelText: string;
  selFill: string;
  selBorder: string;
  cellText: string;
  cellTextDark: string;
  frozenLine: string;
  spillBorder: string;
  errorMark: string;
  noteMark: string;
  linkText: string;
  accent: string;
  /** Adapt light fills / dark text so imported spreadsheets read well in dark mode. */
  adaptColors: boolean;
}

export function readTheme(adaptColors: boolean): GridTheme {
  const cs = getComputedStyle(document.documentElement);
  const v = (n: string) => cs.getPropertyValue(n).trim();
  return {
    dark: document.documentElement.dataset.theme !== 'light',
    cellBg: v('--cell-bg'),
    gridLine: v('--grid-line'),
    headerBg: v('--header-bg'),
    headerText: v('--header-text'),
    headerLine: v('--header-line'),
    headerSelBg: v('--header-sel-bg'),
    headerSelText: v('--header-sel-text'),
    selFill: v('--sel-fill'),
    selBorder: v('--sel-border'),
    cellText: v('--cell-text'),
    cellTextDark: v('--cell-text-dark'),
    frozenLine: v('--frozen-line'),
    spillBorder: v('--spill-border'),
    errorMark: v('--error-mark'),
    noteMark: v('--note-mark'),
    linkText: v('--link-text'),
    accent: v('--accent'),
    adaptColors,
  };
}

const fillCache = new Map<string, string>();
const textCache = new Map<string, string>();

/** How a stored fill color is displayed. In dark mode, near-white fills become dark tints of the same hue. */
export function displayFill(theme: GridTheme, hex: string): string {
  if (!theme.dark || !theme.adaptColors) return hex;
  let out = fillCache.get(hex);
  if (out) return out;
  const lum = luminance(hex);
  if (lum < 0.62) out = hex;
  else {
    const [r, g, b] = parseHex(hex);
    const [hh, s, l] = rgbToHsl(r, g, b);
    const nl = 0.1 + (1 - l) * 0.9 + (s > 0.15 ? 0.1 : 0.02);
    const [nr, ng, nb] = hslToRgb(hh, Math.min(0.55, s * 0.75), Math.min(0.32, nl));
    out = toHex(nr, ng, nb);
  }
  fillCache.set(hex, out);
  return out;
}

/** Picks a readable text color for a cell. */
export function displayText(theme: GridTheme, color: string | undefined, bg: string): string {
  const key = `${color ?? ''}|${bg}|${theme.dark}|${theme.adaptColors}`;
  let out = textCache.get(key);
  if (out) return out;
  const bgDark = luminance(bg) < 0.4;
  if (!color) out = bgDark ? (theme.dark ? theme.cellText : '#eef0f3') : theme.dark ? '#16181c' : theme.cellText;
  else if (!theme.adaptColors || contrastRatio(color, bg) >= 2.4) out = color;
  else {
    // Flip lightness, keeping the hue
    const [r, g, b] = parseHex(color);
    const [hh, s] = rgbToHsl(r, g, b);
    const [nr, ng, nb] = hslToRgb(hh, s, bgDark ? 0.78 : 0.25);
    out = toHex(nr, ng, nb);
    if (contrastRatio(out, bg) < 2.4) out = bgDark ? theme.cellText : '#16181c';
  }
  if (textCache.size > 4000) textCache.clear();
  textCache.set(key, out);
  return out;
}

export function clearColorCaches(): void {
  fillCache.clear();
  textCache.clear();
}
