// Color math and theme parsing shared by the spreadsheet and presentation readers.
import { attr, child, children, numAttr, path, type Document, type Element } from './xml';

export interface Rgb {
  r: number;
  g: number;
  b: number;
  /** 0..1 */
  a: number;
}

export interface Theme {
  /** Scheme colors keyed by name: dk1, lt1, dk2, lt2, accent1..6, hlink, folHlink. */
  colors: Record<string, string>;
  majorFont: ThemeFont;
  minorFont: ThemeFont;
}

export interface ThemeFont {
  latin?: string;
  ea?: string;
}

export const DEFAULT_THEME: Theme = {
  colors: {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '4472C4',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '5B9BD5',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
  },
  majorFont: { latin: 'Calibri Light' },
  minorFont: { latin: 'Calibri' },
};

export function parseTheme(doc: Document | undefined): Theme {
  const root = doc?.documentElement;
  const elements = path(root, 'themeElements');
  if (!elements) return DEFAULT_THEME;
  const colors: Record<string, string> = { ...DEFAULT_THEME.colors };
  for (const el of children(child(elements, 'clrScheme'))) {
    const value = child(el, 'srgbClr') ? attr(child(el, 'srgbClr'), 'val') : attr(child(el, 'sysClr'), 'lastClr');
    if (el.localName && value) colors[el.localName] = value.toUpperCase();
  }
  const fontScheme = child(elements, 'fontScheme');
  return {
    colors,
    majorFont: parseThemeFont(child(fontScheme, 'majorFont')) ?? DEFAULT_THEME.majorFont,
    minorFont: parseThemeFont(child(fontScheme, 'minorFont')) ?? DEFAULT_THEME.minorFont,
  };
}

function parseThemeFont(el: Element | undefined): ThemeFont | undefined {
  if (!el) return undefined;
  const latin = attr(child(el, 'latin'), 'typeface') || undefined;
  let ea = attr(child(el, 'ea'), 'typeface') || undefined;
  if (!ea) {
    const jpan = children(el, 'font').find((f) => attr(f, 'script') === 'Jpan');
    ea = attr(jpan, 'typeface') || undefined;
  }
  return { latin, ea };
}

export function hexToRgb(hex: string): Rgb {
  const h = hex.length === 8 ? hex.slice(2) : hex;
  const n = parseInt(h.padStart(6, '0').slice(0, 6), 16) || 0;
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

export function rgbToCss({ r, g, b, a }: Rgb): string {
  const hex = (v: number) => Math.round(clamp(v, 0, 255)).toString(16).padStart(2, '0');
  if (a >= 1) return `#${hex(r)}${hex(g)}${hex(b)}`;
  return `rgba(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)}, ${+a.toFixed(3)})`;
}

function clamp(v: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, v));
}

interface Hsl {
  h: number;
  s: number;
  l: number;
}

function rgbToHsl({ r, g, b }: Rgb): Hsl {
  const rn = r / 255, gn = g / 255, bn = b / 255;
  const max = Math.max(rn, gn, bn), min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return { h: 0, s: 0, l };
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === rn) h = (gn - bn) / d + (gn < bn ? 6 : 0);
  else if (max === gn) h = (bn - rn) / d + 2;
  else h = (rn - gn) / d + 4;
  return { h: h / 6, s, l };
}

function hslToRgb({ h, s, l }: Hsl, a: number): Rgb {
  if (s === 0) return { r: l * 255, g: l * 255, b: l * 255, a };
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return { r: hue(p, q, h + 1 / 3) * 255, g: hue(p, q, h) * 255, b: hue(p, q, h - 1 / 3) * 255, a };
}

/** SpreadsheetML tint: -1 darkens to black, +1 lightens to white (applied in HSL space). */
export function applyTint(rgb: Rgb, tint: number): Rgb {
  const hsl = rgbToHsl(rgb);
  hsl.l = tint < 0 ? hsl.l * (1 + tint) : hsl.l * (1 - tint) + tint;
  return hslToRgb(hsl, rgb.a);
}

/** Applies DrawingML color transforms (lumMod, lumOff, tint, shade, satMod, alpha). */
export function applyDrawingModifiers(rgb: Rgb, modifiers: Element[]): Rgb {
  let color = { ...rgb };
  for (const mod of modifiers) {
    const val = (numAttr(mod, 'val') ?? 0) / 100_000;
    switch (mod.localName) {
      case 'alpha':
        color.a = val;
        break;
      case 'lumMod': {
        const hsl = rgbToHsl(color);
        hsl.l = clamp(hsl.l * val, 0, 1);
        color = hslToRgb(hsl, color.a);
        break;
      }
      case 'lumOff': {
        const hsl = rgbToHsl(color);
        hsl.l = clamp(hsl.l + val, 0, 1);
        color = hslToRgb(hsl, color.a);
        break;
      }
      case 'satMod': {
        const hsl = rgbToHsl(color);
        hsl.s = clamp(hsl.s * val, 0, 1);
        color = hslToRgb(hsl, color.a);
        break;
      }
      case 'tint':
        color = { r: 255 - (255 - color.r) * val, g: 255 - (255 - color.g) * val, b: 255 - (255 - color.b) * val, a: color.a };
        break;
      case 'shade':
        color = { r: color.r * val, g: color.g * val, b: color.b * val, a: color.a };
        break;
    }
  }
  return color;
}

/** Legacy SpreadsheetML indexed palette (indices 0-63). */
export const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const PRESET_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
  cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080', orange: 'FFA500', purple: '800080',
  dkGray: 'A9A9A9', ltGray: 'D3D3D3', dkBlue: '00008B', dkRed: '8B0000', dkGreen: '006400', navy: '000080',
};

/** Hex value of a DrawingML `prstClr` name (common names only; unknown names map to black). */
export function presetColor(name: string): string {
  return PRESET_COLORS[name] ?? '000000';
}
