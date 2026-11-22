// Appearance model: what a "look" contains, the pre-built themes, validation, and turning a look into CSS and grid colours.
import type { Theme } from './grid';

export type ColorKey =
  | 'accent' | 'bar' | 'barText' | 'ribbon' | 'chrome' | 'status' | 'panel' | 'text' | 'muted' | 'border' | 'hover'
  | 'sheet' | 'cellText' | 'grid' | 'hdr' | 'hdrText' | 'sel';
export type Colors = Record<ColorKey, string>;

export interface Grad {
  on: boolean;
  to: string;
  angle: number;
}

export interface Look {
  id: string;
  name: string;
  c: Colors;
  barGrad: Grad;
  ribbonGrad: Grad;
  wall: { mode: 'none' | 'gradient' | 'image'; a: string; b: string; angle: number; radial: boolean };
  texture: { id: string; strength: number; scale: number };
  anim: { id: string; speed: number; amount: number; quality: 'eco' | 'balanced' | 'smooth'; pauseUnfocused: boolean };
  sheetAlpha: number; // 15..100
  uiAlpha: number; // 30..100
  font: string;
  fontSize: number;
  radius: number;
  motion: boolean;
}

export const FONTS: [string, string][] = [
  ['segoe', 'Segoe UI'], ['variable', 'Segoe UI Variable'], ['calibri', 'Calibri'], ['arial', 'Arial'], ['verdana', 'Verdana'],
  ['tahoma', 'Tahoma'], ['trebuchet', 'Trebuchet MS'], ['georgia', 'Georgia'], ['cambria', 'Cambria'], ['consolas', 'Consolas'], ['cascadia', 'Cascadia Code'],
];
const FONT_CSS: Record<string, string> = {
  segoe: "'Segoe UI', system-ui, Arial, sans-serif", variable: "'Segoe UI Variable Text', 'Segoe UI', sans-serif", calibri: "Calibri, Carlito, 'Segoe UI', sans-serif",
  arial: 'Arial, sans-serif', verdana: 'Verdana, sans-serif', tahoma: 'Tahoma, sans-serif', trebuchet: "'Trebuchet MS', sans-serif",
  georgia: 'Georgia, serif', cambria: 'Cambria, Georgia, serif', consolas: 'Consolas, monospace', cascadia: "'Cascadia Code', Consolas, monospace",
};

// ---------------- colour helpers ----------------
const HEX = /^#[0-9a-f]{6}$/i;
export const isHex = (v: unknown): v is string => typeof v === 'string' && HEX.test(v);
export function rgb(h: string): [number, number, number] {
  const n = parseInt(h.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
export function mix(a: string, b: string, t: number): string {
  const x = rgb(a);
  const y = rgb(b);
  return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('');
}
export function alpha(h: string, a: number): string {
  if (a >= 1) return h;
  const [r, g, b] = rgb(h);
  return `rgba(${r},${g},${b},${+a.toFixed(3)})`;
}
export function lum(h: string): number {
  const [r, g, b] = rgb(h).map((v) => {
    const s = v / 255;
    return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
const contrast = (a: string, b: string) => {
  const x = lum(a);
  const y = lum(b);
  return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
};
export const readable = (bg: string) => (contrast(bg, '#ffffff') >= contrast(bg, '#1a1a1a') ? '#ffffff' : '#1a1a1a');
/** The colour nudged towards black or white until it can be read on bg. */
function ink(color: string, bg: string): string {
  const to = lum(bg) > 0.3 ? '#000000' : '#ffffff';
  for (let t = 0; t <= 1; t += 0.1) {
    const c = mix(color, to, t);
    if (contrast(c, bg) >= 3.6) return c;
  }
  return to;
}
export const isDark = (l: Look) => lum(l.c.panel) < 0.3;

/** A complete palette from one colour. */
export function palette(accent: string, dark: boolean, tint = 0.07): Colors {
  const m = (base: string, k = 1) => mix(base, accent, tint * k);
  const bar = dark ? mix(accent, '#000000', 0.22) : accent;
  return dark
    ? {
        accent, bar, barText: readable(bar), ribbon: m('#2d2d2d'), chrome: m('#292929'), status: m('#292929'), panel: m('#2b2b2b'),
        text: '#ececec', muted: '#a4a4a4', border: m('#4a4a4a', 1.5), hover: m('#3d3d3d', 2),
        sheet: m('#1e1e1e'), cellText: '#e8e8e8', grid: m('#3a3a3a', 1.3), hdr: m('#2b2b2b'), hdrText: '#c8c8c8', sel: ink(accent, '#1e1e1e'),
      }
    : {
        accent, bar, barText: readable(bar), ribbon: m('#f8f8f8'), chrome: m('#f3f3f3'), status: m('#f3f3f3'), panel: m('#ffffff', 0.4),
        text: '#242424', muted: '#666666', border: m('#d0d0d0', 1.5), hover: m('#e6e6e6', 1.6),
        sheet: m('#ffffff', 0.25), cellText: '#000000', grid: m('#d4d4d4', 1.2), hdr: m('#f5f5f5'), hdrText: '#444444', sel: ink(accent, '#ffffff'),
      };
}

// ---------------- pre-built themes ----------------
const CLASSIC: Colors = {
  accent: '#217346', bar: '#217346', barText: '#ffffff', ribbon: '#f8f8f8', chrome: '#f3f3f3', status: '#f3f3f3', panel: '#ffffff',
  text: '#242424', muted: '#6b6b6b', border: '#d4d4d4', hover: '#e6e6e6', sheet: '#ffffff', cellText: '#000000', grid: '#d4d4d4',
  hdr: '#f5f5f5', hdrText: '#444444', sel: '#217346',
};
const DARK_C: Colors = {
  accent: '#2d8a58', bar: '#2d8a58', barText: '#ffffff', ribbon: '#2f2f2f', chrome: '#2b2b2b', status: '#2b2b2b', panel: '#2b2b2b',
  text: '#ececec', muted: '#a0a0a0', border: '#484848', hover: '#3c3c3c', sheet: '#1f1f1f', cellText: '#e8e8e8', grid: '#3a3a3a',
  hdr: '#2b2b2b', hdrText: '#c8c8c8', sel: '#3fb57a',
};
// The two original themes keep their exact hand-picked secondary colours.
const EXACT: Record<string, Record<string, string>> = {
  classic: { '--green-dark': '#185c37', '--green-light': '#d3f0e0', '--pressed': '#cfe8da', '--border-soft': '#e1e1e1', '--accent-ink': '#217346', '--thumb': '#c1c1c1', '--thumb-hover': '#a0a0a0' },
  dark: { '--green-dark': '#1f6d43', '--green-light': '#27493a', '--pressed': '#2f5a45', '--border-soft': '#3a3a3a', '--accent-ink': '#8fe0b2', '--thumb': '#5a5a5a', '--thumb-hover': '#a0a0a0' },
};
const EXACT_GRID: Record<string, Partial<Theme>> = {
  classic: { hdrSelBg: '#d3f0e0', hdrSelText: '#0e5c2f', hdrAllSelBg: '#a8d7bc', hdrLine: '#c8c8c8', selFill: 'rgba(33,115,70,0.13)', freeze: '#8c8c8c', filterBg: '#f0f0f0', filterLine: '#a0a0a0', corner: '#b4b4b4' },
  dark: { hdrSelBg: '#27493a', hdrSelText: '#a6e3bf', hdrAllSelBg: '#2f6b4b', hdrLine: '#454545', selFill: 'rgba(63,181,122,0.18)', freeze: '#7a7a7a', filterBg: '#3a3a3a', filterLine: '#6a6a6a', corner: '#5a5a5a' },
};

function base(id: string, name: string, c: Colors): Look {
  return {
    id, name, c,
    barGrad: { on: false, to: mix(c.bar, '#000000', 0.25), angle: 90 },
    ribbonGrad: { on: false, to: mix(c.ribbon, c.accent, 0.12), angle: 180 },
    wall: { mode: 'none', a: mix(c.sheet, c.accent, 0.3), b: c.sheet, angle: 135, radial: false },
    texture: { id: 'none', strength: 40, scale: 100 },
    anim: { id: 'none', speed: 100, amount: 100, quality: 'balanced', pauseUnfocused: true },
    sheetAlpha: 100, uiAlpha: 100, font: 'segoe', fontSize: 12, radius: 0, motion: true,
  };
}

interface Opts {
  tint?: number;
  bar?: [string, string, number?];
  wall?: [string, string, number?, boolean?];
  tex?: [string, number?];
  anim?: string;
  sheet?: number;
  ui?: number;
  radius?: number;
  font?: string;
  over?: Partial<Colors>;
}
function make(id: string, name: string, accent: string, dark: boolean, o: Opts = {}): Look {
  const c = { ...palette(accent, dark, o.tint ?? 0.08), ...o.over };
  if (o.bar) {
    c.bar = o.bar[0];
    c.barText = readable(mix(o.bar[0], o.bar[1], 0.5));
  }
  const l = base(id, name, c);
  if (o.bar) l.barGrad = { on: true, to: o.bar[1], angle: o.bar[2] ?? 90 };
  if (o.wall) l.wall = { mode: 'gradient', a: o.wall[0], b: o.wall[1], angle: o.wall[2] ?? 135, radial: !!o.wall[3] };
  if (o.tex) l.texture = { id: o.tex[0], strength: o.tex[1] ?? 40, scale: 100 };
  if (o.anim) l.anim.id = o.anim;
  const backdrop = o.wall || o.tex || o.anim;
  l.sheetAlpha = o.sheet ?? (backdrop ? 76 : 100);
  l.uiAlpha = o.ui ?? (backdrop ? 84 : 100);
  l.radius = o.radius ?? (backdrop ? 4 : 2);
  if (o.font) l.font = o.font;
  return l;
}

export interface PresetGroup {
  title: string;
  items: Look[];
}
export const PRESETS: PresetGroup[] = [
  {
    title: 'Animated',
    items: [
      make('blackhole', 'Black Hole', '#ff9d3c', true, { anim: 'blackhole', bar: ['#2a1405', '#0b0710'], wall: ['#0a0612', '#020205'], sheet: 62, ui: 72, tint: 0.05 }),
      make('warpdrive', 'Warp Drive', '#58a6ff', true, { anim: 'warp', bar: ['#0d2547', '#070b1a'], wall: ['#070b1a', '#02030a'], sheet: 66, ui: 76 }),
      make('northern', 'Northern Lights', '#3ddc97', true, { anim: 'aurora', bar: ['#0d3b3a', '#1b1642'], wall: ['#06121f', '#0b0824'], sheet: 66, ui: 76 }),
      make('retrowave', 'Retro Wave', '#ff4fa3', true, { anim: 'retro', bar: ['#5b1a8c', '#c42a7b'], wall: ['#14052b', '#3a0a4f'], sheet: 68, ui: 78 }),
      make('deepsea', 'Deep Sea', '#23c4d8', true, { anim: 'waves', bar: ['#074a63', '#0a2342'], wall: ['#04253a', '#020d1c', 180], sheet: 70, ui: 80 }),
      make('fireflies', 'Firefly Forest', '#c6e04a', true, { anim: 'fireflies', bar: ['#1f3a1a', '#10241c'], wall: ['#10241c', '#050d0b', 180], sheet: 68, ui: 78 }),
      make('rainy', 'Rainy Night', '#7f9cf5', true, { anim: 'rain', bar: ['#262f4f', '#141827'], wall: ['#1a2036', '#090b14', 180], sheet: 70, ui: 80 }),
      make('snowday', 'Snow Day', '#4a90d9', false, { anim: 'snow', bar: ['#4a90d9', '#7fb6ea'], wall: ['#b9d6ef', '#eaf3fb', 180], sheet: 74, ui: 82 }),
      make('lavalamp', 'Lava Lamp', '#ff5e3a', true, { anim: 'lava', bar: ['#7a1d12', '#3b0f2c'], wall: ['#25060f', '#0d0309', 180], sheet: 68, ui: 78 }),
      make('network', 'Constellation', '#a78bfa', true, { anim: 'constellation', bar: ['#33246b', '#141033'], wall: ['#17123a', '#07061a'], sheet: 70, ui: 80 }),
    ],
  },
  {
    title: 'Gradient',
    items: [
      make('sunset', 'Sunset', '#f2643d', false, { bar: ['#f2643d', '#d8347f'], wall: ['#ffd8b8', '#f7c1d9'], sheet: 80, ui: 88 }),
      make('ocean', 'Ocean', '#0a84c6', false, { bar: ['#0a84c6', '#16b3a8'], wall: ['#c3e8f7', '#c9f2e8'], sheet: 80, ui: 88 }),
      make('grape', 'Grape Soda', '#8b4fe0', true, { bar: ['#8b4fe0', '#d1408f'], wall: ['#2a1650', '#3f1238'], sheet: 72, ui: 82 }),
      make('peach', 'Peach', '#ef7f6a', false, { bar: ['#f6a07a', '#ef6f8c'], wall: ['#ffe6d6', '#ffd9e2'], sheet: 82, ui: 90 }),
      make('mint', 'Fresh Mint', '#14a98a', false, { bar: ['#14a98a', '#5ac85a'], wall: ['#d3f6ea', '#e9f8cf'], sheet: 82, ui: 90 }),
      make('ember', 'Ember', '#ff7a2f', true, { bar: ['#b5300f', '#ef8a1e'], wall: ['#3a1208', '#160807'], sheet: 74, ui: 84 }),
    ],
  },
  {
    title: 'Dark',
    items: [
      make('dark', 'Sheetly Dark', '#2d8a58', true, { over: DARK_C, radius: 0 }),
      make('midnight', 'Midnight', '#4c8df6', true, { tint: 0.12 }),
      make('graphite', 'Graphite', '#8a94a6', true, { tint: 0.05 }),
      make('oled', 'True Black', '#34d399', true, { tint: 0, over: { sheet: '#000000', ribbon: '#0a0a0a', chrome: '#0a0a0a', status: '#0a0a0a', panel: '#111111', hdr: '#0d0d0d', grid: '#242424', border: '#2c2c2c', hover: '#1f1f1f', bar: '#000000', barText: '#ffffff' } }),
      make('plum', 'Plum', '#c678dd', true, { tint: 0.13 }),
      make('neon', 'Neon', '#00e5c7', true, { tint: 0.06, bar: ['#00a896', '#7b2ff7'], font: 'consolas' }),
    ],
  },
  {
    title: 'Light',
    items: [
      make('classic', 'Sheetly Classic', '#217346', false, { over: CLASSIC, radius: 0 }),
      make('sky', 'Sky', '#2f6fd0', false),
      make('rose', 'Rose', '#d6336c', false),
      make('violet', 'Violet', '#7048e8', false),
      make('sepia', 'Paper', '#9a6b3f', false, { tint: 0.1, tex: ['paper', 45], sheet: 92, ui: 94, font: 'georgia', over: { sheet: '#fbf5e9', ribbon: '#f3ead8', chrome: '#efe5d1', status: '#efe5d1', panel: '#fbf5e9', hdr: '#f1e8d6' } }),
      make('mono', 'Monochrome', '#404040', false, { tint: 0 }),
    ],
  },
];
export const allPresets = () => PRESETS.flatMap((g) => g.items);
export const preset = (id: string) => clone(allPresets().find((p) => p.id === id) ?? allPresets().find((p) => p.id === 'classic')!);
export const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v));

// ---------------- persistence & validation ----------------
const KEY = 'sheetly.look';
const SAVED = 'sheetly.look.saved';
const IMAGE = 'sheetly.look.image';
const num = (v: unknown, lo: number, hi: number, d: number) => (typeof v === 'number' && isFinite(v) ? Math.max(lo, Math.min(hi, v)) : d);
const oneOf = <T extends string>(v: unknown, list: readonly T[], d: T): T => (list.includes(v as T) ? (v as T) : d);

/** Stored looks are untrusted: every field is checked before it reaches CSS. */
export function sanitize(raw: any): Look {
  const d = preset(typeof raw?.id === 'string' ? raw.id : 'classic');
  if (!raw || typeof raw !== 'object') return d;
  const c = { ...d.c };
  for (const k of Object.keys(c) as ColorKey[]) if (isHex(raw.c?.[k])) c[k] = raw.c[k];
  const grad = (g: any, f: Grad): Grad => ({ on: typeof g?.on === 'boolean' ? g.on : f.on, to: isHex(g?.to) ? g.to : f.to, angle: num(g?.angle, 0, 360, f.angle) });
  return {
    id: typeof raw.id === 'string' ? raw.id.slice(0, 40) : 'custom',
    name: typeof raw.name === 'string' ? raw.name.slice(0, 40) : 'Custom',
    c,
    barGrad: grad(raw.barGrad, d.barGrad),
    ribbonGrad: grad(raw.ribbonGrad, d.ribbonGrad),
    wall: {
      mode: oneOf(raw.wall?.mode, ['none', 'gradient', 'image'] as const, d.wall.mode),
      a: isHex(raw.wall?.a) ? raw.wall.a : d.wall.a,
      b: isHex(raw.wall?.b) ? raw.wall.b : d.wall.b,
      angle: num(raw.wall?.angle, 0, 360, d.wall.angle),
      radial: !!raw.wall?.radial,
    },
    texture: { id: typeof raw.texture?.id === 'string' ? raw.texture.id : 'none', strength: num(raw.texture?.strength, 0, 100, 40), scale: num(raw.texture?.scale, 50, 250, 100) },
    anim: {
      id: typeof raw.anim?.id === 'string' ? raw.anim.id : 'none',
      speed: num(raw.anim?.speed, 25, 250, 100),
      amount: num(raw.anim?.amount, 25, 150, 100),
      quality: oneOf(raw.anim?.quality, ['eco', 'balanced', 'smooth'] as const, 'balanced'),
      pauseUnfocused: raw.anim?.pauseUnfocused !== false,
    },
    sheetAlpha: num(raw.sheetAlpha, 15, 100, 100),
    uiAlpha: num(raw.uiAlpha, 30, 100, 100),
    font: FONTS.some(([f]) => f === raw.font) ? raw.font : 'segoe',
    fontSize: num(raw.fontSize, 11, 14, 12),
    radius: num(raw.radius, 0, 12, 0),
    motion: raw.motion !== false,
  };
}

export function loadLook(): Look {
  try {
    const s = localStorage.getItem(KEY);
    if (s) return sanitize(JSON.parse(s));
    return preset(localStorage.getItem('sheetly.theme') === 'dark' ? 'dark' : 'classic');
  } catch {
    return preset('classic');
  }
}
export function saveLook(l: Look): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(l));
    localStorage.setItem('sheetly.theme', isDark(l) ? 'dark' : 'light');
  } catch {
    /* storage full or blocked: the look still applies for this session */
  }
}
export function loadSaved(): Look[] {
  try {
    const a = JSON.parse(localStorage.getItem(SAVED) ?? '[]');
    return Array.isArray(a) ? a.slice(0, 40).map(sanitize) : [];
  } catch {
    return [];
  }
}
export function storeSaved(list: Look[]): void {
  try {
    localStorage.setItem(SAVED, JSON.stringify(list));
  } catch {
    /* ignore */
  }
}
export function loadImage(): string | null {
  try {
    const s = localStorage.getItem(IMAGE);
    return s && s.startsWith('data:image/') ? s : null;
  } catch {
    return null;
  }
}
export function storeImage(data: string): boolean {
  try {
    localStorage.setItem(IMAGE, data);
    return true;
  } catch {
    return false;
  }
}

export const hasBackdrop = (l: Look) => l.wall.mode !== 'none' || l.texture.id !== 'none' || l.anim.id !== 'none';
/** A background is only visible through see-through surfaces, so turn that on the first time one is chosen. */
export function revealBackdrop(l: Look): void {
  if (l.sheetAlpha >= 100) l.sheetAlpha = 78;
  if (l.uiAlpha >= 100) l.uiAlpha = 86;
}

// ---------------- look -> CSS / grid ----------------
const gradCss = (from: string, g: Grad, a: number) => (g.on ? `linear-gradient(${g.angle}deg, ${alpha(from, a)}, ${alpha(g.to, a)})` : alpha(from, a));

export function cssVars(l: Look): Record<string, string> {
  const c = l.c;
  const back = hasBackdrop(l);
  const ui = back ? l.uiAlpha / 100 : 1;
  const sheet = back ? l.sheetAlpha / 100 : 1;
  const dark = isDark(l);
  const vars: Record<string, string> = {
    '--green': c.accent,
    '--green-dark': mix(c.accent, '#000000', 0.2),
    '--green-light': mix(c.sheet, c.accent, 0.24),
    '--on-accent': readable(c.accent),
    '--accent-ink': ink(c.sel, c.panel),
    '--accent-soft': alpha(c.sel, 0.3),
    '--bg': alpha(c.sheet, sheet),
    '--bg-solid': c.sheet,
    '--chrome': alpha(c.chrome, ui),
    '--ribbon': alpha(c.ribbon, ui),
    '--ribbon-bg': gradCss(c.ribbon, l.ribbonGrad, ui),
    '--ribbon-solid': c.ribbon,
    '--panel': c.panel,
    '--border': c.border,
    '--border-soft': mix(c.border, c.ribbon, 0.4),
    '--text': c.text,
    '--muted': c.muted,
    '--hover': alpha(c.hover, Math.max(ui, 0.75)),
    '--pressed': mix(c.ribbon, c.accent, 0.26),
    '--input-bg': alpha(c.sheet, Math.max(ui, 0.8)),
    '--tab-active': c.ribbon,
    '--status': alpha(c.status, ui),
    '--bar': gradCss(c.bar, l.barGrad, Math.max(ui, 0.55)),
    '--bar-solid': l.barGrad.on ? `linear-gradient(${l.barGrad.angle}deg, ${c.bar}, ${l.barGrad.to})` : c.bar,
    '--bar-text': c.barText,
    '--bar-hover': alpha(c.barText, 0.18),
    '--thumb': dark ? mix(c.chrome, '#ffffff', 0.25) : mix(c.chrome, '#000000', 0.22),
    '--thumb-hover': dark ? mix(c.chrome, '#ffffff', 0.45) : mix(c.chrome, '#000000', 0.36),
    '--shadow': dark ? '0 4px 14px rgba(0, 0, 0, 0.6)' : '0 4px 14px rgba(0, 0, 0, 0.22)',
    '--radius': l.radius + 'px',
    '--ui-font': FONT_CSS[l.font] ?? FONT_CSS.segoe,
    '--ui-size': l.fontSize + 'px',
  };
  return { ...vars, ...(EXACT[l.id] ?? {}) };
}

export function gridTheme(l: Look): Theme {
  const c = l.c;
  const back = hasBackdrop(l);
  const sheet = back ? l.sheetAlpha / 100 : 1;
  const hdrA = back ? Math.max(l.uiAlpha / 100, 0.6) : 1;
  const dark = isDark(l);
  const t: Theme = {
    bg: alpha(c.sheet, sheet), solid: c.sheet, grid: back ? alpha(c.grid, Math.max(0.55, sheet)) : c.grid,
    hdrBg: alpha(c.hdr, hdrA), hdrText: c.hdrText, hdrLine: mix(c.hdr, c.hdrText, 0.25),
    hdrSelBg: mix(c.hdr, c.sel, 0.28), hdrSelText: ink(c.sel, mix(c.hdr, c.sel, 0.28)), hdrAllSelBg: mix(c.hdr, c.sel, 0.48),
    text: c.cellText, accent: c.sel, selFill: alpha(c.sel, dark ? 0.2 : 0.14), freeze: mix(c.sheet, c.cellText, 0.45),
    filterBg: mix(c.hdr, c.cellText, 0.06), filterLine: mix(c.hdr, c.cellText, 0.35), corner: mix(c.hdr, c.hdrText, 0.4),
    lightAuto: lum(c.sheet) < 0.3, transparent: back && sheet < 1,
  };
  return { ...t, ...(EXACT_GRID[l.id] ?? {}) };
}

export function applyCss(l: Look): void {
  const root = document.documentElement;
  for (const [k, v] of Object.entries(cssVars(l))) root.style.setProperty(k, v);
  root.dataset.theme = isDark(l) ? 'dark' : 'light';
  root.dataset.motion = l.motion ? 'on' : 'off';
  root.style.colorScheme = isDark(l) ? 'dark' : 'light';
}
