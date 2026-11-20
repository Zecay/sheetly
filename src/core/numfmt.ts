// Excel-style number format engine: General, 0.00, #,##0, percent, dates, times, text, custom codes.
import { DAYS, MONTHS, serialToTime, serialToYmd, weekdayOf, FErr } from './values';

/** Marks the point where "*" fill splits left-aligned and right-aligned text (accounting formats). */
export const FILL_MARK = '';

export interface Formatted {
  text: string;
  color?: string; // css colour from [Red] etc
  isNumber: boolean;
}

const COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00aaaa',
  green: '#007a00',
  magenta: '#cc00cc',
  red: '#e00000',
  white: '#ffffff',
  yellow: '#c8a800',
};

type Tok =
  | { t: 'lit'; s: string }
  | { t: 'digit'; c: '0' | '#' | '?' }
  | { t: 'dec' }
  | { t: 'comma' }
  | { t: 'pct' }
  | { t: 'exp'; sign: boolean }
  | { t: 'text' } // @
  | { t: 'fill'; c: string }
  | { t: 'slash' }
  | { t: 'date'; k: string } // y,m,d,h,s tokens with the full run text e.g. "yyyy", "mmm", "[h]"
  | { t: 'ampm'; s: string }
  | { t: 'general' };

interface Section {
  toks: Tok[];
  cond?: { op: string; v: number };
  color?: string;
  isDate: boolean;
  src: string;
}

const sectionCache = new Map<string, Section[]>();

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQ = false;
  let inB = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '\\' && i + 1 < fmt.length && !inQ) {
      cur += ch + fmt[++i];
      continue;
    }
    if (ch === '"') inQ = !inQ;
    else if (!inQ && ch === '[') inB = true;
    else if (!inQ && ch === ']') inB = false;
    if (ch === ';' && !inQ && !inB) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

function parseSection(src: string): Section {
  const toks: Tok[] = [];
  let cond: Section['cond'];
  let color: string | undefined;
  let isDate = false;
  let i = 0;
  const n = src.length;
  while (i < n) {
    const ch = src[i];
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      while (j < n && src[j] !== '"') s += src[j++];
      toks.push({ t: 'lit', s });
      i = j + 1;
      continue;
    }
    if (ch === '\\') {
      toks.push({ t: 'lit', s: src[i + 1] ?? '' });
      i += 2;
      continue;
    }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      if (j < 0) break;
      const body = src.slice(i + 1, j);
      i = j + 1;
      const lower = body.toLowerCase();
      if (COLORS[lower]) {
        color = COLORS[lower];
        continue;
      }
      const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(body);
      if (m) {
        cond = { op: m[1], v: parseFloat(m[2]) };
        continue;
      }
      if (/^(h+|m+|s+)$/i.test(body)) {
        toks.push({ t: 'date', k: '[' + lower + ']' });
        isDate = true;
        continue;
      }
      if (body.startsWith('$')) {
        // [$-409] or [$€-407]: currency symbol part
        const cur = body.slice(1).split('-')[0];
        if (cur) toks.push({ t: 'lit', s: cur });
        continue;
      }
      continue; // unknown bracket (locale etc.)
    }
    if (ch === '_') {
      toks.push({ t: 'lit', s: ' ' });
      i += 2;
      continue;
    }
    if (ch === '*') {
      toks.push({ t: 'fill', c: src[i + 1] ?? ' ' });
      i += 2;
      continue;
    }
    if (ch === '0' || ch === '#' || ch === '?') {
      toks.push({ t: 'digit', c: ch });
      i++;
      continue;
    }
    if (ch === '.') {
      toks.push({ t: 'dec' });
      i++;
      continue;
    }
    if (ch === ',') {
      toks.push({ t: 'comma' });
      i++;
      continue;
    }
    if (ch === '%') {
      toks.push({ t: 'pct' });
      i++;
      continue;
    }
    if (ch === '@') {
      toks.push({ t: 'text' });
      i++;
      continue;
    }
    if (ch === '/') {
      toks.push({ t: 'slash' });
      i++;
      continue;
    }
    if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) {
      toks.push({ t: 'exp', sign: src[i + 1] === '+' });
      i += 2;
      continue;
    }
    const rest = src.slice(i);
    if (/^general/i.test(rest)) {
      toks.push({ t: 'general' });
      i += 7;
      continue;
    }
    if (/^(AM\/PM|A\/P)/i.test(rest)) {
      const m = /^(AM\/PM|A\/P)/i.exec(rest)!;
      toks.push({ t: 'ampm', s: m[1] });
      isDate = true;
      i += m[1].length;
      continue;
    }
    if (/[ymdhsYMDHS]/.test(ch)) {
      let j = i;
      const lc = ch.toLowerCase();
      while (j < n && src[j].toLowerCase() === lc) j++;
      toks.push({ t: 'date', k: lc.repeat(j - i) });
      isDate = true;
      i = j;
      continue;
    }
    toks.push({ t: 'lit', s: ch });
    i++;
  }
  return { toks, cond, color, isDate, src };
}

function getSections(fmt: string): Section[] {
  let s = sectionCache.get(fmt);
  if (!s) {
    s = splitSections(fmt).map(parseSection);
    if (sectionCache.size > 500) sectionCache.clear();
    sectionCache.set(fmt, s);
  }
  return s;
}

export function isDateFormat(fmt: string): boolean {
  if (!fmt || fmt === 'General') return false;
  const secs = getSections(fmt);
  return secs.length > 0 && secs[0].isDate;
}

export function isTextFormat(fmt: string): boolean {
  return fmt === '@';
}

function roundHalfUp(x: number, d: number): number {
  if (d > 15) return x;
  const neg = x < 0;
  const a = Math.abs(x);
  const s = a.toString();
  if (s.includes('e')) return neg ? -a : a;
  const r = Number(Math.round(Number(a + 'e' + d)) + 'e-' + d);
  return neg ? -r : r;
}

export function generalFormat(x: number): string {
  if (x === 0) return '0';
  if (!isFinite(x)) return String(x);
  const a = Math.abs(x);
  if (a >= 1e11 || a < 1e-9) {
    let s = x.toExponential(5);
    let [m, e] = s.split('e');
    if (m.includes('.')) m = m.replace(/0+$/, '').replace(/\.$/, '');
    const sign = e[0] === '-' ? '-' : '+';
    const digits = e.replace(/^[+-]/, '').padStart(2, '0');
    return m + 'E' + sign + digits;
  }
  const intDigits = a >= 1 ? Math.floor(Math.log10(a)) + 1 : 1;
  const dec = Math.max(0, 10 - intDigits);
  let s = roundHalfUp(x, dec).toFixed(dec);
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  if (s === '-0') s = '0';
  return s;
}

function condOk(sec: Section, x: number): boolean {
  if (!sec.cond) return true;
  const { op, v } = sec.cond;
  switch (op) {
    case '<': return x < v;
    case '>': return x > v;
    case '<=': return x <= v;
    case '>=': return x >= v;
    case '=': return x === v;
    case '<>': return x !== v;
  }
  return true;
}

/** Pick the section for a number and say whether the minus sign should be dropped. */
function pickSection(secs: Section[], x: number): { sec: Section; abs: boolean } {
  const hasCond = secs.some((s) => s.cond);
  if (hasCond) {
    for (let i = 0; i < Math.min(secs.length, 2); i++) {
      const s = secs[i];
      if (s.cond && condOk(s, x)) return { sec: s, abs: i > 0 && x < 0 ? true : false };
    }
    // fallback: the first section without a condition, else last
    const rest = secs.find((s) => !s.cond);
    return { sec: rest ?? secs[secs.length - 1], abs: false };
  }
  if (secs.length === 1) return { sec: secs[0], abs: false };
  if (secs.length === 2) return x < 0 ? { sec: secs[1], abs: true } : { sec: secs[0], abs: false };
  if (x < 0) return { sec: secs[1], abs: true };
  if (x === 0) return { sec: secs[2], abs: false };
  return { sec: secs[0], abs: false };
}

export function formatValue(v: number | string | boolean | FErr | null, fmt: string, opts?: { width?: number }): Formatted {
  if (v === null) return { text: '', isNumber: false };
  if (v instanceof FErr) return { text: v.code, isNumber: false };
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', isNumber: false };
  if (typeof v === 'string') {
    if (!fmt || fmt === 'General') return { text: v, isNumber: false };
    const secs = getSections(fmt);
    const sec = secs.length >= 4 ? secs[3] : secs.find((s) => s.toks.some((t) => t.t === 'text'));
    if (!sec) return { text: v, isNumber: false };
    let out = '';
    for (const t of sec.toks) {
      if (t.t === 'text') out += v;
      else if (t.t === 'lit') out += t.s;
    }
    return { text: out, color: sec.color, isNumber: false };
  }
  // number
  if (!fmt || fmt === 'General') return { text: generalFormat(v), isNumber: true };
  const secs = getSections(fmt);
  const { sec, abs } = pickSection(secs, v);
  let x = abs ? Math.abs(v) : v;
  if (sec.toks.length === 1 && sec.toks[0].t === 'general') return { text: generalFormat(x), color: sec.color, isNumber: true };
  if (sec.toks.length === 0) return { text: '', color: sec.color, isNumber: true };
  let text: string;
  if (sec.isDate) text = formatDate(sec, x);
  else text = formatNumberSection(sec, x, secs.length === 1 && v < 0);
  return { text, color: sec.color, isNumber: true };
}

// ---------------- numbers ----------------

function formatNumberSection(sec: Section, x: number, addMinus: boolean): string {
  const toks = sec.toks;
  let neg = x < 0;
  if (addMinus) neg = x < 0;
  else if (x < 0) neg = false;
  let a = Math.abs(x);

  // fractions
  if (toks.some((t) => t.t === 'slash') && toks.some((t) => t.t === 'digit')) return formatFraction(toks, a, addMinus && x < 0);

  // percent scaling
  let pcts = 0;
  for (const t of toks) if (t.t === 'pct') pcts++;
  if (pcts) a *= Math.pow(100, pcts);

  // trailing commas scale by 1000
  const firstDigit = toks.findIndex((t) => t.t === 'digit');
  let lastDigit = -1;
  for (let i = toks.length - 1; i >= 0; i--) if (toks[i].t === 'digit') { lastDigit = i; break; }
  if (firstDigit < 0) {
    // no placeholders: just literals (e.g. "N/A" or text only section)
    let out = '';
    for (const t of toks) out += tokLit(t);
    return (addMinus && x < 0 ? '-' : '') + out;
  }
  let scale = 0;
  for (let i = lastDigit + 1; i < toks.length && toks[i].t === 'comma'; i++) scale++;
  if (scale) a /= Math.pow(1000, scale);

  const decIdx = toks.findIndex((t, i) => t.t === 'dec' && i > firstDigit - 1 && i < lastDigit + 1 && i >= firstDigit);
  const expIdx = toks.findIndex((t) => t.t === 'exp');
  const endInt = decIdx >= 0 ? decIdx : expIdx >= 0 ? expIdx : lastDigit + 1;

  let decPlaces = 0;
  if (decIdx >= 0) {
    const stop = expIdx >= 0 ? expIdx : toks.length;
    for (let i = decIdx + 1; i < stop; i++) if (toks[i].t === 'digit') decPlaces++;
  }

  let expStr = '';
  if (expIdx >= 0) {
    const et = toks[expIdx] as { t: 'exp'; sign: boolean };
    let expDigits = 0;
    for (let i = expIdx + 1; i < toks.length; i++) if (toks[i].t === 'digit') expDigits++;
    let intDigitsPattern = 0;
    for (let i = firstDigit; i < endInt; i++) if (toks[i].t === 'digit') intDigitsPattern++;
    let e = a === 0 ? 0 : Math.floor(Math.log10(a));
    if (intDigitsPattern > 1) e = Math.floor(e / intDigitsPattern) * intDigitsPattern;
    let mant = a === 0 ? 0 : a / Math.pow(10, e);
    mant = roundHalfUp(mant, decPlaces);
    if (mant >= Math.pow(10, Math.max(1, intDigitsPattern))) {
      mant /= 10;
      e += 1;
    }
    a = mant;
    const es = String(Math.abs(e)).padStart(expDigits, '0');
    expStr = 'E' + (e < 0 ? '-' : et.sign ? '+' : '') + es;
  } else {
    a = roundHalfUp(a, decPlaces);
  }

  const fixed = a.toFixed(Math.min(decPlaces, 100));
  let [ip, fp = ''] = fixed.split('.');
  // integer pattern contains '0'?
  let hasZeroInt = false;
  for (let i = firstDigit; i < endInt; i++) {
    const t = toks[i];
    if (t.t === 'digit' && t.c === '0') hasZeroInt = true;
  }
  if (ip === '0' && !hasZeroInt) ip = '';
  const grouping = (() => {
    for (let i = firstDigit; i < lastDigit; i++) if (toks[i].t === 'comma') return true;
    return false;
  })();

  // prefix
  let out = '';
  let negDone = false;
  const putNeg = () => {
    if (addMinus && x < 0 && !negDone) {
      out += '-';
      negDone = true;
    }
  };
  for (let i = 0; i < firstDigit; i++) {
    const t = toks[i];
    if (t.t === 'lit' && !negDone && addMinus && x < 0 && !/[\s]/.test(t.s)) {
      // minus goes before currency symbol etc.
      putNeg();
    }
    out += tokLit(t);
  }
  putNeg();

  // integer region, right to left
  const intToks = toks.slice(firstDigit, endInt);
  let di = ip.length - 1;
  let count = 0;
  const rev: string[] = [];
  const pushDigit = (ch: string) => {
    if (grouping && count > 0 && count % 3 === 0 && /\d/.test(ch)) rev.push(',');
    rev.push(ch);
    count++;
  };
  for (let k = intToks.length - 1; k >= 0; k--) {
    const t = intToks[k];
    if (t.t === 'digit') {
      if (di >= 0) pushDigit(ip[di--]);
      else if (t.c === '0') pushDigit('0');
      else if (t.c === '?') {
        rev.push(' ');
        count++;
      }
    } else if (t.t === 'comma') {
      // grouping separator: nothing
    } else rev.push(tokLit(t));
  }
  while (di >= 0) pushDigit(ip[di--]);
  out += rev.reverse().join('');

  // decimal region
  if (decIdx >= 0) {
    const stop = expIdx >= 0 ? expIdx : lastDigit + 1;
    const decToks = toks.slice(decIdx + 1, stop);
    const digitToks = decToks.filter((t) => t.t === 'digit') as { t: 'digit'; c: string }[];
    // trim trailing '#'
    let keep = fp.length;
    for (let k = digitToks.length - 1; k >= 0 && digitToks[k].c === '#' && keep > 0 && fp[keep - 1] === '0'; k--) keep--;
    const kept = fp.slice(0, keep);
    let res = '';
    let di2 = 0;
    for (const t of decToks) {
      if (t.t === 'digit') {
        if (di2 < kept.length) res += kept[di2];
        else if (t.c === '?') res += ' ';
        di2++;
      } else if (t.t !== 'comma') res += tokLit(t);
    }
    if (kept.length > 0 || decToks.some((t) => t.t === 'digit' && t.c !== '#') || true) out += '.' + res;
    // Excel hides the point when no decimals remain and pattern is all '#'
    if (kept.length === 0 && digitToks.every((t) => t.c === '#')) out = out.slice(0, out.length - 1 - res.length) + res;
  }
  out += expStr;
  // suffix
  for (let i = (expIdx >= 0 ? toks.length : lastDigit + 1); i < toks.length; i++) {
    const t = toks[i];
    if (t.t === 'comma') continue;
    out += tokLit(t);
  }
  if (expIdx < 0) {
    // tokens between lastDigit+1.. already included above; remove duplicates handled by loop
  }
  return out;
}

function tokLit(t: Tok): string {
  switch (t.t) {
    case 'lit': return t.s;
    case 'pct': return '%';
    case 'fill': return FILL_MARK;
    case 'comma': return '';
    case 'slash': return '/';
    case 'dec': return '.';
    case 'text': return '';
    default: return '';
  }
}

function formatFraction(toks: Tok[], a: number, neg: boolean): string {
  const slashIdx = toks.findIndex((t) => t.t === 'slash');
  let denDigits = 0;
  for (let i = slashIdx + 1; i < toks.length; i++) if (toks[i].t === 'digit') denDigits++;
  const maxDen = Math.pow(10, Math.max(1, denDigits)) - 1;
  const whole = Math.floor(a);
  const frac = a - whole;
  let bn = 0;
  let bd = 1;
  let best = frac;
  for (let d = 1; d <= maxDen; d++) {
    const n = Math.round(frac * d);
    const err = Math.abs(frac - n / d);
    if (err < best - 1e-12) {
      best = err;
      bn = n;
      bd = d;
    }
  }
  let w = whole;
  if (bn === bd) {
    w += 1;
    bn = 0;
    bd = 1;
  }
  const hasWhole = toks.slice(0, slashIdx).filter((t) => t.t === 'digit').length > 1;
  let out = neg ? '-' : '';
  if (!hasWhole) {
    bn += w * bd;
    w = 0;
  }
  if (hasWhole && w > 0) out += w + ' ';
  else if (hasWhole && bn === 0) out += w;
  if (bn !== 0) out += bn + '/' + bd;
  else if (!hasWhole) out += '0';
  return out;
}

// ---------------- dates ----------------

function formatDate(sec: Section, x: number): string {
  const toks = sec.toks;
  const hasAmPm = toks.some((t) => t.t === 'ampm');
  if (x < 0) return '#'.repeat(10);
  const ymd = serialToYmd(x);
  const tm = serialToTime(x);
  const total = Math.round(x * 86400);
  // look ahead for fractional seconds, e.g. ss.00
  let out = '';
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === 'date') {
      const k = t.k;
      switch (k[0]) {
        case 'y':
          out += k.length <= 2 ? String(ymd.y % 100).padStart(2, '0') : String(ymd.y).padStart(4, '0');
          break;
        case 'd':
          if (k.length === 1) out += ymd.d;
          else if (k.length === 2) out += String(ymd.d).padStart(2, '0');
          else if (k.length === 3) out += DAYS[weekdayOf(x)].slice(0, 3);
          else out += DAYS[weekdayOf(x)];
          break;
        case 'h': {
          let h = tm.h;
          if (hasAmPm) h = h % 12 === 0 ? 12 : h % 12;
          out += k.length === 1 ? h : String(h).padStart(2, '0');
          break;
        }
        case 's':
          out += k.length === 1 ? tm.s : String(tm.s).padStart(2, '0');
          break;
        case 'm': {
          // minute if previous date token is h or next is s
          let isMin = false;
          for (let j = i - 1; j >= 0; j--) {
            const p = toks[j];
            if (p.t === 'date') { isMin = p.k[0] === 'h' || p.k === '[h]'; break; }
            if (p.t === 'ampm') break;
          }
          if (!isMin) {
            for (let j = i + 1; j < toks.length; j++) {
              const p = toks[j];
              if (p.t === 'date') { isMin = p.k[0] === 's'; break; }
              if (p.t === 'ampm') break;
            }
          }
          if (isMin) out += k.length === 1 ? tm.mi : String(tm.mi).padStart(2, '0');
          else if (k.length === 1) out += ymd.m;
          else if (k.length === 2) out += String(ymd.m).padStart(2, '0');
          else if (k.length === 3) out += MONTHS[ymd.m - 1].slice(0, 3);
          else if (k.length === 5) out += MONTHS[ymd.m - 1][0];
          else out += MONTHS[ymd.m - 1];
          break;
        }
        case '[': {
          const unit = k[1];
          const w = k.length - 2;
          const v = unit === 'h' ? Math.floor(total / 3600) : unit === 'm' ? Math.floor(total / 60) : total;
          out += String(v).padStart(w, '0');
          break;
        }
      }
    } else if (t.t === 'ampm') {
      const pm = tm.h >= 12;
      if (/^A\/P$/i.test(t.s)) out += pm ? 'P' : 'A';
      else out += pm ? 'PM' : 'AM';
    } else if (t.t === 'dec') {
      // fractional seconds
      let cnt = 0;
      let j = i + 1;
      while (j < toks.length && toks[j].t === 'digit') { cnt++; j++; }
      if (cnt > 0) {
        out += '.' + String(tm.ms).padStart(3, '0').slice(0, cnt).padEnd(cnt, '0');
        i = j - 1;
      } else out += '.';
    } else if (t.t === 'comma') out += ',';
    else out += tokLit(t);
  }
  return out;
}

// ---------------- helpers for UI ----------------

export const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

export const FMT_CURRENCY = '"$"#,##0.00';
export const FMT_ACCOUNTING = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
export const FMT_PERCENT = '0%';
export const FMT_DATE = 'm/d/yyyy';
export const FMT_TIME = 'h:mm:ss AM/PM';

/** Add or remove a decimal place from a format code (Increase / Decrease Decimal buttons). */
export function changeDecimals(fmt: string, delta: number): string {
  if (!fmt || fmt === 'General') {
    return delta > 0 ? '0.0' : '0';
  }
  const secs = splitSections(fmt);
  const out = secs.map((s) => {
    const m = /^(.*?)(0|#|\?)(\.(?:0|#|\?)*)?((?:[^0#?.][\s\S]*)?)$/.exec(s);
    // find the last digit placeholder run manually
    const re = /[0#?](?:\.[0#?]*)?/g;
    let last: RegExpExecArray | null = null;
    let mm: RegExpExecArray | null;
    let sc = s;
    // skip quoted text for matching by masking it
    const masked = s.replace(/"[^"]*"/g, (q) => '\u0001'.repeat(q.length)).replace(/\[[^\]]*\]/g, (q) => '\u0001'.repeat(q.length));
    while ((mm = re.exec(masked))) last = mm;
    if (!last) return s;
    void m; void sc;
    const start = last.index;
    const run = s.slice(start, start + last[0].length);
    let newRun: string;
    const dot = run.indexOf('.');
    if (delta > 0) newRun = dot >= 0 ? run + '0' : run + '.0';
    else {
      if (dot < 0) return s;
      newRun = run.slice(0, run.length - 1);
      if (newRun.endsWith('.')) newRun = newRun.slice(0, -1);
    }
    return s.slice(0, start) + newRun + s.slice(start + run.length);
  });
  return out.join(';');
}
