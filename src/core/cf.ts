// Conditional formatting evaluation.
import { cellKey } from './addr';
import { Ctx, compareValues, evaluate, parseNumberText } from './eval';
import { CFRule, Sheet, Workbook } from './model';
import { Node, parseFormula, shiftFormula } from './parser';
import { FErr, Value } from './values';
import { displayText } from './ops';

export interface CFResult {
  color?: string;
  fill?: string;
  bold?: boolean;
  italic?: boolean;
  bar?: { pct: number; color: string };
}

interface RuleStats {
  version: number;
  min: number;
  max: number;
  mid: number;
  avg: number;
  counts?: Map<string, number>;
  sorted?: number[];
}

const astCache = new Map<string, Node | null>();
function parseCached(f: string): Node | null {
  let a = astCache.get(f);
  if (a === undefined) {
    try {
      a = parseFormula(f);
    } catch {
      a = null;
    }
    if (astCache.size > 2000) astCache.clear();
    astCache.set(f, a);
  }
  return a;
}

function evalText(wb: Workbook, sheet: Sheet, r: number, c: number, text: string): Value {
  const t = text.trim();
  if (t === '') return null;
  const f = t.startsWith('=') ? t.slice(1) : t;
  if (!t.startsWith('=')) {
    const n = parseNumberText(t);
    if (n !== null) return n;
    if (/^".*"$/.test(t)) return t.slice(1, -1);
    if (/^(true|false)$/i.test(t)) return t.toLowerCase() === 'true';
    return t;
  }
  const ast = parseCached(f);
  if (!ast) return null;
  const ctx: Ctx = { wb, sheet, row: r, col: c };
  try {
    return evaluate(ast, ctx);
  } catch {
    return null;
  }
}

const hexToRgb = (h: string): [number, number, number] => {
  const m = /^#?([0-9a-f]{6})$/i.exec(h);
  if (!m) return [255, 255, 255];
  const n = parseInt(m[1], 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};
const lerp = (a: string, b: string, t: number): string => {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  const o = x.map((v, i) => Math.round(v + (y[i] - v) * t));
  return '#' + o.map((v) => v.toString(16).padStart(2, '0')).join('');
};

export class CFEngine {
  private stats = new Map<number, RuleStats>();
  constructor(private wb: Workbook) {}

  private statsFor(sheet: Sheet, rule: CFRule): RuleStats {
    let st = this.stats.get(rule.id);
    if (st && st.version === this.wb.dataVersion) return st;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    const { r1, c1, r2, c2 } = rule.rect;
    const area = (r2 - r1 + 1) * (c2 - c1 + 1);
    const visit = (cell: { v: Value } | undefined) => {
      if (!cell || cell.v === null) return;
      if (typeof cell.v === 'number') nums.push(cell.v);
      if (rule.type === 'dup' || rule.type === 'unique') {
        const k = typeof cell.v + ':' + String(cell.v instanceof FErr ? cell.v.code : cell.v).toLowerCase();
        counts.set(k, (counts.get(k) ?? 0) + 1);
      }
    };
    if (area <= sheet.cells.size * 2 || area < 5000) {
      const rr2 = Math.min(r2, Math.max(r1, sheet.maxRow));
      const cc2 = Math.min(c2, Math.max(c1, sheet.maxCol));
      for (let r = r1; r <= rr2; r++) for (let c = c1; c <= cc2; c++) visit(sheet.cells.get(cellKey(r, c)));
    } else {
      for (const [k, cell] of sheet.cells) {
        const r = Math.floor(k / 16384);
        const c = k % 16384;
        if (r >= r1 && r <= r2 && c >= c1 && c <= c2) visit(cell);
      }
    }
    nums.sort((a, b) => a - b);
    const min = nums.length ? nums[0] : 0;
    const max = nums.length ? nums[nums.length - 1] : 0;
    const mid = nums.length ? (nums.length % 2 ? nums[nums.length >> 1] : (nums[(nums.length >> 1) - 1] + nums[nums.length >> 1]) / 2) : 0;
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    st = { version: this.wb.dataVersion, min, max, mid, avg, counts, sorted: nums };
    this.stats.set(rule.id, st);
    this.pruneStats();
    return st;
  }

  /** Drop stats of rules that no longer exist on any sheet. */
  private pruneStats(): void {
    const live = new Set<number>();
    for (const sh of this.wb.sheets) for (const r of sh.cf) live.add(r.id);
    if (this.stats.size <= live.size) return;
    for (const id of Array.from(this.stats.keys())) if (!live.has(id)) this.stats.delete(id);
  }

  /** Combine all rules affecting a cell. */
  evaluate(sheet: Sheet, r: number, c: number): CFResult | null {
    if (!sheet.cf.length) return null;
    let out: CFResult | null = null;
    const cell = sheet.cells.get(cellKey(r, c));
    const v: Value = cell ? cell.v : null;
    for (const rule of sheet.cf) {
      const { r1, c1, r2, c2 } = rule.rect;
      if (r < r1 || r > r2 || c < c1 || c > c2) continue;
      const res = this.applyRule(sheet, rule, r, c, v);
      if (!res) continue;
      if (!out) out = {};
      if (res.color && !out.color) out.color = res.color;
      if (res.fill && !out.fill) out.fill = res.fill;
      if (res.bold && out.bold === undefined) out.bold = res.bold;
      if (res.italic && out.italic === undefined) out.italic = res.italic;
      if (res.bar && !out.bar) out.bar = res.bar;
    }
    return out;
  }

  private applyRule(sheet: Sheet, rule: CFRule, r: number, c: number, v: Value): CFResult | null {
    const style = rule.style ?? {};
    const hit = (): CFResult => ({ color: style.color, fill: style.fill, bold: style.bold, italic: style.italic });
    switch (rule.type) {
      case 'cell': {
        if (v === null || v instanceof FErr) return rule.op === '<>' && v === null ? hit() : null;
        const a = evalText(this.wb, sheet, rule.rect.r1, rule.rect.c1, rule.v1 ?? '');
        const op = rule.op ?? '>';
        if (op === 'between' || op === 'notbetween') {
          const b = evalText(this.wb, sheet, rule.rect.r1, rule.rect.c1, rule.v2 ?? '');
          if (typeof v === 'number' && typeof a === 'number' && typeof b === 'number') {
            const inside = v >= Math.min(a, b) && v <= Math.max(a, b);
            return (op === 'between') === inside ? hit() : null;
          }
          return null;
        }
        if (typeof v !== typeof a && !(typeof v === 'number' && typeof a === 'number')) {
          return op === '<>' ? hit() : null;
        }
        const cmp = compareValues(v, a);
        const ok = op === '>' ? cmp > 0 : op === '<' ? cmp < 0 : op === '>=' ? cmp >= 0 : op === '<=' ? cmp <= 0 : op === '=' ? cmp === 0 : cmp !== 0;
        return ok ? hit() : null;
      }
      case 'text': {
        const text = displayText(this.wb, sheet.cells.get(cellKey(r, c))).toLowerCase();
        const q = (rule.v1 ?? '').toLowerCase();
        const op = rule.op ?? 'contains';
        const ok = op === 'contains' ? text.includes(q) : op === 'notcontains' ? !text.includes(q) : op === 'begins' ? text.startsWith(q) : text.endsWith(q);
        return ok && v !== null ? hit() : null;
      }
      case 'dup':
      case 'unique': {
        if (v === null) return null;
        const st = this.statsFor(sheet, rule);
        const k = typeof v + ':' + String(v instanceof FErr ? v.code : v).toLowerCase();
        const n = st.counts!.get(k) ?? 0;
        return (rule.type === 'dup' ? n > 1 : n === 1) ? hit() : null;
      }
      case 'top':
      case 'bottom': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(sheet, rule);
        const n = Math.max(1, rule.n ?? 10);
        const arr = st.sorted!;
        if (rule.type === 'top') return v >= arr[Math.max(0, arr.length - n)] ? hit() : null;
        return v <= arr[Math.min(arr.length - 1, n - 1)] ? hit() : null;
      }
      case 'above':
      case 'below': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(sheet, rule);
        return (rule.type === 'above' ? v > st.avg : v < st.avg) ? hit() : null;
      }
      case 'formula': {
        const base = (rule.v1 ?? '').replace(/^=/, '');
        const f = shiftFormula(base, r - rule.rect.r1, c - rule.rect.c1);
        const ast = parseCached(f);
        if (!ast) return null;
        const ctx: Ctx = { wb: this.wb, sheet, row: r, col: c };
        let res: Value;
        try {
          res = evaluate(ast, ctx);
        } catch {
          return null;
        }
        const truthy = typeof res === 'boolean' ? res : typeof res === 'number' ? res !== 0 : false;
        return truthy ? hit() : null;
      }
      case 'scale': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(sheet, rule);
        const cols = rule.colors ?? ['#f8696b', '#63be7b'];
        if (st.max === st.min) return { fill: cols[0] };
        if (cols.length === 3) {
          if (v <= st.mid) return { fill: lerp(cols[0], cols[1], st.mid === st.min ? 1 : (v - st.min) / (st.mid - st.min)) };
          return { fill: lerp(cols[1], cols[2], st.max === st.mid ? 1 : (v - st.mid) / (st.max - st.mid)) };
        }
        return { fill: lerp(cols[0], cols[1], (v - st.min) / (st.max - st.min)) };
      }
      case 'bar': {
        if (typeof v !== 'number') return null;
        const st = this.statsFor(sheet, rule);
        const lo = Math.min(0, st.min);
        const pct = st.max === lo ? 1 : Math.max(0.03, (v - lo) / (st.max - lo));
        return { bar: { pct: Math.min(1, pct), color: rule.barColor ?? '#638ec6' } };
      }
    }
    return null;
  }
}
