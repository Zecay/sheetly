// Formula evaluator: AST -> value. Ranges are lazy RangeVal objects.
import { MAX_COLS, MAX_ROWS, cellKey } from './addr';
import type { Node } from './parser';
import { E, FErr, Value, isErr, ymdToSerial } from './values';

export interface CellLike {
  v: Value;
}
export interface SheetLike {
  id: number;
  name: string;
  cells: Map<number, CellLike>;
  maxRow: number; // upper bound of used rows (0-based)
  maxCol: number;
}
export interface NameDef {
  name: string;
  refers: string; // formula text, no "="
  scope?: number; // sheet id for a sheet-scoped name; undefined = workbook scope
}
export interface WbLike {
  sheetByName(name: string): SheetLike | undefined;
  getName(name: string, sheet?: { id: number }): NameDef | undefined;
  nameAst(def: NameDef): Node | null;
}

export interface RangeVal {
  k: 'rng';
  sheet: SheetLike;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}
export interface ArrVal {
  k: 'arr';
  rows: Value[][];
}
export type Arg = Value | RangeVal | ArrVal;

export interface Ctx {
  wb: WbLike;
  sheet: SheetLike;
  row: number;
  col: number;
}

export function isRange(a: Arg): a is RangeVal {
  return typeof a === 'object' && a !== null && (a as RangeVal).k === 'rng';
}
export function isArr(a: Arg): a is ArrVal {
  return typeof a === 'object' && a !== null && (a as ArrVal).k === 'arr';
}

export function getCellValue(sheet: SheetLike, r: number, c: number): Value {
  const cell = sheet.cells.get(cellKey(r, c));
  return cell ? cell.v : null;
}

export function rangeDims(r: RangeVal): { rows: number; cols: number } {
  return { rows: r.r2 - r.r1 + 1, cols: r.c2 - r.c1 + 1 };
}

/** Clamp a range to the sheet's used extent (for iterating whole-column refs). */
export function clampRange(r: RangeVal): RangeVal {
  const r2 = Math.min(r.r2, Math.max(r.r1, r.sheet.maxRow));
  const c2 = Math.min(r.c2, Math.max(r.c1, r.sheet.maxCol));
  return r2 === r.r2 && c2 === r.c2 ? r : { ...r, r2, c2 };
}

/** Visit existing (non-empty) cells of a range. */
export function forEachCell(r: RangeVal, cb: (v: Value, row: number, col: number) => void): void {
  const area = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
  if (area <= r.sheet.cells.size || area <= 64) {
    for (let row = r.r1; row <= r.r2; row++)
      for (let col = r.c1; col <= r.c2; col++) {
        const cell = r.sheet.cells.get(cellKey(row, col));
        if (cell && cell.v !== null) cb(cell.v, row, col);
      }
  } else {
    // clamped dense walk is cheaper than scanning every cell when the range is mostly outside the data
    const cr = clampRange(r);
    const carea = (cr.r2 - cr.r1 + 1) * (cr.c2 - cr.c1 + 1);
    if (carea <= r.sheet.cells.size) {
      for (let row = cr.r1; row <= cr.r2; row++)
        for (let col = cr.c1; col <= cr.c2; col++) {
          const cell = r.sheet.cells.get(cellKey(row, col));
          if (cell && cell.v !== null) cb(cell.v, row, col);
        }
    } else {
      const found: [number, number, Value][] = [];
      for (const [k, cell] of r.sheet.cells) {
        if (cell.v === null) continue;
        const row = Math.floor(k / MAX_COLS);
        const col = k % MAX_COLS;
        if (row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2) found.push([row, col, cell.v]);
      }
      found.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      for (const [row, col, v] of found) cb(v, row, col);
    }
  }
}

/** Dense 2-D values of an arg (clamped to used extent for ranges). */
export function toMatrix(a: Arg): Value[][] {
  if (isArr(a)) return a.rows;
  if (isRange(a)) {
    const cr = clampRange(a);
    const out: Value[][] = [];
    for (let r = cr.r1; r <= cr.r2; r++) {
      const row: Value[] = [];
      for (let c = cr.c1; c <= cr.c2; c++) row.push(getCellValue(cr.sheet, r, c));
      out.push(row);
    }
    return out;
  }
  return [[a]];
}

// ---------------- coercion ----------------

const NUM_TEXT = /^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$/;

export function parseNumberText(s: string): number | null {
  if (NUM_TEXT.test(s)) return parseFloat(s);
  const t = s.trim();
  let m = /^([-+]?)\$?\s*([\d,]+(?:\.\d*)?|\.\d+)\s*(%?)$/.exec(t);
  if (m && /\d/.test(m[2])) {
    const body = m[2];
    if (body.includes(',') && !/^\d{1,3}(,\d{3})+(\.\d*)?$/.test(body)) return null;
    let n = parseFloat(body.replace(/,/g, ''));
    if (m[3]) n /= 100;
    return m[1] === '-' ? -n : n;
  }
  m = /^\(\$?([\d,]+(?:\.\d*)?)\)$/.exec(t);
  if (m) return -parseFloat(m[1].replace(/,/g, ''));
  return null;
}

/** Parse date/time text (m/d/yyyy, yyyy-mm-dd, h:mm[:ss] [AM/PM]) to a serial, or null. */
export function parseDateText(s: string): number | null {
  const t = s.trim();
  let date: number | null = null;
  let rest = t;
  let m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(.*))?$/.exec(t);
  if (m) {
    let y = parseInt(m[3], 10);
    if (m[3].length <= 2) y += y < 30 ? 2000 : 1900;
    const mo = parseInt(m[1], 10);
    const d = parseInt(m[2], 10);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    date = ymdToSerial(y, mo, d);
    rest = m[4] ?? '';
  } else if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s]+(.*))?$/.exec(t))) {
    const mo = parseInt(m[2], 10);
    const d = parseInt(m[3], 10);
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    date = ymdToSerial(parseInt(m[1], 10), mo, d);
    rest = m[4] ?? '';
  } else if ((m = /^(\d{1,2})-([A-Za-z]{3,9})-(\d{2,4})(?:\s+(.*))?$/.exec(t))) {
    const mi = monthIndex(m[2]);
    if (mi < 0) return null;
    let y = parseInt(m[3], 10);
    if (m[3].length <= 2) y += y < 30 ? 2000 : 1900;
    date = ymdToSerial(y, mi + 1, parseInt(m[1], 10));
    rest = m[4] ?? '';
  } else if ((m = /^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})(?:\s+(.*))?$/.exec(t))) {
    const mi = monthIndex(m[1]);
    if (mi < 0) return null;
    date = ymdToSerial(parseInt(m[3], 10), mi + 1, parseInt(m[2], 10));
    rest = m[4] ?? '';
  }
  let time: number | null = null;
  if (rest) {
    const tm = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM|A|P)?$/i.exec(rest.trim());
    if (!tm) return null;
    let h = parseInt(tm[1], 10);
    const mi = parseInt(tm[2], 10);
    const sec = tm[3] ? parseFloat(tm[3]) : 0;
    if (tm[4]) {
      const pm = tm[4].toUpperCase().startsWith('P');
      if (h < 1 || h > 12) return null;
      if (pm && h < 12) h += 12;
      if (!pm && h === 12) h = 0;
    }
    if (h > 23 || mi > 59 || sec >= 60) return null;
    time = (h * 3600 + mi * 60 + sec) / 86400;
  }
  if (date === null && time === null) return null;
  return (date ?? 0) + (time ?? 0);
}

const MONTH_ABBR = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
export function monthIndex(s: string): number {
  return MONTH_ABBR.indexOf(s.slice(0, 3).toLowerCase());
}

export function toNum(v: Value): number | FErr {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof FErr) return v;
  const n = parseNumberText(v);
  if (n !== null) return n;
  const d = parseDateText(v);
  if (d !== null) return d;
  return E.VALUE;
}

export function toStr(v: Value): string | FErr {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numToStr(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v;
}

export function numToStr(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  let s = parseFloat(n.toPrecision(15)).toString();
  if (s.includes('e')) {
    const [m, e] = s.split('e');
    s = m + 'E' + (e[0] === '-' ? '-' : '+') + e.replace(/^[+-]/, '').padStart(2, '0');
  }
  return s;
}

export function toBool(v: Value): boolean | FErr {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v instanceof FErr) return v;
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return E.VALUE;
}

/** Implicit intersection: turn a range/array arg into one scalar value. */
export function scalar(a: Arg, ctx: Ctx): Value {
  if (isRange(a)) {
    if (a.r1 === a.r2 && a.c1 === a.c2) return getCellValue(a.sheet, a.r1, a.c1);
    if (a.c1 === a.c2 && ctx.row >= a.r1 && ctx.row <= a.r2) return getCellValue(a.sheet, ctx.row, a.c1);
    if (a.r1 === a.r2 && ctx.col >= a.c1 && ctx.col <= a.c2) return getCellValue(a.sheet, a.r1, ctx.col);
    return E.VALUE;
  }
  if (isArr(a)) return a.rows.length ? a.rows[0][0] ?? null : null;
  return a;
}

// ---------------- comparison ----------------

function typeRank(v: Value): number {
  if (v === null) return 0;
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  return 3; // boolean
}

/** Excel comparison: returns negative/0/positive. Empty compares as 0 or "". */
export function compareValues(a: Value, b: Value): number {
  if (a === null && b === null) return 0;
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number') return a < (b as number) ? -1 : a > (b as number) ? 1 : 0;
  if (typeof a === 'string') {
    const x = a.toLowerCase();
    const y = (b as string).toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return (a === b ? 0 : a ? 1 : -1);
}

// ---------------- evaluation ----------------

export type FuncImpl = (args: Arg[], ctx: Ctx) => Arg;
export type LazyImpl = (args: Node[], ctx: Ctx, ev: (n: Node) => Arg) => Arg;

export interface FuncDef {
  name: string;
  min: number;
  max: number; // Infinity allowed
  sig: string; // e.g. "number1, [number2], ..."
  desc: string;
  fn?: FuncImpl;
  lazy?: LazyImpl;
  volatile?: boolean;
  noScalar?: boolean; // arguments are passed through as ranges/arrays
}

export type FuncTable = Map<string, FuncDef>;
export const FUNCS: FuncTable = new Map();

function resolveSheet(ctx: Ctx, name?: string): SheetLike | undefined {
  if (name === undefined) return ctx.sheet;
  return ctx.wb.sheetByName(name);
}

export function evalNode(n: Node, ctx: Ctx, depth = 0): Arg {
  if (depth > 1000) return E.VALUE;
  switch (n.k) {
    case 'num':
    case 'str':
    case 'bool':
      return n.v;
    case 'err':
      return n.v;
    case 'missing':
      return null;
    case 'ref': {
      const sh = resolveSheet(ctx, n.sheet);
      if (!sh) return E.REF;
      return getCellValue(sh, n.a.r, n.a.c);
    }
    case 'range': {
      const sh = resolveSheet(ctx, n.sheet);
      if (!sh) return E.REF;
      return {
        k: 'rng',
        sheet: sh,
        r1: Math.min(n.a.r, n.b.r),
        c1: Math.min(n.a.c, n.b.c),
        r2: Math.max(n.a.r, n.b.r),
        c2: Math.max(n.a.c, n.b.c),
      };
    }
    case 'name': {
      const def = ctx.wb.getName(n.name, ctx.sheet);
      if (!def) return E.NAME;
      const ast = ctx.wb.nameAst(def);
      if (!ast) return E.NAME;
      return evalNode(ast, ctx, depth + 1);
    }
    case 'un': {
      const x = evalNode(n.x, ctx, depth + 1);
      if (n.op === '+') return x;
      return mapUnary(x, ctx, (v) => {
        const nn = toNum(v);
        return nn instanceof FErr ? nn : -nn;
      });
    }
    case 'pct': {
      const x = evalNode(n.x, ctx, depth + 1);
      return mapUnary(x, ctx, (v) => {
        const nn = toNum(v);
        return nn instanceof FErr ? nn : nn / 100;
      });
    }
    case 'bin': {
      // walk the left spine iteratively so long chains (1+1+1+...) don't cost depth
      const spine: Extract<Node, { k: 'bin' }>[] = [];
      let cur: Node = n;
      while (cur.k === 'bin') {
        spine.push(cur);
        cur = cur.l;
      }
      let acc = evalNode(cur, ctx, depth + 1);
      for (let i = spine.length - 1; i >= 0; i--) {
        const b = spine[i];
        acc = binary(b.op, acc, evalNode(b.r, ctx, depth + 1), ctx);
      }
      return acc;
    }
    case 'call': {
      const def = FUNCS.get(n.name);
      if (!def) return E.NAME;
      if (n.args.length < def.min || n.args.length > def.max) return E.VALUE;
      if (def.lazy) return def.lazy(n.args, ctx, (x) => evalNode(x, ctx, depth + 1));
      const args = n.args.map((a) => evalNode(a, ctx, depth + 1));
      return def.fn!(args, ctx);
    }
  }
}

function mapUnary(x: Arg, ctx: Ctx, f: (v: Value) => Value): Arg {
  if (isRange(x) || isArr(x)) {
    const m = toMatrix(x);
    if (m.length === 1 && m[0].length === 1) return f(m[0][0]);
    return { k: 'arr', rows: m.map((row) => row.map(f)) };
  }
  return f(x);
}

function binScalar(op: string, a: Value, b: Value): Value {
  if (a instanceof FErr) return a;
  if (b instanceof FErr) return b;
  switch (op) {
    case '+':
    case '-':
    case '*':
    case '/':
    case '^': {
      const x = toNum(a);
      if (x instanceof FErr) return x;
      const y = toNum(b);
      if (y instanceof FErr) return y;
      switch (op) {
        case '+': return fix(x + y);
        case '-': return fix(x - y);
        case '*': return x * y;
        case '/': return y === 0 ? E.DIV0 : x / y;
        default: {
          if (x === 0 && y === 0) return E.NUM;
          if (x === 0 && y < 0) return E.DIV0;
          const r = Math.pow(x, y);
          return Number.isNaN(r) || !isFinite(r) ? E.NUM : r;
        }
      }
    }
    case '&': {
      const x = toStr(a);
      if (x instanceof FErr) return x;
      const y = toStr(b);
      if (y instanceof FErr) return y;
      return x + y;
    }
    default: {
      const c = compareValues(a, b);
      switch (op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '<': return c < 0;
        case '>': return c > 0;
        case '<=': return c <= 0;
        case '>=': return c >= 0;
      }
      return E.VALUE;
    }
  }
}

/** Remove float noise from add/subtract like Excel does near zero (0.1+0.2-0.3). */
function fix(x: number): number {
  if (x === 0 || !isFinite(x)) return x;
  const r = parseFloat(x.toPrecision(15));
  return r;
}

function binary(op: string, l: Arg, r: Arg, ctx: Ctx): Arg {
  const lArr = isRange(l) || isArr(l);
  const rArr = isRange(r) || isArr(r);
  if (!lArr && !rArr) return binScalar(op, l as Value, r as Value);
  const tooBig = (a: Arg) => isRange(a) && (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1) > 4_000_000;
  if ((lArr && tooBig(l)) || (rArr && tooBig(r))) return E.VALUE;
  const lm = lArr ? toMatrixFull(l) : null;
  const rm = rArr ? toMatrixFull(r) : null;
  // 1x1 ranges act as scalars
  const lv = lm && lm.length === 1 && lm[0].length === 1 ? lm[0][0] : null;
  const rv = rm && rm.length === 1 && rm[0].length === 1 ? rm[0][0] : null;
  if (lm && rm) {
    if (lv !== null || (lm.length === 1 && lm[0].length === 1)) return mapUnary(r, ctx, (v) => binScalar(op, lm[0][0], v));
    if (rm.length === 1 && rm[0].length === 1) return mapUnary(l, ctx, (v) => binScalar(op, v, rm[0][0]));
    const rows = Math.max(lm.length, rm.length);
    const cols = Math.max(lm[0].length, rm[0].length);
    const out: Value[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: Value[] = [];
      for (let j = 0; j < cols; j++) {
        const a = lm[lm.length === 1 ? 0 : i]?.[lm[0].length === 1 ? 0 : j];
        const b = rm[rm.length === 1 ? 0 : i]?.[rm[0].length === 1 ? 0 : j];
        row.push(a === undefined || b === undefined ? E.NA : binScalar(op, a, b));
      }
      out.push(row);
    }
    return { k: 'arr', rows: out };
  }
  if (lm) {
    if (lm.length === 1 && lm[0].length === 1) return binScalar(op, lm[0][0], r as Value);
    return { k: 'arr', rows: lm.map((row) => row.map((v) => binScalar(op, v, r as Value))) };
  }
  const rmm = rm!;
  if (rmm.length === 1 && rmm[0].length === 1) return binScalar(op, l as Value, rmm[0][0]);
  return { k: 'arr', rows: rmm.map((row) => row.map((v) => binScalar(op, l as Value, v))) };
}

/** Like toMatrix but never clamps below the requested shape for ranges (needed for elementwise ops). */
function toMatrixFull(a: Arg): Value[][] {
  if (isRange(a)) {
    const rows = Math.min(a.r2 - a.r1 + 1, 1000000);
    const cols = Math.min(a.c2 - a.c1 + 1, 16384);
    const cr = clampRange(a);
    const out: Value[][] = [];
    for (let r = 0; r < rows; r++) {
      const row: Value[] = [];
      for (let c = 0; c < cols; c++) {
        const rr = a.r1 + r;
        const cc = a.c1 + c;
        row.push(rr <= cr.r2 && cc <= cr.c2 ? getCellValue(a.sheet, rr, cc) : null);
      }
      out.push(row);
    }
    return out;
  }
  return toMatrix(a);
}

export function evaluate(ast: Node, ctx: Ctx): Value {
  const res = evalNode(ast, ctx);
  if (isRange(res) || isArr(res)) {
    const v = scalar(res, ctx);
    return v === null ? 0 : v;
  }
  return res === undefined || res === null ? 0 : res;
}

// ---------------- dependency collection ----------------

export interface DepRef {
  sheet?: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function collectRefs(n: Node, out: DepRef[], names: string[], volatile: { v: boolean }): void {
  switch (n.k) {
    case 'ref':
      out.push({ sheet: n.sheet, r1: n.a.r, c1: n.a.c, r2: n.a.r, c2: n.a.c });
      break;
    case 'range':
      out.push({
        sheet: n.sheet,
        r1: Math.min(n.a.r, n.b.r),
        c1: Math.min(n.a.c, n.b.c),
        r2: Math.max(n.a.r, n.b.r),
        c2: Math.max(n.a.c, n.b.c),
      });
      break;
    case 'name':
      names.push(n.name);
      break;
    case 'un':
    case 'pct':
      collectRefs(n.x, out, names, volatile);
      break;
    case 'bin': {
      let cur: Node = n;
      while (cur.k === 'bin') {
        collectRefs(cur.r, out, names, volatile);
        cur = cur.l;
      }
      collectRefs(cur, out, names, volatile);
      break;
    }
    case 'call': {
      const def = FUNCS.get(n.name);
      if (def && def.volatile) volatile.v = true;
      for (const a of n.args) collectRefs(a, out, names, volatile);
      break;
    }
    default:
      break;
  }
}

export { MAX_ROWS };
export { isErr };
