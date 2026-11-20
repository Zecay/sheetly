// Built-in worksheet functions.
import { formatValue, FILL_MARK } from './numfmt';
import {
  Arg, Ctx, FUNCS, FuncDef, RangeVal, clampRange, compareValues, evalNode, forEachCell, getCellValue, isArr, isRange,
  numToStr, parseDateText, scalar, toBool, toMatrix, toNum, toStr, monthIndex, parseNumberText,
} from './eval';
import { E, FErr, Value, isErr, nowSerial, serialToTime, serialToYmd, todaySerial, weekdayOf, ymdToSerial } from './values';
import type { Node } from './parser';
import { MAX_COLS, MAX_ROWS, keyCol, keyRow } from './addr';

type Opts = Partial<Pick<FuncDef, 'volatile' | 'noScalar'>>;

function def(name: string, min: number, max: number, sig: string, desc: string, fn: (a: Arg[], c: Ctx) => Arg, opts: Opts = {}) {
  FUNCS.set(name, { name, min, max, sig, desc, fn, ...opts });
}
function lazy(name: string, min: number, max: number, sig: string, desc: string, l: (a: Node[], c: Ctx, ev: (n: Node) => Arg) => Arg) {
  FUNCS.set(name, { name, min, max, sig, desc, lazy: l });
}
const INF = Infinity;

// ---------- helpers ----------

/** Scalar (implicit-intersection) value of an argument. */
const sv = (a: Arg, c: Ctx): Value => scalar(a, c);

function num(a: Arg, c: Ctx): number | FErr {
  return toNum(sv(a, c));
}
function str(a: Arg, c: Ctx): string | FErr {
  return toStr(sv(a, c));
}

/** Numbers from args following SUM/AVERAGE rules. */
function collectNums(args: Arg[], c: Ctx): number[] | FErr {
  const out: number[] = [];
  let err: FErr | null = null;
  for (const a of args) {
    if (isRange(a) || isArr(a)) {
      if (isArr(a)) {
        for (const row of a.rows)
          for (const v of row) {
            if (v instanceof FErr) return v;
            if (typeof v === 'number') out.push(v);
          }
      } else
        forEachCell(a, (v) => {
          if (err) return;
          if (v instanceof FErr) err = v;
          else if (typeof v === 'number') out.push(v);
        });
      if (err) return err;
    } else {
      if (a instanceof FErr) return a;
      if (a === null) {
        out.push(0);
        continue;
      }
      const n = toNum(a);
      if (n instanceof FErr) return n;
      out.push(n);
    }
  }
  return out;
}

function wildcardRe(p: string): RegExp {
  let re = '^';
  for (let i = 0; i < p.length; i++) {
    const ch = p[i];
    if (ch === '~' && i + 1 < p.length) {
      re += p[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(re + '$', 'i');
}

export function makeCriteria(crit: Value): (v: Value) => boolean {
  if (typeof crit === 'number') return (v) => typeof v === 'number' && v === crit;
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => v === null || v === 0 ? v === null : false;
  if (crit instanceof FErr) return (v) => v instanceof FErr && v.code === crit.code;
  let op = '=';
  let rest = crit;
  const m = /^(<=|>=|<>|<|>|=)/.exec(crit);
  if (m) {
    op = m[1];
    rest = crit.slice(op.length);
  }
  const asNum = parseNumberText(rest);
  const asDate = asNum === null ? parseDateText(rest) : null;
  const target = asNum ?? asDate;
  if (target !== null && rest.trim() !== '') {
    return (v) => {
      if (typeof v !== 'number') {
        if (op === '<>') return true;
        return false;
      }
      switch (op) {
        case '=': return v === target;
        case '<>': return v !== target;
        case '<': return v < target;
        case '>': return v > target;
        case '<=': return v <= target;
        default: return v >= target;
      }
    };
  }
  const upper = rest.toUpperCase();
  if (op === '=' || op === '<>') {
    if (rest === '') {
      return op === '='
        ? (v) => v === null || v === ''
        : (v) => !(v === null || v === '');
    }
    const hasWild = /[*?~]/.test(rest);
    const re = hasWild ? wildcardRe(rest) : null;
    const test = (v: Value): boolean => {
      if (typeof v === 'string') return re ? re.test(v) : v.toUpperCase() === upper;
      if (typeof v === 'boolean') return (v ? 'TRUE' : 'FALSE') === upper;
      return false;
    };
    return op === '=' ? test : (v) => !test(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const c2 = v.toUpperCase() < upper ? -1 : v.toUpperCase() > upper ? 1 : 0;
    switch (op) {
      case '<': return c2 < 0;
      case '>': return c2 > 0;
      case '<=': return c2 <= 0;
      default: return c2 >= 0;
    }
  };
}

function needRange(a: Arg): RangeVal | FErr {
  if (isRange(a)) return a;
  return E.VALUE;
}

/** Iterate cells matched by (range, criteria) pairs. cb receives the offset (row,col) inside the range. */
function matchIf(pairs: [Arg, Arg][], c: Ctx, cb: (dr: number, dc: number) => void): FErr | null {
  const preds: { r: RangeVal; f: (v: Value) => boolean }[] = [];
  let rows = -1;
  let cols = -1;
  for (const [ra, ca] of pairs) {
    const r = needRange(ra);
    if (r instanceof FErr) return r;
    const crit = sv(ca, c);
    const h = r.r2 - r.r1;
    const w = r.c2 - r.c1;
    if (rows >= 0 && (rows !== h || cols !== w)) return E.VALUE;
    rows = h;
    cols = w;
    preds.push({ r, f: makeCriteria(crit) });
  }
  if (!preds.length) return null;
  // walk the first range (existing cells + empties for blank criteria)
  const first = preds[0];
  const needsEmpties = pairs.some(([, ca]) => {
    const cv = sv(ca, c);
    return cv === null || cv === '' || (typeof cv === 'string' && /^(=|<>)$/.test(cv)) || (typeof cv === 'string' && cv.startsWith('<>'));
  });
  const visit = (dr: number, dc: number) => {
    for (const p of preds) {
      const v = getCellValue(p.r.sheet, p.r.r1 + dr, p.r.c1 + dc);
      if (!p.f(v)) return;
    }
    cb(dr, dc);
  };
  const fr = first.r;
  const area = (fr.r2 - fr.r1 + 1) * (fr.c2 - fr.c1 + 1);
  if (!needsEmpties && area > 64 && area > fr.sheet.cells.size * 2) {
    // sparse: only populated cells of the first range can match a non-blank criterion
    const hits: [number, number][] = [];
    for (const k of fr.sheet.cells.keys()) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (r >= fr.r1 && r <= fr.r2 && c >= fr.c1 && c <= fr.c2) hits.push([r - fr.r1, c - fr.c1]);
    }
    hits.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    for (const [dr, dc] of hits) visit(dr, dc);
    return null;
  }
  // dense walk limited to the populated extent of the first range
  let maxR = fr.r1;
  let maxC = fr.c1;
  if (area > fr.sheet.cells.size) {
    for (const k of fr.sheet.cells.keys()) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (r >= fr.r1 && r <= fr.r2 && c >= fr.c1 && c <= fr.c2) {
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
      }
    }
  } else {
    maxR = fr.r2;
    maxC = fr.c2;
  }
  if ((maxR - fr.r1 + 1) * (maxC - fr.c1 + 1) > 5_000_000) return E.VALUE;
  for (let dr = 0; dr <= maxR - fr.r1; dr++) for (let dc = 0; dc <= maxC - fr.c1; dc++) visit(dr, dc);
  return null;
}

function fl(v: number): number {
  return Math.floor(v);
}

// ---------- math ----------

def('SUM', 1, INF, 'number1, [number2], ...', 'Adds all the numbers in a range of cells.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  let s = 0;
  for (const x of n) s += x;
  return parseFloat(s.toPrecision(15)) === s ? s : parseFloat(s.toPrecision(15));
});
def('PRODUCT', 1, INF, 'number1, [number2], ...', 'Multiplies all the numbers given as arguments.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  return n.reduce((p, x) => p * x, 1);
});
def('SUMSQ', 1, INF, 'number1, [number2], ...', 'Returns the sum of the squares of the arguments.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  return n.reduce((p, x) => p + x * x, 0);
});
def('AVERAGE', 1, INF, 'number1, [number2], ...', 'Returns the average (arithmetic mean) of the arguments.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  if (!n.length) return E.DIV0;
  return n.reduce((p, x) => p + x, 0) / n.length;
});
def('MIN', 1, INF, 'number1, [number2], ...', 'Returns the smallest number in a set of values.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  return n.length ? n.reduce((p, x) => (x < p ? x : p), Infinity) : 0;
});
def('MAX', 1, INF, 'number1, [number2], ...', 'Returns the largest value in a set of values.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  return n.length ? n.reduce((p, x) => (x > p ? x : p), -Infinity) : 0;
});
def('MEDIAN', 1, INF, 'number1, [number2], ...', 'Returns the median of the given numbers.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  if (!n.length) return E.NUM;
  n.sort((x, y) => x - y);
  const m = n.length >> 1;
  return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
});
def('MODE', 1, INF, 'number1, [number2], ...', 'Returns the most frequent value in a data set.', (a, c) => {
  const n = collectNums(a, c);
  if (n instanceof FErr) return n;
  const cnt = new Map<number, number>();
  let best = 0;
  let bv: number | null = null;
  for (const x of n) {
    const k = (cnt.get(x) ?? 0) + 1;
    cnt.set(x, k);
    if (k > best) {
      best = k;
      bv = x;
    }
  }
  return best < 2 || bv === null ? E.NA : bv;
});
const variance = (n: number[], sample: boolean): number | FErr => {
  const d = sample ? n.length - 1 : n.length;
  if (d <= 0) return E.DIV0;
  const m = n.reduce((p, x) => p + x, 0) / n.length;
  return n.reduce((p, x) => p + (x - m) * (x - m), 0) / d;
};
for (const [name, sample, root] of [
  ['VAR', true, false], ['VAR.S', true, false], ['VAR.P', false, false], ['VARP', false, false],
  ['STDEV', true, true], ['STDEV.S', true, true], ['STDEV.P', false, true], ['STDEVP', false, true],
] as [string, boolean, boolean][]) {
  def(name, 1, INF, 'number1, [number2], ...', root ? 'Estimates standard deviation.' : 'Estimates variance.', (a, c) => {
    const n = collectNums(a, c);
    if (n instanceof FErr) return n;
    const v = variance(n, sample);
    return v instanceof FErr ? v : root ? Math.sqrt(v) : v;
  });
}
def('COUNT', 1, INF, 'value1, [value2], ...', 'Counts how many numbers are in the list of arguments.', (a, c) => {
  let n = 0;
  for (const x of a) {
    if (isRange(x)) forEachCell(x, (v) => { if (typeof v === 'number') n++; });
    else if (isArr(x)) { for (const r of x.rows) for (const v of r) if (typeof v === 'number') n++; }
    else if (typeof x === 'number' || (typeof x === 'string' && toNum(x) !== E.VALUE && !isErr(toNum(x)))) n++;
    else if (typeof x === 'boolean') n++;
  }
  return n;
});
def('COUNTA', 1, INF, 'value1, [value2], ...', 'Counts how many values are in the list of arguments.', (a) => {
  let n = 0;
  for (const x of a) {
    if (isRange(x)) forEachCell(x, () => { n++; });
    else if (isArr(x)) { for (const r of x.rows) for (const v of r) if (v !== null) n++; }
    else n++;
  }
  return n;
});
def('COUNTBLANK', 1, 1, 'range', 'Counts the empty cells in a specified range.', (a) => {
  const r = needRange(a[0]);
  if (r instanceof FErr) return r;
  const total = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
  let filled = 0;
  forEachCell(r, (v) => { if (v !== '') filled++; });
  return total - filled;
});
def('LARGE', 2, 2, 'array, k', 'Returns the k-th largest value in a data set.', (a, c) => {
  const n = collectNums([a[0]], c);
  if (n instanceof FErr) return n;
  const k = num(a[1], c);
  if (k instanceof FErr) return k;
  if (k < 1 || k > n.length) return E.NUM;
  n.sort((x, y) => y - x);
  return n[Math.ceil(k) - 1];
});
def('SMALL', 2, 2, 'array, k', 'Returns the k-th smallest value in a data set.', (a, c) => {
  const n = collectNums([a[0]], c);
  if (n instanceof FErr) return n;
  const k = num(a[1], c);
  if (k instanceof FErr) return k;
  if (k < 1 || k > n.length) return E.NUM;
  n.sort((x, y) => x - y);
  return n[Math.ceil(k) - 1];
});
def('RANK', 2, 3, 'number, ref, [order]', 'Returns the rank of a number in a list of numbers.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  const n = collectNums([a[1]], c);
  if (n instanceof FErr) return n;
  const asc = a[2] !== undefined && num(a[2], c) !== 0;
  if (!n.includes(x)) return E.NA;
  return 1 + n.filter((v) => (asc ? v < x : v > x)).length;
});

def('ABS', 1, 1, 'number', 'Returns the absolute value of a number.', (a, c) => {
  const x = num(a[0], c);
  return x instanceof FErr ? x : Math.abs(x);
});
def('SIGN', 1, 1, 'number', 'Returns the sign of a number.', (a, c) => {
  const x = num(a[0], c);
  return x instanceof FErr ? x : Math.sign(x);
});
def('INT', 1, 1, 'number', 'Rounds a number down to the nearest integer.', (a, c) => {
  const x = num(a[0], c);
  return x instanceof FErr ? x : Math.floor(x);
});
def('TRUNC', 1, 2, 'number, [num_digits]', 'Truncates a number to an integer.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  const d = a[1] === undefined ? 0 : num(a[1], c);
  if (d instanceof FErr) return d;
  const p = Math.pow(10, Math.trunc(d));
  return Math.trunc(x * p) / p;
});
def('MOD', 2, 2, 'number, divisor', 'Returns the remainder after a number is divided by a divisor.', (a, c) => {
  const x = num(a[0], c);
  const y = num(a[1], c);
  if (x instanceof FErr) return x;
  if (y instanceof FErr) return y;
  if (y === 0) return E.DIV0;
  return x - y * Math.floor(x / y);
});
def('POWER', 2, 2, 'number, power', 'Returns the result of a number raised to a power.', (a, c) => {
  const x = num(a[0], c);
  const y = num(a[1], c);
  if (x instanceof FErr) return x;
  if (y instanceof FErr) return y;
  if (x === 0 && y === 0) return E.NUM;
  if (x === 0 && y < 0) return E.DIV0;
  const r = Math.pow(x, y);
  return isFinite(r) ? r : E.NUM;
});
def('SQRT', 1, 1, 'number', 'Returns a positive square root.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  return x < 0 ? E.NUM : Math.sqrt(x);
});
def('EXP', 1, 1, 'number', 'Returns e raised to the power of a given number.', (a, c) => {
  const x = num(a[0], c);
  return x instanceof FErr ? x : Math.exp(x);
});
def('LN', 1, 1, 'number', 'Returns the natural logarithm of a number.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  return x <= 0 ? E.NUM : Math.log(x);
});
def('LOG10', 1, 1, 'number', 'Returns the base-10 logarithm of a number.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  return x <= 0 ? E.NUM : Math.log10(x);
});
def('LOG', 1, 2, 'number, [base]', 'Returns the logarithm of a number to a specified base.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  const b = a[1] === undefined ? 10 : num(a[1], c);
  if (b instanceof FErr) return b;
  if (x <= 0 || b <= 0 || b === 1) return E.NUM;
  return Math.log(x) / Math.log(b);
});
def('PI', 0, 0, '', 'Returns the value of pi.', () => Math.PI);
for (const [n, f] of [['SIN', Math.sin], ['COS', Math.cos], ['TAN', Math.tan], ['ATAN', Math.atan], ['DEGREES', (x: number) => (x * 180) / Math.PI], ['RADIANS', (x: number) => (x * Math.PI) / 180]] as [string, (x: number) => number][]) {
  def(n, 1, 1, 'number', 'Trigonometric function.', (a, c) => {
    const x = num(a[0], c);
    return x instanceof FErr ? x : f(x);
  });
}

function roundTo(x: number, d: number, mode: 'near' | 'up' | 'down'): number {
  const p = Math.pow(10, d);
  const sgn = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  let y = ax * p;
  // cleanup float noise like 1.005*100 = 100.49999999999999
  const yc = parseFloat(y.toPrecision(15));
  let r: number;
  if (mode === 'near') r = Math.floor(yc + 0.5);
  else if (mode === 'up') r = Math.ceil(yc);
  else r = Math.floor(yc);
  return (sgn * r) / p;
}
for (const [n, mode, desc] of [
  ['ROUND', 'near', 'Rounds a number to a specified number of digits.'],
  ['ROUNDUP', 'up', 'Rounds a number up, away from zero.'],
  ['ROUNDDOWN', 'down', 'Rounds a number down, toward zero.'],
] as ['ROUND' | 'ROUNDUP' | 'ROUNDDOWN', 'near' | 'up' | 'down', string][]) {
  def(n, 2, 2, 'number, num_digits', desc, (a, c) => {
    const x = num(a[0], c);
    const d = num(a[1], c);
    if (x instanceof FErr) return x;
    if (d instanceof FErr) return d;
    return roundTo(x, Math.trunc(d), mode);
  });
}
def('CEILING', 1, 2, 'number, [significance]', 'Rounds a number up to the nearest multiple of significance.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  const s = a[1] === undefined ? 1 : num(a[1], c);
  if (s instanceof FErr) return s;
  if (s === 0) return 0;
  return Math.ceil(x / s - 1e-12) * s;
});
def('FLOOR', 1, 2, 'number, [significance]', 'Rounds a number down to the nearest multiple of significance.', (a, c) => {
  const x = num(a[0], c);
  if (x instanceof FErr) return x;
  const s = a[1] === undefined ? 1 : num(a[1], c);
  if (s instanceof FErr) return s;
  if (s === 0) return 0;
  return Math.floor(x / s + 1e-12) * s;
});
def('RAND', 0, 0, '', 'Returns a random number between 0 and 1.', () => Math.random(), { volatile: true });
def('RANDBETWEEN', 2, 2, 'bottom, top', 'Returns a random integer between the numbers you specify.', (a, c) => {
  const lo = num(a[0], c);
  const hi = num(a[1], c);
  if (lo instanceof FErr) return lo;
  if (hi instanceof FErr) return hi;
  const l = Math.ceil(lo);
  const h = Math.floor(hi);
  if (l > h) return E.NUM;
  return l + Math.floor(Math.random() * (h - l + 1));
}, { volatile: true });
def('SUMPRODUCT', 1, INF, 'array1, [array2], ...', 'Returns the sum of the products of corresponding ranges or arrays.', (a) => {
  const ms = a.map((x) => {
    if (isRange(x)) {
      const rows = x.r2 - x.r1 + 1;
      const cols = x.c2 - x.c1 + 1;
      const cr = clampRange(x);
      const m: Value[][] = [];
      for (let r = 0; r < rows; r++) {
        const row: Value[] = [];
        for (let cc = 0; cc < cols; cc++) row.push(x.r1 + r <= cr.r2 && x.c1 + cc <= cr.c2 ? getCellValue(x.sheet, x.r1 + r, x.c1 + cc) : null);
        m.push(row);
      }
      return m;
    }
    return toMatrix(x);
  });
  const rows = ms[0].length;
  const cols = ms[0][0].length;
  for (const m of ms) if (m.length !== rows || m[0].length !== cols) return E.VALUE;
  let sum = 0;
  for (let r = 0; r < rows; r++)
    for (let cc = 0; cc < cols; cc++) {
      let p = 1;
      for (const m of ms) {
        const v = m[r][cc];
        if (v instanceof FErr) return v;
        p *= typeof v === 'number' ? v : 0;
      }
      sum += p;
    }
  return sum;
});

// ---------- conditional aggregates ----------

def('COUNTIF', 2, 2, 'range, criteria', 'Counts the number of cells that meet a criterion.', (a, c) => {
  let n = 0;
  const e = matchIf([[a[0], a[1]]], c, () => { n++; });
  return e ?? n;
});
def('COUNTIFS', 2, INF, 'criteria_range1, criteria1, ...', 'Counts the cells that meet multiple criteria.', (a, c) => {
  if (a.length % 2) return E.VALUE;
  const pairs: [Arg, Arg][] = [];
  for (let i = 0; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]);
  let n = 0;
  const e = matchIf(pairs, c, () => { n++; });
  return e ?? n;
});
function sumLike(name: string, withIfs: boolean, kind: 'sum' | 'avg' | 'min' | 'max') {
  const sig = withIfs ? `${kind === 'sum' ? 'sum_range' : 'average_range'}, criteria_range1, criteria1, ...` : 'range, criteria, [sum_range]';
  def(name, withIfs ? 3 : 2, withIfs ? INF : 3, sig, 'Conditionally aggregates the cells that meet the criteria.', (a, c) => {
    let target: Arg;
    const pairs: [Arg, Arg][] = [];
    if (withIfs) {
      if ((a.length - 1) % 2) return E.VALUE;
      target = a[0];
      for (let i = 1; i < a.length; i += 2) pairs.push([a[i], a[i + 1]]);
    } else {
      target = a[2] ?? a[0];
      pairs.push([a[0], a[1]]);
    }
    const tr = needRange(target);
    if (tr instanceof FErr) return tr;
    const first = pairs[0][0];
    if (isRange(first) && (first.r2 - first.r1 !== tr.r2 - tr.r1 || first.c2 - first.c1 !== tr.c2 - tr.c1)) {
      if (!withIfs) {
        // SUMIF allows different-sized sum_range: it is resized to the criteria range
        tr.r2 = tr.r1 + (first.r2 - first.r1);
        tr.c2 = tr.c1 + (first.c2 - first.c1);
      } else return E.VALUE;
    }
    let sum = 0;
    let cnt = 0;
    let mn = Infinity;
    let mx = -Infinity;
    let err: FErr | null = null;
    const e = matchIf(pairs, c, (dr, dc) => {
      const v = getCellValue(tr.sheet, tr.r1 + dr, tr.c1 + dc);
      if (v instanceof FErr) {
        err = v;
        return;
      }
      if (typeof v !== 'number') return;
      sum += v;
      cnt++;
      if (v < mn) mn = v;
      if (v > mx) mx = v;
    });
    if (e) return e;
    if (err) return err;
    switch (kind) {
      case 'sum': return sum;
      case 'avg': return cnt ? sum / cnt : E.DIV0;
      case 'min': return cnt ? mn : 0;
      default: return cnt ? mx : 0;
    }
  });
}
sumLike('SUMIF', false, 'sum');
sumLike('SUMIFS', true, 'sum');
sumLike('AVERAGEIF', false, 'avg');
sumLike('AVERAGEIFS', true, 'avg');
sumLike('MINIFS', true, 'min');
sumLike('MAXIFS', true, 'max');

// ---------- logical ----------

lazy('IF', 2, 3, 'logical_test, value_if_true, [value_if_false]', 'Checks a condition and returns one value if TRUE, another if FALSE.', (a, c, ev) => {
  const t = sv(ev(a[0]), c);
  const b = toBool(t);
  if (b instanceof FErr) return b;
  if (b) return ev(a[1]);
  return a[2] === undefined ? false : ev(a[2]);
});
lazy('IFS', 2, INF, 'logical_test1, value_if_true1, ...', 'Checks multiple conditions and returns the first match.', (a, c, ev) => {
  if (a.length % 2) return E.VALUE;
  for (let i = 0; i < a.length; i += 2) {
    const b = toBool(sv(ev(a[i]), c));
    if (b instanceof FErr) return b;
    if (b) return ev(a[i + 1]);
  }
  return E.NA;
});
lazy('IFERROR', 2, 2, 'value, value_if_error', 'Returns a value you specify if a formula evaluates to an error.', (a, c, ev) => {
  const v = ev(a[0]);
  const s = isRange(v) || isArr(v) ? sv(v, c) : v;
  return s instanceof FErr ? ev(a[1]) : v;
});
lazy('IFNA', 2, 2, 'value, value_if_na', 'Returns a value you specify if the formula returns #N/A.', (a, c, ev) => {
  const v = ev(a[0]);
  const s = isRange(v) || isArr(v) ? sv(v, c) : v;
  return s instanceof FErr && s.code === '#N/A' ? ev(a[1]) : v;
});
lazy('CHOOSE', 2, INF, 'index_num, value1, [value2], ...', 'Chooses a value from a list of values.', (a, c, ev) => {
  const i = toNum(sv(ev(a[0]), c));
  if (i instanceof FErr) return i;
  const k = Math.floor(i);
  if (k < 1 || k >= a.length) return E.VALUE;
  return ev(a[k]);
});
lazy('SWITCH', 3, INF, 'expression, value1, result1, [default_or_value2, result2], ...', 'Evaluates an expression against a list of values and returns the matching result.', (a, c, ev) => {
  const x = sv(ev(a[0]), c);
  if (x instanceof FErr) return x;
  let i = 1;
  for (; i + 1 < a.length; i += 2) {
    const v = sv(ev(a[i]), c);
    if (compareValues(x, v) === 0) return ev(a[i + 1]);
  }
  return i < a.length ? ev(a[i]) : E.NA;
});
function boolArgs(a: Arg[], c: Ctx): boolean[] | FErr {
  const out: boolean[] = [];
  for (const x of a) {
    if (isRange(x) || isArr(x)) {
      let err: FErr | null = null;
      const f = (v: Value) => {
        if (v instanceof FErr) err = err ?? v;
        else if (typeof v === 'boolean') out.push(v);
        else if (typeof v === 'number') out.push(v !== 0);
      };
      if (isRange(x)) forEachCell(x, f);
      else for (const r of x.rows) for (const v of r) f(v);
      if (err) return err;
    } else {
      const b = toBool(x);
      if (b instanceof FErr) return b;
      out.push(b);
    }
  }
  return out;
}
def('AND', 1, INF, 'logical1, [logical2], ...', 'Returns TRUE if all arguments are TRUE.', (a, c) => {
  const b = boolArgs(a, c);
  if (b instanceof FErr) return b;
  return b.length ? b.every(Boolean) : E.VALUE;
});
def('OR', 1, INF, 'logical1, [logical2], ...', 'Returns TRUE if any argument is TRUE.', (a, c) => {
  const b = boolArgs(a, c);
  if (b instanceof FErr) return b;
  return b.length ? b.some(Boolean) : E.VALUE;
});
def('XOR', 1, INF, 'logical1, [logical2], ...', 'Returns a logical exclusive OR of all arguments.', (a, c) => {
  const b = boolArgs(a, c);
  if (b instanceof FErr) return b;
  return b.filter(Boolean).length % 2 === 1;
});
def('NOT', 1, 1, 'logical', 'Reverses the logic of its argument.', (a, c) => {
  const b = toBool(sv(a[0], c));
  return b instanceof FErr ? b : !b;
});
def('TRUE', 0, 0, '', 'Returns the logical value TRUE.', () => true);
def('FALSE', 0, 0, '', 'Returns the logical value FALSE.', () => false);

// ---------- info ----------

const isE = (v: Value) => v instanceof FErr;
def('ISNUMBER', 1, 1, 'value', 'Returns TRUE if the value is a number.', (a, c) => typeof sv(a[0], c) === 'number');
def('ISTEXT', 1, 1, 'value', 'Returns TRUE if the value is text.', (a, c) => typeof sv(a[0], c) === 'string');
def('ISNONTEXT', 1, 1, 'value', 'Returns TRUE if the value is not text.', (a, c) => typeof sv(a[0], c) !== 'string');
def('ISLOGICAL', 1, 1, 'value', 'Returns TRUE if the value is a logical value.', (a, c) => typeof sv(a[0], c) === 'boolean');
def('ISBLANK', 1, 1, 'value', 'Returns TRUE if the value is blank.', (a, c) => sv(a[0], c) === null);
def('ISERROR', 1, 1, 'value', 'Returns TRUE if the value is any error value.', (a, c) => isE(sv(a[0], c)));
def('ISERR', 1, 1, 'value', 'Returns TRUE if the value is an error other than #N/A.', (a, c) => {
  const v = sv(a[0], c);
  return v instanceof FErr && v.code !== '#N/A';
});
def('ISNA', 1, 1, 'value', 'Returns TRUE if the value is #N/A.', (a, c) => {
  const v = sv(a[0], c);
  return v instanceof FErr && v.code === '#N/A';
});
def('ISEVEN', 1, 1, 'number', 'Returns TRUE if the number is even.', (a, c) => {
  const x = num(a[0], c);
  return x instanceof FErr ? x : Math.floor(Math.abs(x)) % 2 === 0;
});
def('ISODD', 1, 1, 'number', 'Returns TRUE if the number is odd.', (a, c) => {
  const x = num(a[0], c);
  return x instanceof FErr ? x : Math.floor(Math.abs(x)) % 2 === 1;
});
def('NA', 0, 0, '', 'Returns the error value #N/A.', () => E.NA);
def('N', 1, 1, 'value', 'Returns a value converted to a number.', (a, c) => {
  const v = sv(a[0], c);
  return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : v instanceof FErr ? v : 0;
});
def('T', 1, 1, 'value', 'Converts its arguments to text.', (a, c) => {
  const v = sv(a[0], c);
  return typeof v === 'string' ? v : v instanceof FErr ? v : '';
});

// ---------- text ----------

def('CONCAT', 1, INF, 'text1, [text2], ...', 'Combines the text from multiple ranges and/or strings.', (a, c) => {
  let out = '';
  for (const x of a) {
    if (isRange(x) || isArr(x)) {
      const m = toMatrix(x);
      for (const row of m)
        for (const v of row) {
          const s = toStr(v);
          if (s instanceof FErr) return s;
          out += s;
        }
    } else {
      const s = toStr(x);
      if (s instanceof FErr) return s;
      out += s;
    }
  }
  return out;
});
FUNCS.set('CONCATENATE', { ...FUNCS.get('CONCAT')!, name: 'CONCATENATE', desc: 'Joins several text strings into one text string.' });
def('TEXTJOIN', 3, INF, 'delimiter, ignore_empty, text1, [text2], ...', 'Combines text from multiple ranges using a delimiter.', (a, c) => {
  const d = str(a[0], c);
  if (d instanceof FErr) return d;
  const ig = toBool(sv(a[1], c));
  if (ig instanceof FErr) return ig;
  const parts: string[] = [];
  for (const x of a.slice(2)) {
    const m = toMatrix(x);
    for (const row of m)
      for (const v of row) {
        if (v instanceof FErr) return v;
        const s = toStr(v) as string;
        if (ig && s === '') continue;
        parts.push(s);
      }
  }
  return parts.join(d);
});
def('LEFT', 1, 2, 'text, [num_chars]', 'Returns the leftmost characters from a text value.', (a, c) => {
  const s = str(a[0], c);
  if (s instanceof FErr) return s;
  const n = a[1] === undefined ? 1 : num(a[1], c);
  if (n instanceof FErr) return n;
  if (n < 0) return E.VALUE;
  return s.slice(0, Math.floor(n));
});
def('RIGHT', 1, 2, 'text, [num_chars]', 'Returns the rightmost characters from a text value.', (a, c) => {
  const s = str(a[0], c);
  if (s instanceof FErr) return s;
  const n = a[1] === undefined ? 1 : num(a[1], c);
  if (n instanceof FErr) return n;
  if (n < 0) return E.VALUE;
  return n === 0 ? '' : s.slice(-Math.floor(n));
});
def('MID', 3, 3, 'text, start_num, num_chars', 'Returns a specific number of characters from a text string.', (a, c) => {
  const s = str(a[0], c);
  const st = num(a[1], c);
  const n = num(a[2], c);
  if (s instanceof FErr) return s;
  if (st instanceof FErr) return st;
  if (n instanceof FErr) return n;
  if (st < 1 || n < 0) return E.VALUE;
  return s.substr(Math.floor(st) - 1, Math.floor(n));
});
def('LEN', 1, 1, 'text', 'Returns the number of characters in a text string.', (a, c) => {
  const s = str(a[0], c);
  return s instanceof FErr ? s : s.length;
});
def('UPPER', 1, 1, 'text', 'Converts text to uppercase.', (a, c) => {
  const s = str(a[0], c);
  return s instanceof FErr ? s : s.toUpperCase();
});
def('LOWER', 1, 1, 'text', 'Converts text to lowercase.', (a, c) => {
  const s = str(a[0], c);
  return s instanceof FErr ? s : s.toLowerCase();
});
def('PROPER', 1, 1, 'text', 'Capitalizes the first letter in each word of a text value.', (a, c) => {
  const s = str(a[0], c);
  return s instanceof FErr ? s : s.toLowerCase().replace(/(^|[^A-Za-zÀ-￿])([A-Za-zÀ-￿])/g, (_m, p, l) => p + l.toUpperCase());
});
def('TRIM', 1, 1, 'text', 'Removes extra spaces from text.', (a, c) => {
  const s = str(a[0], c);
  return s instanceof FErr ? s : s.replace(/ +/g, ' ').replace(/^ | $/g, '');
});
def('CLEAN', 1, 1, 'text', 'Removes all nonprintable characters from text.', (a, c) => {
  const s = str(a[0], c);
  return s instanceof FErr ? s : s.replace(/[\x00-\x1f]/g, '');
});
def('EXACT', 2, 2, 'text1, text2', 'Checks whether two text strings are exactly the same.', (a, c) => {
  const x = str(a[0], c);
  const y = str(a[1], c);
  if (x instanceof FErr) return x;
  if (y instanceof FErr) return y;
  return x === y;
});
def('REPT', 2, 2, 'text, number_times', 'Repeats text a given number of times.', (a, c) => {
  const s = str(a[0], c);
  const n = num(a[1], c);
  if (s instanceof FErr) return s;
  if (n instanceof FErr) return n;
  if (n < 0 || s.length * n > 32767) return E.VALUE;
  return s.repeat(Math.floor(n));
});
def('CHAR', 1, 1, 'number', 'Returns the character specified by the code number.', (a, c) => {
  const n = num(a[0], c);
  if (n instanceof FErr) return n;
  return n < 1 || n > 255 ? E.VALUE : String.fromCharCode(Math.floor(n));
});
def('CODE', 1, 1, 'text', 'Returns a numeric code for the first character in a text string.', (a, c) => {
  const s = str(a[0], c);
  if (s instanceof FErr) return s;
  return s.length ? s.charCodeAt(0) : E.VALUE;
});
def('VALUE', 1, 1, 'text', 'Converts a text argument to a number.', (a, c) => {
  const v = sv(a[0], c);
  if (typeof v === 'number') return v;
  const n = toNum(v);
  return n;
});
def('FIND', 2, 3, 'find_text, within_text, [start_num]', 'Finds one text value within another (case-sensitive).', (a, c) => {
  const f = str(a[0], c);
  const t = str(a[1], c);
  if (f instanceof FErr) return f;
  if (t instanceof FErr) return t;
  const st = a[2] === undefined ? 1 : num(a[2], c);
  if (st instanceof FErr) return st;
  if (st < 1 || st > t.length + 1) return E.VALUE;
  const i = t.indexOf(f, Math.floor(st) - 1);
  return i < 0 ? E.VALUE : i + 1;
});
def('SEARCH', 2, 3, 'find_text, within_text, [start_num]', 'Finds one text value within another (not case-sensitive).', (a, c) => {
  const f = str(a[0], c);
  const t = str(a[1], c);
  if (f instanceof FErr) return f;
  if (t instanceof FErr) return t;
  const st = a[2] === undefined ? 1 : num(a[2], c);
  if (st instanceof FErr) return st;
  if (st < 1 || st > t.length + 1) return E.VALUE;
  const re = new RegExp(wildcardRe(f).source.slice(1, -1), 'i');
  const m = re.exec(t.slice(Math.floor(st) - 1));
  return m ? m.index + Math.floor(st) : E.VALUE;
});
def('SUBSTITUTE', 3, 4, 'text, old_text, new_text, [instance_num]', 'Substitutes new text for old text in a text string.', (a, c) => {
  const t = str(a[0], c);
  const o = str(a[1], c);
  const n = str(a[2], c);
  if (t instanceof FErr) return t;
  if (o instanceof FErr) return o;
  if (n instanceof FErr) return n;
  if (o === '') return t;
  if (a[3] === undefined) return t.split(o).join(n);
  const inst = num(a[3], c);
  if (inst instanceof FErr) return inst;
  let idx = -1;
  for (let i = 0; i < inst; i++) {
    idx = t.indexOf(o, idx + 1);
    if (idx < 0) return t;
  }
  return t.slice(0, idx) + n + t.slice(idx + o.length);
});
def('REPLACE', 4, 4, 'old_text, start_num, num_chars, new_text', 'Replaces characters within text.', (a, c) => {
  const t = str(a[0], c);
  const s = num(a[1], c);
  const l = num(a[2], c);
  const n = str(a[3], c);
  if (t instanceof FErr) return t;
  if (s instanceof FErr) return s;
  if (l instanceof FErr) return l;
  if (n instanceof FErr) return n;
  if (s < 1 || l < 0) return E.VALUE;
  return t.slice(0, s - 1) + n + t.slice(s - 1 + l);
});
def('TEXT', 2, 2, 'value, format_text', 'Formats a number and converts it to text.', (a, c) => {
  const v = sv(a[0], c);
  const f = str(a[1], c);
  if (v instanceof FErr) return v;
  if (f instanceof FErr) return f;
  let val: Value = v;
  if (typeof v === 'string') {
    const n = toNum(v);
    if (!(n instanceof FErr)) val = n;
  }
  return formatValue(val, f).text.split(FILL_MARK).join('');
});
def('NUMBERVALUE', 1, 3, 'text, [decimal_separator], [group_separator]', 'Converts text to number in a locale-independent way.', (a, c) => {
  const s = str(a[0], c);
  if (s instanceof FErr) return s;
  const n = parseNumberText(s.replace(/\s/g, ''));
  return n ?? E.VALUE;
});

// ---------- date & time ----------

def('TODAY', 0, 0, '', 'Returns the current date.', () => todaySerial(), { volatile: true });
def('NOW', 0, 0, '', 'Returns the current date and time.', () => nowSerial(), { volatile: true });
def('DATE', 3, 3, 'year, month, day', 'Returns the serial number of a particular date.', (a, c) => {
  let y = num(a[0], c);
  const m = num(a[1], c);
  const d = num(a[2], c);
  if (y instanceof FErr) return y;
  if (m instanceof FErr) return m;
  if (d instanceof FErr) return d;
  y = Math.floor(y);
  if (y < 0 || y > 9999) return E.NUM;
  if (y < 1900) y += 1900;
  const s = ymdToSerial(y, Math.floor(m), Math.floor(d));
  return s < 0 ? E.NUM : s;
});
def('TIME', 3, 3, 'hour, minute, second', 'Returns the serial number of a particular time.', (a, c) => {
  const h = num(a[0], c);
  const m = num(a[1], c);
  const s = num(a[2], c);
  if (h instanceof FErr) return h;
  if (m instanceof FErr) return m;
  if (s instanceof FErr) return s;
  const t = (h * 3600 + m * 60 + s) / 86400;
  return t < 0 ? E.NUM : t - Math.floor(t);
});
function dateArg(a: Arg, c: Ctx): number | FErr {
  const v = sv(a, c);
  const n = toNum(v);
  if (n instanceof FErr) return n;
  if (n < 0) return E.NUM;
  return n;
}
def('YEAR', 1, 1, 'serial_number', 'Converts a serial number to a year.', (a, c) => {
  const n = dateArg(a[0], c);
  return n instanceof FErr ? n : serialToYmd(n).y;
});
def('MONTH', 1, 1, 'serial_number', 'Converts a serial number to a month.', (a, c) => {
  const n = dateArg(a[0], c);
  return n instanceof FErr ? n : serialToYmd(n).m;
});
def('DAY', 1, 1, 'serial_number', 'Converts a serial number to a day of the month.', (a, c) => {
  const n = dateArg(a[0], c);
  return n instanceof FErr ? n : serialToYmd(n).d;
});
def('HOUR', 1, 1, 'serial_number', 'Converts a serial number to an hour.', (a, c) => {
  const n = dateArg(a[0], c);
  return n instanceof FErr ? n : serialToTime(n).h;
});
def('MINUTE', 1, 1, 'serial_number', 'Converts a serial number to a minute.', (a, c) => {
  const n = dateArg(a[0], c);
  return n instanceof FErr ? n : serialToTime(n).mi;
});
def('SECOND', 1, 1, 'serial_number', 'Converts a serial number to a second.', (a, c) => {
  const n = dateArg(a[0], c);
  return n instanceof FErr ? n : serialToTime(n).s;
});
def('WEEKDAY', 1, 2, 'serial_number, [return_type]', 'Converts a serial number to a day of the week.', (a, c) => {
  const n = dateArg(a[0], c);
  if (n instanceof FErr) return n;
  const t = a[1] === undefined ? 1 : num(a[1], c);
  if (t instanceof FErr) return t;
  const w = weekdayOf(n); // 0 = Sunday
  if (t === 1) return w + 1;
  if (t === 2) return ((w + 6) % 7) + 1;
  if (t === 3) return (w + 6) % 7;
  return E.NUM;
});
def('DATEVALUE', 1, 1, 'date_text', 'Converts a date in the form of text to a serial number.', (a, c) => {
  const s = str(a[0], c);
  if (s instanceof FErr) return s;
  const d = parseDateText(s);
  return d === null ? E.VALUE : Math.floor(d);
});
def('TIMEVALUE', 1, 1, 'time_text', 'Converts a time in the form of text to a serial number.', (a, c) => {
  const s = str(a[0], c);
  if (s instanceof FErr) return s;
  const d = parseDateText(s);
  return d === null ? E.VALUE : d - Math.floor(d);
});
def('DAYS', 2, 2, 'end_date, start_date', 'Returns the number of days between two dates.', (a, c) => {
  const e = dateArg(a[0], c);
  const s = dateArg(a[1], c);
  if (e instanceof FErr) return e;
  if (s instanceof FErr) return s;
  return Math.floor(e) - Math.floor(s);
});
def('EDATE', 2, 2, 'start_date, months', 'Returns the date that is the indicated number of months before or after the start date.', (a, c) => {
  const s = dateArg(a[0], c);
  const m = num(a[1], c);
  if (s instanceof FErr) return s;
  if (m instanceof FErr) return m;
  const { y, m: mo, d } = serialToYmd(s);
  const total = y * 12 + (mo - 1) + Math.trunc(m);
  const ny = Math.floor(total / 12);
  const nm = (total % 12) + 1;
  const dim = new Date(Date.UTC(ny, nm, 0)).getUTCDate();
  return ymdToSerial(ny, nm, Math.min(d, dim));
});
def('EOMONTH', 2, 2, 'start_date, months', 'Returns the last day of the month before or after a specified number of months.', (a, c) => {
  const s = dateArg(a[0], c);
  const m = num(a[1], c);
  if (s instanceof FErr) return s;
  if (m instanceof FErr) return m;
  const { y, m: mo } = serialToYmd(s);
  const total = y * 12 + (mo - 1) + Math.trunc(m);
  return ymdToSerial(Math.floor(total / 12), (total % 12) + 2, 0);
});
def('DATEDIF', 3, 3, 'start_date, end_date, unit', 'Calculates the difference between two dates.', (a, c) => {
  const s = dateArg(a[0], c);
  const e = dateArg(a[1], c);
  const u = str(a[2], c);
  if (s instanceof FErr) return s;
  if (e instanceof FErr) return e;
  if (u instanceof FErr) return u;
  if (e < s) return E.NUM;
  const a1 = serialToYmd(s);
  const b1 = serialToYmd(e);
  let months = (b1.y - a1.y) * 12 + (b1.m - a1.m);
  if (b1.d < a1.d) months--;
  switch (u.toUpperCase()) {
    case 'D': return Math.floor(e) - Math.floor(s);
    case 'M': return months;
    case 'Y': return Math.floor(months / 12);
    case 'MD': {
      let d = b1.d - a1.d;
      if (d < 0) d += new Date(Date.UTC(b1.y, b1.m - 1, 0)).getUTCDate();
      return d;
    }
    case 'YM': return ((months % 12) + 12) % 12;
    default: return E.NUM;
  }
});

// ---------- lookup & reference ----------

function wantRange(a: Arg): RangeVal | Value[][] {
  if (isRange(a)) return a;
  return toMatrix(a);
}
function cellAt(src: RangeVal | Value[][], r: number, c: number): Value {
  if (Array.isArray(src)) return src[r]?.[c] ?? null;
  return getCellValue(src.sheet, src.r1 + r, src.c1 + c);
}
function dimsOf(src: RangeVal | Value[][]): { rows: number; cols: number } {
  if (Array.isArray(src)) return { rows: src.length, cols: src[0]?.length ?? 0 };
  return { rows: src.r2 - src.r1 + 1, cols: src.c2 - src.c1 + 1 };
}
function usedRows(src: RangeVal | Value[][]): number {
  if (Array.isArray(src)) return src.length;
  return Math.min(src.r2 - src.r1 + 1, Math.max(0, src.sheet.maxRow - src.r1 + 1));
}
function usedCols(src: RangeVal | Value[][]): number {
  if (Array.isArray(src)) return src[0]?.length ?? 0;
  return Math.min(src.c2 - src.c1 + 1, Math.max(0, src.sheet.maxCol - src.c1 + 1));
}

function eqLookup(x: Value, v: Value, wild: boolean): boolean {
  if (v === null) return false;
  if (typeof x === 'string' && typeof v === 'string') {
    if (wild && /[*?~]/.test(x)) return wildcardRe(x).test(v);
    return x.toLowerCase() === v.toLowerCase();
  }
  if (typeof x !== typeof v) return false;
  return x === v;
}

/** Binary-ish approximate match for sorted data: index of largest value <= x (asc) or smallest >= x (desc). */
function approxIndex(n: number, get: (i: number) => Value, x: Value, asc: boolean): number {
  let lo = 0;
  let hi = n - 1;
  let res = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const v = get(mid);
    if (v === null || typeof v !== typeof x) {
      // skip type mismatches: linear fallback
      return linearApprox(n, get, x, asc);
    }
    const cmp = compareValues(v, x);
    if (asc ? cmp <= 0 : cmp >= 0) {
      res = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return res;
}
function linearApprox(n: number, get: (i: number) => Value, x: Value, asc: boolean): number {
  let res = -1;
  for (let i = 0; i < n; i++) {
    const v = get(i);
    if (v === null || typeof v !== typeof x) continue;
    const cmp = compareValues(v, x);
    if (asc ? cmp <= 0 : cmp >= 0) res = i;
    else break;
  }
  return res;
}

function lookupCore(
  x: Value, n: number, get: (i: number) => Value, mode: 'exact' | 'approx' | 'approxDesc', wild = true,
): number {
  if (mode === 'exact') {
    for (let i = 0; i < n; i++) if (eqLookup(x, get(i), wild)) return i;
    return -1;
  }
  return approxIndex(n, get, x, mode === 'approx');
}

def('VLOOKUP', 3, 4, 'lookup_value, table_array, col_index_num, [range_lookup]', 'Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.', (a, c) => {
  const x = sv(a[0], c);
  if (x instanceof FErr) return x;
  const t = wantRange(a[1]);
  const ci = num(a[2], c);
  if (ci instanceof FErr) return ci;
  const exact = a[3] !== undefined && toBool(sv(a[3], c)) === false;
  const d = dimsOf(t);
  if (ci < 1) return E.VALUE;
  if (ci > d.cols) return E.REF;
  const n = usedRows(t);
  const i = lookupCore(x, n, (k) => cellAt(t, k, 0), exact ? 'exact' : 'approx');
  if (i < 0) return E.NA;
  const v = cellAt(t, i, Math.floor(ci) - 1);
  return v === null ? 0 : v;
});
def('HLOOKUP', 3, 4, 'lookup_value, table_array, row_index_num, [range_lookup]', 'Looks in the top row of a table and returns the value in the same column from a row you specify.', (a, c) => {
  const x = sv(a[0], c);
  if (x instanceof FErr) return x;
  const t = wantRange(a[1]);
  const ri = num(a[2], c);
  if (ri instanceof FErr) return ri;
  const exact = a[3] !== undefined && toBool(sv(a[3], c)) === false;
  const d = dimsOf(t);
  if (ri < 1) return E.VALUE;
  if (ri > d.rows) return E.REF;
  const n = usedCols(t);
  const i = lookupCore(x, n, (k) => cellAt(t, 0, k), exact ? 'exact' : 'approx');
  if (i < 0) return E.NA;
  const v = cellAt(t, Math.floor(ri) - 1, i);
  return v === null ? 0 : v;
});
def('MATCH', 2, 3, 'lookup_value, lookup_array, [match_type]', 'Returns the relative position of an item in an array.', (a, c) => {
  const x = sv(a[0], c);
  if (x instanceof FErr) return x;
  const t = wantRange(a[1]);
  const mt = a[2] === undefined ? 1 : num(a[2], c);
  if (mt instanceof FErr) return mt;
  const d = dimsOf(t);
  if (d.rows !== 1 && d.cols !== 1) return E.NA;
  const vertical = d.cols === 1;
  const n = vertical ? usedRows(t) : usedCols(t);
  const get = (k: number) => (vertical ? cellAt(t, k, 0) : cellAt(t, 0, k));
  const i = lookupCore(x, n, get, mt === 0 ? 'exact' : mt > 0 ? 'approx' : 'approxDesc');
  return i < 0 ? E.NA : i + 1;
});
def('XMATCH', 2, 4, 'lookup_value, lookup_array, [match_mode], [search_mode]', 'Returns the relative position of an item in an array or range.', (a, c) => {
  const x = sv(a[0], c);
  if (x instanceof FErr) return x;
  const t = wantRange(a[1]);
  const mm = a[2] === undefined ? 0 : num(a[2], c);
  const sm = a[3] === undefined ? 1 : num(a[3], c);
  if (mm instanceof FErr) return mm;
  if (sm instanceof FErr) return sm;
  const d = dimsOf(t);
  const vertical = d.cols === 1;
  const n = vertical ? usedRows(t) : usedCols(t);
  const get = (k: number) => (vertical ? cellAt(t, k, 0) : cellAt(t, 0, k));
  const i = xFind(x, n, get, mm, sm);
  return i < 0 ? E.NA : i + 1;
});
function xFind(x: Value, n: number, get: (i: number) => Value, mm: number, sm: number): number {
  const idx = (k: number) => (sm === -1 ? n - 1 - k : k);
  if (mm === 0 || mm === 2) {
    for (let k = 0; k < n; k++) {
      const i = idx(k);
      if (eqLookup(x, get(i), mm === 2)) return i;
    }
    return -1;
  }
  // nearest match
  let best = -1;
  for (let i = 0; i < n; i++) {
    const v = get(i);
    if (v === null || typeof v !== typeof x) continue;
    const cmp = compareValues(v, x);
    if (cmp === 0) return i;
    if (mm === -1 && cmp < 0 && (best < 0 || compareValues(v, get(best)) > 0)) best = i;
    if (mm === 1 && cmp > 0 && (best < 0 || compareValues(v, get(best)) < 0)) best = i;
  }
  return best;
}
def('XLOOKUP', 3, 6, 'lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]', 'Searches a range or an array and returns an item corresponding to the first match found.', (a, c) => {
  const x = sv(a[0], c);
  if (x instanceof FErr) return x;
  const la = wantRange(a[1]);
  const ra = wantRange(a[2]);
  const mm = a[4] === undefined ? 0 : num(a[4], c);
  const sm = a[5] === undefined ? 1 : num(a[5], c);
  if (mm instanceof FErr) return mm;
  if (sm instanceof FErr) return sm;
  const d = dimsOf(la);
  if (d.rows !== 1 && d.cols !== 1) return E.VALUE;
  const vertical = d.cols === 1;
  const n = vertical ? usedRows(la) : usedCols(la);
  const get = (k: number) => (vertical ? cellAt(la, k, 0) : cellAt(la, 0, k));
  const i = xFind(x, n, get, mm, sm);
  if (i < 0) return a[3] !== undefined ? a[3] : E.NA;
  const rd = dimsOf(ra);
  if (vertical) {
    if (rd.rows !== d.rows) return E.VALUE;
    if (rd.cols === 1) {
      const v = cellAt(ra, i, 0);
      return v === null ? 0 : v;
    }
    const row: Value[] = [];
    for (let k = 0; k < rd.cols; k++) row.push(cellAt(ra, i, k));
    return { k: 'arr', rows: [row] };
  }
  if (rd.cols !== d.cols) return E.VALUE;
  if (rd.rows === 1) {
    const v = cellAt(ra, 0, i);
    return v === null ? 0 : v;
  }
  const col: Value[][] = [];
  for (let k = 0; k < rd.rows; k++) col.push([cellAt(ra, k, i)]);
  return { k: 'arr', rows: col };
});
def('LOOKUP', 2, 3, 'lookup_value, lookup_vector, [result_vector]', 'Looks up a value in a vector.', (a, c) => {
  const x = sv(a[0], c);
  if (x instanceof FErr) return x;
  const t = wantRange(a[1]);
  const d = dimsOf(t);
  const vert = d.cols === 1 || d.rows > 1;
  const n = vert ? usedRows(t) : usedCols(t);
  const get = (k: number) => (vert ? cellAt(t, k, 0) : cellAt(t, 0, k));
  const i = approxIndex(n, get, x, true);
  if (i < 0) return E.NA;
  const rv = a[2] === undefined ? t : wantRange(a[2]);
  const rd = dimsOf(rv);
  const v = rd.cols === 1 ? cellAt(rv, i, 0) : cellAt(rv, 0, i);
  return v === null ? 0 : v;
});
def('INDEX', 2, 3, 'array, row_num, [column_num]', 'Returns a value or the reference to a value from within a table or range.', (a, c) => {
  const src = a[0];
  const rn = num(a[1], c);
  if (rn instanceof FErr) return rn;
  const cn = a[2] === undefined ? null : num(a[2], c);
  if (cn instanceof FErr) return cn;
  const m = isRange(src) ? null : toMatrix(src);
  const rows = isRange(src) ? src.r2 - src.r1 + 1 : m!.length;
  const cols = isRange(src) ? src.c2 - src.c1 + 1 : m![0].length;
  let r = Math.floor(rn);
  let cc: number;
  if (cn === null) {
    if (rows === 1 && cols > 1) {
      cc = r;
      r = 1;
    } else cc = cols === 1 ? 1 : 0;
  } else cc = Math.floor(cn);
  if (r < 0 || cc < 0 || r > rows || cc > cols) return E.REF;
  if (isRange(src)) {
    const r1 = r === 0 ? src.r1 : src.r1 + r - 1;
    const r2 = r === 0 ? src.r2 : r1;
    const c1 = cc === 0 ? src.c1 : src.c1 + cc - 1;
    const c2 = cc === 0 ? src.c2 : c1;
    return { k: 'rng', sheet: src.sheet, r1, c1, r2, c2 };
  }
  if (r === 0 && cc === 0) return { k: 'arr', rows: m! };
  if (r === 0) return { k: 'arr', rows: m!.map((row) => [row[cc - 1]]) };
  if (cc === 0) return { k: 'arr', rows: [m![r - 1]] };
  return m![r - 1][cc - 1];
});
def('ROWS', 1, 1, 'array', 'Returns the number of rows in a reference.', (a) => {
  if (isRange(a[0])) return a[0].r2 - a[0].r1 + 1;
  return toMatrix(a[0]).length;
});
def('COLUMNS', 1, 1, 'array', 'Returns the number of columns in a reference.', (a) => {
  if (isRange(a[0])) return a[0].c2 - a[0].c1 + 1;
  return toMatrix(a[0])[0].length;
});
lazy('ROW', 0, 1, '[reference]', 'Returns the row number of a reference.', (a, c) => {
  if (!a.length) return c.row + 1;
  const n = a[0];
  if (n.k === 'ref') return n.a.r + 1;
  if (n.k === 'range') return Math.min(n.a.r, n.b.r) + 1;
  return E.VALUE;
});
lazy('COLUMN', 0, 1, '[reference]', 'Returns the column number of a reference.', (a, c) => {
  if (!a.length) return c.col + 1;
  const n = a[0];
  if (n.k === 'ref') return n.a.c + 1;
  if (n.k === 'range') return Math.min(n.a.c, n.b.c) + 1;
  return E.VALUE;
});
void evalNode;
void fl;
void MAX_COLS;
void MAX_ROWS;
void monthIndex;
void numToStr;
