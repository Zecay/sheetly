// Formula tokenizer + parser (Excel syntax) and reference rewriting.
import { colIndex, colName, MAX_COLS, MAX_ROWS } from './addr';
import { errFromCode, FErr } from './values';

export interface RefPart {
  r: number; // 0-based; for whole-column refs r is 0 / MAX_ROWS-1
  c: number;
  absR: boolean;
  absC: boolean;
}

export type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'err'; v: FErr }
  | { k: 'ref'; sheet?: string; a: RefPart }
  | { k: 'range'; sheet?: string; a: RefPart; b: RefPart; whole?: 'col' | 'row' }
  | { k: 'name'; name: string }
  | { k: 'un'; op: '-' | '+'; x: Node }
  | { k: 'pct'; x: Node }
  | { k: 'bin'; op: string; l: Node; r: Node }
  | { k: 'call'; name: string; args: Node[] }
  | { k: 'missing' };

export type TokKind = 'num' | 'str' | 'bool' | 'err' | 'ref' | 'name' | 'func' | 'op' | 'lparen' | 'rparen' | 'comma' | 'ws' | 'bad';

export interface Token {
  kind: TokKind;
  text: string;
  start: number;
  end: number;
  // for refs
  sheet?: string;
  range?: { a: RefPart; b?: RefPart; whole?: 'col' | 'row' };
}

export class ParseError extends Error {}

const SHEET_PREFIX = "(?:'(?:[^']|'')+'|[A-Za-z_\\u00C0-\\uFFFF][\\w.\\u00C0-\\uFFFF]*)!";
const CELL = '\\$?[A-Za-z]{1,3}\\$?\\d+';
const REF_RE = new RegExp(
  `^(${SHEET_PREFIX})?(${CELL}(?::${CELL})?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?\\d+:\\$?\\d+)`,
);
const NUM_RE = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const NAME_RE = /^[A-Za-z_\\À-￿][\w.À-￿]*/;
const ERR_RE = /^#(?:DIV\/0!|REF!|NAME\?|VALUE!|N\/A|NUM!|NULL!|CIRC!)/i;

function unquoteSheet(s: string): string {
  s = s.slice(0, -1); // strip !
  if (s.startsWith("'")) return s.slice(1, -1).replace(/''/g, "'");
  return s;
}

export function quoteSheet(name: string): string {
  if (/^[A-Za-z_][\w.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^[Rr]\d*[Cc]\d*$/.test(name)) return name;
  return "'" + name.replace(/'/g, "''") + "'";
}

function parseRefPart(s: string): RefPart | null {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(s);
  if (!m) return null;
  const c = colIndex(m[2]);
  const r = parseInt(m[4], 10) - 1;
  if (c >= MAX_COLS || r >= MAX_ROWS || r < 0) return null;
  return { r, c, absC: m[1] === '$', absR: m[3] === '$' };
}

function parseRefText(txt: string): { a: RefPart; b?: RefPart; whole?: 'col' | 'row' } | null {
  const parts = txt.split(':');
  if (parts.length === 1) {
    const a = parseRefPart(parts[0]);
    return a ? { a } : null;
  }
  const a = parseRefPart(parts[0]);
  const b = parseRefPart(parts[1]);
  if (a && b) return { a, b };
  const mc1 = /^(\$?)([A-Za-z]{1,3})$/.exec(parts[0]);
  const mc2 = /^(\$?)([A-Za-z]{1,3})$/.exec(parts[1]);
  if (mc1 && mc2) {
    const c1 = colIndex(mc1[2]);
    const c2 = colIndex(mc2[2]);
    if (c1 >= MAX_COLS || c2 >= MAX_COLS) return null;
    return {
      a: { r: 0, c: c1, absR: true, absC: mc1[1] === '$' },
      b: { r: MAX_ROWS - 1, c: c2, absR: true, absC: mc2[1] === '$' },
      whole: 'col',
    };
  }
  const mr1 = /^(\$?)(\d+)$/.exec(parts[0]);
  const mr2 = /^(\$?)(\d+)$/.exec(parts[1]);
  if (mr1 && mr2) {
    const r1 = parseInt(mr1[2], 10) - 1;
    const r2 = parseInt(mr2[2], 10) - 1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return null;
    return {
      a: { r: r1, c: 0, absR: mr1[1] === '$', absC: true },
      b: { r: r2, c: MAX_COLS - 1, absR: mr2[1] === '$', absC: true },
      whole: 'row',
    };
  }
  return null;
}

export function tokenize(src: string, keepWs = false): Token[] {
  const out: Token[] = [];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') {
      let j = i;
      while (j < n && /\s/.test(src[j])) j++;
      if (keepWs) out.push({ kind: 'ws', text: src.slice(i, j), start: i, end: j });
      else out.push({ kind: 'ws', text: src.slice(i, j), start: i, end: j });
      i = j;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j++];
      }
      if (j >= n) {
        out.push({ kind: 'bad', text: src.slice(i), start: i, end: n });
        i = n;
        continue;
      }
      out.push({ kind: 'str', text: s, start: i, end: j + 1 });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const m = ERR_RE.exec(src.slice(i));
      if (m) {
        out.push({ kind: 'err', text: m[0].toUpperCase(), start: i, end: i + m[0].length });
        i += m[0].length;
        continue;
      }
    }
    const rest = i === 0 ? src : src.slice(i);
    // references first (they begin with a letter, $, digit or quote)
    if (/[A-Za-z_$'\dÀ-￿]/.test(ch)) {
      const m = REF_RE.exec(rest);
      if (m) {
        const after = rest[m[0].length];
        const isCall = after === '(';
        // "A1" followed by identifier chars is not a reference (e.g. A1B)
        const idChar = after !== undefined && /[\w.À-￿]/.test(after);
        // a number like 1:2 is fine; but plain digits followed by e.g. ".5" not matched here anyway
        if (!isCall && !idChar) {
          const range = parseRefText(m[2]);
          if (range) {
            out.push({
              kind: 'ref',
              text: m[0],
              start: i,
              end: i + m[0].length,
              sheet: m[1] ? unquoteSheet(m[1]) : undefined,
              range,
            });
            i += m[0].length;
            continue;
          }
        }
      }
    }
    if (/[\d.]/.test(ch)) {
      const m = NUM_RE.exec(rest);
      if (m) {
        out.push({ kind: 'num', text: m[0], start: i, end: i + m[0].length });
        i += m[0].length;
        continue;
      }
    }
    if (/[A-Za-z_\\À-￿]/.test(ch) || (ch === "'" && false)) {
      const m = NAME_RE.exec(rest);
      if (m) {
        const text = m[0];
        const next = rest[text.length];
        // skip whitespace to find '('
        if (next === '(') {
          out.push({ kind: 'func', text, start: i, end: i + text.length });
        } else if (/^(TRUE|FALSE)$/i.test(text)) {
          out.push({ kind: 'bool', text, start: i, end: i + text.length });
        } else out.push({ kind: 'name', text, start: i, end: i + text.length });
        i += text.length;
        continue;
      }
    }
    if (ch === '(') {
      out.push({ kind: 'lparen', text: '(', start: i, end: i + 1 });
      i++;
      continue;
    }
    if (ch === ')') {
      out.push({ kind: 'rparen', text: ')', start: i, end: i + 1 });
      i++;
      continue;
    }
    if (ch === ',' || ch === ';') {
      out.push({ kind: 'comma', text: ',', start: i, end: i + 1 });
      i++;
      continue;
    }
    const two = src.substr(i, 2);
    if (two === '<=' || two === '>=' || two === '<>') {
      out.push({ kind: 'op', text: two, start: i, end: i + 2 });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%:'.includes(ch)) {
      out.push({ kind: 'op', text: ch, start: i, end: i + 1 });
      i++;
      continue;
    }
    out.push({ kind: 'bad', text: ch, start: i, end: i + 1 });
    i++;
  }
  return out;
}

// ---------------- parser ----------------

const PREC: Record<string, number> = {
  '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1,
  '&': 2,
  '+': 3, '-': 3,
  '*': 4, '/': 4,
  '^': 5,
};

const MAX_PARSE_DEPTH = 500;
const MAX_PARSE_OPS = 2000;

export function parseFormula(src: string): Node {
  // src without the leading "="
  const toks = tokenize(src).filter((t) => t.kind !== 'ws');
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  let depth = 0;
  let ops = 0;
  const enter = () => {
    if (++depth > MAX_PARSE_DEPTH) throw new ParseError('Formula is nested too deeply');
  };

  function parseExpr(minPrec: number): Node {
    enter();
    try {
      return parseExprInner(minPrec);
    } finally {
      depth--;
    }
  }

  function parseExprInner(minPrec: number): Node {
    let left = parseUnary();
    for (;;) {
      const t = peek();
      if (!t || t.kind !== 'op' || !(t.text in PREC)) break;
      const prec = PREC[t.text];
      if (prec < minPrec) break;
      next();
      if (++ops > MAX_PARSE_OPS) throw new ParseError('Formula is too long');
      const right = parseExpr(prec + 1); // left associative
      left = { k: 'bin', op: t.text, l: left, r: right };
    }
    return left;
  }

  function parseUnary(): Node {
    const t = peek();
    if (t && t.kind === 'op' && (t.text === '-' || t.text === '+')) {
      next();
      enter();
      try {
        const x = parseUnary();
        return { k: 'un', op: t.text as '-' | '+', x };
      } finally {
        depth--;
      }
    }
    return parsePostfix();
  }

  function parsePostfix(): Node {
    let x = parsePrimary();
    while (peek() && peek().kind === 'op' && peek().text === '%') {
      next();
      x = { k: 'pct', x };
    }
    // exponent binds tighter than unary minus on its left but handled by PREC loop; unary above wins like Excel
    return x;
  }

  function parseArgs(): Node[] {
    const args: Node[] = [];
    if (peek() && peek().kind === 'rparen') {
      next();
      return args;
    }
    for (;;) {
      const t = peek();
      if (!t) throw new ParseError('Missing )');
      if (t.kind === 'comma' || t.kind === 'rparen') {
        args.push({ k: 'missing' });
      } else args.push(parseExpr(1));
      const s = next();
      if (!s) throw new ParseError('Missing )');
      if (s.kind === 'rparen') break;
      if (s.kind !== 'comma') throw new ParseError('Unexpected ' + s.text);
    }
    return args;
  }

  function parsePrimary(): Node {
    const t = next();
    if (!t) throw new ParseError('Unexpected end');
    switch (t.kind) {
      case 'num':
        return { k: 'num', v: parseFloat(t.text) };
      case 'str':
        return { k: 'str', v: t.text };
      case 'bool':
        return { k: 'bool', v: t.text.toUpperCase() === 'TRUE' };
      case 'err':
        return { k: 'err', v: errFromCode(t.text)! };
      case 'ref': {
        const rg = t.range!;
        // allow A1:B2 or A1:INDEX(...)? keep simple
        if (rg.b) return { k: 'range', sheet: t.sheet, a: rg.a, b: rg.b, whole: rg.whole };
        // A1 : B2 written with the colon operator is handled below
        const nx = peek();
        if (nx && nx.kind === 'op' && nx.text === ':') {
          const save = p;
          next();
          const t2 = peek();
          if (t2 && t2.kind === 'ref' && t2.range && !t2.range.b) {
            next();
            return { k: 'range', sheet: t.sheet, a: rg.a, b: t2.range.a };
          }
          p = save;
        }
        return { k: 'ref', sheet: t.sheet, a: rg.a };
      }
      case 'name':
        return { k: 'name', name: t.text };
      case 'func': {
        const l = next();
        if (!l || l.kind !== 'lparen') throw new ParseError('Expected (');
        return { k: 'call', name: t.text.toUpperCase(), args: parseArgs() };
      }
      case 'lparen': {
        const e = parseExpr(1);
        const r = next();
        if (!r || r.kind !== 'rparen') throw new ParseError('Missing )');
        return e;
      }
      default:
        throw new ParseError('Unexpected ' + t.text);
    }
  }

  if (toks.length === 0) throw new ParseError('Empty formula');
  let node: Node;
  try {
    node = parseExpr(1);
  } catch (e) {
    if (e instanceof RangeError) throw new ParseError('Formula is nested too deeply');
    throw e;
  }
  if (p < toks.length) throw new ParseError('Unexpected ' + toks[p].text);
  return node;
}

// ---------------- reference rewriting ----------------

export interface RefEditor {
  /** Return the new reference text (without sheet prefix) or null to turn the whole token into #REF!. */
  (sheet: string | undefined, a: RefPart, b: RefPart | undefined, whole: 'col' | 'row' | undefined): { a: RefPart; b?: RefPart; sheet?: string } | null;
}

export function refText(a: RefPart, b?: RefPart, whole?: 'col' | 'row'): string {
  if (whole === 'col' && b) return (a.absC ? '$' : '') + colName(a.c) + ':' + (b.absC ? '$' : '') + colName(b.c);
  if (whole === 'row' && b) return (a.absR ? '$' : '') + (a.r + 1) + ':' + (b.absR ? '$' : '') + (b.r + 1);
  const one = (x: RefPart) => (x.absC ? '$' : '') + colName(x.c) + (x.absR ? '$' : '') + (x.r + 1);
  return b ? one(a) + ':' + one(b) : one(a);
}

/** Rewrite every reference in a formula string (no leading "="). */
export function rewriteRefs(formula: string, edit: RefEditor): string {
  const toks = tokenize(formula);
  let out = '';
  let pos = 0;
  for (const t of toks) {
    if (t.kind !== 'ref') continue;
    out += formula.slice(pos, t.start);
    pos = t.end;
    const rg = t.range!;
    const res = edit(t.sheet, rg.a, rg.b, rg.whole);
    const target = res?.sheet ?? t.sheet;
    const prefix = target !== undefined ? quoteSheet(target) + '!' : '';
    if (!res) out += '#REF!';
    else out += prefix + refText(res.a, res.b, rg.whole);
  }
  out += formula.slice(pos);
  return out;
}

/** Shift relative references by (dr, dc), like copy/paste or fill. */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula;
  return rewriteRefs(formula, (_s, a, b, whole) => {
    const sh = (x: RefPart, isEndOfWhole: boolean): RefPart | null => {
      const r = x.absR || (whole === 'col') ? x.r : x.r + dr;
      const c = x.absC || (whole === 'row') ? x.c : x.c + dc;
      void isEndOfWhole;
      if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null;
      return { ...x, r, c };
    };
    const na = sh(a, false);
    if (!na) return null;
    if (!b) return { a: na };
    const nb = sh(b, true);
    if (!nb) return null;
    return { a: na, b: nb };
  });
}

/** Rename a sheet inside formula text. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const toks = tokenize(formula);
  let out = '';
  let pos = 0;
  for (const t of toks) {
    if (t.kind !== 'ref' || t.sheet === undefined || t.sheet.toLowerCase() !== oldName.toLowerCase()) continue;
    out += formula.slice(pos, t.start);
    pos = t.end;
    out += quoteSheet(newName) + '!' + t.text.slice(t.text.lastIndexOf('!') + 1);
  }
  out += formula.slice(pos);
  return out;
}
