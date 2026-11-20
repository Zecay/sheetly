// Cell address helpers. Rows and columns are 0-based internally.
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export function colName(c: number): string {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(name: string): number {
  let n = 0;
  for (let i = 0; i < name.length; i++) {
    const ch = name.charCodeAt(i);
    const v = ch >= 97 ? ch - 96 : ch - 64;
    n = n * 26 + v;
  }
  return n - 1;
}

/** Numeric cell key: row-major, so sorting keys sorts cells by row then column. */
export function cellKey(r: number, c: number): number {
  return r * MAX_COLS + c;
}
export function keyRow(k: number): number {
  return Math.floor(k / MAX_COLS);
}
export function keyCol(k: number): number {
  return k % MAX_COLS;
}

export function a1(r: number, c: number, absR = false, absC = false): string {
  return (absC ? '$' : '') + colName(c) + (absR ? '$' : '') + (r + 1);
}

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function normRect(r1: number, c1: number, r2: number, c2: number): Rect {
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

export function rectText(r: Rect): string {
  if (r.r1 === r.r2 && r.c1 === r.c2) return a1(r.r1, r.c1);
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return colName(r.c1) + ':' + colName(r.c2);
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return r.r1 + 1 + ':' + (r.r2 + 1);
  return a1(r.r1, r.c1) + ':' + a1(r.r2, r.c2);
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

export function parseCell(s: string): { r: number; c: number } | null {
  const m = CELL_RE.exec(s.trim());
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = parseInt(m[2], 10) - 1;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
  return { r, c };
}

/** Parse "A1", "A1:B5", "A:A", "1:3" (with optional $) into a Rect. */
export function parseRect(s: string): Rect | null {
  s = s.trim().replace(/\$/g, '');
  const parts = s.split(':');
  if (parts.length === 1) {
    const a = parseCell(parts[0]);
    return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null;
  }
  if (parts.length !== 2) return null;
  const a = parseCell(parts[0]);
  const b = parseCell(parts[1]);
  if (a && b) return normRect(a.r, a.c, b.r, b.c);
  if (/^[A-Za-z]{1,3}$/.test(parts[0]) && /^[A-Za-z]{1,3}$/.test(parts[1])) {
    const c1 = colIndex(parts[0]);
    const c2 = colIndex(parts[1]);
    if (c1 < MAX_COLS && c2 < MAX_COLS) return normRect(0, c1, MAX_ROWS - 1, c2);
  }
  if (/^\d+$/.test(parts[0]) && /^\d+$/.test(parts[1])) {
    const r1 = parseInt(parts[0], 10) - 1;
    const r2 = parseInt(parts[1], 10) - 1;
    if (r1 >= 0 && r2 >= 0 && r1 < MAX_ROWS && r2 < MAX_ROWS) return normRect(r1, 0, r2, MAX_COLS - 1);
  }
  return null;
}
