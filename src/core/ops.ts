// Editing operations on a workbook: input, clear, insert/delete, fill, sort, copy/move, find.
import { MAX_COLS, MAX_ROWS, Rect, cellKey, keyCol, keyRow, normRect } from './addr';
import { Cell, CFRule, Sheet, StylePatch, Workbook } from './model';
import { parseInput } from './input';
import { ParseError, RefEditor, RefPart, parseFormula, rewriteRefs, shiftFormula } from './parser';
import { FErr, Value, DAYS, MONTHS, isErr } from './values';
import { compareValues } from './eval';
import { formatValue, FILL_MARK } from './numfmt';

export class FormulaError extends Error {}

/** Validate formula text (without "="). */
export function checkFormula(f: string): void {
  try {
    parseFormula(f);
  } catch (e) {
    if (e instanceof ParseError || e instanceof RangeError) throw new FormulaError(e.message);
    throw e;
  }
}

/** Enter text into a cell like a user typing into it. Throws FormulaError for bad formulas. */
export function applyInput(wb: Workbook, sheet: Sheet, r: number, c: number, text: string): void {
  const old = sheet.get(r, c);
  const style = wb.styleOf(old?.s);
  const p = parseInput(text, style.fmt);
  if (p.f !== undefined) checkFormula(p.f);
  let s = old?.s;
  if (p.fmt && (!style.fmt || style.fmt === 'General')) s = wb.patchStyle(s, { fmt: p.fmt });
  if (p.v === null && p.f === undefined) {
    if (!old) return;
    wb.setCell(sheet, r, c, s ? { v: null, s } : undefined);
    return;
  }
  const cell: Cell = { v: p.v, s: s || undefined };
  if (p.f !== undefined) cell.f = p.f;
  wb.setCell(sheet, r, c, cell);
}

export function rectCells(rect: Rect): number {
  return (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
}

/** Iterate existing cells inside rect (sparse-aware). */
export function eachCell(sheet: Sheet, rect: Rect, cb: (r: number, c: number, cell: Cell) => void): void {
  if (rectCells(rect) <= sheet.cells.size) {
    for (let r = rect.r1; r <= rect.r2; r++)
      for (let c = rect.c1; c <= rect.c2; c++) {
        const cell = sheet.cells.get(cellKey(r, c));
        if (cell) cb(r, c, cell);
      }
  } else {
    const hits: [number, number, Cell][] = [];
    for (const [k, cell] of sheet.cells) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (r >= rect.r1 && r <= rect.r2 && c >= rect.c1 && c <= rect.c2) hits.push([r, c, cell]);
    }
    hits.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    for (const h of hits) cb(h[0], h[1], h[2]);
  }
}

export function clearRange(wb: Workbook, sheet: Sheet, rect: Rect, what: 'all' | 'contents' | 'formats' = 'contents'): void {
  const targets: [number, number, Cell][] = [];
  eachCell(sheet, rect, (r, c, cell) => targets.push([r, c, cell]));
  for (const [r, c, cell] of targets) {
    if (what === 'all') wb.setCell(sheet, r, c, undefined);
    else if (what === 'contents') wb.setCell(sheet, r, c, cell.s ? { v: null, s: cell.s } : undefined);
    else if (cell.v !== null || cell.f !== undefined) wb.setCell(sheet, r, c, { v: cell.v, f: cell.f, ast: cell.ast });
    else wb.setCell(sheet, r, c, undefined);
  }
}

export function applyStyle(wb: Workbook, sheet: Sheet, rect: Rect, patch: StylePatch): void {
  // For huge ranges (whole rows/columns) only touch existing cells plus visible-ish area is not possible
  // in a sparse model, so whole-row/column styling creates cells up to the used extent plus a margin.
  const area = rectCells(rect);
  if (area > 200000) {
    const used = sheet.usedRect();
    const r2 = Math.min(rect.r2, Math.max(used ? used.r2 : 0, rect.r1 + 200) + 0);
    const c2 = Math.min(rect.c2, Math.max(used ? used.c2 : 0, rect.c1 + 30));
    rect = { r1: rect.r1, c1: rect.c1, r2: rect.r2 === MAX_ROWS - 1 ? r2 : rect.r2, c2: rect.c2 === MAX_COLS - 1 ? c2 : rect.c2 };
    if (rectCells(rect) > 400000) rect = { ...rect, r2: Math.min(rect.r2, rect.r1 + Math.floor(400000 / (rect.c2 - rect.c1 + 1))) };
  }
  const cache = new Map<number, number>();
  for (let r = rect.r1; r <= rect.r2; r++)
    for (let c = rect.c1; c <= rect.c2; c++) {
      const old = sheet.get(r, c);
      const os = old?.s ?? 0;
      let ns = cache.get(os);
      if (ns === undefined) {
        ns = wb.patchStyle(os, patch);
        cache.set(os, ns);
      }
      if (ns === os) continue;
      wb.setCell(sheet, r, c, { ...(old ?? { v: null }), s: ns || undefined, ast: old?.ast });
    }
}

// ---------------- structure: insert / delete ----------------

type Axis = 'row' | 'col';

function mapCoord(i: number, at: number, delta: number): number | null {
  // delta > 0 insert, delta < 0 delete
  if (delta > 0) return i >= at ? i + delta : i;
  const n = -delta;
  if (i < at) return i;
  if (i < at + n) return null;
  return i - n;
}

function mapSpan(s: number, e: number, at: number, delta: number): [number, number] | null {
  if (delta > 0) return [s >= at ? s + delta : s, e >= at ? e + delta : e];
  const n = -delta;
  const ns = s < at ? s : s < at + n ? at : s - n;
  const ne = e < at ? e : e < at + n ? at - 1 : e - n;
  return ns > ne ? null : [ns, ne];
}

function structEditor(target: Sheet, home: string, axis: Axis, at: number, delta: number): RefEditor {
  const tl = target.name.toLowerCase();
  const max = axis === 'row' ? MAX_ROWS : MAX_COLS;
  return (sheetName, a, b, whole) => {
    const sh = (sheetName === undefined ? home : sheetName).toLowerCase();
    if (sh !== tl) return { a, b };
    if (whole === 'col' && axis === 'row') return { a, b };
    if (whole === 'row' && axis === 'col') return { a, b };
    const get = (p: RefPart) => (axis === 'row' ? p.r : p.c);
    const put = (p: RefPart, v: number): RefPart => (axis === 'row' ? { ...p, r: Math.min(v, max - 1) } : { ...p, c: Math.min(v, max - 1) });
    if (!b) {
      const v = mapCoord(get(a), at, delta);
      return v === null ? null : { a: put(a, v) };
    }
    const ga = get(a);
    const gb = get(b);
    const lo = Math.min(ga, gb);
    const hi = Math.max(ga, gb);
    const sp = mapSpan(lo, hi, at, delta);
    if (!sp) return null;
    const [nl, nh] = sp;
    return ga <= gb ? { a: put(a, nl), b: put(b, nh) } : { a: put(a, nh), b: put(b, nl) };
  };
}

function adjustRect(rect: Rect, axis: Axis, at: number, delta: number): Rect | null {
  if (axis === 'row') {
    const sp = mapSpan(rect.r1, rect.r2, at, delta);
    return sp ? { ...rect, r1: sp[0], r2: sp[1] } : null;
  }
  const sp = mapSpan(rect.c1, rect.c2, at, delta);
  return sp ? { ...rect, c1: sp[0], c2: sp[1] } : null;
}

function restructure(wb: Workbook, sheet: Sheet, axis: Axis, at: number, delta: number): void {
  const max = axis === 'row' ? MAX_ROWS : MAX_COLS;
  // move cells
  const moved = new Map<number, Cell>();
  for (const [k, cell] of sheet.cells) {
    const r = keyRow(k);
    const c = keyCol(k);
    const v = mapCoord(axis === 'row' ? r : c, at, delta);
    if (v === null || v >= max) continue;
    const nk = axis === 'row' ? cellKey(v, c) : cellKey(r, v);
    moved.set(nk, cell);
  }
  sheet.cells = moved;
  sheet.recomputeBounds();
  // sizes
  const ax = axis === 'row' ? sheet.rows : sheet.cols;
  const ns = new Map<number, number>();
  for (const [i, size] of ax.sizes) {
    const v = mapCoord(i, at, delta);
    if (v !== null && v < max) ns.set(v, size);
  }
  ax.sizes = ns;
  ax.touch();
  // merges, cf, filter
  sheet.merges = sheet.merges.map((m) => adjustRect(m, axis, at, delta)).filter((m): m is Rect => !!m);
  sheet.cf = sheet.cf.map((r) => ({ ...r, rect: adjustRect(r.rect, axis, at, delta)! })).filter((r) => r.rect);
  if (sheet.filter) {
    const nr = adjustRect(sheet.filter.rect, axis, at, delta);
    if (nr) sheet.filter.rect = nr;
    else sheet.filter = undefined;
  }
  if (axis === 'row' && delta > 0 && at < sheet.freezeRows) sheet.freezeRows += delta;
  if (axis === 'col' && delta > 0 && at < sheet.freezeCols) sheet.freezeCols += delta;
  // rewrite formulas everywhere
  for (const s of wb.sheets) {
    const edit = structEditor(sheet, s.name, axis, at, delta);
    for (const [k, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const nf = rewriteRefs(cell.f, edit);
      if (nf !== cell.f) s.cells.set(k, { v: cell.v, f: nf, s: cell.s });
    }
  }
  const nameEdit = structEditor(sheet, '', axis, at, delta);
  for (const [k, def] of Array.from(wb.names.entries())) {
    const nf = rewriteRefs(def.refers, nameEdit);
    if (nf !== def.refers) wb.names.set(k, { name: def.name, refers: nf });
  }
  wb.rebuildDeps();
  wb.requestFullRecalc();
}

export function insertRows(wb: Workbook, sheet: Sheet, at: number, count: number): void {
  wb.transact('Insert rows', () => restructure(wb, sheet, 'row', at, count), true);
}
export function deleteRows(wb: Workbook, sheet: Sheet, at: number, count: number): void {
  wb.transact('Delete rows', () => restructure(wb, sheet, 'row', at, -count), true);
}
export function insertCols(wb: Workbook, sheet: Sheet, at: number, count: number): void {
  wb.transact('Insert columns', () => restructure(wb, sheet, 'col', at, count), true);
}
export function deleteCols(wb: Workbook, sheet: Sheet, at: number, count: number): void {
  wb.transact('Delete columns', () => restructure(wb, sheet, 'col', at, -count), true);
}

/** Rename a sheet and fix every formula and name that mentions it. */
export function renameSheet(wb: Workbook, sheet: Sheet, newName: string): void {
  wb.transact('Rename sheet', () => {
    const old = sheet.name;
    sheet.name = newName;
    const ol = old.toLowerCase();
    const edit: RefEditor = (_s, a, b) => ({ a, b });
    void edit;
    for (const s of wb.sheets)
      for (const [k, cell] of s.cells) {
        if (cell.f === undefined) continue;
        const nf = renameSheetRefs(cell.f, ol, newName);
        if (nf !== cell.f) s.cells.set(k, { v: cell.v, f: nf, s: cell.s });
      }
    for (const [k, def] of Array.from(wb.names.entries())) {
      const nf = renameSheetRefs(def.refers, ol, newName);
      if (nf !== def.refers) wb.names.set(k, { name: def.name, refers: nf });
    }
    wb.rebuildDeps();
    wb.requestFullRecalc();
  }, true);
}

import { renameSheetInFormula } from './parser';
function renameSheetRefs(f: string, oldLower: string, newName: string): string {
  return renameSheetInFormula(f, oldLower, newName);
}

export function deleteSheet(wb: Workbook, sheet: Sheet): void {
  if (wb.sheets.length <= 1) return;
  wb.transact('Delete sheet', () => {
    const idx = wb.sheets.indexOf(sheet);
    const nameL = sheet.name.toLowerCase();
    wb.sheets.splice(idx, 1);
    // formulas pointing at the deleted sheet become #REF!
    for (const s of wb.sheets)
      for (const [k, cell] of s.cells) {
        if (cell.f === undefined) continue;
        const nf = rewriteRefs(cell.f, (sn, a, b) => (sn !== undefined && sn.toLowerCase() === nameL ? null : { a, b }));
        if (nf !== cell.f) s.cells.set(k, { v: cell.v, f: nf, s: cell.s });
      }
    wb.rebuildDeps();
    wb.requestFullRecalc();
  }, true);
}

export function moveSheet(wb: Workbook, from: number, to: number): void {
  if (from === to) return;
  wb.transact('Move sheet', () => {
    const [s] = wb.sheets.splice(from, 1);
    wb.sheets.splice(to, 0, s);
  }, true);
}

export function duplicateSheet(wb: Workbook, sheet: Sheet): Sheet {
  return wb.transact('Duplicate sheet', () => {
    const idx = wb.sheets.indexOf(sheet);
    const copy = wb.addSheet(wb.uniqueSheetName(sheet.name), idx + 1);
    for (const [k, cell] of sheet.cells) copy.cells.set(k, { v: cell.v, f: cell.f, s: cell.s });
    copy.cols.sizes = new Map(sheet.cols.sizes);
    copy.rows.sizes = new Map(sheet.rows.sizes);
    copy.cols.touch();
    copy.rows.touch();
    copy.merges = sheet.merges.map((m) => ({ ...m }));
    copy.cf = sheet.cf.map((c) => ({ ...c, rect: { ...c.rect }, id: wb.nextRuleId++ }));
    copy.freezeRows = sheet.freezeRows;
    copy.freezeCols = sheet.freezeCols;
    copy.tabColor = sheet.tabColor;
    copy.recomputeBounds();
    wb.rebuildDeps();
    wb.requestFullRecalc();
    return copy;
  }, true);
}

// ---------------- copy / move ----------------

export interface ClipCell {
  v: Value;
  f?: string;
  style: number; // style index in the *source workbook*
}
export interface ClipData {
  rows: number;
  cols: number;
  cells: (ClipCell | null)[][];
  sheetName: string;
  origin: { r: number; c: number };
  tsv: string;
  cut: boolean;
  rect: Rect;
}

export function displayText(wb: Workbook, cell: Cell | undefined): string {
  if (!cell || cell.v === null) return '';
  const fmt = wb.styleOf(cell.s).fmt;
  return formatValue(cell.v, fmt ?? 'General').text.split(FILL_MARK).join('');
}

export function tsvEscape(s: string): string {
  return /[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

export function copyRange(wb: Workbook, sheet: Sheet, rect: Rect, cut = false): ClipData {
  const rows = rect.r2 - rect.r1 + 1;
  const cols = rect.c2 - rect.c1 + 1;
  const cells: (ClipCell | null)[][] = [];
  const lines: string[] = [];
  for (let r = 0; r < rows; r++) {
    const row: (ClipCell | null)[] = [];
    const parts: string[] = [];
    for (let c = 0; c < cols; c++) {
      const cell = sheet.get(rect.r1 + r, rect.c1 + c);
      row.push(cell ? { v: cell.v, f: cell.f, style: cell.s ?? 0 } : null);
      parts.push(tsvEscape(displayText(wb, cell)));
    }
    cells.push(row);
    lines.push(parts.join('\t'));
  }
  return { rows, cols, cells, sheetName: sheet.name, origin: { r: rect.r1, c: rect.c1 }, tsv: lines.join('\r\n') + '\r\n', cut, rect };
}

export function parseTSV(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let inQ = false;
  let i = 0;
  const n = text.length;
  let atFieldStart = true;
  while (i < n) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i += 2;
          continue;
        }
        inQ = false;
        i++;
        continue;
      }
      cur += ch;
      i++;
      continue;
    }
    if (ch === '"' && atFieldStart) {
      inQ = true;
      atFieldStart = false;
      i++;
      continue;
    }
    if (ch === '\t') {
      row.push(cur);
      cur = '';
      atFieldStart = true;
      i++;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
      atFieldStart = true;
      if (ch === '\r' && text[i + 1] === '\n') i++;
      i++;
      continue;
    }
    cur += ch;
    atFieldStart = false;
    i++;
  }
  if (cur !== '' || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

export type PasteMode = 'all' | 'values' | 'formulas' | 'formats' | 'transpose';

/** Paste internal clipboard data at (r, c). Returns the pasted rect. */
export function pasteClip(wb: Workbook, dst: Sheet, r: number, c: number, clip: ClipData, srcWb: Workbook, mode: PasteMode = 'all', src?: Sheet): Rect {
  const transpose = mode === 'transpose';
  const rows = transpose ? clip.cols : clip.rows;
  const cols = transpose ? clip.rows : clip.cols;
  const rect: Rect = { r1: r, c1: c, r2: Math.min(r + rows - 1, MAX_ROWS - 1), c2: Math.min(c + cols - 1, MAX_COLS - 1) };
  const dr = r - clip.origin.r;
  const dc = c - clip.origin.c;
  for (let i = 0; i < rows; i++)
    for (let j = 0; j < cols; j++) {
      const tr = r + i;
      const tc = c + j;
      if (tr >= MAX_ROWS || tc >= MAX_COLS) continue;
      const cc = transpose ? clip.cells[j][i] : clip.cells[i][j];
      const old = dst.get(tr, tc);
      if (!cc) {
        if (mode === 'formats') {
          wb.setCell(dst, tr, tc, old && (old.v !== null || old.f !== undefined) ? { ...old, s: undefined } : undefined);
        } else if (mode === 'all' || mode === 'transpose') wb.setCell(dst, tr, tc, undefined);
        else if (mode === 'values' || mode === 'formulas') {
          /* source empty: clear contents, keep format */
          if (old) wb.setCell(dst, tr, tc, old.s ? { v: null, s: old.s } : undefined);
        }
        continue;
      }
      let style = cc.style;
      if (srcWb !== wb && style) style = wb.styleIndex(srcWb.styleOf(style));
      let cell: Cell;
      if (mode === 'formats') {
        cell = { v: old?.v ?? null, f: old?.f, s: style || undefined, ast: old?.ast };
        if (cell.v === null && cell.f === undefined && !cell.s) {
          wb.setCell(dst, tr, tc, undefined);
          continue;
        }
      } else if (mode === 'values') {
        cell = { v: cc.v, s: old?.s };
      } else {
        const f = cc.f !== undefined ? (transpose ? cc.f : shiftFormula(cc.f, dr, dc)) : undefined;
        cell = { v: cc.v, s: mode === 'formulas' ? old?.s : style || undefined };
        if (f !== undefined) cell.f = f;
      }
      wb.setCell(dst, tr, tc, cell);
    }
  void src;
  return rect;
}

/** Paste plain text (TSV) at (r,c). */
export function pasteText(wb: Workbook, dst: Sheet, r: number, c: number, text: string): Rect {
  const grid = parseTSV(text);
  let maxCols = 1;
  for (const row of grid) maxCols = Math.max(maxCols, row.length);
  for (let i = 0; i < grid.length; i++)
    for (let j = 0; j < grid[i].length; j++) {
      const tr = r + i;
      const tc = c + j;
      if (tr >= MAX_ROWS || tc >= MAX_COLS) continue;
      try {
        applyInput(wb, dst, tr, tc, grid[i][j]);
      } catch {
        // a malformed formula pasted from outside becomes plain text
        const old = dst.get(tr, tc);
        wb.setCell(dst, tr, tc, { v: grid[i][j], s: old?.s });
      }
    }
  return { r1: r, c1: c, r2: Math.min(r + grid.length - 1, MAX_ROWS - 1), c2: Math.min(c + maxCols - 1, MAX_COLS - 1) };
}

/** Cut & paste: move cells and update every reference that pointed into the moved block. */
export function moveRange(wb: Workbook, srcSheet: Sheet, src: Rect, dstSheet: Sheet, r: number, c: number): Rect {
  const dr = r - src.r1;
  const dc = c - src.c1;
  const dst: Rect = { r1: r, c1: c, r2: src.r2 + dr, c2: src.c2 + dc };
  if (dst.r2 >= MAX_ROWS || dst.c2 >= MAX_COLS) return src;
  const grabbed: [number, number, Cell][] = [];
  eachCell(srcSheet, src, (rr, cc, cell) => grabbed.push([rr, cc, cell]));
  // clear source and destination
  const clearDst: [number, number][] = [];
  eachCell(dstSheet, dst, (rr, cc) => clearDst.push([rr, cc]));
  for (const [rr, cc] of clearDst) wb.setCell(dstSheet, rr, cc, undefined);
  for (const [rr, cc] of grabbed) wb.setCell(srcSheet, rr, cc, undefined);
  for (const [rr, cc, cell] of grabbed) wb.setCell(dstSheet, rr + dr, cc + dc, { v: cell.v, f: cell.f, s: cell.s });
  // fix references into the moved block, everywhere
  const sl = srcSheet.name.toLowerCase();
  for (const s of wb.sheets) {
    const edit: RefEditor = (sheetName, a, b, whole) => {
      const sh = (sheetName === undefined ? s.name : sheetName).toLowerCase();
      if (sh !== sl || whole) return { a, b };
      const inside = (p: RefPart) => p.r >= src.r1 && p.r <= src.r2 && p.c >= src.c1 && p.c <= src.c2;
      const mv = (p: RefPart): RefPart => ({ ...p, r: p.r + dr, c: p.c + dc });
      if (!b) return inside(a) ? { a: mv(a) } : { a, b };
      if (inside(a) && inside(b)) return { a: mv(a), b: mv(b) };
      return { a, b };
    };
    const updates: [number, Cell][] = [];
    for (const [k, cell] of s.cells) {
      if (cell.f === undefined) continue;
      const nf = rewriteRefs(cell.f, edit);
      if (nf !== cell.f) updates.push([k, { v: cell.v, f: nf, s: cell.s }]);
    }
    for (const [k, nc] of updates) wb.setCell(s, keyRow(k), keyCol(k), nc);
  }
  return dst;
}

// ---------------- fill ----------------

const DAY_ABBR = DAYS.map((d) => d.slice(0, 3));
const MONTH_ABBR = MONTHS.map((m) => m.slice(0, 3));

interface ListMatch {
  list: string[];
  idx: number;
  caps: 'upper' | 'lower' | 'title';
}
function matchList(s: string): ListMatch | null {
  const lists = [DAYS, DAY_ABBR, MONTHS, MONTH_ABBR];
  for (const list of lists) {
    const i = list.findIndex((x) => x.toLowerCase() === s.toLowerCase());
    if (i >= 0) {
      const caps = s === s.toUpperCase() && s.length > 1 ? 'upper' : s === s.toLowerCase() ? 'lower' : 'title';
      // "May" appears in both months lists: prefer the abbreviated list only for 3-letter inputs
      return { list, idx: i, caps };
    }
  }
  return null;
}
function applyCaps(s: string, caps: ListMatch['caps']): string {
  return caps === 'upper' ? s.toUpperCase() : caps === 'lower' ? s.toLowerCase() : s;
}

function trailingNumber(s: string): { prefix: string; num: number; width: number; suffix: string } | null {
  const m = /^(.*?)(\d+)(\D*)$/.exec(s);
  if (!m || m[1] === undefined) return null;
  return { prefix: m[1], num: parseInt(m[2], 10), width: m[2].length, suffix: m[3] };
}

/** Linear-regression step for a list of numbers (equal to the common step for arithmetic sequences). */
function trendStep(nums: number[]): number {
  const n = nums.length;
  if (n < 2) return 0;
  const mx = (n - 1) / 2;
  const my = nums.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let den = 0;
  for (let i = 0; i < n; i++) {
    num += (i - mx) * (nums[i] - my);
    den += (i - mx) * (i - mx);
  }
  return num / den;
}

/**
 * Compute the values for the cells continuing a source sequence.
 * `src` is in the direction of travel (reversed for up/left fills).
 * Returns the generator for destination index i (0-based).
 */
export function seriesGenerator(src: (Cell | undefined)[], isDate: boolean[]): ((i: number) => { v: Value; copyOf: number }) | null {
  const k = src.length;
  const allNum = src.every((c) => c && c.f === undefined && typeof c.v === 'number');
  if (allNum) {
    const nums = src.map((c) => c!.v as number);
    if (k >= 2) {
      const step = trendStep(nums);
      const mean = nums.reduce((a, b) => a + b, 0) / k;
      const mx = (k - 1) / 2;
      // continue the fitted line from the last point: for exact arithmetic sequences this equals last + step*(i+1)
      const exact = nums.every((x, i) => Math.abs(x - (nums[0] + step * i)) < 1e-9 * Math.max(1, Math.abs(x)));
      return (i) => ({ v: exact ? nums[k - 1] + step * (i + 1) : mean + step * (k + i - mx), copyOf: i % k });
    }
    // single number: copy (Excel behaviour)
    return (i) => ({ v: nums[0], copyOf: i % k });
  }
  if (k === 1 || src.every((c) => c && c.f === undefined && typeof c.v === 'string')) {
    const first = src[0];
    if (first && first.f === undefined && typeof first.v === 'string') {
      const strs = src.map((c) => c!.v as string);
      // month / weekday names
      const lm = strs.map(matchList);
      if (lm.every((m) => m && m.list === lm[0]!.list)) {
        const L = lm[0]!.list;
        const idxs = lm.map((m) => m!.idx);
        const step = k >= 2 ? idxs[1] - idxs[0] : 1;
        return (i) => ({ v: applyCaps(L[(((idxs[k - 1] + step * (i + 1)) % L.length) + L.length) % L.length], lm[0]!.caps), copyOf: i % k });
      }
      // text with trailing number
      const tn = strs.map(trailingNumber);
      if (tn.every((t) => t && t.prefix === tn[0]!.prefix && t.suffix === tn[0]!.suffix)) {
        const nums = tn.map((t) => t!.num);
        const step = k >= 2 ? nums[1] - nums[0] : 1;
        const w = tn[0]!.width;
        const pad = nums[0] >= 0 && tn[0]!.width > 1 && String(nums[0]).length < w;
        return (i) => {
          const n = nums[k - 1] + step * (i + 1);
          const ns = n < 0 ? String(n) : pad ? String(n).padStart(w, '0') : String(n);
          return { v: tn[0]!.prefix + ns + tn[0]!.suffix, copyOf: i % k };
        };
      }
    }
  }
  void isDate;
  return null;
}

/**
 * Fill from a source rect into a target extension.
 * dir: direction of fill. `count` = number of rows/cols to add beyond the source.
 */
export function fillRange(wb: Workbook, sheet: Sheet, src: Rect, dir: 'down' | 'up' | 'right' | 'left', count: number, mode: 'series' | 'copy' = 'series'): Rect | null {
  if (count <= 0) return null;
  const vertical = dir === 'down' || dir === 'up';
  const forward = dir === 'down' || dir === 'right';
  const lanes = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
  const k = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
  let dst: Rect;
  if (dir === 'down') dst = { r1: src.r2 + 1, c1: src.c1, r2: Math.min(src.r2 + count, MAX_ROWS - 1), c2: src.c2 };
  else if (dir === 'up') dst = { r1: Math.max(src.r1 - count, 0), c1: src.c1, r2: src.r1 - 1, c2: src.c2 };
  else if (dir === 'right') dst = { r1: src.r1, c1: src.c2 + 1, r2: src.r2, c2: Math.min(src.c2 + count, MAX_COLS - 1) };
  else dst = { r1: src.r1, c1: Math.max(src.c1 - count, 0), r2: src.r2, c2: src.c1 - 1 };
  if (dst.r2 < dst.r1 || dst.c2 < dst.c1) return null;
  const total = vertical ? dst.r2 - dst.r1 + 1 : dst.c2 - dst.c1 + 1;
  for (let lane = 0; lane < lanes; lane++) {
    // source cells in direction of travel
    const pos = (i: number) => (vertical ? { r: src.r1 + i, c: src.c1 + lane } : { r: src.r1 + lane, c: src.c1 + i });
    const order: number[] = [];
    for (let i = 0; i < k; i++) order.push(forward ? i : k - 1 - i);
    const cells = order.map((i) => {
      const p = pos(i);
      return sheet.get(p.r, p.c);
    });
    const gen = mode === 'series' ? seriesGenerator(cells, cells.map(() => false)) : null;
    for (let t = 0; t < total; t++) {
      const srcIdx = t % k;
      const sp = pos(order[srcIdx]);
      const dpos = vertical
        ? { r: forward ? dst.r1 + t : dst.r2 - t, c: src.c1 + lane }
        : { r: src.r1 + lane, c: forward ? dst.c1 + t : dst.c2 - t };
      const sc = cells[srcIdx];
      let cell: Cell | undefined;
      if (!sc) cell = undefined;
      else if (sc.f !== undefined) cell = { v: sc.v, f: shiftFormula(sc.f, dpos.r - sp.r, dpos.c - sp.c), s: sc.s };
      else {
        const g = gen ? gen(t) : null;
        cell = { v: g ? g.v : sc.v, s: sc.s };
      }
      wb.setCell(sheet, dpos.r, dpos.c, cell);
    }
  }
  return dst;
}

// ---------------- sort ----------------

export interface SortKey {
  col: number; // absolute column
  asc: boolean;
}

function sortRank(v: Value): number {
  if (v === null) return 5;
  if (typeof v === 'number') return 0;
  if (typeof v === 'string') return 1;
  if (typeof v === 'boolean') return 2;
  return 3;
}

export function sortRange(wb: Workbook, sheet: Sheet, rect: Rect, keys: SortKey[], hasHeader: boolean): void {
  const r1 = hasHeader ? rect.r1 + 1 : rect.r1;
  if (r1 >= rect.r2) return;
  const rows: { orig: number; cells: (Cell | undefined)[] }[] = [];
  for (let r = r1; r <= rect.r2; r++) {
    const cells: (Cell | undefined)[] = [];
    for (let c = rect.c1; c <= rect.c2; c++) cells.push(sheet.get(r, c));
    rows.push({ orig: r, cells });
  }
  const valAt = (row: { cells: (Cell | undefined)[] }, col: number): Value => row.cells[col - rect.c1]?.v ?? null;
  const sorted = rows.slice().sort((a, b) => {
    for (const k of keys) {
      const va = valAt(a, k.col);
      const vb = valAt(b, k.col);
      const ra = sortRank(va);
      const rb = sortRank(vb);
      if (ra === 5 || rb === 5) {
        if (ra !== rb) return ra === 5 ? 1 : -1; // blanks always last
        continue;
      }
      let cmp: number;
      if (ra !== rb) cmp = ra - rb;
      else cmp = compareValues(va, vb);
      if (cmp !== 0) return k.asc ? cmp : -cmp;
    }
    return a.orig - b.orig;
  });
  wb.transact('Sort', () => {
    sorted.forEach((row, i) => {
      const tr = r1 + i;
      for (let j = 0; j < row.cells.length; j++) {
        const cell = row.cells[j];
        const tc = rect.c1 + j;
        let nc: Cell | undefined = cell;
        if (cell && cell.f !== undefined && row.orig !== tr) nc = { v: cell.v, f: shiftFormula(cell.f, tr - row.orig, 0), s: cell.s };
        else if (cell) nc = { v: cell.v, f: cell.f, s: cell.s };
        wb.setCell(sheet, tr, tc, nc);
      }
    });
  });
}

// ---------------- filter ----------------

export function filterValueText(wb: Workbook, cell: Cell | undefined): string {
  return displayText(wb, cell);
}

export function applyFilter(wb: Workbook, sheet: Sheet): void {
  const f = sheet.filter;
  if (!f) return;
  wb.transact('Filter', () => {
    for (let r = f.rect.r1 + 1; r <= f.rect.r2; r++) {
      let ok = true;
      for (const [col, allowed] of f.cols) {
        if (!allowed.has(filterValueText(wb, sheet.get(r, col)))) {
          ok = false;
          break;
        }
      }
      const hidden = sheet.rows.sizes.get(r) === 0;
      if (!ok && !hidden) sheet.rows.set(r, 0);
      else if (ok && hidden) sheet.rows.set(r, sheet.rows.def);
    }
  }, true);
}

export function clearFilter(wb: Workbook, sheet: Sheet): void {
  if (!sheet.filter) return;
  sheet.filter.cols.clear();
  applyFilter(wb, sheet);
}

export function removeFilter(wb: Workbook, sheet: Sheet): void {
  if (!sheet.filter) return;
  wb.transact('Remove filter', () => {
    const f = sheet.filter!;
    for (let r = f.rect.r1 + 1; r <= f.rect.r2; r++) if (sheet.rows.sizes.get(r) === 0) sheet.rows.set(r, sheet.rows.def);
    sheet.filter = undefined;
  }, true);
}

/** Detect the contiguous data region around a cell (like Ctrl+A / Ctrl+T / sort). */
export function currentRegion(sheet: Sheet, r: number, c: number): Rect {
  const has = (rr: number, cc: number) => {
    if (rr < 0 || cc < 0 || rr >= MAX_ROWS || cc >= MAX_COLS) return false;
    const cell = sheet.get(rr, cc);
    return !!cell && (cell.v !== null || cell.f !== undefined);
  };
  let r1 = r, r2 = r, c1 = c, c2 = c;
  let grew = true;
  const rowHas = (rr: number) => {
    for (let cc = c1 - 1; cc <= c2 + 1; cc++) if (has(rr, cc)) return true;
    return false;
  };
  const colHas = (cc: number) => {
    for (let rr = r1 - 1; rr <= r2 + 1; rr++) if (has(rr, cc)) return true;
    return false;
  };
  let guard = 0;
  while (grew && guard++ < 100000) {
    grew = false;
    if (r1 > 0 && rowHas(r1 - 1)) { r1--; grew = true; }
    if (r2 < MAX_ROWS - 1 && rowHas(r2 + 1)) { r2++; grew = true; }
    if (c1 > 0 && colHas(c1 - 1)) { c1--; grew = true; }
    if (c2 < MAX_COLS - 1 && colHas(c2 + 1)) { c2++; grew = true; }
  }
  return { r1, c1, r2, c2 };
}

// ---------------- find & replace ----------------

export interface FindOpts {
  what: string;
  matchCase: boolean;
  wholeCell: boolean;
  inFormulas: boolean; // look in formulas, else in displayed values
  allSheets: boolean;
  byRows: boolean;
}

export interface FoundCell {
  sheet: Sheet;
  r: number;
  c: number;
}

export function cellMatches(wb: Workbook, cell: Cell, o: FindOpts): boolean {
  const text = o.inFormulas ? (cell.f !== undefined ? '=' + cell.f : displayText(wb, cell)) : displayText(wb, cell);
  const a = o.matchCase ? text : text.toLowerCase();
  const b = o.matchCase ? o.what : o.what.toLowerCase();
  if (o.wholeCell) return a === b;
  // wildcards
  if (/[*?]/.test(b)) {
    const re = new RegExp((o.wholeCell ? '^' : '') + b.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + (o.wholeCell ? '$' : ''), o.matchCase ? '' : 'i');
    return re.test(text);
  }
  return a.includes(b);
}

export function findAll(wb: Workbook, active: Sheet, o: FindOpts): FoundCell[] {
  const out: FoundCell[] = [];
  const sheets = o.allSheets ? wb.sheets : [active];
  for (const sh of sheets) {
    const hits: [number, number][] = [];
    for (const [k, cell] of sh.cells) {
      if (cell.v === null && cell.f === undefined) continue;
      if (cellMatches(wb, cell, o)) hits.push([keyRow(k), keyCol(k)]);
    }
    hits.sort(o.byRows ? (a, b) => a[0] - b[0] || a[1] - b[1] : (a, b) => a[1] - b[1] || a[0] - b[0]);
    for (const [r, c] of hits) out.push({ sheet: sh, r, c });
  }
  return out;
}

export function replaceInCell(wb: Workbook, sheet: Sheet, r: number, c: number, o: FindOpts, repl: string): boolean {
  const cell = sheet.get(r, c);
  if (!cell) return false;
  const flags = o.matchCase ? 'g' : 'gi';
  const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  let text = cell.f !== undefined ? '=' + cell.f : displayText(wb, cell);
  const re = new RegExp(o.wholeCell ? '^' + esc(o.what) + '$' : esc(o.what), flags);
  if (!re.test(text)) return false;
  text = text.replace(new RegExp(re.source, flags), () => repl);
  try {
    applyInput(wb, sheet, r, c, text);
  } catch {
    return false;
  }
  return true;
}

export type { CFRule, FErr };
export { isErr, normRect };
