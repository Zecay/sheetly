// Canvas-based virtualized grid: rendering, selection, keyboard and mouse interaction, scrolling, editing.
import { MAX_COLS, MAX_ROWS, Rect, a1, cellKey, colName, normRect, rectText } from '../core/addr';
import { CFEngine } from '../core/cf';
import { formatValue, FILL_MARK, generalFormat } from '../core/numfmt';
import { Cell, Sheet, Workbook } from '../core/model';
import { FErr, Value } from '../core/values';
import { currentRegion } from '../core/ops';
import { CellEditor, EditorHost, RefSpan, refSpans } from './editor';

export interface Theme {
  bg: string;
  grid: string;
  hdrBg: string;
  hdrText: string;
  hdrLine: string;
  hdrSelBg: string;
  hdrSelText: string;
  hdrAllSelBg: string;
  text: string;
  accent: string;
  selFill: string;
  freeze: string;
}

export const LIGHT: Theme = {
  bg: '#ffffff', grid: '#d4d4d4', hdrBg: '#f5f5f5', hdrText: '#444444', hdrLine: '#c8c8c8',
  hdrSelBg: '#d3f0e0', hdrSelText: '#0e5c2f', hdrAllSelBg: '#a8d7bc', text: '#000000', accent: '#217346',
  selFill: 'rgba(33,115,70,0.13)', freeze: '#8c8c8c',
};
export const DARK: Theme = {
  bg: '#1f1f1f', grid: '#3a3a3a', hdrBg: '#2b2b2b', hdrText: '#c8c8c8', hdrLine: '#454545',
  hdrSelBg: '#27493a', hdrSelText: '#a6e3bf', hdrAllSelBg: '#2f6b4b', text: '#e8e8e8', accent: '#3fb57a',
  selFill: 'rgba(63,181,122,0.18)', freeze: '#7a7a7a',
};

export interface GridHost {
  wb: Workbook;
  cf: CFEngine;
  dark: boolean;
  gridlines: boolean;
  headings: boolean;
  showFormulas: boolean;
  selectionChanged(): void;
  viewChanged(): void;
  /** Return an error message to keep editing, or null on success. */
  commitCell(r: number, c: number, text: string): string | null;
  contextMenu(x: number, y: number, kind: 'cell' | 'col' | 'row' | 'all'): void;
  fillDone(src: Rect, dir: 'down' | 'up' | 'right' | 'left', count: number): void;
  moveDone(src: Rect, r: number, c: number, copy: boolean): void;
  filterButton(col: number, anchor: DOMRect): void;
  definedNames(): string[];
  clipboardRect(): { sheet: Sheet; rect: Rect } | null;
  painterActive(): boolean;
  painterApply(rect: Rect): void;
  editStateChanged(): void;
  editorFont(r: number, c: number): { css: string };
  editorText(r: number, c: number): string;
  message(text: string): void;
  clearSelection(): void;
}

interface Seg {
  i: number;
  p: number; // screen position (px)
  s: number; // size (px)
}

const ROW_HDR_MIN = 34;
const HDR_H = 20;
const DEFAULT_FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';

type Drag =
  | { kind: 'select'; anchor: { r: number; c: number }; ctrl: boolean }
  | { kind: 'col'; anchor: number; ctrl: boolean }
  | { kind: 'row'; anchor: number; ctrl: boolean }
  | { kind: 'colResize'; c: number; startX: number; startW: number; sel: number[] }
  | { kind: 'rowResize'; r: number; startY: number; startH: number; sel: number[] }
  | { kind: 'fill'; src: Rect }
  | { kind: 'move'; src: Rect; grab: { r: number; c: number } }
  | { kind: 'point'; anchor: { r: number; c: number } };

export class Grid {
  sheet: Sheet;
  zoom = 1;
  theme: Theme = LIGHT;
  ranges: Rect[] = [{ r1: 0, c1: 0, r2: 0, c2: 0 }];
  active = { r: 0, c: 0 };
  anchor = { r: 0, c: 0 };
  focus = { r: 0, c: 0 };
  editor: CellEditor;
  canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private dpr = 1;
  private W = 0;
  private H = 0;
  private raf = 0;
  private drag: Drag | null = null;
  private dragPreview: Rect | null = null;
  private mouse = { x: 0, y: 0 };
  private autoScroll = 0;
  private textCache = new Map<string, number>();
  private fontCache = new Map<number, string>();
  private mergeMap = new Map<number, Rect>();
  private mergeStamp = '';
  private dispCache = new WeakMap<Cell, { v: Value; fmt: string; res: { text: string; color?: string; isNumber: boolean } }>();
  private vthumb: HTMLElement;
  private hthumb: HTMLElement;
  private selAllNext = false;
  cursor = 'cell';
  private hostEditorRect: EditorHost;

  constructor(
    public host: GridHost,
    public area: HTMLElement,
    private vscroll: HTMLElement,
    private hscroll: HTMLElement,
    sheet: Sheet,
  ) {
    this.sheet = sheet;
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'grid-canvas';
    this.area.append(this.canvas);
    this.ctx = this.canvas.getContext('2d', { alpha: false })!;
    this.vthumb = vscroll.querySelector('.thumb') as HTMLElement;
    this.hthumb = hscroll.querySelector('.thumb') as HTMLElement;
    this.hostEditorRect = {
      onEditChange: () => {
        this.invalidate();
        this.host.editStateChanged();
      },
      definedNames: () => this.host.definedNames(),
      anchorRect: () => this.editorAnchor(),
    };
    this.editor = new CellEditor(this.hostEditorRect, this.area);
    this.bind();
    const ro = new ResizeObserver(() => this.resize());
    ro.observe(this.area);
    this.resize();
  }

  // ------------------------------------------------------------------ geometry
  private hwCache = { key: '', v: 0 };
  get hw(): number {
    if (!this.host.headings) return 0;
    const key = `${this.sheet.id}|${this.sheet.scrollY}|${this.zoom}|${this.H}|${this.sheet.freezeRows}`;
    if (this.hwCache.key === key) return this.hwCache.v;
    const last = Math.min(MAX_ROWS, this.rowAt(this.H) + 2);
    const digits = String(last).length;
    const v = Math.max(ROW_HDR_MIN, digits * 8 + 14) * Math.max(0.5, Math.min(this.zoom, 2));
    this.hwCache = { key, v };
    return v;
  }
  get hh(): number {
    return this.host.headings ? HDR_H * Math.max(0.5, Math.min(this.zoom, 2)) : 0;
  }
  private get fc(): number {
    return this.sheet.freezeCols;
  }
  private get fr(): number {
    return this.sheet.freezeRows;
  }
  get fW(): number {
    return this.fc ? this.sheet.cols.top(this.fc) * this.zoom : 0;
  }
  get fH(): number {
    return this.fr ? this.sheet.rows.top(this.fr) * this.zoom : 0;
  }
  colX(c: number): number {
    const s = this.sheet;
    if (c < this.fc) return this.hw + s.cols.top(c) * this.zoom;
    return this.hw + this.fW + (s.cols.top(c) - s.cols.top(this.fc) - s.scrollX) * this.zoom;
  }
  rowY(r: number): number {
    const s = this.sheet;
    if (r < this.fr) return this.hh + s.rows.top(r) * this.zoom;
    return this.hh + this.fH + (s.rows.top(r) - s.rows.top(this.fr) - s.scrollY) * this.zoom;
  }
  colW(c: number): number {
    return this.sheet.cols.size(c) * this.zoom;
  }
  rowH(r: number): number {
    return this.sheet.rows.size(r) * this.zoom;
  }
  /** Column under screen x (clamped); frozen-aware. */
  colAt(x: number): number {
    const s = this.sheet;
    if (x < this.hw + this.fW && this.fc) return Math.min(this.fc - 1, s.cols.at((x - this.hw) / this.zoom, MAX_COLS));
    const pos = (x - this.hw - this.fW) / this.zoom + s.cols.top(this.fc) + s.scrollX;
    return Math.min(MAX_COLS - 1, s.cols.at(pos, MAX_COLS));
  }
  rowAt(y: number): number {
    const s = this.sheet;
    if (y < this.hh + this.fH && this.fr) return Math.min(this.fr - 1, s.rows.at((y - this.hh) / this.zoom, MAX_ROWS));
    const pos = (y - this.hh - this.fH) / this.zoom + s.rows.top(this.fr) + s.scrollY;
    return Math.min(MAX_ROWS - 1, s.rows.at(pos, MAX_ROWS));
  }

  /** Rectangle of a cell (or its merged area) in area coordinates. */
  cellRect(r: number, c: number): { x: number; y: number; w: number; h: number } {
    const m = this.mergeAt(r, c);
    if (m) {
      const x = this.colX(m.c1);
      const y = this.rowY(m.r1);
      const x2 = m.c2 + 1 < MAX_COLS ? this.colX(m.c2 + 1) : this.colX(m.c2) + this.colW(m.c2);
      const y2 = m.r2 + 1 < MAX_ROWS ? this.rowY(m.r2 + 1) : this.rowY(m.r2) + this.rowH(m.r2);
      return { x, y, w: x2 - x, h: y2 - y };
    }
    return { x: this.colX(c), y: this.rowY(r), w: this.colW(c), h: this.rowH(r) };
  }
  rectPx(r: Rect): { x: number; y: number; w: number; h: number } {
    const x = this.colX(r.c1);
    const y = this.rowY(r.r1);
    let x2 = r.c2 + 1 < MAX_COLS ? this.colX(r.c2 + 1) : x + this.sheet.cols.top(MAX_COLS) * this.zoom;
    let y2 = r.r2 + 1 < MAX_ROWS ? this.rowY(r.r2 + 1) : y + 1e7;
    if (r.c2 === MAX_COLS - 1) x2 = this.W + 20;
    if (r.r2 === MAX_ROWS - 1) y2 = this.H + 20;
    return { x, y, w: x2 - x, h: y2 - y };
  }

  private mergeAt(r: number, c: number): Rect | undefined {
    this.ensureMergeMap();
    return this.mergeMap.get(cellKey(r, c));
  }
  private ensureMergeMap() {
    const stamp = this.sheet.id + ':' + this.sheet.merges.length + ':' + this.sheet.merges.map((m) => m.r1 + ',' + m.c1 + ',' + m.r2 + ',' + m.c2).join(';');
    if (stamp === this.mergeStamp) return;
    this.mergeStamp = stamp;
    this.mergeMap.clear();
    let budget = 300000;
    for (const m of this.sheet.merges) {
      for (let r = m.r1; r <= m.r2 && budget > 0; r++)
        for (let c = m.c1; c <= m.c2 && budget > 0; c++) {
          this.mergeMap.set(cellKey(r, c), m);
          budget--;
        }
    }
  }
  expandForMerges(rect: Rect): Rect {
    if (!this.sheet.merges.length) return rect;
    let r = { ...rect };
    let changed = true;
    let guard = 0;
    while (changed && guard++ < 20) {
      changed = false;
      for (const m of this.sheet.merges) {
        if (m.r1 <= r.r2 && m.r2 >= r.r1 && m.c1 <= r.c2 && m.c2 >= r.c1) {
          const n = { r1: Math.min(r.r1, m.r1), c1: Math.min(r.c1, m.c1), r2: Math.max(r.r2, m.r2), c2: Math.max(r.c2, m.c2) };
          if (n.r1 !== r.r1 || n.c1 !== r.c1 || n.r2 !== r.r2 || n.c2 !== r.c2) {
            r = n;
            changed = true;
          }
        }
      }
    }
    return r;
  }

  // ------------------------------------------------------------------ public state
  get range(): Rect {
    return this.ranges[this.ranges.length - 1];
  }

  setSheet(sheet: Sheet): void {
    this.cancelEdit();
    this.sheet = sheet;
    this.ranges = [{ r1: sheet.sel.r, c1: sheet.sel.c, r2: sheet.sel.r2, c2: sheet.sel.c2 }];
    this.active = { r: sheet.sel.r, c: sheet.sel.c };
    this.anchor = { ...this.active };
    this.focus = { r: sheet.sel.r2, c: sheet.sel.c2 };
    this.mergeStamp = '';
    this.invalidate();
    this.host.selectionChanged();
    this.host.viewChanged();
  }

  select(r: number, c: number, r2 = r, c2 = c, keepActive = false): void {
    r = clamp(r, 0, MAX_ROWS - 1);
    c = clamp(c, 0, MAX_COLS - 1);
    r2 = clamp(r2, 0, MAX_ROWS - 1);
    c2 = clamp(c2, 0, MAX_COLS - 1);
    if (!keepActive) {
      this.active = { r, c };
      this.anchor = { r, c };
    }
    this.focus = { r: r2, c: c2 };
    this.ranges = [this.expandForMerges(normRect(this.anchor.r, this.anchor.c, r2, c2))];
    this.saveSel();
    this.invalidate();
    this.host.selectionChanged();
  }

  private saveSel() {
    const rg = this.range;
    this.sheet.sel = { r: this.active.r, c: this.active.c, r2: this.focus.r, c2: this.focus.c };
    void rg;
  }

  selectRect(rect: Rect): void {
    this.active = { r: rect.r1, c: rect.c1 };
    this.anchor = { r: rect.r1, c: rect.c1 };
    this.focus = { r: rect.r2, c: rect.c2 };
    this.ranges = [rect];
    this.saveSel();
    this.invalidate();
    this.host.selectionChanged();
  }

  /** All selected cells as a single bounding rect plus count helper. */
  selectionRects(): Rect[] {
    return this.ranges;
  }

  invalidate(): void {
    if (this.raf) return;
    this.raf = requestAnimationFrame(() => {
      this.raf = 0;
      this.render();
    });
  }

  setZoom(z: number): void {
    z = clamp(Math.round(z * 100) / 100, 0.1, 4);
    if (z === this.zoom) return;
    this.zoom = z;
    this.clampScroll();
    this.invalidate();
    this.host.viewChanged();
    if (this.editor.active) this.editor.layout();
  }

  focusGrid(): void {
    if (this.editor.active) return;
    (this.area as HTMLElement).focus({ preventScroll: true });
  }

  // ------------------------------------------------------------------ scrolling
  private viewW(): number {
    return Math.max(0, this.W - this.hw - this.fW);
  }
  private viewH(): number {
    return Math.max(0, this.H - this.hh - this.fH);
  }
  private maxScrollX(): number {
    const s = this.sheet;
    return Math.max(0, s.cols.total(MAX_COLS) - s.cols.top(this.fc) - this.viewW() / this.zoom);
  }
  private maxScrollY(): number {
    const s = this.sheet;
    return Math.max(0, s.rows.total(MAX_ROWS) - s.rows.top(this.fr) - this.viewH() / this.zoom);
  }
  clampScroll(): void {
    const s = this.sheet;
    s.scrollX = clamp(s.scrollX, 0, this.maxScrollX());
    s.scrollY = clamp(s.scrollY, 0, this.maxScrollY());
  }
  scrollBy(dx: number, dy: number): void {
    const s = this.sheet;
    const ox = s.scrollX;
    const oy = s.scrollY;
    s.scrollX += dx / this.zoom;
    s.scrollY += dy / this.zoom;
    this.clampScroll();
    if (ox !== s.scrollX || oy !== s.scrollY) {
      this.invalidate();
      if (this.editor.active) this.editor.layout();
    }
  }
  scrollTo(x: number, y: number): void {
    this.sheet.scrollX = x;
    this.sheet.scrollY = y;
    this.clampScroll();
    this.invalidate();
  }
  /** Bring a cell into view (keeps frozen rows/cols in mind). */
  scrollIntoView(r: number, c: number): void {
    const s = this.sheet;
    const m = this.mergeAt(r, c);
    if (m) {
      r = m.r1;
      c = m.c1;
    }
    if (c >= this.fc) {
      const left = s.cols.top(c) - s.cols.top(this.fc);
      const right = s.cols.top(c + 1) - s.cols.top(this.fc);
      const vw = this.viewW() / this.zoom;
      if (left < s.scrollX) s.scrollX = left;
      else if (right > s.scrollX + vw) s.scrollX = right - vw;
    }
    if (r >= this.fr) {
      const top = s.rows.top(r) - s.rows.top(this.fr);
      const bottom = s.rows.top(r + 1) - s.rows.top(this.fr);
      const vh = this.viewH() / this.zoom;
      if (top < s.scrollY) s.scrollY = top;
      else if (bottom > s.scrollY + vh) s.scrollY = bottom - vh;
    }
    this.clampScroll();
    this.invalidate();
  }

  private updateScrollbars(): void {
    const s = this.sheet;
    const vh = this.viewH() / this.zoom;
    const vw = this.viewW() / this.zoom;
    const used = s.usedRect();
    const usedBottom = used ? s.rows.top(used.r2 + 1) - s.rows.top(this.fr) : 0;
    const usedRight = used ? s.cols.top(used.c2 + 1) - s.cols.top(this.fc) : 0;
    const extentY = Math.min(s.rows.total(MAX_ROWS), Math.max(usedBottom + vh, s.scrollY + vh * 2, vh * 2));
    const extentX = Math.min(s.cols.total(MAX_COLS), Math.max(usedRight + vw, s.scrollX + vw * 2, vw * 2));
    const vTrack = this.vscroll.clientHeight - 2;
    const hTrack = this.hscroll.clientWidth - 2;
    const vSize = Math.max(24, Math.min(vTrack, (vh / extentY) * vTrack));
    const hSize = Math.max(24, Math.min(hTrack, (vw / extentX) * hTrack));
    const vPos = extentY - vh > 0 ? (s.scrollY / (extentY - vh)) * (vTrack - vSize) : 0;
    const hPos = extentX - vw > 0 ? (s.scrollX / (extentX - vw)) * (hTrack - hSize) : 0;
    this.vthumb.style.height = vSize + 'px';
    this.vthumb.style.top = Math.max(0, vPos) + 1 + 'px';
    this.hthumb.style.width = hSize + 'px';
    this.hthumb.style.left = Math.max(0, hPos) + 1 + 'px';
    (this as any)._ext = { extentX, extentY, vSize, hSize, vTrack, hTrack, vh, vw };
  }

  // ------------------------------------------------------------------ fonts & text
  private fontFor(styleIdx: number | undefined): string {
    const key = (styleIdx ?? 0) * 1000 + Math.round(this.zoom * 100);
    let f = this.fontCache.get(key);
    if (!f) {
      const st = this.host.wb.styleOf(styleIdx);
      const px = (st.size ?? 11) * (4 / 3) * this.zoom;
      f = `${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}${px}px ${st.font ? `"${st.font}", ` : ''}${DEFAULT_FONT}`;
      if (this.fontCache.size > 2000) this.fontCache.clear();
      this.fontCache.set(key, f);
    }
    return f;
  }
  private textCacheChars = 0;
  private measure(font: string, text: string): number {
    if (text.length > 500) {
      // long text is not cached (keeps the cache small)
      this.ctx.font = font;
      return this.ctx.measureText(text).width;
    }
    const k = font + '|' + text;
    let w = this.textCache.get(k);
    if (w === undefined) {
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      if (this.textCache.size > 6000 || this.textCacheChars > 1_000_000) {
        this.textCache.clear();
        this.textCacheChars = 0;
      }
      this.textCache.set(k, w);
      this.textCacheChars += k.length;
    }
    return w;
  }
  private wrapCache = new Map<string, string[]>();
  private wrapLines(font: string, text: string, maxW: number): string[] {
    if (text.length > 2000) text = text.slice(0, 2000); // only the visible start of huge cells is laid out
    const key = font + '|' + Math.round(maxW) + '|' + text;
    const hit = this.wrapCache.get(key);
    if (hit) return hit;
    const res = this.wrapLinesUncached(font, text, maxW);
    if (this.wrapCache.size > 1500) this.wrapCache.clear();
    this.wrapCache.set(key, res);
    return res;
  }
  private wrapLinesUncached(font: string, text: string, maxW: number): string[] {
    const out: string[] = [];
    for (const para of text.split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        if (w === '') continue;
        const t = line + w;
        if (this.measure(font, t.trimEnd()) <= maxW || line === '') {
          line = t;
          // very long single word: break by characters
          if (this.measure(font, line) > maxW && !/\s/.test(line)) {
            let cur = '';
            for (const ch of line) {
              if (this.measure(font, cur + ch) > maxW && cur) {
                out.push(cur);
                cur = ch;
              } else cur += ch;
            }
            line = cur;
          }
        } else {
          out.push(line.trimEnd());
          line = w.trimStart();
        }
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  private display(cell: Cell, fmt: string): { text: string; color?: string; isNumber: boolean } {
    if (this.host.showFormulas && cell.f !== undefined) return { text: '=' + cell.f, isNumber: false };
    const hit = this.dispCache.get(cell);
    if (hit && hit.v === cell.v && hit.fmt === fmt) return hit.res;
    const res = formatValue(cell.v, fmt || 'General');
    this.dispCache.set(cell, { v: cell.v, fmt, res });
    return res;
  }

  // ------------------------------------------------------------------ resize & render
  resize(): void {
    const r = this.area.getBoundingClientRect();
    this.dpr = window.devicePixelRatio || 1;
    this.W = Math.max(1, Math.floor(r.width));
    this.H = Math.max(1, Math.floor(r.height));
    this.canvas.width = Math.floor(this.W * this.dpr);
    this.canvas.height = Math.floor(this.H * this.dpr);
    this.canvas.style.width = this.W + 'px';
    this.canvas.style.height = this.H + 'px';
    this.clampScroll();
    this.invalidate();
    this.host.viewChanged();
  }

  private px(v: number): number {
    return (Math.round(v * this.dpr) + 0.5) / this.dpr;
  }

  private listCols(): { frozen: Seg[]; scroll: Seg[] } {
    const s = this.sheet;
    const frozen: Seg[] = [];
    for (let c = 0; c < this.fc; c++) {
      const w = this.colW(c);
      if (w > 0) frozen.push({ i: c, p: this.colX(c), s: w });
    }
    const scroll: Seg[] = [];
    const start = Math.max(this.fc, s.cols.at(s.cols.top(this.fc) + s.scrollX, MAX_COLS));
    let x = this.colX(start);
    for (let c = start; c < MAX_COLS && x < this.W; c++) {
      const w = this.colW(c);
      if (w > 0) scroll.push({ i: c, p: x, s: w });
      x += w;
    }
    return { frozen, scroll };
  }
  private listRows(): { frozen: Seg[]; scroll: Seg[] } {
    const s = this.sheet;
    const frozen: Seg[] = [];
    for (let r = 0; r < this.fr; r++) {
      const h = this.rowH(r);
      if (h > 0) frozen.push({ i: r, p: this.rowY(r), s: h });
    }
    const scroll: Seg[] = [];
    const start = Math.max(this.fr, s.rows.at(s.rows.top(this.fr) + s.scrollY, MAX_ROWS));
    let y = this.rowY(start);
    for (let r = start; r < MAX_ROWS && y < this.H; r++) {
      const h = this.rowH(r);
      if (h > 0) scroll.push({ i: r, p: y, s: h });
      y += h;
    }
    return { frozen, scroll };
  }

  render(): void {
    const ctx = this.ctx;
    const t = this.theme;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = t.bg;
    ctx.fillRect(0, 0, this.W, this.H);
    this.ensureMergeMap();
    const cols = this.listCols();
    const rows = this.listRows();
    const hw = this.hw;
    const hh = this.hh;
    const fW = this.fW;
    const fH = this.fH;
    const quads = [
      { c: cols.scroll, r: rows.scroll, x: hw + fW, y: hh + fH, w: this.W - hw - fW, h: this.H - hh - fH },
      { c: cols.frozen, r: rows.scroll, x: hw, y: hh + fH, w: fW, h: this.H - hh - fH },
      { c: cols.scroll, r: rows.frozen, x: hw + fW, y: hh, w: this.W - hw - fW, h: fH },
      { c: cols.frozen, r: rows.frozen, x: hw, y: hh, w: fW, h: fH },
    ];
    const refs = this.editor.active ? refSpans(this.editor.text) : [];
    for (const q of quads) {
      if (q.w <= 0 || q.h <= 0 || !q.c.length || !q.r.length) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(q.x, q.y, q.w, q.h);
      ctx.clip();
      this.drawCells(q.c, q.r, q);
      this.drawOverlays(refs);
      ctx.restore();
    }
    if (this.host.headings) this.drawHeaders(cols, rows);
    // frozen separators
    ctx.strokeStyle = t.freeze;
    ctx.lineWidth = 1;
    if (this.fc) {
      ctx.beginPath();
      ctx.moveTo(this.px(hw + fW), hh);
      ctx.lineTo(this.px(hw + fW), this.H);
      ctx.stroke();
    }
    if (this.fr) {
      ctx.beginPath();
      ctx.moveTo(hw, this.px(hh + fH));
      ctx.lineTo(this.W, this.px(hh + fH));
      ctx.stroke();
    }
    this.updateScrollbars();
  }

  private drawCells(cols: Seg[], rows: Seg[], q: { x: number; y: number; w: number; h: number }): void {
    const ctx = this.ctx;
    const t = this.theme;
    const s = this.sheet;
    const wb = this.host.wb;
    const z = this.zoom;
    const x0 = cols[0].p;
    const x1 = cols[cols.length - 1].p + cols[cols.length - 1].s;
    const y0 = rows[0].p;
    const y1 = rows[rows.length - 1].p + rows[rows.length - 1].s;
    // grid lines
    if (this.host.gridlines && s.showGrid) {
      ctx.strokeStyle = t.grid;
      ctx.lineWidth = 1 / this.dpr;
      ctx.beginPath();
      for (const c of cols) {
        const x = this.px(c.p + c.s);
        ctx.moveTo(x, y0);
        ctx.lineTo(x, y1);
      }
      for (const r of rows) {
        const y = this.px(r.p + r.s);
        ctx.moveTo(x0, y);
        ctx.lineTo(x1, y);
      }
      ctx.stroke();
    }
    const hasCF = s.cf.length > 0;
    const editing = this.editor.active ? this.editor : null;
    const painted: { r: number; c: number; x: number; y: number; w: number; h: number; cell: Cell; cfc?: string; cf?: { bold?: boolean; italic?: boolean } }[] = [];
    // pass 1: fills, bars and borders
    for (const row of rows) {
      for (const col of cols) {
        const key = cellKey(row.i, col.i);
        const cell = s.cells.get(key);
        const cf = hasCF ? this.host.cf.evaluate(s, row.i, col.i) : null;
        if (!cell && !cf) continue;
        if (this.mergeMap.has(key)) continue;
        const st = wb.styleOf(cell?.s);
        const fill = cf?.fill ?? st.fill;
        if (fill) {
          ctx.fillStyle = fill;
          ctx.fillRect(col.p, row.p, col.s + 1, row.s + 1);
        }
        if (cf?.bar) {
          ctx.fillStyle = cf.bar.color;
          ctx.globalAlpha = 0.75;
          ctx.fillRect(col.p + 2, row.p + 2, Math.max(1, (col.s - 4) * cf.bar.pct), row.s - 4);
          ctx.globalAlpha = 1;
        }
        if (st.bl || st.br || st.bt || st.bb) this.drawBorders(st, col.p, row.p, col.s, row.s);
        if (cell && (cell.v !== null || cell.f !== undefined)) painted.push({ r: row.i, c: col.i, x: col.p, y: row.p, w: col.s, h: row.s, cell, cfc: cf?.color, cf: cf ?? undefined });
      }
    }
    // merged cells intersecting this quadrant
    for (const m of s.merges) {
      const r = this.cellRect(m.r1, m.c1);
      if (r.x > q.x + q.w || r.x + r.w < q.x || r.y > q.y + q.h || r.y + r.h < q.y) continue;
      const cell = s.cells.get(cellKey(m.r1, m.c1));
      const st = wb.styleOf(cell?.s);
      const cf = hasCF ? this.host.cf.evaluate(s, m.r1, m.c1) : null;
      const fill = cf?.fill ?? st.fill;
      ctx.fillStyle = fill ?? t.bg;
      ctx.fillRect(r.x, r.y, r.w + 1, r.h + 1);
      if (this.host.gridlines && !fill) {
        /* merged area has no inner grid lines */
      }
      // Borders may live on the other cells of the merged area (bottom row / right column).
      const stB = wb.styleOf(s.cells.get(cellKey(m.r2, m.c1))?.s);
      const stR = wb.styleOf(s.cells.get(cellKey(m.r1, m.c2))?.s);
      const stBR = wb.styleOf(s.cells.get(cellKey(m.r2, m.c2))?.s);
      const bst = { ...st, bb: stB.bb ?? stBR.bb ?? st.bb, br: stR.br ?? stBR.br ?? st.br };
      if (bst.bl || bst.br || bst.bt || bst.bb) this.drawBorders(bst, r.x, r.y, r.w, r.h);
      if (cell && (cell.v !== null || cell.f !== undefined)) painted.push({ r: m.r1, c: m.c1, x: r.x, y: r.y, w: r.w, h: r.h, cell, cfc: cf?.color, cf: cf ?? undefined });
    }
    // pass 2: text
    for (const p of painted) {
      if (editing && editing.r === p.r && editing.c === p.c) continue;
      this.drawText(p.r, p.c, p.x, p.y, p.w, p.h, p.cell, p.cfc, p.cf);
    }
    // filter buttons
    if (s.filter) {
      const f = s.filter;
      for (const col of cols) {
        if (col.i < f.rect.c1 || col.i > f.rect.c2) continue;
        for (const row of rows) {
          if (row.i !== f.rect.r1) continue;
          this.drawFilterButton(col.p + col.s, row.p + row.s, !!f.cols.get(col.i));
        }
      }
    }
    void z;
  }

  private drawFilterButton(right: number, bottom: number, active: boolean): void {
    const ctx = this.ctx;
    const z = Math.max(0.7, Math.min(this.zoom, 1.5));
    const sz = 16 * z;
    const x = right - sz - 1;
    const y = bottom - sz - 1;
    ctx.fillStyle = this.theme === DARK ? '#3a3a3a' : '#f0f0f0';
    ctx.fillRect(x, y, sz, sz);
    ctx.strokeStyle = this.theme === DARK ? '#6a6a6a' : '#a0a0a0';
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, sz - 1, sz - 1);
    ctx.fillStyle = active ? this.theme.accent : this.theme.text;
    ctx.beginPath();
    ctx.moveTo(x + sz * 0.28, y + sz * 0.4);
    ctx.lineTo(x + sz * 0.72, y + sz * 0.4);
    ctx.lineTo(x + sz * 0.5, y + sz * 0.68);
    ctx.closePath();
    ctx.fill();
  }

  private drawBorders(st: import('../core/model').CellStyle, x: number, y: number, w: number, h: number): void {
    const ctx = this.ctx;
    const side = (b: import('../core/model').Border | undefined, x1: number, y1: number, x2: number, y2: number, horiz: boolean) => {
      if (!b) return;
      ctx.strokeStyle = b.color || this.theme.text;
      ctx.setLineDash([]);
      let lw = 1;
      if (b.style === 'medium') lw = 2;
      else if (b.style === 'thick') lw = 3;
      else if (b.style === 'dashed') ctx.setLineDash([4, 2]);
      else if (b.style === 'dotted') ctx.setLineDash([1, 1.5]);
      else if (b.style === 'hair') ctx.setLineDash([1, 1]);
      ctx.lineWidth = lw;
      const off = lw % 2 === 1 ? 0.5 : 0;
      const draw = (dx: number, dy: number) => {
        ctx.beginPath();
        ctx.moveTo(Math.round(x1) + (horiz ? 0 : off) + dx, Math.round(y1) + (horiz ? off : 0) + dy);
        ctx.lineTo(Math.round(x2) + (horiz ? 0 : off) + dx, Math.round(y2) + (horiz ? off : 0) + dy);
        ctx.stroke();
      };
      if (b.style === 'double') {
        ctx.lineWidth = 1;
        draw(horiz ? 0 : -1, horiz ? -1 : 0);
        draw(horiz ? 0 : 1, horiz ? 1 : 0);
      } else draw(0, 0);
      ctx.setLineDash([]);
    };
    side(st.bt, x, y, x + w, y, true);
    side(st.bb, x, y + h, x + w, y + h, true);
    side(st.bl, x, y, x, y + h, false);
    side(st.br, x + w, y, x + w, y + h, false);
  }

  private drawText(r: number, c: number, x: number, y: number, w: number, h: number, cell: Cell, cfColor: string | undefined, cf?: { bold?: boolean; italic?: boolean }): void {
    const ctx = this.ctx;
    const wb = this.host.wb;
    const st = wb.styleOf(cell.s);
    const z = this.zoom;
    const disp = this.display(cell, st.fmt ?? 'General');
    let text = disp.text;
    if (text === '') return;
    let font = this.fontFor(cell.s);
    if (cf?.bold && !st.bold) font = 'bold ' + font;
    if (cf?.italic && !st.italic) font = 'italic ' + font;
    const v = cell.v;
    const ha = st.h ?? (typeof v === 'number' ? 'right' : typeof v === 'boolean' || v instanceof FErr ? 'center' : 'left');
    const fontPx = (st.size ?? 11) * (4 / 3) * z;
    const pad = 3 * z;
    const indent = (st.indent ?? 0) * 9 * z;
    ctx.font = font;
    ctx.fillStyle = cfColor ?? disp.color ?? st.color ?? this.theme.text;
    ctx.textBaseline = 'alphabetic';
    const pos = st.v ?? 'bottom';
    const lineH = fontPx * 1.2;
    const baseFor = (lineTop: number) => lineTop + fontPx * 0.95;

    // accounting-style fill split
    const hasFill = text.includes(FILL_MARK);
    let left = '';
    let right = '';
    if (hasFill) {
      const i = text.indexOf(FILL_MARK);
      left = text.slice(0, i);
      right = text.slice(i + 1).split(FILL_MARK).join('');
      text = left + right;
    }

    if (st.wrap && !hasFill && typeof v !== 'number') {
      const maxW = Math.max(4, w - 2 * pad - indent);
      const lines = this.wrapLines(font, text, maxW);
      const totalH = lines.length * lineH;
      let top = pos === 'top' ? y + 1 : pos === 'middle' ? y + (h - totalH) / 2 : y + h - totalH - 1;
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      ctx.clip();
      for (const ln of lines) {
        const tw = this.measure(font, ln);
        const tx = ha === 'right' ? x + w - pad - tw : ha === 'center' ? x + (w - tw) / 2 : x + pad + indent;
        ctx.fillText(ln, tx, baseFor(top));
        this.decorate(st, tx, baseFor(top), tw, fontPx);
        top += lineH;
      }
      ctx.restore();
      return;
    }
    if (st.wrap && typeof v === 'number') {
      /* numbers never wrap */
    }

    let tw = this.measure(font, text);
    const avail = w - 2 * pad - indent;
    if (disp.isNumber && tw > avail + 0.5) {
      // try to squeeze General numbers, otherwise ####
      let shown: string | null = null;
      if ((!st.fmt || st.fmt === 'General') && typeof v === 'number') {
        for (let p = 9; p >= 1; p--) {
          const alt = Number(v.toPrecision(p));
          const s2 = Math.abs(alt) >= 1e11 || (Math.abs(alt) < 1e-4 && alt !== 0) ? alt.toExponential(Math.max(0, p - 1)).replace(/e\+?/, 'E+').replace('E+-', 'E-') : generalFormat(alt);
          if (this.measure(font, s2) <= avail) {
            shown = s2;
            break;
          }
        }
      }
      if (shown) text = shown;
      else {
        const hashW = this.measure(font, '#');
        text = '#'.repeat(Math.max(1, Math.floor(avail / hashW)));
      }
      tw = this.measure(font, text);
    }
    const baseY = baseFor(pos === 'top' ? y + 1 : pos === 'middle' ? y + (h - lineH) / 2 : y + h - lineH - 1);

    if (hasFill) {
      const lw = this.measure(font, left);
      const rw = this.measure(font, right);
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      ctx.clip();
      ctx.fillText(left, x + pad, baseY);
      ctx.fillText(right, x + w - pad - rw, baseY);
      this.decorate(st, x + pad, baseY, lw, fontPx);
      ctx.restore();
      return;
    }

    let tx: number;
    let clipX = x;
    let clipW = w;
    if (ha === 'right') tx = x + w - pad - tw;
    else if (ha === 'center') tx = x + (w - tw) / 2;
    else tx = x + pad + indent;
    // text overflow into empty neighbours
    if (typeof v === 'string' && tw > avail + 0.5 && !this.mergeAt(r, c)) {
      const s = this.sheet;
      const free = (cc: number, rr: number) => {
        if (cc < 0 || cc >= MAX_COLS) return false;
        const n = s.cells.get(cellKey(rr, cc));
        return !n || (n.v === null && n.f === undefined);
      };
      if (ha === 'left') {
        let extra = 0;
        for (let k = c + 1; k < c + 60 && free(k, r) && !this.mergeAt(r, k) && x + w + extra < this.W + 50; k++) {
          extra += this.colW(k);
          if (tx + tw <= x + w + extra - pad) break;
        }
        clipW = w + extra;
      } else if (ha === 'right') {
        let extra = 0;
        for (let k = c - 1; k >= Math.max(0, c - 60) && free(k, r) && !this.mergeAt(r, k); k--) {
          extra += this.colW(k);
          if (tx >= x - extra + pad) break;
        }
        clipX = x - extra;
        clipW = w + extra;
      } else {
        let l = 0;
        let rr = 0;
        const need = (tw - avail) / 2;
        for (let k = c - 1; k >= Math.max(0, c - 60) && free(k, r) && l < need; k--) l += this.colW(k);
        for (let k = c + 1; k < c + 60 && free(k, r) && rr < need; k++) rr += this.colW(k);
        clipX = x - l;
        clipW = w + l + rr;
      }
    }
    const needClip = tw > clipW - 2 * pad + 0.5 || fontPx * 1.1 > h;
    if (needClip) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, y, clipW, h);
      ctx.clip();
      ctx.fillText(text, tx, baseY);
      this.decorate(st, tx, baseY, tw, fontPx);
      ctx.restore();
    } else {
      ctx.fillText(text, tx, baseY);
      this.decorate(st, tx, baseY, tw, fontPx);
    }
  }

  private decorate(st: import('../core/model').CellStyle, x: number, base: number, w: number, fontPx: number): void {
    if (!st.underline && !st.strike) return;
    const ctx = this.ctx;
    ctx.save();
    ctx.strokeStyle = ctx.fillStyle as string;
    ctx.lineWidth = Math.max(1, fontPx / 14);
    if (st.underline) {
      const y = Math.round(base + fontPx * 0.12) + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.stroke();
    }
    if (st.strike) {
      const y = Math.round(base - fontPx * 0.3) + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + w, y);
      ctx.stroke();
    }
    ctx.restore();
  }

  private clipRectFor(r: Rect) {
    return this.rectPx(r);
  }

  private drawOverlays(refs: RefSpan[]): void {
    const ctx = this.ctx;
    const t = this.theme;
    // editing references
    for (const ref of refs) {
      if (ref.sheet && ref.sheet.toLowerCase() !== this.sheet.name.toLowerCase()) continue;
      const p = this.rectPx(this.expandForMerges(ref.rect));
      ctx.strokeStyle = ref.color;
      ctx.lineWidth = 2;
      ctx.strokeRect(p.x + 1, p.y + 1, p.w - 2, p.h - 2);
      ctx.fillStyle = ref.color;
      ctx.globalAlpha = 0.07;
      ctx.fillRect(p.x, p.y, p.w, p.h);
      ctx.globalAlpha = 1;
    }
    // copy marquee
    const cb = this.host.clipboardRect();
    if (cb && cb.sheet === this.sheet) {
      const p = this.rectPx(cb.rect);
      ctx.strokeStyle = t.accent;
      ctx.lineWidth = 1.5;
      ctx.setLineDash([5, 4]);
      ctx.strokeRect(p.x + 1, p.y + 1, p.w - 2, p.h - 2);
      ctx.setLineDash([]);
    }
    if (this.editor.active && this.editor.mode !== 'point' && !this.editor.text.startsWith('=')) {
      // nothing: editor box covers the cell
    }
    // selection
    const n = this.ranges.length;
    this.ranges.forEach((rg, idx) => {
      const p = this.rectPx(rg);
      const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
      const a = this.expandForMerges({ r1: this.active.r, c1: this.active.c, r2: this.active.r, c2: this.active.c });
      const hasActive = idx === n - 1 && a.r1 >= rg.r1 && a.r2 <= rg.r2 && a.c1 >= rg.c1 && a.c2 <= rg.c2;
      const isSingle = !multi || (hasActive && a.r1 === rg.r1 && a.r2 === rg.r2 && a.c1 === rg.c1 && a.c2 === rg.c2);
      if (!isSingle) {
        ctx.fillStyle = t.selFill;
        if (hasActive) {
          const ap = this.rectPx(a);
          ctx.beginPath();
          ctx.rect(p.x, p.y, p.w, p.h);
          ctx.rect(ap.x, ap.y, ap.w, ap.h);
          ctx.fill('evenodd');
        } else ctx.fillRect(p.x, p.y, p.w, p.h);
      }
      if (this.editor.active && this.editor.mode === 'point') return;
      ctx.strokeStyle = t.accent;
      ctx.lineWidth = 2;
      ctx.strokeRect(p.x + 0.5, p.y + 0.5, p.w - 1, p.h - 1);
      if (hasActive && !isSingle) {
        // active cell outline inside a range
      }
    });
    // fill handle
    const last = this.range;
    if (!this.editor.active && last.r2 < MAX_ROWS - 1 && last.c2 < MAX_COLS - 1) {
      const p = this.rectPx(last);
      const hx = p.x + p.w - 4;
      const hy = p.y + p.h - 4;
      ctx.fillStyle = t.bg;
      ctx.fillRect(hx - 1, hy - 1, 8, 8);
      ctx.fillStyle = t.accent;
      ctx.fillRect(hx, hy, 6, 6);
    }
    // drag preview
    if (this.dragPreview) {
      const p = this.rectPx(this.dragPreview);
      ctx.strokeStyle = this.drag?.kind === 'move' ? t.text : t.accent;
      ctx.lineWidth = 2;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(p.x + 1, p.y + 1, p.w - 2, p.h - 2);
      ctx.setLineDash([]);
    }
  }

  private drawHeaders(cols: { frozen: Seg[]; scroll: Seg[] }, rows: { frozen: Seg[]; scroll: Seg[] }): void {
    const ctx = this.ctx;
    const t = this.theme;
    const hw = this.hw;
    const hh = this.hh;
    const z = Math.max(0.6, Math.min(this.zoom, 2));
    const fontPx = 11 * (4 / 3) * 0.95 * z;
    const font = `${fontPx}px "Segoe UI", Calibri, Arial, sans-serif`;
    // selection sets
    const colSel = (c: number) => this.ranges.some((r) => c >= r.c1 && c <= r.c2);
    const rowSel = (r: number) => this.ranges.some((rg) => r >= rg.r1 && r <= rg.r2);
    const colFull = (c: number) => this.ranges.some((r) => c >= r.c1 && c <= r.c2 && r.r1 === 0 && r.r2 === MAX_ROWS - 1);
    const rowFull = (r: number) => this.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && rg.c1 === 0 && rg.c2 === MAX_COLS - 1);
    ctx.font = font;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    // column headers
    for (const [segs, cx, cw] of [
      [cols.scroll, hw + this.fW, this.W - hw - this.fW],
      [cols.frozen, hw, this.fW],
    ] as [Seg[], number, number][]) {
      if (cw <= 0) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(cx, 0, cw, hh);
      ctx.clip();
      ctx.fillStyle = t.hdrBg;
      ctx.fillRect(cx, 0, cw, hh);
      for (const s of segs) {
        const sel = colSel(s.i);
        if (sel) {
          ctx.fillStyle = colFull(s.i) ? t.hdrAllSelBg : t.hdrSelBg;
          ctx.fillRect(s.p, 0, s.s, hh);
        }
        ctx.fillStyle = sel ? t.hdrSelText : t.hdrText;
        ctx.fillText(colName(s.i), s.p + s.s / 2, hh / 2 + 1);
        ctx.strokeStyle = t.hdrLine;
        ctx.lineWidth = 1 / this.dpr;
        ctx.beginPath();
        ctx.moveTo(this.px(s.p + s.s), 0);
        ctx.lineTo(this.px(s.p + s.s), hh);
        ctx.stroke();
        if (sel) {
          ctx.fillStyle = t.accent;
          ctx.fillRect(s.p, hh - 2, s.s, 2);
        }
      }
      ctx.restore();
    }
    // row headers
    for (const [segs, cy, ch] of [
      [rows.scroll, hh + this.fH, this.H - hh - this.fH],
      [rows.frozen, hh, this.fH],
    ] as [Seg[], number, number][]) {
      if (ch <= 0) continue;
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, cy, hw, ch);
      ctx.clip();
      ctx.fillStyle = t.hdrBg;
      ctx.fillRect(0, cy, hw, ch);
      for (const s of segs) {
        const sel = rowSel(s.i);
        if (sel) {
          ctx.fillStyle = rowFull(s.i) ? t.hdrAllSelBg : t.hdrSelBg;
          ctx.fillRect(0, s.p, hw, s.s);
        }
        ctx.fillStyle = sel ? t.hdrSelText : t.hdrText;
        ctx.fillText(String(s.i + 1), hw / 2, s.p + s.s / 2 + 1);
        ctx.strokeStyle = t.hdrLine;
        ctx.lineWidth = 1 / this.dpr;
        ctx.beginPath();
        ctx.moveTo(0, this.px(s.p + s.s));
        ctx.lineTo(hw, this.px(s.p + s.s));
        ctx.stroke();
        if (sel) {
          ctx.fillStyle = t.accent;
          ctx.fillRect(hw - 2, s.p, 2, s.s);
        }
      }
      ctx.restore();
    }
    // corner
    ctx.fillStyle = t.hdrBg;
    ctx.fillRect(0, 0, hw, hh);
    ctx.fillStyle = this.theme === DARK ? '#5a5a5a' : '#b4b4b4';
    ctx.beginPath();
    ctx.moveTo(hw - 4, hh - 14 * z);
    ctx.lineTo(hw - 4, hh - 4);
    ctx.lineTo(hw - 14 * z, hh - 4);
    ctx.closePath();
    ctx.fill();
    ctx.strokeStyle = t.hdrLine;
    ctx.lineWidth = 1 / this.dpr;
    ctx.beginPath();
    ctx.moveTo(0, this.px(hh));
    ctx.lineTo(this.W, this.px(hh));
    ctx.moveTo(this.px(hw), 0);
    ctx.lineTo(this.px(hw), this.H);
    ctx.stroke();
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
  }

  // ------------------------------------------------------------------ hit testing
  private hit(x: number, y: number) {
    const hw = this.hw;
    const hh = this.hh;
    const inCorner = x < hw && y < hh;
    if (inCorner) return { zone: 'corner' as const };
    if (y < hh) {
      const c = this.colAt(x);
      const right = this.colX(c) + this.colW(c);
      let resize: number | null = null;
      if (Math.abs(x - right) <= 4) resize = c;
      else {
        // left edge of next col means resize of the previous visible col
        const left = this.colX(c);
        if (Math.abs(x - left) <= 4 && c > 0) {
          let p = c - 1;
          while (p > 0 && this.colW(p) === 0) p--;
          resize = p;
        }
      }
      return { zone: 'colHeader' as const, c, resize };
    }
    if (x < hw) {
      const r = this.rowAt(y);
      const bottom = this.rowY(r) + this.rowH(r);
      let resize: number | null = null;
      if (Math.abs(y - bottom) <= 3) resize = r;
      else {
        const top = this.rowY(r);
        if (Math.abs(y - top) <= 3 && r > 0) {
          let p = r - 1;
          while (p > 0 && this.rowH(p) === 0) p--;
          resize = p;
        }
      }
      return { zone: 'rowHeader' as const, r, resize };
    }
    const r = this.rowAt(y);
    const c = this.colAt(x);
    // fill handle & selection border
    const last = this.range;
    let onHandle = false;
    let onBorder = false;
    if (!this.editor.active) {
      const p = this.rectPx(last);
      const hx = p.x + p.w;
      const hy = p.y + p.h;
      if (Math.abs(x - hx) <= 5 && Math.abs(y - hy) <= 5 && last.r2 < MAX_ROWS - 1 && last.c2 < MAX_COLS - 1) onHandle = true;
      else if (this.ranges.length === 1) {
        const inX = x >= p.x - 3 && x <= p.x + p.w + 3;
        const inY = y >= p.y - 3 && y <= p.y + p.h + 3;
        const nearV = Math.abs(x - p.x) <= 3 || Math.abs(x - (p.x + p.w)) <= 3;
        const nearH = Math.abs(y - p.y) <= 3 || Math.abs(y - (p.y + p.h)) <= 3;
        if ((nearV && inY) || (nearH && inX)) onBorder = true;
      }
    }
    // filter button
    let filterCol: number | null = null;
    const f = this.sheet.filter;
    if (f && r === f.rect.r1 && c >= f.rect.c1 && c <= f.rect.c2) {
      const right = this.colX(c) + this.colW(c);
      const bottom = this.rowY(r) + this.rowH(r);
      const sz = 16 * Math.max(0.7, Math.min(this.zoom, 1.5));
      if (x >= right - sz - 1 && x <= right && y >= bottom - sz - 1 && y <= bottom) filterCol = c;
    }
    return { zone: 'cell' as const, r, c, onHandle, onBorder, filterCol };
  }

  private local(e: MouseEvent): { x: number; y: number } {
    const r = this.area.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  // ------------------------------------------------------------------ mouse
  private bind(): void {
    const area = this.area;
    area.tabIndex = 0;
    area.addEventListener('mousedown', (e) => this.onMouseDown(e));
    area.addEventListener('dblclick', (e) => this.onDblClick(e));
    area.addEventListener('mousemove', (e) => this.onMouseMove(e));
    area.addEventListener('contextmenu', (e) => this.onContext(e));
    area.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
    this.bindScrollbars();
  }

  private onWheel(e: WheelEvent): void {
    e.preventDefault();
    if (e.ctrlKey) {
      this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
      return;
    }
    let dx = e.deltaX;
    let dy = e.deltaY;
    if (e.deltaMode === 1) {
      dx *= 20;
      dy *= 20;
    }
    if (e.shiftKey && dx === 0) {
      dx = dy;
      dy = 0;
    }
    this.scrollBy(dx * 0.8, dy * 0.6);
  }

  private onContext(e: MouseEvent): void {
    e.preventDefault();
    const p = this.local(e);
    const h = this.hit(p.x, p.y);
    if (h.zone === 'cell') {
      const inSel = this.ranges.some((r) => h.r >= r.r1 && h.r <= r.r2 && h.c >= r.c1 && h.c <= r.c2);
      if (!inSel) this.select(h.r, h.c);
      this.host.contextMenu(e.clientX, e.clientY, 'cell');
    } else if (h.zone === 'colHeader') {
      const inSel = this.ranges.some((r) => h.c >= r.c1 && h.c <= r.c2 && r.r1 === 0 && r.r2 === MAX_ROWS - 1);
      if (!inSel) this.selectCols(h.c, h.c);
      this.host.contextMenu(e.clientX, e.clientY, 'col');
    } else if (h.zone === 'rowHeader') {
      const inSel = this.ranges.some((r) => h.r >= r.r1 && h.r <= r.r2 && r.c1 === 0 && r.c2 === MAX_COLS - 1);
      if (!inSel) this.selectRows(h.r, h.r);
      this.host.contextMenu(e.clientX, e.clientY, 'row');
    }
  }

  selectCols(c1: number, c2: number): void {
    this.active = { r: this.active.r, c: Math.min(c1, c2) };
    this.anchor = { r: 0, c: c1 };
    this.focus = { r: MAX_ROWS - 1, c: c2 };
    this.ranges = [normRect(0, c1, MAX_ROWS - 1, c2)];
    this.saveSel();
    this.invalidate();
    this.host.selectionChanged();
  }
  selectRows(r1: number, r2: number): void {
    this.active = { r: Math.min(r1, r2), c: this.active.c };
    this.anchor = { r: r1, c: 0 };
    this.focus = { r: r2, c: MAX_COLS - 1 };
    this.ranges = [normRect(r1, 0, r2, MAX_COLS - 1)];
    this.saveSel();
    this.invalidate();
    this.host.selectionChanged();
  }
  selectAll(): void {
    this.anchor = { r: 0, c: 0 };
    this.focus = { r: MAX_ROWS - 1, c: MAX_COLS - 1 };
    this.ranges = [{ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }];
    this.saveSel();
    this.invalidate();
    this.host.selectionChanged();
  }

  private onMouseDown(e: MouseEvent): void {
    if (e.button !== 0) {
      if (e.button === 2) return;
      return;
    }
    if ((e.target as HTMLElement).closest('.cell-editor')) return;
    const p = this.local(e);
    this.mouse = p;
    const h = this.hit(p.x, p.y);
    // formula point mode
    if (this.editor.active && h.zone === 'cell' && this.editor.canPoint()) {
      e.preventDefault();
      this.pointAt(h.r, h.c, false);
      this.drag = { kind: 'point', anchor: { r: h.r, c: h.c } };
      this.startDragListeners();
      return;
    }
    if (this.editor.active) {
      if (!this.commitEdit('none')) {
        e.preventDefault();
        return;
      }
    }
    e.preventDefault();
    this.area.focus({ preventScroll: true });
    if (h.zone === 'corner') {
      this.selectAll();
      return;
    }
    if (h.zone === 'colHeader') {
      if (h.resize !== null) {
        const sel = this.ranges.some((r) => h.resize! >= r.c1 && h.resize! <= r.c2 && r.r1 === 0 && r.r2 === MAX_ROWS - 1) ? this.colsOfSelection() : [h.resize];
        this.drag = { kind: 'colResize', c: h.resize, startX: p.x, startW: this.sheet.cols.size(h.resize), sel };
        this.startDragListeners();
        return;
      }
      if (e.shiftKey) this.selectCols(this.anchor.c, h.c);
      else if (e.ctrlKey || e.metaKey) {
        this.ranges.push(normRect(0, h.c, MAX_ROWS - 1, h.c));
        this.active = { r: this.active.r, c: h.c };
        this.invalidate();
        this.host.selectionChanged();
      } else this.selectCols(h.c, h.c);
      this.drag = { kind: 'col', anchor: e.shiftKey ? this.anchor.c : h.c, ctrl: e.ctrlKey };
      this.startDragListeners();
      return;
    }
    if (h.zone === 'rowHeader') {
      if (h.resize !== null) {
        const sel = this.ranges.some((r) => h.resize! >= r.r1 && h.resize! <= r.r2 && r.c1 === 0 && r.c2 === MAX_COLS - 1) ? this.rowsOfSelection() : [h.resize];
        this.drag = { kind: 'rowResize', r: h.resize, startY: p.y, startH: this.sheet.rows.size(h.resize), sel };
        this.startDragListeners();
        return;
      }
      if (e.shiftKey) this.selectRows(this.anchor.r, h.r);
      else if (e.ctrlKey || e.metaKey) {
        this.ranges.push(normRect(h.r, 0, h.r, MAX_COLS - 1));
        this.active = { r: h.r, c: this.active.c };
        this.invalidate();
        this.host.selectionChanged();
      } else this.selectRows(h.r, h.r);
      this.drag = { kind: 'row', anchor: e.shiftKey ? this.anchor.r : h.r, ctrl: e.ctrlKey };
      this.startDragListeners();
      return;
    }
    if (h.zone !== 'cell') return;
    if (h.filterCol !== null) {
      const cr = this.cellRect(h.r, h.filterCol);
      const ar = this.area.getBoundingClientRect();
      this.host.filterButton(h.filterCol, new DOMRect(ar.left + cr.x, ar.top + cr.y, cr.w, cr.h));
      return;
    }
    if (h.onHandle) {
      this.drag = { kind: 'fill', src: { ...this.range } };
      this.startDragListeners();
      return;
    }
    if (h.onBorder && !e.shiftKey) {
      this.drag = { kind: 'move', src: { ...this.range }, grab: { r: h.r, c: h.c } };
      this.startDragListeners();
      return;
    }
    if (e.shiftKey) {
      this.select(this.anchor.r, this.anchor.c, h.r, h.c, true);
    } else if (e.ctrlKey || e.metaKey) {
      const m = this.expandForMerges({ r1: h.r, c1: h.c, r2: h.r, c2: h.c });
      this.ranges.push(m);
      this.active = { r: h.r, c: h.c };
      this.anchor = { r: h.r, c: h.c };
      this.focus = { r: h.r, c: h.c };
      this.invalidate();
      this.host.selectionChanged();
    } else {
      this.select(h.r, h.c);
    }
    this.drag = { kind: 'select', anchor: e.shiftKey ? { ...this.anchor } : { r: h.r, c: h.c }, ctrl: e.ctrlKey };
    this.startDragListeners();
  }

  private colsOfSelection(): number[] {
    const out: number[] = [];
    for (const r of this.ranges) if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 2000); c++) out.push(c);
    return out;
  }
  private rowsOfSelection(): number[] {
    const out: number[] = [];
    for (const r of this.ranges) if (r.c1 === 0 && r.c2 === MAX_COLS - 1) for (let k = r.r1; k <= Math.min(r.r2, r.r1 + 5000); k++) out.push(k);
    return out;
  }

  private startDragListeners(): void {
    const move = (ev: MouseEvent) => {
      this.mouse = this.local(ev);
      this.onDragMove(ev);
    };
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      clearInterval(this.autoScroll);
      this.autoScroll = 0;
      this.onDragEnd(ev);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
    clearInterval(this.autoScroll);
    this.autoScroll = window.setInterval(() => this.tickAutoScroll(), 40);
  }

  private tickAutoScroll(): void {
    const d = this.drag;
    if (!d || d.kind === 'colResize' || d.kind === 'rowResize') return;
    const { x, y } = this.mouse;
    const minX = this.hw + this.fW;
    const minY = this.hh + this.fH;
    let dx = 0;
    let dy = 0;
    if (x > this.W - 6) dx = Math.min(80, (x - this.W + 6) * 0.8 + 14);
    else if (x < minX && (d.kind !== 'row')) dx = -Math.min(80, (minX - x) * 0.8 + 14);
    if (y > this.H - 6) dy = Math.min(80, (y - this.H + 6) * 0.8 + 14);
    else if (y < minY && (d.kind !== 'col')) dy = -Math.min(80, (minY - y) * 0.8 + 14);
    if (dx || dy) {
      this.scrollBy(dx, dy);
      this.onDragMove(null);
    }
  }

  private onDragMove(_e: MouseEvent | null): void {
    const d = this.drag;
    if (!d) return;
    const { x, y } = this.mouse;
    const cx = clamp(x, this.hw + 1, this.W - 1);
    const cy = clamp(y, this.hh + 1, this.H - 1);
    switch (d.kind) {
      case 'select': {
        const r = this.rowAt(cy);
        const c = this.colAt(cx);
        if (d.ctrl && this.ranges.length > 1) {
          this.ranges[this.ranges.length - 1] = this.expandForMerges(normRect(d.anchor.r, d.anchor.c, r, c));
          this.focus = { r, c };
          this.invalidate();
          this.host.selectionChanged();
        } else this.select(d.anchor.r, d.anchor.c, r, c, true);
        break;
      }
      case 'col': {
        const c = this.colAt(clamp(x, this.hw + 1, this.W - 1));
        if (d.ctrl && this.ranges.length > 1) this.ranges[this.ranges.length - 1] = normRect(0, d.anchor, MAX_ROWS - 1, c);
        else this.selectCols(d.anchor, c);
        this.invalidate();
        this.host.selectionChanged();
        break;
      }
      case 'row': {
        const r = this.rowAt(clamp(y, this.hh + 1, this.H - 1));
        if (d.ctrl && this.ranges.length > 1) this.ranges[this.ranges.length - 1] = normRect(d.anchor, 0, r, MAX_COLS - 1);
        else this.selectRows(d.anchor, r);
        this.invalidate();
        this.host.selectionChanged();
        break;
      }
      case 'colResize': {
        const w = Math.max(0, Math.round(d.startW + (x - d.startX) / this.zoom));
        this.sheet.cols.set(d.c, w);
        for (const c of d.sel) if (c !== d.c) this.sheet.cols.set(c, w);
        this.invalidate();
        this.host.viewChanged();
        this.host.message(`Width: ${(Math.max(0, w - 5) / 7).toFixed(2)} (${w} pixels)`);
        break;
      }
      case 'rowResize': {
        const h = Math.max(0, Math.round(d.startH + (y - d.startY) / this.zoom));
        this.sheet.rows.set(d.r, h);
        for (const r of d.sel) if (r !== d.r) this.sheet.rows.set(r, h);
        this.invalidate();
        this.host.viewChanged();
        this.host.message(`Height: ${(h * 0.75).toFixed(2)} (${h} pixels)`);
        break;
      }
      case 'fill': {
        const r = this.rowAt(cy);
        const c = this.colAt(cx);
        const src = d.src;
        const below = r - src.r2;
        const above = src.r1 - r;
        const right = c - src.c2;
        const left = src.c1 - c;
        const vert = Math.max(below, above);
        const horiz = Math.max(right, left);
        if (vert <= 0 && horiz <= 0) this.dragPreview = null;
        else if (vert >= horiz) this.dragPreview = below >= above ? { ...src, r2: r } : { ...src, r1: r };
        else this.dragPreview = right >= left ? { ...src, c2: c } : { ...src, c1: c };
        this.invalidate();
        break;
      }
      case 'move': {
        const r = this.rowAt(cy);
        const c = this.colAt(cx);
        const dr = r - d.grab.r;
        const dc = c - d.grab.c;
        const s = d.src;
        this.dragPreview = { r1: Math.max(0, s.r1 + dr), c1: Math.max(0, s.c1 + dc), r2: Math.max(0, s.r2 + dr), c2: Math.max(0, s.c2 + dc) };
        this.invalidate();
        break;
      }
      case 'point': {
        const r = this.rowAt(cy);
        const c = this.colAt(cx);
        this.pointAt(d.anchor.r, d.anchor.c, true, r, c);
        break;
      }
    }
  }

  private onDragEnd(e: MouseEvent): void {
    const d = this.drag;
    this.drag = null;
    const preview = this.dragPreview;
    this.dragPreview = null;
    if (!d) return;
    switch (d.kind) {
      case 'colResize': {
        const w = this.sheet.cols.size(d.c);
        // commit through the undo history
        const sel = d.sel;
        this.sheet.cols.set(d.c, d.startW);
        for (const c of sel) this.sheet.cols.set(c, d.startW);
        this.host.wb.transact('Column width', () => {
          this.sheet.cols.set(d.c, w);
          for (const c of sel) this.sheet.cols.set(c, w);
        }, true);
        this.host.message('');
        break;
      }
      case 'rowResize': {
        const hgt = this.sheet.rows.size(d.r);
        const sel = d.sel;
        this.sheet.rows.set(d.r, d.startH);
        for (const r of sel) this.sheet.rows.set(r, d.startH);
        this.host.wb.transact('Row height', () => {
          this.sheet.rows.set(d.r, hgt);
          for (const r of sel) this.sheet.rows.set(r, hgt);
        }, true);
        this.host.message('');
        break;
      }
      case 'fill': {
        if (!preview) break;
        const src = d.src;
        if (preview.r2 > src.r2) this.host.fillDone(src, 'down', preview.r2 - src.r2);
        else if (preview.r1 < src.r1) this.host.fillDone(src, 'up', src.r1 - preview.r1);
        else if (preview.c2 > src.c2) this.host.fillDone(src, 'right', preview.c2 - src.c2);
        else if (preview.c1 < src.c1) this.host.fillDone(src, 'left', src.c1 - preview.c1);
        break;
      }
      case 'move': {
        if (!preview) break;
        if (preview.r1 !== d.src.r1 || preview.c1 !== d.src.c1) this.host.moveDone(d.src, preview.r1, preview.c1, e.ctrlKey);
        break;
      }
      case 'select':
        if (this.host.painterActive()) this.host.painterApply(this.range);
        break;
      default:
        break;
    }
    this.invalidate();
    this.host.viewChanged();
  }

  private onMouseMove(e: MouseEvent): void {
    if (this.drag) return;
    const p = this.local(e);
    const h = this.hit(p.x, p.y);
    let cur = 'cell';
    if (h.zone === 'colHeader') cur = h.resize !== null ? 'col-resize' : 'default';
    else if (h.zone === 'rowHeader') cur = h.resize !== null ? 'row-resize' : 'default';
    else if (h.zone === 'corner') cur = 'default';
    else if (h.zone === 'cell') cur = h.filterCol !== null ? 'default' : h.onHandle ? 'crosshair' : h.onBorder ? 'move' : 'cell';
    if (cur !== this.cursor) {
      this.cursor = cur;
      this.area.style.cursor = cur;
    }
  }

  private onDblClick(e: MouseEvent): void {
    if ((e.target as HTMLElement).closest('.cell-editor')) return;
    const p = this.local(e);
    const h = this.hit(p.x, p.y);
    if (h.zone === 'colHeader' && h.resize !== null) {
      const cols = this.ranges.some((r) => h.resize! >= r.c1 && h.resize! <= r.c2 && r.r1 === 0 && r.r2 === MAX_ROWS - 1) ? this.colsOfSelection() : [h.resize];
      this.autofitCols(cols);
    } else if (h.zone === 'rowHeader' && h.resize !== null) {
      const rows = this.ranges.some((r) => h.resize! >= r.r1 && h.resize! <= r.r2 && r.c1 === 0 && r.c2 === MAX_COLS - 1) ? this.rowsOfSelection() : [h.resize];
      this.autofitRows(rows);
    } else if (h.zone === 'cell') {
      if (h.onHandle) {
        this.fillDown();
        return;
      }
      if (h.onBorder) {
        // double click on the selection border jumps to the edge of data
        return;
      }
      if (h.filterCol !== null) return;
      this.beginEdit('edit');
    }
  }

  private fillDown(): void {
    const rg = this.range;
    const s = this.sheet;
    const has = (r: number, c: number) => {
      if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return false;
      const cell = s.get(r, c);
      return !!cell && (cell.v !== null || cell.f !== undefined);
    };
    // the neighbouring column (left first, then right) decides how far to fill
    let ref = -1;
    if (has(rg.r2 + 1, rg.c1 - 1)) ref = rg.c1 - 1;
    else if (has(rg.r2 + 1, rg.c2 + 1)) ref = rg.c2 + 1;
    if (ref < 0) return;
    let r = rg.r2 + 1;
    while (has(r + 1, ref)) r++;
    this.host.fillDone(rg, 'down', r - rg.r2);
  }

  // ------------------------------------------------------------------ scrollbars
  private bindScrollbars(): void {
    const drag = (thumb: HTMLElement, vertical: boolean) => {
      thumb.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const ext = (this as any)._ext;
        if (!ext) return;
        const startPos = vertical ? e.clientY : e.clientX;
        const startScroll = vertical ? this.sheet.scrollY : this.sheet.scrollX;
        const extent = vertical ? ext.extentY : ext.extentX;
        const view = vertical ? ext.vh : ext.vw;
        const track = vertical ? ext.vTrack - ext.vSize : ext.hTrack - ext.hSize;
        const move = (ev: MouseEvent) => {
          const d = (vertical ? ev.clientY : ev.clientX) - startPos;
          const v = startScroll + (d / Math.max(1, track)) * (extent - view);
          if (vertical) this.sheet.scrollY = v;
          else this.sheet.scrollX = v;
          this.clampScroll();
          this.invalidate();
          if (this.editor.active) this.editor.layout();
        };
        const up = () => {
          window.removeEventListener('mousemove', move);
          window.removeEventListener('mouseup', up);
        };
        window.addEventListener('mousemove', move);
        window.addEventListener('mouseup', up);
      });
    };
    drag(this.vthumb, true);
    drag(this.hthumb, false);
    const track = (el: HTMLElement, vertical: boolean) => {
      el.addEventListener('mousedown', (e) => {
        if (e.target !== el) return;
        e.preventDefault();
        const r = el.getBoundingClientRect();
        const th = (vertical ? this.vthumb : this.hthumb).getBoundingClientRect();
        const pos = vertical ? e.clientY : e.clientX;
        const before = vertical ? pos < th.top : pos < th.left;
        const page = (vertical ? this.viewH() : this.viewW()) * 0.9;
        this.scrollBy(vertical ? 0 : before ? -page : page, vertical ? (before ? -page : page) : 0);
        void r;
      });
    };
    track(this.vscroll, true);
    track(this.hscroll, false);
  }

  // ------------------------------------------------------------------ sizing helpers
  autofitCols(cols: number[]): void {
    const s = this.sheet;
    const wb = this.host.wb;
    const colSet = new Set(cols);
    const widths = new Map<number, number>();
    // bucket the populated cells once for all requested columns
    for (const [k, cell] of s.cells) {
      const c = k % MAX_COLS;
      if (!colSet.has(c) || cell.v === null) continue;
      const st = wb.styleOf(cell.s);
      const disp = this.display(cell, st.fmt ?? 'General');
      if (disp.text === '') continue;
      const font = this.fontFor(cell.s).replace(/[\d.]+px/, (m) => parseFloat(m) / this.zoom + 'px');
      const text = disp.text.split(FILL_MARK).join('');
      const widest = st.wrap ? 0 : this.measureUnzoomed(font, text);
      const w = widest + 8 + (st.indent ?? 0) * 9 + (s.filter && s.filter.rect.r1 === Math.floor(k / MAX_COLS) ? 18 : 0);
      if (w > (widths.get(c) ?? 0)) widths.set(c, w);
    }
    wb.transact('AutoFit column width', () => {
      for (const c of cols) {
        const maxW = widths.get(c) ?? 0;
        s.cols.set(c, maxW > 0 ? Math.max(12, Math.ceil(maxW)) : 64);
      }
    }, true);
    this.invalidate();
  }

  private measureUnzoomed(font: string, text: string): number {
    this.ctx.font = font;
    let w = 0;
    for (const line of text.split('\n')) w = Math.max(w, this.ctx.measureText(line).width);
    return w;
  }

  autofitRows(rows: number[], record = true): void {
    const s = this.sheet;
    const wb = this.host.wb;
    const rowSet = new Set(rows);
    const needs = new Map<number, number>();
    // bucket the populated cells once for all requested rows
    for (const [k, cell] of s.cells) {
      const r = Math.floor(k / MAX_COLS);
      if (!rowSet.has(r) || cell.v === null) continue;
      const st = wb.styleOf(cell.s);
      const px = (st.size ?? 11) * (4 / 3);
      let lines = 1;
      if (st.wrap) {
        const c = k % MAX_COLS;
        const font = this.fontFor(cell.s).replace(/[\d.]+px/, () => px + 'px');
        const disp = this.display(cell, st.fmt ?? 'General');
        const w = s.cols.size(c) - 8;
        this.ctx.font = font;
        lines = this.wrapLinesUnzoomed(font, disp.text, w).length;
      }
      const n = Math.ceil(lines * px * 1.25 + 4);
      if (n > (needs.get(r) ?? 20)) needs.set(r, n);
    }
    const run = (label: string, fn: () => void) => (record ? wb.transact(label, fn, true) : fn());
    run('AutoFit row height', () => {
      for (const r of rows) s.rows.set(r, needs.get(r) ?? 20);
    });
    this.invalidate();
  }

  private wrapLinesUnzoomed(font: string, text: string, maxW: number): string[] {
    const out: string[] = [];
    this.ctx.font = font;
    const m = (t: string) => this.ctx.measureText(t).width;
    for (const para of text.split('\n')) {
      let line = '';
      for (const w of para.split(/(\s+)/)) {
        if (w === '') continue;
        if (m((line + w).trimEnd()) <= maxW || line === '') line += w;
        else {
          out.push(line.trimEnd());
          line = w.trimStart();
        }
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  // ------------------------------------------------------------------ editing
  private editorAnchor() {
    if (!this.editor.active) return null;
    const cr = this.cellRect(this.editor.r, this.editor.c);
    // keep inside the visible area
    const maxW = this.W - cr.x - 2;
    const maxH = this.H - cr.y - 2;
    return { x: cr.x, y: cr.y, w: cr.w, h: cr.h, maxW: Math.max(cr.w, maxW), maxH: Math.max(cr.h, maxH) };
  }

  beginEdit(mode: 'enter' | 'edit', initial?: string): void {
    if (this.editor.active) return;
    const { r, c } = this.active;
    const m = this.mergeAt(r, c);
    const rr = m ? m.r1 : r;
    const cc = m ? m.c1 : c;
    this.scrollIntoView(rr, cc);
    const font = this.host.editorFont(rr, cc);
    this.editor.fontCss = font.css;
    this.editor.zoom = this.zoom;
    const text = initial !== undefined ? initial : this.host.editorText(rr, cc);
    this.editor.start(rr, cc, text, mode);
    this.host.editStateChanged();
    this.invalidate();
  }

  /** Commit current edit. Returns false when the edit was rejected (error shown). */
  commitEdit(move: 'down' | 'up' | 'right' | 'left' | 'none'): boolean {
    if (!this.editor.active) return true;
    const { r, c } = this.editor;
    const text = this.editor.text;
    const err = this.host.commitCell(r, c, text);
    if (err) {
      this.host.message(err);
      return false;
    }
    this.editor.stop();
    this.host.editStateChanged();
    this.focusGrid();
    if (move !== 'none') this.moveInSelection(move);
    else this.host.selectionChanged();
    this.invalidate();
    return true;
  }

  cancelEdit(): void {
    if (!this.editor.active) return;
    this.editor.stop();
    this.host.editStateChanged();
    this.focusGrid();
    this.invalidate();
    this.host.selectionChanged();
  }

  private pointAt(r1: number, c1: number, dragging: boolean, r2 = r1, c2 = c1): void {
    const rect = this.expandForMerges(normRect(r1, c1, r2, c2));
    let text = rect.r1 === rect.r2 && rect.c1 === rect.c2 ? a1(rect.r1, rect.c1) : a1(rect.r1, rect.c1) + ':' + a1(rect.r2, rect.c2);
    if (dragging && this.editor.pointSpan === null) return;
    // reference to another sheet is prefixed by the App; same sheet here
    this.editor.pointTo(text);
    this.invalidate();
  }

  // ------------------------------------------------------------------ keyboard
  private isFilled(r: number, c: number): boolean {
    const cell = this.sheet.cells.get(cellKey(r, c));
    return !!cell && (cell.v !== null || cell.f !== undefined);
  }

  private step(r: number, c: number, dr: number, dc: number): { r: number; c: number } {
    let nr = r;
    let nc = c;
    const m = this.mergeAt(r, c);
    if (m) {
      if (dr > 0) nr = m.r2;
      if (dr < 0) nr = m.r1;
      if (dc > 0) nc = m.c2;
      if (dc < 0) nc = m.c1;
    }
    nr += dr;
    nc += dc;
    while (nr >= 0 && nr < MAX_ROWS && this.sheet.hiddenRow(nr) && dr !== 0) nr += dr;
    while (nc >= 0 && nc < MAX_COLS && this.sheet.hiddenCol(nc) && dc !== 0) nc += dc;
    nr = clamp(nr, 0, MAX_ROWS - 1);
    nc = clamp(nc, 0, MAX_COLS - 1);
    const m2 = this.mergeAt(nr, nc);
    if (m2) return { r: m2.r1, c: m2.c1 };
    return { r: nr, c: nc };
  }

  private jump(r: number, c: number, dr: number, dc: number): { r: number; c: number } {
    const maxR = MAX_ROWS - 1;
    const maxC = MAX_COLS - 1;
    const inb = (rr: number, cc: number) => rr >= 0 && rr <= maxR && cc >= 0 && cc <= maxC;
    let nr = r + dr;
    let nc = c + dc;
    if (!inb(nr, nc)) return { r, c };
    const cur = this.isFilled(r, c);
    const nxt = this.isFilled(nr, nc);
    if (cur && nxt) {
      while (inb(nr + dr, nc + dc) && this.isFilled(nr + dr, nc + dc)) {
        nr += dr;
        nc += dc;
      }
      return { r: nr, c: nc };
    }
    // search for the next filled cell (bounded by the used extent)
    const used = this.sheet.usedRect();
    const limitR = dr > 0 ? (used ? used.r2 : 0) : dr < 0 ? (used ? used.r1 : 0) : nr;
    const limitC = dc > 0 ? (used ? used.c2 : 0) : dc < 0 ? (used ? used.c1 : 0) : nc;
    while (inb(nr, nc)) {
      if (this.isFilled(nr, nc)) return { r: nr, c: nc };
      if ((dr > 0 && nr >= limitR) || (dr < 0 && nr <= limitR) || (dc > 0 && nc >= limitC) || (dc < 0 && nc <= limitC)) break;
      nr += dr;
      nc += dc;
    }
    return { r: dr > 0 ? maxR : dr < 0 ? 0 : r, c: dc > 0 ? maxC : dc < 0 ? 0 : c };
  }

  moveInSelection(dir: 'down' | 'up' | 'right' | 'left'): void {
    const rg = this.range;
    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    const a = this.active;
    if (multi && this.ranges.length === 1) {
      let { r, c } = a;
      const m = this.mergeAt(r, c);
      void m;
      if (dir === 'down') {
        r++;
        if (r > rg.r2) {
          r = rg.r1;
          c++;
          if (c > rg.c2) c = rg.c1;
        }
      } else if (dir === 'up') {
        r--;
        if (r < rg.r1) {
          r = rg.r2;
          c--;
          if (c < rg.c1) c = rg.c2;
        }
      } else if (dir === 'right') {
        c++;
        if (c > rg.c2) {
          c = rg.c1;
          r++;
          if (r > rg.r2) r = rg.r1;
        }
      } else {
        c--;
        if (c < rg.c1) {
          c = rg.c2;
          r--;
          if (r < rg.r1) r = rg.r2;
        }
      }
      this.active = { r, c };
      this.scrollIntoView(r, c);
      this.saveSel();
      this.invalidate();
      this.host.selectionChanged();
      return;
    }
    const d = dir === 'down' ? [1, 0] : dir === 'up' ? [-1, 0] : dir === 'right' ? [0, 1] : [0, -1];
    const n = this.step(a.r, a.c, d[0], d[1]);
    this.select(n.r, n.c);
    this.scrollIntoView(n.r, n.c);
  }

  /** Returns true if the key was handled. */
  handleKey(e: KeyboardEvent): boolean {
    if (this.editor.active) return this.handleEditKey(e);
    const ctrl = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;
    const k = e.key;
    const nav = (dr: number, dc: number) => {
      const from = shift ? this.focus : this.active;
      let t = ctrl ? this.jump(from.r, from.c, dr, dc) : this.step(from.r, from.c, dr, dc);
      if (shift) {
        this.select(this.anchor.r, this.anchor.c, t.r, t.c, true);
      } else this.select(t.r, t.c);
      this.scrollIntoView(t.r, t.c);
      return true;
    };
    switch (k) {
      case 'ArrowUp': return nav(-1, 0);
      case 'ArrowDown': return nav(1, 0);
      case 'ArrowLeft': return nav(0, -1);
      case 'ArrowRight': return nav(0, 1);
      case 'Tab':
        this.moveInSelection(shift ? 'left' : 'right');
        return true;
      case 'Enter':
        this.moveInSelection(shift ? 'up' : 'down');
        return true;
      case 'Home': {
        if (ctrl) {
          const t = { r: this.fr, c: this.fc };
          if (shift) this.select(this.anchor.r, this.anchor.c, 0, 0, true);
          else this.select(0, 0);
          this.scrollTo(0, 0);
          void t;
          return true;
        }
        const r = shift ? this.focus.r : this.active.r;
        if (shift) this.select(this.anchor.r, this.anchor.c, r, 0, true);
        else this.select(r, 0);
        this.scrollIntoView(r, 0);
        return true;
      }
      case 'End': {
        if (ctrl) {
          const u = this.sheet.usedRect();
          const r = u ? u.r2 : 0;
          const c = u ? u.c2 : 0;
          if (shift) this.select(this.anchor.r, this.anchor.c, r, c, true);
          else this.select(r, c);
          this.scrollIntoView(r, c);
          return true;
        }
        return false;
      }
      case 'PageDown':
      case 'PageUp': {
        const dir = k === 'PageDown' ? 1 : -1;
        const from = shift ? this.focus : this.active;
        if (e.altKey) {
          const n = Math.max(1, this.visibleColCount() - 1);
          const c = clamp(from.c + dir * n, 0, MAX_COLS - 1);
          if (shift) this.select(this.anchor.r, this.anchor.c, from.r, c, true);
          else this.select(from.r, c);
          this.scrollTo(this.sheet.scrollX + dir * this.viewW() / this.zoom, this.sheet.scrollY);
          return true;
        }
        const n = Math.max(1, this.visibleRowCount() - 1);
        const r = clamp(from.r + dir * n, 0, MAX_ROWS - 1);
        if (shift) this.select(this.anchor.r, this.anchor.c, r, from.c, true);
        else this.select(r, from.c);
        this.scrollTo(this.sheet.scrollX, this.sheet.scrollY + (dir * this.viewH()) / this.zoom);
        this.scrollIntoView(r, from.c);
        return true;
      }
      case 'F2':
        this.beginEdit('edit');
        return true;
      case 'Delete':
        this.host.clearSelection();
        return true;
      case 'Backspace':
        if (ctrl) return false;
        this.host.clearSelection();
        this.beginEdit('enter', '');
        return true;
      case ' ':
        if (ctrl && shift) {
          this.selectAll();
          return true;
        }
        if (ctrl) {
          this.selectCols(this.range.c1, this.range.c2);
          return true;
        }
        if (shift) {
          this.selectRows(this.range.r1, this.range.r2);
          return true;
        }
        break;
      case 'a':
      case 'A':
        if (ctrl && !shift) {
          const reg = currentRegion(this.sheet, this.active.r, this.active.c);
          const cur = this.range;
          const same = cur.r1 === reg.r1 && cur.c1 === reg.c1 && cur.r2 === reg.r2 && cur.c2 === reg.c2;
          if (same || (reg.r1 === reg.r2 && reg.c1 === reg.c2) || this.selAllNext) {
            this.selectAll();
            this.selAllNext = false;
          } else {
            this.anchor = { r: reg.r1, c: reg.c1 };
            this.select(reg.r1, reg.c1, reg.r2, reg.c2, true);
            this.selAllNext = false;
          }
          return true;
        }
        break;
    }
    // printable character starts editing
    if (!ctrl && !e.altKey && k.length === 1) {
      this.beginEdit('enter', '');
      return true; // do not preventDefault: the character lands in the editor
    }
    return false;
  }

  private visibleRowCount(): number {
    return Math.max(1, Math.floor(this.viewH() / (this.sheet.rows.def * this.zoom)));
  }
  private visibleColCount(): number {
    return Math.max(1, Math.floor(this.viewW() / (this.sheet.cols.def * this.zoom)));
  }

  private handleEditKey(e: KeyboardEvent): boolean {
    const ed = this.editor;
    const k = e.key;
    if (k === 'Escape') {
      this.cancelEdit();
      return true;
    }
    if (k === 'Enter') {
      if (e.altKey) {
        const ta = e.target as HTMLTextAreaElement;
        if (ta && ta.setRangeText) {
          const s = ta.selectionStart;
          ta.setRangeText('\n', s, ta.selectionEnd, 'end');
          ed.setText(ta.value, s + 1);
        }
        return true;
      }
      this.commitEdit(e.shiftKey ? 'up' : 'down');
      return true;
    }
    if (k === 'Tab') {
      this.commitEdit(e.shiftKey ? 'left' : 'right');
      return true;
    }
    if (ed.mode !== 'edit' && (k === 'ArrowUp' || k === 'ArrowDown' || k === 'ArrowLeft' || k === 'ArrowRight') && !e.ctrlKey) {
      // enter / point mode: arrows commit (or point into a formula)
      if (ed.text.startsWith('=') && ed.canPoint()) {
        const base = ed.mode === 'point' && this.pointCursor ? this.pointCursor : { r: this.active.r, c: this.active.c };
        const d = k === 'ArrowUp' ? [-1, 0] : k === 'ArrowDown' ? [1, 0] : k === 'ArrowLeft' ? [0, -1] : [0, 1];
        const n = e.shiftKey && this.pointCursor ? { r: clamp(this.pointCursor.r + d[0], 0, MAX_ROWS - 1), c: clamp(this.pointCursor.c + d[1], 0, MAX_COLS - 1) } : this.step(base.r, base.c, d[0], d[1]);
        e.preventDefault();
        if (e.shiftKey && this.pointAnchor) this.pointAt(this.pointAnchor.r, this.pointAnchor.c, false, n.r, n.c);
        else {
          this.pointAnchor = { ...n };
          this.pointAt(n.r, n.c, false);
        }
        this.pointCursor = n;
        this.scrollIntoView(n.r, n.c);
        return true;
      }
      e.preventDefault();
      this.commitEdit(k === 'ArrowUp' ? 'up' : k === 'ArrowDown' ? 'down' : k === 'ArrowLeft' ? 'left' : 'right');
      return true;
    }
    this.pointCursor = null;
    this.pointAnchor = null;
    if (k === 'F2') {
      ed.mode = ed.mode === 'edit' ? 'enter' : 'edit';
      return true;
    }
    return false;
  }
  private pointCursor: { r: number; c: number } | null = null;
  private pointAnchor: { r: number; c: number } | null = null;
}

function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

export { rectText };
