// Cell editor + formula bar editing: coloured references, function autocomplete and argument hints.
import { Rect, normRect } from '../core/addr';
import { FUNCS } from '../core/eval';
import '../core/functions';
import { tokenize } from '../core/parser';
import { h } from './dom';

export const REF_COLORS = ['#2e75b6', '#548235', '#7030a0', '#c00000', '#c55a11', '#00859b', '#bf8f00', '#6b6b6b'];

export interface RefSpan {
  start: number; // index in full text (including "=")
  end: number;
  rect: Rect;
  sheet?: string;
  color: string;
  text: string;
}

export function refSpans(text: string): RefSpan[] {
  if (!text.startsWith('=')) return [];
  const out: RefSpan[] = [];
  const colorOf = new Map<string, string>();
  let ci = 0;
  for (const t of tokenize(text.slice(1))) {
    if (t.kind !== 'ref' || !t.range) continue;
    const a = t.range.a;
    const b = t.range.b ?? a;
    const key = (t.sheet ?? '') + '!' + t.text.toUpperCase().replace(/\$/g, '');
    let col = colorOf.get(key);
    if (!col) {
      col = REF_COLORS[ci++ % REF_COLORS.length];
      colorOf.set(key, col);
    }
    out.push({ start: t.start + 1, end: t.end + 1, rect: normRect(a.r, a.c, b.r, b.c), sheet: t.sheet, color: col, text: t.text });
  }
  return out;
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function mirrorHtml(text: string): string {
  const spans = refSpans(text);
  if (!spans.length) return esc(text) + '\n';
  let out = '';
  let pos = 0;
  for (const s of spans) {
    out += esc(text.slice(pos, s.start));
    out += `<span style="color:${s.color}">${esc(text.slice(s.start, s.end))}</span>`;
    pos = s.end;
  }
  out += esc(text.slice(pos));
  return out + '\n';
}

/** Rotate absolute/relative state of the reference at caret (F4). */
export function cycleReference(text: string, caret: number): { text: string; start: number; end: number } | null {
  if (!text.startsWith('=')) return null;
  const toks = tokenize(text.slice(1)).filter((t) => t.kind === 'ref' && t.range);
  const tok = toks.find((t) => caret >= t.start + 1 && caret <= t.end + 1);
  if (!tok) return null;
  const sheetPrefix = tok.text.includes('!') ? tok.text.slice(0, tok.text.lastIndexOf('!') + 1) : '';
  const body = tok.text.slice(sheetPrefix.length);
  const parts = body.split(':').map((p) => {
    const m = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(p);
    if (!m) return p;
    const absC = m[1] === '$';
    const absR = m[3] === '$';
    let nc: boolean, nr: boolean;
    if (!absC && !absR) { nc = true; nr = true; }
    else if (absC && absR) { nc = false; nr = true; }
    else if (!absC && absR) { nc = true; nr = false; }
    else { nc = false; nr = false; }
    return (nc ? '$' : '') + m[2] + (nr ? '$' : '') + m[4];
  });
  const nt = sheetPrefix + parts.join(':');
  const start = tok.start + 1;
  const end = tok.end + 1;
  return { text: text.slice(0, start) + nt + text.slice(end), start, end: start + nt.length };
}

// ---------------- argument hints ----------------

interface CallFrame {
  name: string | null;
  idx: number;
}

export function callAt(text: string, caret: number): CallFrame | null {
  if (!text.startsWith('=')) return null;
  const toks = tokenize(text.slice(1));
  const stack: CallFrame[] = [];
  let pendingFunc: string | null = null;
  for (const t of toks) {
    const s = t.start + 1;
    if (s >= caret) break;
    if (t.kind === 'func') pendingFunc = t.text.toUpperCase();
    else if (t.kind === 'lparen') {
      stack.push({ name: pendingFunc, idx: 0 });
      pendingFunc = null;
    } else if (t.kind === 'rparen') {
      // only pop when the paren closes before caret
      if (t.end + 1 <= caret) stack.pop();
    } else if (t.kind === 'comma') {
      if (stack.length) stack[stack.length - 1].idx++;
    } else if (t.kind !== 'ws') pendingFunc = null;
  }
  for (let i = stack.length - 1; i >= 0; i--) if (stack[i].name) return stack[i];
  return null;
}

export function hintHtml(name: string, idx: number): string | null {
  const def = FUNCS.get(name);
  if (!def) return null;
  const args = def.sig ? def.sig.split(',').map((s) => s.trim()) : [];
  let cur = idx;
  if (cur >= args.length) {
    const last = args.length - 1;
    cur = args[last] === '...' ? Math.max(0, last - 1) : last;
  }
  const parts = args.map((a, i) => (i === cur && a !== '...' ? `<b>${esc(a)}</b>` : esc(a)));
  return `${esc(def.name)}(${parts.join(', ')})`;
}

// ---------------- the editor ----------------

export type EditMode = 'enter' | 'edit' | 'point';

export interface EditorHost {
  /** called whenever text/caret changes so the grid can redraw reference boxes */
  onEditChange(): void;
  definedNames(): string[];
  /** Rendering anchor: cell rectangle in container (px) */
  anchorRect(): { x: number; y: number; w: number; h: number; maxW: number; maxH: number } | null;
}

export class CellEditor {
  active = false;
  mode: EditMode = 'enter';
  r = 0;
  c = 0;
  /** last reference inserted by pointing: replaced on next point action */
  pointSpan: { start: number; end: number } | null = null;
  readonly box: HTMLDivElement;
  readonly ta: HTMLTextAreaElement;
  private mirror: HTMLDivElement;
  bar: HTMLTextAreaElement | null = null;
  barMirror: HTMLDivElement | null = null;
  private suggestEl: HTMLDivElement;
  private hintEl: HTMLDivElement;
  private sugg: { label: string; insert: string; desc: string; fn: boolean }[] = [];
  private suggIdx = 0;
  private fragStart = 0;
  private measure: CanvasRenderingContext2D;
  fontCss = '14.67px Calibri';
  zoom = 1;

  constructor(private host: EditorHost, parent: HTMLElement) {
    this.ta = h('textarea', { class: 'cell-ta', spellcheck: 'false', autocomplete: 'off', autocapitalize: 'off', rows: '1', wrap: 'off' });
    this.mirror = h('div', { class: 'cell-mirror' });
    this.box = h('div', { class: 'cell-editor', style: { display: 'none' } }, this.mirror, this.ta);
    this.suggestEl = h('div', { class: 'suggest', style: { display: 'none' } });
    this.hintEl = h('div', { class: 'fx-hint', style: { display: 'none' } });
    parent.append(this.box);
    document.body.append(this.suggestEl, this.hintEl);
    this.measure = document.createElement('canvas').getContext('2d')!;
    this.ta.addEventListener('input', () => this.onInput(this.ta));
    this.ta.addEventListener('keydown', (e) => this.onKey(e, this.ta));
    this.ta.addEventListener('keyup', () => this.refresh());
    this.ta.addEventListener('mouseup', () => this.refresh());
    this.ta.addEventListener('blur', () => setTimeout(() => this.hideHelpers(), 120));
    this.suggestEl.addEventListener('mousedown', (e) => e.preventDefault());
  }

  attachBar(bar: HTMLTextAreaElement, mirror: HTMLDivElement): void {
    this.bar = bar;
    this.barMirror = mirror;
    bar.addEventListener('input', () => this.onInput(bar));
    bar.addEventListener('keydown', (e) => this.onKey(e, bar));
    bar.addEventListener('keyup', () => this.refresh());
    bar.addEventListener('mouseup', () => this.refresh());
    bar.addEventListener('blur', () => setTimeout(() => this.hideHelpers(), 120));
  }

  get text(): string {
    return this.ta.value;
  }

  private focused(): HTMLTextAreaElement {
    return document.activeElement === this.bar && this.bar ? this.bar : this.ta;
  }

  /** Caret position in the active textarea. */
  caret(): number {
    return this.focused().selectionStart ?? this.text.length;
  }

  start(r: number, c: number, text: string, mode: EditMode, caretAtEnd = true): void {
    this.active = true;
    this.r = r;
    this.c = c;
    this.mode = mode;
    this.pointSpan = null;
    this.ta.value = text;
    this.box.style.display = 'block';
    this.layout();
    this.ta.focus();
    if (caretAtEnd) this.ta.setSelectionRange(text.length, text.length);
    this.syncBar();
    this.refresh();
  }

  stop(): void {
    this.active = false;
    this.box.style.display = 'none';
    this.hideHelpers();
    this.ta.value = '';
    this.pointSpan = null;
  }

  setText(text: string, caret?: number): void {
    this.ta.value = text;
    if (caret !== undefined) {
      this.ta.setSelectionRange(caret, caret);
      this.bar?.setSelectionRange(caret, caret);
    }
    this.syncBar();
    this.layout();
    this.refresh();
  }

  /** Insert text at the caret of whichever textarea has focus. */
  replaceRange(start: number, end: number, ins: string, caretAfter = start + ins.length): void {
    const t = this.text;
    this.setText(t.slice(0, start) + ins + t.slice(end), caretAfter);
  }

  private onInput(src: HTMLTextAreaElement): void {
    if (!this.active) {
      return; // App handles starting edit from the formula bar
    }
    if (src !== this.ta) this.ta.value = src.value;
    this.pointSpan = null;
    if (this.mode === 'point') this.mode = 'edit';
    this.syncBar();
    this.layout();
    this.refresh();
  }

  syncBar(): void {
    if (this.bar && this.bar.value !== this.ta.value) this.bar.value = this.ta.value;
    if (this.barMirror) this.barMirror.innerHTML = mirrorHtml(this.ta.value);
    this.mirror.innerHTML = mirrorHtml(this.ta.value);
  }

  /** Size/position the editor box over the cell. */
  layout(): void {
    const a = this.host.anchorRect();
    if (!a) return;
    const text = this.ta.value;
    this.measure.font = this.fontCss;
    const lines = text.split('\n');
    let w = 0;
    for (const l of lines) w = Math.max(w, this.measure.measureText(l).width);
    const pad = 8 * this.zoom;
    let width = Math.max(a.w, w + pad + 14 * this.zoom);
    let wrap = false;
    if (width > a.maxW) {
      width = a.maxW;
      wrap = true;
    }
    const fontPx = parseFloat(/(\d+(?:\.\d+)?)px/.exec(this.fontCss)?.[1] ?? '');
    const lh = Math.round((Number.isFinite(fontPx) ? fontPx : 15) * 1.25);
    const lineCount = wrap ? Math.max(lines.length, Math.ceil((w + pad) / Math.max(10, width - pad))) : lines.length;
    const height = Math.min(a.maxH, Math.max(a.h, lineCount * lh + 4));
    Object.assign(this.box.style, { left: a.x + 'px', top: a.y + 'px', width: width + 'px', height: height + 'px', font: this.fontCss });
    const ws = wrap ? 'pre-wrap' : 'pre';
    this.ta.style.whiteSpace = ws;
    this.mirror.style.whiteSpace = ws;
    this.ta.setAttribute('wrap', wrap ? 'soft' : 'off');
    this.ta.style.font = this.fontCss;
    this.mirror.style.font = this.fontCss;
  }

  // ---------- keys ----------
  private onKey(e: KeyboardEvent, src: HTMLTextAreaElement): void {
    if (!this.active) return;
    // suggestion navigation
    if (this.suggestEl.style.display !== 'none' && this.sugg.length) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        this.suggIdx = (this.suggIdx + 1) % this.sugg.length;
        this.renderSuggest();
        return;
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        this.suggIdx = (this.suggIdx - 1 + this.sugg.length) % this.sugg.length;
        this.renderSuggest();
        return;
      }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey && !e.ctrlKey)) {
        e.preventDefault();
        e.stopPropagation();
        this.acceptSuggest();
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        this.suggestEl.style.display = 'none';
        this.sugg = [];
        return;
      }
    }
    if (e.key === 'F4') {
      e.preventDefault();
      const caret = src.selectionStart ?? 0;
      const res = cycleReference(this.text, caret);
      if (res) this.setText(res.text, res.end);
    }
  }

  // ---------- suggestions & hints ----------
  refresh(): void {
    this.host.onEditChange();
    if (!this.active) return;
    const text = this.text;
    const caret = this.caret();
    this.updateSuggest(text, caret);
    this.updateHint(text, caret);
  }

  private updateSuggest(text: string, caret: number): void {
    if (!text.startsWith('=') || caret < 2) return this.hideSuggest();
    const before = text.slice(0, caret);
    const m = /[A-Za-z_][A-Za-z0-9_.]*$/.exec(before);
    if (!m) return this.hideSuggest();
    const start = caret - m[0].length;
    const prev = before[start - 1];
    if (prev && /[A-Za-z0-9_.!'$]/.test(prev)) return this.hideSuggest();
    const frag = m[0].toUpperCase();
    // looks like a cell reference being typed (A1, B12): no suggestions
    if (/^[A-Z]{1,3}\d+$/.test(frag)) return this.hideSuggest();
    const next = text[caret];
    if (next && /[A-Za-z0-9_.]/.test(next)) return this.hideSuggest();
    const items: { label: string; insert: string; desc: string; fn: boolean }[] = [];
    for (const [name, def] of FUNCS) {
      if (name.startsWith(frag)) items.push({ label: name, insert: name + '(', desc: def.desc, fn: true });
    }
    items.sort((a, b) => a.label.length - b.label.length || a.label.localeCompare(b.label));
    for (const n of this.host.definedNames()) if (n.toUpperCase().startsWith(frag)) items.push({ label: n, insert: n, desc: 'Defined name', fn: false });
    if (!items.length || (items.length === 1 && items[0].label === frag && text[caret] === '(')) return this.hideSuggest();
    this.sugg = items.slice(0, 12);
    this.suggIdx = Math.min(this.suggIdx, this.sugg.length - 1);
    this.fragStart = start;
    this.renderSuggest();
  }

  private renderSuggest(): void {
    const el = this.suggestEl;
    el.innerHTML = '';
    this.sugg.forEach((s, i) => {
      const row = h('div', { class: 'sg-row' + (i === this.suggIdx ? ' sel' : '') }, h('span', { class: 'sg-ic' }, s.fn ? 'ƒx' : '▣'), h('span', { class: 'sg-name' }, s.label));
      row.addEventListener('mousedown', (e) => {
        e.preventDefault();
        this.suggIdx = i;
        this.acceptSuggest();
      });
      el.append(row);
    });
    const cur = this.sugg[this.suggIdx];
    if (cur) el.append(h('div', { class: 'sg-desc' }, cur.desc));
    el.style.display = 'block';
    const base = this.focused() === this.bar && this.bar ? this.bar.getBoundingClientRect() : this.box.getBoundingClientRect();
    el.style.left = base.left + 'px';
    el.style.top = base.bottom + 2 + 'px';
    const sel = el.querySelector('.sel') as HTMLElement | null;
    sel?.scrollIntoView({ block: 'nearest' });
  }

  private acceptSuggest(): void {
    const s = this.sugg[this.suggIdx];
    if (!s) return;
    const caret = this.caret();
    this.replaceRange(this.fragStart, caret, s.insert);
    this.hideSuggest();
    this.focused().focus();
  }

  private hideSuggest(): void {
    this.suggestEl.style.display = 'none';
    this.sugg = [];
    this.suggIdx = 0;
  }

  private updateHint(text: string, caret: number): void {
    const call = callAt(text, caret);
    const html = call?.name ? hintHtml(call.name, call.idx) : null;
    if (!html) {
      this.hintEl.style.display = 'none';
      return;
    }
    this.hintEl.innerHTML = html;
    this.hintEl.style.display = 'block';
    const base = this.focused() === this.bar && this.bar ? this.bar.getBoundingClientRect() : this.box.getBoundingClientRect();
    this.hintEl.style.left = base.left + 'px';
    const sugVisible = this.suggestEl.style.display !== 'none';
    this.hintEl.style.top = (sugVisible ? base.bottom + this.suggestEl.offsetHeight + 6 : base.bottom + 2) + 'px';
  }

  hideHelpers(): void {
    this.hideSuggest();
    this.hintEl.style.display = 'none';
  }

  /** Whether a click/arrow key should insert a reference at the caret right now. */
  canPoint(): boolean {
    const t = this.text;
    if (!t.startsWith('=')) return false;
    const caret = this.caret();
    if (this.pointSpan && this.pointSpan.end === caret) return true;
    let i = caret - 1;
    while (i >= 0 && t[i] === ' ') i--;
    if (i < 1) return i === 0; // right after "="
    return '=+-*/^&<>,(;:'.includes(t[i]) || (t[i] === '=' && true);
  }

  /** Insert/replace a reference produced by clicking or arrowing in the grid. */
  pointTo(refText: string): void {
    const caret = this.caret();
    const span = this.pointSpan && this.pointSpan.end === caret ? this.pointSpan : { start: caret, end: caret };
    this.replaceRange(span.start, span.end, refText);
    this.pointSpan = { start: span.start, end: span.start + refText.length };
    this.mode = 'point';
    this.focused().focus();
    this.ta.focus();
  }

  destroy(): void {
    this.suggestEl.remove();
    this.hintEl.remove();
  }
}
