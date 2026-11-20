// Workbook model: sheets, sparse cells, styles, undo history and the recalculation engine.
import { MAX_COLS, MAX_ROWS, Rect, cellKey, keyCol, keyRow } from './addr';
import { Ctx, DepRef, NameDef, SheetLike, WbLike, collectRefs, evaluate } from './eval';
import { FErr, E, Value } from './values';
import { Node, ParseError, parseFormula } from './parser';
import './functions';

// ---------------- styles ----------------

export interface Border {
  style: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair';
  color: string;
}
export interface CellStyle {
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fill?: string;
  h?: 'left' | 'center' | 'right' | 'justify';
  v?: 'top' | 'middle' | 'bottom';
  wrap?: boolean;
  indent?: number;
  fmt?: string;
  bl?: Border;
  br?: Border;
  bt?: Border;
  bb?: Border;
}
export type StylePatch = { [K in keyof CellStyle]?: CellStyle[K] | null };

const STYLE_KEYS: (keyof CellStyle)[] = ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color', 'fill', 'h', 'v', 'wrap', 'indent', 'fmt', 'bl', 'br', 'bt', 'bb'];

export function styleKey(s: CellStyle): string {
  const parts: unknown[] = [];
  for (const k of STYLE_KEYS) {
    const v = s[k];
    if (v !== undefined) parts.push(k, v);
  }
  return JSON.stringify(parts);
}

// ---------------- cells ----------------

export interface Cell {
  v: Value;
  f?: string; // formula text without leading "="
  s?: number; // style index
  ast?: Node | null; // cached parse (null = parse error)
}

export interface CFRule {
  id: number;
  rect: Rect;
  type: 'cell' | 'text' | 'dup' | 'unique' | 'top' | 'bottom' | 'scale' | 'bar' | 'formula' | 'above' | 'below';
  op?: string; // > < >= <= = <> between notbetween ; text: contains, begins, ends, notcontains
  v1?: string;
  v2?: string;
  n?: number; // top/bottom N
  style?: { color?: string; fill?: string; bold?: boolean; italic?: boolean };
  colors?: string[]; // color scale
  barColor?: string;
}

export interface AutoFilter {
  rect: Rect; // header row is rect.r1
  cols: Map<number, Set<string>>; // allowed display strings per column
}

export class Axis {
  private idx: number[] = [];
  private cum: number[] = [];
  private dirty = true;
  sizes = new Map<number, number>();
  constructor(public def: number) {}

  size(i: number): number {
    const s = this.sizes.get(i);
    return s === undefined ? this.def : s;
  }
  set(i: number, size: number): void {
    if (size === this.def) this.sizes.delete(i);
    else this.sizes.set(i, size);
    this.dirty = true;
  }
  clear(): void {
    this.sizes.clear();
    this.dirty = true;
  }
  touch(): void {
    this.dirty = true;
  }
  private build() {
    if (!this.dirty) return;
    this.idx = Array.from(this.sizes.keys()).sort((a, b) => a - b);
    this.cum = new Array(this.idx.length);
    let acc = 0;
    for (let k = 0; k < this.idx.length; k++) {
      acc += this.sizes.get(this.idx[k])! - this.def;
      this.cum[k] = acc;
    }
    this.dirty = false;
  }
  /** Offset of the start of item i. */
  top(i: number): number {
    this.build();
    // count customs with index < i
    let lo = 0;
    let hi = this.idx.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.idx[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return i * this.def + (lo > 0 ? this.cum[lo - 1] : 0);
  }
  /** Index of the item containing position pos. */
  at(pos: number, max: number): number {
    if (pos <= 0) return 0;
    let lo = 0;
    let hi = max - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.top(mid) <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
  total(max: number): number {
    return this.top(max);
  }
}

export class Sheet implements SheetLike {
  cells = new Map<number, Cell>();
  maxRow = -1;
  maxCol = -1;
  cols = new Axis(64);
  rows = new Axis(20);
  merges: Rect[] = [];
  cf: CFRule[] = [];
  freezeRows = 0;
  freezeCols = 0;
  tabColor?: string;
  filter?: AutoFilter;
  showGrid = true;
  // view state kept per sheet
  sel = { r: 0, c: 0, r2: 0, c2: 0 };
  scrollX = 0;
  scrollY = 0;
  constructor(public id: number, public name: string) {}

  get(r: number, c: number): Cell | undefined {
    return this.cells.get(cellKey(r, c));
  }
  hiddenRow(r: number): boolean {
    return this.rows.sizes.get(r) === 0;
  }
  hiddenCol(c: number): boolean {
    return this.cols.sizes.get(c) === 0;
  }
  usedRect(): Rect | null {
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    for (const [k, cell] of this.cells) {
      if (cell.v === null && cell.f === undefined && cell.s === undefined) continue;
      const r = keyRow(k);
      const c = keyCol(k);
      if (r < r1) r1 = r;
      if (r > r2) r2 = r;
      if (c < c1) c1 = c;
      if (c > c2) c2 = c;
    }
    return r2 < 0 ? null : { r1, c1, r2, c2 };
  }
  recomputeBounds(): void {
    let mr = -1, mc = -1;
    for (const k of this.cells.keys()) {
      const r = keyRow(k);
      const c = keyCol(k);
      if (r > mr) mr = r;
      if (c > mc) mc = c;
    }
    this.maxRow = mr;
    this.maxCol = mc;
  }
}

// ---------------- history ----------------

interface HistoryEntry {
  label: string;
  swap(): void;
  sheetId?: number;
  rect?: Rect;
  /** approximate number of cells this step keeps in memory */
  size?: number;
}

interface SheetSnap {
  sheet: Sheet;
  name: string;
  cells: Map<number, Cell>;
  cols: Map<number, number>;
  rows: Map<number, number>;
  merges: Rect[];
  cf: CFRule[];
  freezeRows: number;
  freezeCols: number;
  tabColor?: string;
  filter?: { rect: Rect; cols: [number, string[]][] };
}
interface WbSnap {
  sheets: SheetSnap[];
  names: [string, NameDef][];
}

const SHEET_BASE = 2 ** 35;
export function fullKey(sheetId: number, r: number, c: number): number {
  return sheetId * SHEET_BASE + cellKey(r, c);
}

interface RangeDep {
  f: number;
  sheet: number;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}
interface DepRec {
  pts: number[];
  rngs: RangeDep[];
}

const MAX_UNDO = 200;
const MAX_UNDO_CELLS = 2_000_000;

export class Workbook implements WbLike {
  sheets: Sheet[] = [];
  styles: CellStyle[] = [{}];
  private styleMap = new Map<string, number>([[styleKey({}), 0]]);
  names = new Map<string, NameDef>();
  private sheetsById = new Map<number, Sheet>();
  private nextSheetId = 1;
  nextRuleId = 1;
  dataVersion = 0;
  // dependency graph
  private deps = new Map<number, DepRec>();
  private pointDeps = new Map<number, Set<number>>();
  private rangeBuckets = new Map<number, Set<RangeDep>>();
  private tallRanges = new Map<number, Set<RangeDep>>();
  private volatileSet = new Set<number>();
  private nameAstCache = new Map<string, Node | null>();
  // history
  private undoStack: HistoryEntry[] = [];
  private redoStack: HistoryEntry[] = [];
  private tx: { label: string; journal: Map<number, Cell | undefined>; meta?: WbSnap; depth: number; sheetId?: number; rect?: Rect } | null = null;
  private pending: number[] = [];
  private fullRecalcPending = false;
  onChange: (() => void) | null = null;
  autoCalc = true;

  constructor() {
    this.addSheet('Sheet1');
  }

  // ---------- sheets ----------
  sheetByName(name: string): Sheet | undefined {
    const l = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === l);
  }
  sheetById(id: number): Sheet | undefined {
    return this.sheetsById.get(id);
  }
  addSheet(name?: string, at?: number): Sheet {
    if (!name) {
      let i = this.sheets.length + 1;
      while (this.sheetByName('Sheet' + i)) i++;
      name = 'Sheet' + i;
    }
    const s = new Sheet(this.nextSheetId++, name);
    this.sheetsById.set(s.id, s);
    if (at === undefined || at >= this.sheets.length) this.sheets.push(s);
    else this.sheets.splice(at, 0, s);
    return s;
  }
  uniqueSheetName(base: string, except?: Sheet): string {
    let name = base;
    let i = 2;
    while (this.sheets.some((s) => s !== except && s.name.toLowerCase() === name.toLowerCase())) name = base + ' (' + i++ + ')';
    return name;
  }

  // ---------- names ----------
  getName(name: string): NameDef | undefined {
    return this.names.get(name.toLowerCase());
  }
  nameAst(def: NameDef): Node | null {
    let a = this.nameAstCache.get(def.refers);
    if (a === undefined) {
      try {
        a = parseFormula(def.refers);
      } catch {
        a = null;
      }
      this.nameAstCache.set(def.refers, a);
    }
    return a;
  }
  setName(name: string, refers: string): void {
    this.names.set(name.toLowerCase(), { name, refers });
    this.nameAstCache.clear();
  }
  deleteName(name: string): void {
    this.names.delete(name.toLowerCase());
  }

  // ---------- styles ----------
  styleIndex(s: CellStyle): number {
    const k = styleKey(s);
    let i = this.styleMap.get(k);
    if (i === undefined) {
      i = this.styles.length;
      const clean: CellStyle = {};
      for (const key of STYLE_KEYS) if (s[key] !== undefined) (clean as any)[key] = s[key];
      this.styles.push(clean);
      this.styleMap.set(k, i);
    }
    return i;
  }
  styleOf(idx: number | undefined): CellStyle {
    return this.styles[idx ?? 0] ?? this.styles[0];
  }
  patchStyle(idx: number | undefined, patch: StylePatch): number {
    const base = { ...this.styleOf(idx) } as any;
    for (const k of Object.keys(patch)) {
      const v = (patch as any)[k];
      if (v === null || v === undefined) delete base[k];
      else base[k] = v;
    }
    return this.styleIndex(base);
  }

  // ---------- transactions & history ----------
  begin(label: string, meta = false): void {
    if (this.tx) {
      this.tx.depth++;
      if (meta && !this.tx.meta) this.tx.meta = this.snapshot();
      return;
    }
    this.tx = { label, journal: new Map(), depth: 1, meta: meta ? this.snapshot() : undefined };
  }
  /** Declare which sheet/rect the transaction touched (used to move the selection on undo). */
  touch(sheetId: number, rect?: Rect): void {
    if (!this.tx) return;
    this.tx.sheetId = sheetId;
    if (rect) this.tx.rect = this.tx.rect ? union(this.tx.rect, rect) : { ...rect };
  }
  commit(): void {
    const tx = this.tx;
    if (!tx) return;
    if (--tx.depth > 0) return;
    this.tx = null;
    this.flush();
    if (tx.meta) {
      const snap = tx.meta;
      const entry: HistoryEntry = {
        label: tx.label,
        sheetId: tx.sheetId,
        rect: tx.rect,
        size: snap.sheets.reduce((n, sh) => n + sh.cells.size, 0),
        swap: () => {
          const cur = this.snapshot();
          this.restore(snap);
          Object.assign(snap, cur);
        },
      };
      this.pushUndo(entry);
    } else if (tx.journal.size) {
      const states = tx.journal; // key -> state to restore (initially the "before" state)
      const entry: HistoryEntry = {
        label: tx.label,
        sheetId: tx.sheetId,
        rect: tx.rect,
        size: states.size,
        swap: () => {
          const changed: number[] = [];
          for (const [k, prev] of Array.from(states.entries())) {
            const sheet = this.sheetsById.get(Math.floor(k / SHEET_BASE));
            if (!sheet) continue;
            const ck = k % SHEET_BASE;
            const cur = sheet.cells.get(ck);
            this.putRaw(sheet, ck, prev);
            states.set(k, cur);
            changed.push(k);
          }
          this.pending.push(...changed);
          this.flush();
        },
      };
      this.pushUndo(entry);
    }
    this.onChange?.();
  }
  /** Run fn inside a transaction. */
  transact<T>(label: string, fn: () => T, meta = false): T {
    this.begin(label, meta);
    try {
      return fn();
    } finally {
      this.commit();
    }
  }
  private pushUndo(e: HistoryEntry) {
    this.undoStack.push(e);
    if (this.undoStack.length > MAX_UNDO) this.undoStack.shift();
    // cap total journaled cells; always keep the newest step
    let total = 0;
    for (const x of this.undoStack) total += x.size ?? 1;
    while (total > MAX_UNDO_CELLS && this.undoStack.length > 1) total -= this.undoStack.shift()!.size ?? 1;
    this.redoStack = [];
  }
  canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }
  undoLabel(): string | null {
    return this.undoStack.length ? this.undoStack[this.undoStack.length - 1].label : null;
  }
  redoLabel(): string | null {
    return this.redoStack.length ? this.redoStack[this.redoStack.length - 1].label : null;
  }
  undo(): HistoryEntry | null {
    const e = this.undoStack.pop();
    if (!e) return null;
    e.swap();
    this.redoStack.push(e);
    this.onChange?.();
    return e;
  }
  redo(): HistoryEntry | null {
    const e = this.redoStack.pop();
    if (!e) return null;
    e.swap();
    this.undoStack.push(e);
    this.onChange?.();
    return e;
  }
  clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  snapshot(): WbSnap {
    return {
      sheets: this.sheets.map((s) => ({
        sheet: s,
        name: s.name,
        cells: new Map(s.cells),
        cols: new Map(s.cols.sizes),
        rows: new Map(s.rows.sizes),
        merges: s.merges.map((m) => ({ ...m })),
        cf: s.cf.map((c) => ({ ...c, rect: { ...c.rect } })),
        freezeRows: s.freezeRows,
        freezeCols: s.freezeCols,
        tabColor: s.tabColor,
        filter: s.filter ? { rect: { ...s.filter.rect }, cols: Array.from(s.filter.cols).map(([k, v]) => [k, Array.from(v)] as [number, string[]]) } : undefined,
      })),
      names: Array.from(this.names.entries()).map(([k, v]) => [k, { ...v }]),
    };
  }
  restore(snap: WbSnap): void {
    this.sheets = snap.sheets.map((ss) => {
      const s = ss.sheet;
      s.name = ss.name;
      s.cells = new Map(ss.cells);
      s.cols.sizes = new Map(ss.cols);
      s.cols.touch();
      s.rows.sizes = new Map(ss.rows);
      s.rows.touch();
      s.merges = ss.merges.map((m) => ({ ...m }));
      s.cf = ss.cf.map((c) => ({ ...c, rect: { ...c.rect } }));
      s.freezeRows = ss.freezeRows;
      s.freezeCols = ss.freezeCols;
      s.tabColor = ss.tabColor;
      s.filter = ss.filter ? { rect: { ...ss.filter.rect }, cols: new Map(ss.filter.cols.map(([k, v]) => [k, new Set(v)])) } : undefined;
      s.recomputeBounds();
      this.sheetsById.set(s.id, s);
      return s;
    });
    this.names = new Map(snap.names.map(([k, v]) => [k, { ...v }]));
    this.nameAstCache.clear();
    this.rebuildDeps();
    this.recalcAll();
  }

  // ---------- cell access ----------
  getCell(sheet: Sheet, r: number, c: number): Cell | undefined {
    return sheet.cells.get(cellKey(r, c));
  }
  getValue(sheet: Sheet, r: number, c: number): Value {
    const cell = sheet.cells.get(cellKey(r, c));
    return cell ? cell.v : null;
  }

  private putRaw(sheet: Sheet, ck: number, cell: Cell | undefined): void {
    const fk = sheet.id * SHEET_BASE + ck;
    const old = sheet.cells.get(ck);
    if (old && old.f !== undefined) this.unregister(fk);
    if (!cell || (cell.v === null && cell.f === undefined && cell.s === undefined)) {
      if (old && (keyRow(ck) >= sheet.maxRow || keyCol(ck) >= sheet.maxCol)) this.boundsStale.add(sheet);
      sheet.cells.delete(ck);
      return;
    }
    sheet.cells.set(ck, cell);
    const r = keyRow(ck);
    const c = keyCol(ck);
    if (r > sheet.maxRow) sheet.maxRow = r;
    if (c > sheet.maxCol) sheet.maxCol = c;
    if (cell.f !== undefined) this.register(sheet, fk, cell);
  }

  /** Low-level cell write with undo journaling and recalculation bookkeeping. */
  setCell(sheet: Sheet, r: number, c: number, cell: Cell | undefined): void {
    const ck = cellKey(r, c);
    const fk = sheet.id * SHEET_BASE + ck;
    if (this.tx && !this.tx.journal.has(fk)) this.tx.journal.set(fk, sheet.cells.get(ck));
    if (this.tx) {
      this.tx.sheetId = sheet.id;
      this.tx.rect = this.tx.rect ? union(this.tx.rect, { r1: r, c1: c, r2: r, c2: c }) : { r1: r, c1: c, r2: r, c2: c };
    }
    this.putRaw(sheet, ck, cell);
    this.pending.push(fk);
    this.dataVersion++;
    if (!this.tx) this.flush();
  }

  /** Set the value (and optionally formula) keeping the existing style. */
  setValue(sheet: Sheet, r: number, c: number, v: Value, f?: string): void {
    const old = sheet.cells.get(cellKey(r, c));
    const cell: Cell = { v, s: old?.s };
    if (f !== undefined) cell.f = f;
    this.setCell(sheet, r, c, cell);
  }
  setStyle(sheet: Sheet, r: number, c: number, s: number | undefined): void {
    const old = sheet.cells.get(cellKey(r, c));
    if (!old && !s) return;
    this.setCell(sheet, r, c, { ...(old ?? { v: null }), s: s || undefined, ast: old?.ast });
  }
  clearCell(sheet: Sheet, r: number, c: number, keepStyle = true): void {
    const old = sheet.cells.get(cellKey(r, c));
    if (!old) return;
    if (keepStyle && old.s) this.setCell(sheet, r, c, { v: null, s: old.s });
    else this.setCell(sheet, r, c, undefined);
  }

  // ---------- dependency graph ----------
  private register(sheet: Sheet, fk: number, cell: Cell): void {
    if (cell.ast === undefined) {
      try {
        cell.ast = parseFormula(cell.f!);
      } catch (e) {
        if (!(e instanceof ParseError || e instanceof RangeError)) throw e;
        cell.ast = null;
      }
    }
    const rec: DepRec = { pts: [], rngs: [] };
    this.deps.set(fk, rec);
    if (!cell.ast) return;
    const refs: DepRef[] = [];
    const names: string[] = [];
    const vol = { v: false };
    collectRefs(cell.ast, refs, names, vol);
    const seenNames = new Set<string>();
    const addNames = (list: string[], depth: number) => {
      for (const n of list) {
        const key = n.toLowerCase();
        if (seenNames.has(key) || depth > 8) continue;
        seenNames.add(key);
        const def = this.names.get(key);
        if (!def) continue;
        const ast = this.nameAst(def);
        if (!ast) continue;
        const sub: string[] = [];
        collectRefs(ast, refs, sub, vol);
        addNames(sub, depth + 1);
      }
    };
    addNames(names, 0);
    if (vol.v) this.volatileSet.add(fk);
    for (const ref of refs) {
      const sh = ref.sheet === undefined ? sheet : this.sheetByName(ref.sheet);
      if (!sh) continue;
      if (ref.r1 === ref.r2 && ref.c1 === ref.c2) {
        const pk = sh.id * SHEET_BASE + cellKey(ref.r1, ref.c1);
        rec.pts.push(pk);
        let set = this.pointDeps.get(pk);
        if (!set) this.pointDeps.set(pk, (set = new Set()));
        set.add(fk);
      } else {
        const rd: RangeDep = { f: fk, sheet: sh.id, r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 };
        rec.rngs.push(rd);
        const b1 = ref.r1 >> 6;
        const b2 = ref.r2 >> 6;
        if (b2 - b1 > 64) {
          let set = this.tallRanges.get(sh.id);
          if (!set) this.tallRanges.set(sh.id, (set = new Set()));
          set.add(rd);
        } else
          for (let b = b1; b <= b2; b++) {
            const bk = sh.id * 65536 + b;
            let set = this.rangeBuckets.get(bk);
            if (!set) this.rangeBuckets.set(bk, (set = new Set()));
            set.add(rd);
          }
      }
    }
  }

  private unregister(fk: number): void {
    const rec = this.deps.get(fk);
    if (!rec) return;
    this.deps.delete(fk);
    this.volatileSet.delete(fk);
    for (const pk of rec.pts) {
      const set = this.pointDeps.get(pk);
      if (set) {
        set.delete(fk);
        if (!set.size) this.pointDeps.delete(pk);
      }
    }
    for (const rd of rec.rngs) {
      const b1 = rd.r1 >> 6;
      const b2 = rd.r2 >> 6;
      if (b2 - b1 > 64) this.tallRanges.get(rd.sheet)?.delete(rd);
      else
        for (let b = b1; b <= b2; b++) {
          const bk = rd.sheet * 65536 + b;
          const set = this.rangeBuckets.get(bk);
          if (set) {
            set.delete(rd);
            if (!set.size) this.rangeBuckets.delete(bk);
          }
        }
    }
  }

  rebuildDeps(): void {
    this.deps.clear();
    this.pointDeps.clear();
    this.rangeBuckets.clear();
    this.tallRanges.clear();
    this.volatileSet.clear();
    for (const s of this.sheets)
      for (const [ck, cell] of s.cells) {
        if (cell.f === undefined) continue;
        if (this.names.size) cell.ast = undefined; // names may resolve differently now
        this.register(s, s.id * SHEET_BASE + ck, cell);
      }
  }

  private dependentsOf(fk: number, out: (d: number) => void): void {
    const pts = this.pointDeps.get(fk);
    if (pts) for (const d of pts) out(d);
    const sid = Math.floor(fk / SHEET_BASE);
    const ck = fk % SHEET_BASE;
    const r = keyRow(ck);
    const c = keyCol(ck);
    const b = this.rangeBuckets.get(sid * 65536 + (r >> 6));
    if (b)
      for (const rd of b) if (r >= rd.r1 && r <= rd.r2 && c >= rd.c1 && c <= rd.c2) out(rd.f);
    const t = this.tallRanges.get(sid);
    if (t) for (const rd of t) if (r >= rd.r1 && r <= rd.r2 && c >= rd.c1 && c <= rd.c2) out(rd.f);
  }

  // ---------- recalculation ----------
  private boundsStale = new Set<Sheet>();
  private flush(): void {
    if (this.tx) return;
    if (this.boundsStale.size) {
      for (const sh of this.boundsStale) sh.recomputeBounds();
      this.boundsStale.clear();
    }
    if (this.fullRecalcPending) {
      this.fullRecalcPending = false;
      this.pending = [];
      this.recalcAll();
      return;
    }
    if (!this.pending.length && !this.volatileSet.size) return;
    const changed = this.pending;
    this.pending = [];
    if (this.autoCalc) this.recalc(changed);
  }

  /** Recalculate everything (after open, structural changes, or manual F9). */
  recalcAll(): void {
    const dirty = new Set<number>();
    for (const s of this.sheets) for (const [ck, cell] of s.cells) if (cell.f !== undefined) dirty.add(s.id * SHEET_BASE + ck);
    this.evalDirty(dirty);
    this.dataVersion++;
  }

  recalc(changed: number[]): void {
    const dirty = new Set<number>();
    const queue: number[] = [];
    const add = (k: number) => {
      if (dirty.has(k)) return;
      dirty.add(k);
      queue.push(k);
    };
    for (const k of changed) {
      const sheet = this.sheetsById.get(Math.floor(k / SHEET_BASE));
      const cell = sheet?.cells.get(k % SHEET_BASE);
      if (cell && cell.f !== undefined) add(k);
      else queue.push(k);
    }
    for (const k of this.volatileSet) add(k);
    for (let qi = 0; qi < queue.length; qi++) this.dependentsOf(queue[qi], add);
    this.evalDirty(dirty);
  }

  private evalDirty(dirty: Set<number>): void {
    if (!dirty.size) return;
    let dirtyList: number[] | null = null;
    const precsOf = (fk: number): number[] => {
      const rec = this.deps.get(fk);
      if (!rec) return [];
      const out: number[] = [];
      for (const p of rec.pts) if (dirty.has(p)) out.push(p);
      for (const rd of rec.rngs) {
        const area = (rd.r2 - rd.r1 + 1) * (rd.c2 - rd.c1 + 1);
        if (area <= 2048) {
          const base = rd.sheet * SHEET_BASE;
          for (let r = rd.r1; r <= rd.r2; r++) for (let c = rd.c1; c <= rd.c2; c++) {
            const k = base + cellKey(r, c);
            if (dirty.has(k)) out.push(k);
          }
        } else {
          if (!dirtyList) dirtyList = Array.from(dirty);
          for (const k of dirtyList) {
            if (Math.floor(k / SHEET_BASE) !== rd.sheet) continue;
            const ck = k % SHEET_BASE;
            const r = keyRow(ck);
            const c = keyCol(ck);
            if (r >= rd.r1 && r <= rd.r2 && c >= rd.c1 && c <= rd.c2) out.push(k);
          }
        }
      }
      return out;
    };
    const state = new Map<number, number>();
    const circ = new Set<number>();
    interface Frame { k: number; ps: number[]; i: number }
    for (const root of dirty) {
      if (state.has(root)) continue;
      const stack: Frame[] = [{ k: root, ps: precsOf(root), i: 0 }];
      state.set(root, 1);
      while (stack.length) {
        const f = stack[stack.length - 1];
        if (f.i < f.ps.length) {
          const p = f.ps[f.i++];
          const st = state.get(p);
          if (st === undefined) {
            state.set(p, 1);
            stack.push({ k: p, ps: precsOf(p), i: 0 });
          } else if (st === 1) {
            for (let j = stack.length - 1; j >= 0; j--) {
              circ.add(stack[j].k);
              if (stack[j].k === p) break;
            }
          }
        } else {
          stack.pop();
          state.set(f.k, 2);
          this.evalFormulaCell(f.k, circ.has(f.k));
        }
      }
    }
  }

  private evalFormulaCell(fk: number, isCirc: boolean): void {
    const sheet = this.sheetsById.get(Math.floor(fk / SHEET_BASE));
    if (!sheet) return;
    const ck = fk % SHEET_BASE;
    const cell = sheet.cells.get(ck);
    if (!cell || cell.f === undefined) return;
    if (isCirc) {
      cell.v = E.CIRC;
      return;
    }
    if (cell.ast === undefined) {
      try {
        cell.ast = parseFormula(cell.f);
      } catch {
        cell.ast = null;
      }
    }
    if (!cell.ast) {
      cell.v = E.NAME;
      return;
    }
    const ctx: Ctx = { wb: this, sheet, row: keyRow(ck), col: keyCol(ck) };
    try {
      cell.v = evaluate(cell.ast, ctx);
    } catch {
      cell.v = E.VALUE;
    }
  }

  /** Mark that the whole workbook must be recalculated at the end of the current transaction. */
  requestFullRecalc(): void {
    this.fullRecalcPending = true;
    if (!this.tx) this.flush();
  }

  // ---------- helpers ----------
  allFormulaCells(): number {
    let n = 0;
    for (const s of this.sheets) for (const c of s.cells.values()) if (c.f !== undefined) n++;
    return n;
  }
}

function union(a: Rect, b: Rect): Rect {
  return { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
}

export { MAX_COLS, MAX_ROWS, FErr };
