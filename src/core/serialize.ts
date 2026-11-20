// Native Sheetly format (.sheetly, JSON) — also used for autosave recovery.
import { Rect, cellKey, keyCol, keyRow } from './addr';
import { CFRule, CellStyle, Workbook } from './model';
import { E, FErr, Value, errFromCode } from './values';

type JV = number | string | boolean | null | { e: string };

const enc = (v: Value): JV => (v instanceof FErr ? { e: v.code } : typeof v === 'number' && !Number.isFinite(v) ? { e: '#NUM!' } : v);
const dec = (v: JV): Value => (v !== null && typeof v === 'object' ? errFromCode(v.e) ?? E.VALUE : v);

interface SheetJSON {
  name: string;
  cells: [number, number, JV, string | undefined, number | undefined][];
  cols: [number, number][];
  rows: [number, number][];
  defRow: number;
  merges: Rect[];
  cf: CFRule[];
  freeze: [number, number];
  tab?: string;
  grid: boolean;
  filter?: { rect: Rect; cols: [number, string[]][] };
  sel?: [number, number, number, number];
}
interface BookJSON {
  v: 1;
  styles: CellStyle[];
  sheets: SheetJSON[];
  names: [string, string][];
  active: number;
}

export function toJSON(wb: Workbook, active = 0): string {
  const book: BookJSON = {
    v: 1,
    styles: wb.styles,
    names: Array.from(wb.names.values()).map((n) => [n.name, n.refers]),
    active,
    sheets: wb.sheets.map((s) => {
      const cells: SheetJSON['cells'] = [];
      for (const [k, c] of s.cells) cells.push([keyRow(k), keyCol(k), enc(c.v), c.f, c.s]);
      return {
        name: s.name,
        cells,
        cols: Array.from(s.cols.sizes),
        rows: Array.from(s.rows.sizes),
        defRow: s.rows.def,
        merges: s.merges,
        cf: s.cf,
        freeze: [s.freezeRows, s.freezeCols] as [number, number],
        tab: s.tabColor,
        grid: s.showGrid,
        filter: s.filter
          ? { rect: s.filter.rect, cols: Array.from(s.filter.cols).map(([c, set]) => [c, Array.from(set)] as [number, string[]]) }
          : undefined,
        sel: [s.sel.r, s.sel.c, s.sel.r2, s.sel.c2] as [number, number, number, number],
      };
    }),
  };
  return JSON.stringify(book);
}

export function fromJSON(text: string): { wb: Workbook; active: number } {
  const book = JSON.parse(text) as BookJSON;
  if (book.v !== 1) throw new Error('Unsupported Sheetly file version');
  const wb = new Workbook();
  wb.sheets = [];
  const idMap: number[] = [0];
  for (let i = 1; i < book.styles.length; i++) idMap[i] = wb.styleIndex(book.styles[i]);
  for (const sj of book.sheets) {
    const s = wb.addSheet(sj.name);
    for (const [r, c, v, f, st] of sj.cells) {
      const cell: { v: Value; f?: string; s?: number } = { v: dec(v), s: st ? idMap[st] : undefined };
      if (typeof f === "string") cell.f = f;
      s.cells.set(cellKey(r, c), cell);
    }
    for (const [i, w] of sj.cols) s.cols.set(i, w);
    s.rows.def = sj.defRow || 20;
    for (const [i, h] of sj.rows) s.rows.set(i, h);
    s.merges = sj.merges;
    s.cf = sj.cf.map((r) => ({ ...r, id: wb.nextRuleId++ }));
    s.freezeRows = sj.freeze[0];
    s.freezeCols = sj.freeze[1];
    s.tabColor = sj.tab;
    s.showGrid = sj.grid !== false;
    if (sj.filter) s.filter = { rect: sj.filter.rect, cols: new Map(sj.filter.cols.map(([c, v]) => [c, new Set(v)])) };
    if (sj.sel) s.sel = { r: sj.sel[0], c: sj.sel[1], r2: sj.sel[2], c2: sj.sel[3] };
    s.recomputeBounds();
  }
  if (!wb.sheets.length) wb.addSheet('Sheet1');
  for (const [n, r] of book.names) wb.setName(n, r);
  wb.rebuildDeps();
  wb.recalcAll();
  wb.clearHistory();
  return { wb, active: Math.min(book.active ?? 0, wb.sheets.length - 1) };
}
