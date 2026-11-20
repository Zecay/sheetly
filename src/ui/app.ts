// Application controller: document state, commands, formula bar, sheet tabs, status bar, files.
import { MAX_COLS, MAX_ROWS, Rect, a1, keyCol, keyRow, normRect, parseRect, rectText, colName, parseCell } from '../core/addr';
import { CFEngine } from '../core/cf';
import { readCSV, writeCSV } from '../core/csv';
import { FUNCS } from '../core/eval';
import { editText } from '../core/input';
import { Border, Cell, CFRule, Sheet, StylePatch, Workbook } from '../core/model';
import { changeDecimals, FMT_ACCOUNTING, FMT_CURRENCY, FMT_PERCENT } from '../core/numfmt';
import {
  ClipData, FormulaError, PasteMode, SortKey, applyFilter, applyInput, applyStyle, clearFilter, clearRange, copyRange, currentRegion, deleteCols, deleteRows,
  deleteSheet, displayText, duplicateSheet, eachCell, fillRange, insertCols, insertRows, moveRange, moveSheet, pasteClip, pasteText, removeFilter, renameSheet, sortRange,
} from '../core/ops';
import { fromJSON, toJSON } from '../core/serialize';
import { readXlsx, writeXlsx } from '../core/xlsx';
import { FErr, todaySerial, nowSerial } from '../core/values';
import { Grid, GridHost, DARK, LIGHT } from './grid';
import { appMark, icon } from './icons';
import { h, messageBox, openDialog, promptBox, showMenu, showMenuBelow, showColorMenu, MenuItem, toast, closeMenus } from './dom';
import * as fs from './fsapi';
import { Ribbon } from './ribbon';
import * as dlg from './dialogs';
import { Backstage } from './backstage';
import { printSheet } from './print';

export type DocKind = 'xlsx' | 'csv' | 'sheetly';

export interface CellStyleDef {
  name: string;
  group: string;
  patch: StylePatch;
}

const B = (color = '#7f7f7f'): Border => ({ style: 'thin', color });

export const CELL_STYLES: CellStyleDef[] = [
  { group: 'Good, Bad and Neutral', name: 'Normal', patch: { font: null, size: null, bold: null, italic: null, underline: null, strike: null, color: null, fill: null, h: null, v: null, wrap: null, indent: null, fmt: null, bl: null, br: null, bt: null, bb: null } },
  { group: 'Good, Bad and Neutral', name: 'Bad', patch: { color: '#9c0006', fill: '#ffc7ce' } },
  { group: 'Good, Bad and Neutral', name: 'Good', patch: { color: '#006100', fill: '#c6efce' } },
  { group: 'Good, Bad and Neutral', name: 'Neutral', patch: { color: '#9c5700', fill: '#ffeb9c' } },
  { group: 'Data and Model', name: 'Calculation', patch: { color: '#fa7d00', bold: true, fill: '#f2f2f2', bl: B(), br: B(), bt: B(), bb: B() } },
  { group: 'Data and Model', name: 'Check Cell', patch: { color: '#ffffff', bold: true, fill: '#a5a5a5', bl: { style: 'double', color: '#3f3f3f' }, br: { style: 'double', color: '#3f3f3f' }, bt: { style: 'double', color: '#3f3f3f' }, bb: { style: 'double', color: '#3f3f3f' } } },
  { group: 'Data and Model', name: 'Explanatory Text', patch: { color: '#7f7f7f', italic: true } },
  { group: 'Data and Model', name: 'Input', patch: { color: '#3f3f76', fill: '#ffcc99', bl: B('#7f7f7f'), br: B('#7f7f7f'), bt: B('#7f7f7f'), bb: B('#7f7f7f') } },
  { group: 'Data and Model', name: 'Linked Cell', patch: { color: '#fa7d00', bb: { style: 'double', color: '#ff8001' } } },
  { group: 'Data and Model', name: 'Note', patch: { fill: '#ffffcc', bl: B('#b2b2b2'), br: B('#b2b2b2'), bt: B('#b2b2b2'), bb: B('#b2b2b2') } },
  { group: 'Data and Model', name: 'Output', patch: { color: '#3f3f3f', bold: true, fill: '#f2f2f2', bl: B('#3f3f3f'), br: B('#3f3f3f'), bt: B('#3f3f3f'), bb: B('#3f3f3f') } },
  { group: 'Data and Model', name: 'Warning Text', patch: { color: '#ff0000' } },
  { group: 'Titles and Headings', name: 'Heading 1', patch: { bold: true, size: 15, color: '#44546a', bb: { style: 'thick', color: '#4472c4' } } },
  { group: 'Titles and Headings', name: 'Heading 2', patch: { bold: true, size: 13, color: '#44546a', bb: { style: 'thick', color: '#a9bdd6' } } },
  { group: 'Titles and Headings', name: 'Heading 3', patch: { bold: true, size: 11, color: '#44546a', bb: { style: 'medium', color: '#8ea9db' } } },
  { group: 'Titles and Headings', name: 'Heading 4', patch: { bold: true, size: 11, color: '#44546a' } },
  { group: 'Titles and Headings', name: 'Title', patch: { size: 18, color: '#44546a', font: 'Cambria' } },
  { group: 'Titles and Headings', name: 'Total', patch: { bold: true, bt: B('#4472c4'), bb: { style: 'double', color: '#4472c4' } } },
  { group: 'Number Format', name: 'Comma', patch: { fmt: '#,##0.00' } },
  { group: 'Number Format', name: 'Currency', patch: { fmt: FMT_CURRENCY } },
  { group: 'Number Format', name: 'Percent', patch: { fmt: FMT_PERCENT } },
];

export class App implements GridHost {
  wb = new Workbook();
  cf = new CFEngine(this.wb);
  dark = false;
  gridlines = true;
  headings = true;
  showFormulas = false;
  showFormulaBar = true;
  orientation: 'portrait' | 'landscape' = 'portrait';
  printGridlines = false;
  doc = { path: null as string | null, name: 'Book1', dirty: false, kind: 'xlsx' as DocKind };
  clip: { data: ClipData; wb: Workbook; sheet: Sheet; text: string } | null = null;
  painter: { style: number; sticky: boolean } | null = null;
  grid!: Grid;
  ribbon!: Ribbon;
  backstage!: Backstage;
  findState = { what: '', replace: '', matchCase: false, wholeCell: false, inFormulas: false, allSheets: false, byRows: true };
  private statusMsg = '';
  private refreshQueued = false;
  private untitledCount = 1;
  private changeCount = 0; // bumped on every workbook change
  private autosavedCount = 0; // changeCount captured by the last successful autosave/save

  // DOM refs
  el = {
    tabs: document.getElementById('tabs')!,
    ribbon: document.getElementById('ribbon')!,
    namebox: document.getElementById('namebox') as HTMLInputElement,
    fbTa: document.getElementById('fb-ta') as HTMLTextAreaElement,
    fbMirror: document.getElementById('fb-mirror') as HTMLDivElement,
    fbInput: document.getElementById('fb-input')!,
    fbCancel: document.getElementById('fb-cancel')!,
    fbEnter: document.getElementById('fb-enter')!,
    sheetTabs: document.getElementById('sheet-tabs')!,
    stats: document.getElementById('sb-stats')!,
    mode: document.getElementById('sb-mode')!,
    msg: document.getElementById('sb-msg')!,
    zoom: document.getElementById('sb-zoom')!,
    title: document.getElementById('doc-title')!,
    formulabar: document.getElementById('formulabar')!,
  };

  constructor() {
    this.wb.onChange = () => this.onWorkbookChange();
    const dark = localStorage.getItem('sheetly.theme') === 'dark';
    this.setTheme(dark, false);
    const area = document.getElementById('grid-area')!;
    this.grid = new Grid(this, area, document.getElementById('vscroll')!, document.getElementById('hscroll')!, this.wb.sheets[0]);
    this.grid.editor.attachBar(this.el.fbTa, this.el.fbMirror);
    this.ribbon = new Ribbon(this);
    this.backstage = new Backstage(this);
    this.bindChrome();
    this.renderTabs();
    this.refreshAll();
    this.updateTitle();
    this.initStatusZoom();
    this.initWindowControls();
    window.setInterval(() => this.autosave(), 60000);
    (window as any).sheetly = this; // handy for debugging
  }

  get sheet(): Sheet {
    return this.grid.sheet;
  }
  get selRect(): Rect {
    return this.grid.range;
  }
  get activeCell(): { r: number; c: number } {
    return this.grid.active;
  }

  // ------------------------------------------------------------------ GridHost
  selectionChanged(): void {
    this.scheduleRefresh();
  }
  viewChanged(): void {
    this.updateZoomUI();
  }
  commitCell(r: number, c: number, text: string): string | null {
    try {
      this.wb.transact('Typing', () => applyInput(this.wb, this.sheet, r, c, text));
    } catch (e) {
      if (e instanceof FormulaError) return "There's a problem with this formula.\n\nTry clicking the Insert Function button to get help with function names and arguments.";
      throw e;
    }
    const cell = this.sheet.get(r, c);
    const st = this.wb.styleOf(cell?.s);
    if (st.wrap && !this.sheet.rows.sizes.has(r)) this.grid.autofitRows([r], false);
    return null;
  }
  contextMenu(x: number, y: number, kind: 'cell' | 'col' | 'row' | 'all'): void {
    const common: MenuItem[] = [
      { label: 'Cut', icon: 'cut', shortcut: 'Ctrl+X', cmd: () => this.run('cut') },
      { label: 'Copy', icon: 'copy', shortcut: 'Ctrl+C', cmd: () => this.run('copy') },
      { label: 'Paste', icon: 'paste', shortcut: 'Ctrl+V', cmd: () => this.run('paste') },
    ];
    if (kind === 'cell') {
      showMenu(x, y, [
        ...common,
        { label: 'Paste Special', submenu: [
          { label: 'Values', cmd: () => this.run('paste', 'values') },
          { label: 'Formulas', cmd: () => this.run('paste', 'formulas') },
          { label: 'Formats', cmd: () => this.run('paste', 'formats') },
          { label: 'Transpose', cmd: () => this.run('paste', 'transpose') },
          { sep: true },
          { label: 'Paste Special…', cmd: () => this.run('pasteSpecial') },
        ] },
        { sep: true },
        { label: 'Insert…', icon: 'insertCells', cmd: () => this.run('insertDialog') },
        { label: 'Delete…', icon: 'deleteCells', cmd: () => this.run('deleteDialog') },
        { label: 'Clear Contents', cmd: () => this.run('clear', 'contents') },
        { sep: true },
        { label: 'Sort', submenu: [
          { label: 'Sort A to Z', cmd: () => this.run('sort', 'asc') },
          { label: 'Sort Z to A', cmd: () => this.run('sort', 'desc') },
          { label: 'Custom Sort…', cmd: () => this.run('sortDialog') },
        ] },
        { label: 'Filter', icon: 'filter', cmd: () => this.run('filter') },
        { sep: true },
        { label: 'Format Cells…', icon: 'formatCells', shortcut: 'Ctrl+1', cmd: () => this.run('formatCells') },
      ]);
    } else if (kind === 'col') {
      showMenu(x, y, [
        ...common,
        { sep: true },
        { label: 'Insert', icon: 'colsInsert', cmd: () => this.run('insertCols') },
        { label: 'Delete', icon: 'deleteCells', cmd: () => this.run('deleteCols') },
        { label: 'Clear Contents', cmd: () => this.run('clear', 'contents') },
        { sep: true },
        { label: 'Format Cells…', icon: 'formatCells', cmd: () => this.run('formatCells') },
        { label: 'Column Width…', cmd: () => this.run('colWidth') },
        { label: 'AutoFit Column Width', cmd: () => this.run('autofitCols') },
        { label: 'Hide', cmd: () => this.run('hideCols') },
        { label: 'Unhide', cmd: () => this.run('unhideCols') },
      ]);
    } else if (kind === 'row') {
      showMenu(x, y, [
        ...common,
        { sep: true },
        { label: 'Insert', icon: 'rowsInsert', cmd: () => this.run('insertRows') },
        { label: 'Delete', icon: 'deleteCells', cmd: () => this.run('deleteRows') },
        { label: 'Clear Contents', cmd: () => this.run('clear', 'contents') },
        { sep: true },
        { label: 'Format Cells…', icon: 'formatCells', cmd: () => this.run('formatCells') },
        { label: 'Row Height…', cmd: () => this.run('rowHeight') },
        { label: 'AutoFit Row Height', cmd: () => this.run('autofitRows') },
        { label: 'Hide', cmd: () => this.run('hideRows') },
        { label: 'Unhide', cmd: () => this.run('unhideRows') },
      ]);
    }
  }
  fillDone(src: Rect, dir: 'down' | 'up' | 'right' | 'left', count: number): void {
    let dst: Rect | null = null;
    this.wb.transact('Fill', () => {
      dst = fillRange(this.wb, this.sheet, src, dir, count);
    });
    if (dst) {
      const d: Rect = dst;
      this.grid.selectRect({ r1: Math.min(src.r1, d.r1), c1: Math.min(src.c1, d.c1), r2: Math.max(src.r2, d.r2), c2: Math.max(src.c2, d.c2) });
    }
  }
  moveDone(src: Rect, r: number, c: number, copy: boolean): void {
    let dst: Rect = src;
    if (copy) {
      const clip = copyRange(this.wb, this.sheet, src);
      this.wb.transact('Copy cells', () => {
        dst = pasteClip(this.wb, this.sheet, r, c, clip, this.wb);
      });
    } else {
      this.wb.transact('Move cells', () => {
        dst = moveRange(this.wb, this.sheet, src, this.sheet, r, c);
      });
    }
    this.grid.selectRect(dst);
  }
  filterButton(col: number, anchor: DOMRect): void {
    dlg.showFilterPanel(this, col, anchor);
  }
  definedNames(): string[] {
    return Array.from(this.wb.names.values()).map((n) => n.name);
  }
  clipboardRect(): { sheet: Sheet; rect: Rect } | null {
    return this.clip ? { sheet: this.clip.sheet, rect: this.clip.data.rect } : null;
  }
  painterActive(): boolean {
    return !!this.painter;
  }
  /** True (and tells the user) when the rectangles hold more than `limit` cells. */
  tooLarge(rects: Rect[], limit: number, what: string): boolean {
    let n = 0;
    for (const r of rects) n += (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
    if (n <= limit) return false;
    void messageBox('Sheetly', `${what} is limited to ${limit.toLocaleString('en-US')} cells at a time. Select a smaller range and try again.`);
    return true;
  }
  painterApply(rect: Rect): void {
    const p = this.painter;
    if (!p) return;
    if (this.tooLarge([rect], 500_000, 'Format Painter')) return;
    this.wb.transact('Format Painter', () => {
      for (let r = rect.r1; r <= rect.r2; r++)
        for (let c = rect.c1; c <= rect.c2; c++) this.wb.setStyle(this.sheet, r, c, p.style || undefined);
    });
    if (!p.sticky) this.painter = null;
    this.ribbon.update();
    this.grid.area.style.cursor = 'cell';
  }
  editStateChanged(): void {
    const ed = this.grid.editor;
    this.el.fbCancel.hidden = !ed.active;
    this.el.fbEnter.hidden = !ed.active;
    this.el.fbTa.classList.toggle('editing', ed.active);
    this.el.mode.textContent = ed.active ? (ed.mode === 'point' ? 'Point' : ed.mode === 'edit' ? 'Edit' : 'Enter') : 'Ready';
    if (!ed.active) this.refreshFormulaBar();
    else this.el.fbMirror.innerHTML = '';
    this.updateFormulaBarMirror();
  }
  private updateFormulaBarMirror() {
    const ed = this.grid.editor;
    if (ed.active) ed.syncBar();
  }
  editorFont(r: number, c: number): { css: string } {
    const st = this.wb.styleOf(this.sheet.get(r, c)?.s);
    const px = (st.size ?? 11) * (4 / 3) * this.grid.zoom;
    return { css: `${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}${px}px ${st.font ? `"${st.font}", ` : ''}Calibri, Carlito, "Segoe UI", Arial, sans-serif` };
  }
  editorText(r: number, c: number): string {
    const cell = this.sheet.get(r, c);
    return editText(cell, this.wb.styleOf(cell?.s));
  }
  message(text: string): void {
    this.statusMsg = text;
    this.el.msg.textContent = text;
    if (text.startsWith("There's a problem")) void messageBox('Sheetly', text);
  }
  clearSelection(): void {
    this.run('clear', 'contents');
  }

  // ------------------------------------------------------------------ chrome wiring
  private bindChrome(): void {
    const { namebox, fbTa } = this.el;
    // name box
    namebox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.goToText(namebox.value);
        this.grid.focusGrid();
      } else if (e.key === 'Escape') {
        this.refreshNameBox();
        this.grid.focusGrid();
      }
      e.stopPropagation();
    });
    namebox.addEventListener('focus', () => namebox.select());
    document.getElementById('namebox-arrow')!.addEventListener('click', (e) => {
      const names = Array.from(this.wb.names.values());
      const items: MenuItem[] = names.length
        ? names.map((n) => ({ label: n.name, cmd: () => this.goToText(n.name) }))
        : [{ label: '(no names defined)', disabled: true }];
      showMenuBelow(e.currentTarget as HTMLElement, items);
    });
    // formula bar
    const startFromBar = () => {
      if (this.grid.editor.active) return;
      this.grid.beginEdit('edit');
      fbTa.focus();
    };
    fbTa.addEventListener('focus', startFromBar);
    fbTa.addEventListener('mousedown', () => {
      if (!this.grid.editor.active) setTimeout(startFromBar, 0);
    });
    document.getElementById('fb-cancel')!.addEventListener('click', () => this.grid.cancelEdit());
    document.getElementById('fb-enter')!.addEventListener('click', () => this.grid.commitEdit('none'));
    document.getElementById('fb-fx')!.addEventListener('click', () => this.run('insertFunction'));
    document.getElementById('fb-expand')!.addEventListener('click', () => this.toggleFormulaBarExpand());
    document.getElementById('fb-expand')!.innerHTML = icon('chevron', 12);
    document.getElementById('namebox-arrow')!.innerHTML = icon('chevron', 10);
    document.getElementById('fb-cancel')!.innerHTML = icon('close', 14);
    document.getElementById('fb-enter')!.innerHTML = icon('check', 14);
    document.getElementById('fb-fx')!.innerHTML = icon('fx', 18);
    // sheet bar
    document.getElementById('sheet-add')!.innerHTML = icon('plus', 14);
    document.getElementById('sheet-add')!.addEventListener('click', () => this.run('addSheet'));
    document.getElementById('sb-left')!.innerHTML = icon('left', 10);
    document.getElementById('sb-right')!.innerHTML = icon('right', 10);
    document.getElementById('sb-left')!.addEventListener('click', () => (this.el.sheetTabs.scrollLeft -= 120));
    document.getElementById('sb-right')!.addEventListener('click', () => (this.el.sheetTabs.scrollLeft += 120));
    // status bar views
    const views = document.getElementById('sb-views')!;
    const mk = (ic: string, title: string, on: boolean, fn: () => void) => {
      const b = h('button', { type: 'button', title, class: on ? 'on' : '', html: icon(ic, 14) });
      b.addEventListener('click', fn);
      return b;
    };
    views.append(
      mk('gridlines', 'Normal', true, () => {}),
      mk('print', 'Page Layout (print preview)', false, () => this.run('print')),
      mk('table', 'Page Break Preview (not available yet)', false, () => toast('Page Break Preview is not available yet.')),
    );
    // keyboard & clipboard
    window.addEventListener('keydown', (e) => this.onKeyDown(e));
    document.addEventListener('copy', (e) => this.onCopyEvent(e, false));
    document.addEventListener('cut', (e) => this.onCopyEvent(e, true));
    document.addEventListener('paste', (e) => this.onPasteEvent(e));
    // clicking blank chrome returns focus to the grid
    this.el.sheetTabs.addEventListener('mousedown', () => {});
    window.addEventListener('beforeunload', (e) => {
      if (this.doc.dirty && !fs.isTauri) e.preventDefault();
    });
  }

  private isTextTarget(t: EventTarget | null): boolean {
    return t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t instanceof HTMLSelectElement;
  }

  private onCopyEvent(e: ClipboardEvent, cut: boolean): void {
    const t = e.target;
    if (this.grid.editor.active || (this.isTextTarget(t) && t !== this.grid.area)) return;
    if (document.querySelector('.dlg-backdrop:not(.modeless)')) return;
    e.preventDefault();
    const clip = this.makeClip(cut);
    if (!clip) return;
    e.clipboardData?.setData('text/plain', clip.text);
    e.clipboardData?.setData('text/html', this.clipHtml(clip.data));
  }

  private onPasteEvent(e: ClipboardEvent): void {
    const t = e.target;
    if (this.grid.editor.active || (this.isTextTarget(t) && t !== this.grid.area)) return;
    if (document.querySelector('.dlg-backdrop:not(.modeless)')) return;
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') ?? '';
    this.doPaste(text, 'all');
  }

  // ------------------------------------------------------------------ keyboard
  private onKeyDown(e: KeyboardEvent): void {
    if (document.querySelector('.dlg-backdrop:not(.modeless)') || document.querySelector('.menu')) {
      return;
    }
    const target = e.target as HTMLElement;
    const ed = this.grid.editor;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key;
    const inEditor = ed.active && (target === ed.ta || target === this.el.fbTa);
    if (inEditor) {
      if (ctrl && k === 'Enter') {
        e.preventDefault();
        this.fillSelectionWithEdit();
        return;
      }
      if (ctrl && e.shiftKey && (k === 'U' || k === 'u')) {
        e.preventDefault();
        this.toggleFormulaBarExpand();
        return;
      }
      if (this.grid.handleKey(e)) e.preventDefault();
      return;
    }
    if (this.isTextTarget(target) && target !== this.grid.area) return;
    if (ed.active) return;
    // global shortcuts
    if (this.shortcut(e)) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    if (this.grid.handleKey(e)) {
      // printable characters must reach the freshly focused editor
      if (!(k.length === 1 && !ctrl && !e.altKey)) e.preventDefault();
    }
  }

  private shortcut(e: KeyboardEvent): boolean {
    const ctrl = e.ctrlKey || e.metaKey;
    const shift = e.shiftKey;
    const alt = e.altKey;
    const k = e.key;
    const lk = k.length === 1 ? k.toLowerCase() : k;
    if (k === 'F2' || k === 'Delete' || k === 'Backspace' || k === 'Enter' || k === 'Tab') return false;
    if (k === 'Escape') {
      if (this.clip) {
        this.clip = null;
        this.grid.invalidate();
        return true;
      }
      if (this.painter) {
        this.painter = null;
        this.ribbon.update();
        return true;
      }
      return false;
    }
    if (k === 'F5' || (ctrl && lk === 'g')) return this.do('goto');
    if (k === 'F9') return this.do('calcNow');
    if (k === 'F11' && shift) return this.do('addSheet');
    if (alt && k === '=') return this.do('autosum', 'SUM');
    if (alt) return false;
    if (!ctrl) return false;
    if (k === 'PageDown' || k === 'PageUp') {
      const i = this.wb.sheets.indexOf(this.sheet) + (k === 'PageDown' ? 1 : -1);
      if (i >= 0 && i < this.wb.sheets.length) this.setActiveSheet(this.wb.sheets[i]);
      return true;
    }
    switch (lk) {
      case 'z': return this.do('undo');
      case 'y': return this.do('redo');
      // Ctrl+C / X / V are left to the browser so the native copy/cut/paste events fire (no clipboard permission prompt)
      case 'b': return this.do('bold');
      case 'i': return this.do('italic');
      case 'u': return this.do('underline');
      case '5': return this.do('strike');
      case '1': return this.do('formatCells');
      case 'f': return this.do('find');
      case 'h': return this.do('replace');
      case 's': return this.do(shift ? 'saveAs' : 'save');
      case 'o': return this.do('open');
      case 'n': return this.do('new');
      case 'p': return this.do('print');
      case 'd': return this.do('fill', 'down');
      case 'r': return this.do('fill', 'right');
      case ';': return this.do('insertDate');
      case ':': return this.do('insertTime');
      case 'l':
        if (shift) return this.do('filter');
        return false;
      case '`':
      case '~':
        return this.do('showFormulas');
      case '9': return this.do(shift ? 'unhideRows' : 'hideRows');
      case '0': return this.do(shift ? 'unhideCols' : 'hideCols');
      case '-':
        return this.do('deleteDialog');
      case '+':
      case '=':
        if (shift || k === '+') return this.do('insertDialog');
        return false;
      case '$': return this.do('numFmt', FMT_CURRENCY);
      case '%': return this.do('numFmt', FMT_PERCENT);
      case '!': return this.do('numFmt', '#,##0.00');
      case '#': return this.do('numFmt', 'm/d/yyyy');
      case '@': return this.do('numFmt', 'h:mm AM/PM');
      case '^': return this.do('numFmt', 'General');
    }
    // shifted digit keys on layouts that report digits with shift
    if (shift) {
      if (k === '4') return this.do('numFmt', FMT_CURRENCY);
      if (k === '5') return this.do('numFmt', FMT_PERCENT);
    }
    return false;
  }
  private do(cmd: string, arg?: unknown): boolean {
    void this.run(cmd, arg);
    return true;
  }

  private fillSelectionWithEdit(): void {
    const ed = this.grid.editor;
    const text = ed.text;
    const rects = this.grid.ranges.slice();
    ed.stop();
    this.editStateChanged();
    if (this.tooLarge(rects, 500_000, 'Filling a selection')) return;
    try {
      this.wb.transact('Fill', () => {
        for (const rg of rects)
          for (let r = rg.r1; r <= rg.r2; r++)
            for (let c = rg.c1; c <= rg.c2; c++) {
              if (text.startsWith('=')) {
                // relative references adjust for each target cell
                const base = this.grid.active;
                applyInput(this.wb, this.sheet, r, c, '=' + shiftFormulaImpl(text.slice(1), r - base.r, c - base.c));
              } else applyInput(this.wb, this.sheet, r, c, text);
            }
      });
    } catch (e) {
      if (e instanceof FormulaError) void messageBox('Sheetly', "There's a problem with this formula.");
    }
    this.grid.focusGrid();
  }

  // ------------------------------------------------------------------ refresh
  scheduleRefresh(): void {
    if (this.refreshQueued) return;
    this.refreshQueued = true;
    queueMicrotask(() => {
      this.refreshQueued = false;
      this.refreshAll();
    });
  }

  refreshAll(): void {
    this.refreshNameBox();
    if (!this.grid.editor.active) this.refreshFormulaBar();
    this.refreshStats();
    this.ribbon.update();
    this.updateQat();
    this.grid.invalidate();
  }

  private onWorkbookChange(): void {
    this.doc.dirty = true;
    this.changeCount++;
    this.updateTitle();
    this.grid.invalidate();
    this.scheduleRefresh();
    this.renderTabsIfChanged();
  }

  refreshNameBox(): void {
    const { r, c } = this.grid.active;
    let name = a1(r, c);
    const rg = this.grid.range;
    // a defined name that exactly matches the selection wins
    const sel = this.sheet.name + '!' + rectAbs(rg);
    for (const n of this.wb.names.values()) {
      if (n.refers.replace(/'/g, '').toLowerCase() === sel.replace(/'/g, '').toLowerCase()) name = n.name;
    }
    if (document.activeElement !== this.el.namebox) this.el.namebox.value = name;
  }

  refreshFormulaBar(): void {
    const { r, c } = this.grid.active;
    const cell = this.sheet.get(r, c);
    const text = editText(cell, this.wb.styleOf(cell?.s));
    if (this.el.fbTa.value !== text) this.el.fbTa.value = text;
    this.el.fbMirror.innerHTML = '';
  }

  private toggleFormulaBarExpand(): void {
    const exp = this.el.fbInput.classList.toggle('expanded');
    document.getElementById('fb-expand')!.innerHTML = icon(exp ? 'chevron' : 'chevron', 12);
    (document.getElementById('fb-expand') as HTMLElement).style.transform = exp ? 'rotate(180deg)' : '';
  }

  private statsTimer = 0;
  refreshStats(): void {
    let total = 0;
    for (const rg of this.grid.ranges) total += (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    if (total > 100_000) {
      // big selections: coalesce and compute later (only populated cells are visited)
      if (this.statsTimer) return;
      this.el.stats.textContent = '';
      this.statsTimer = window.setTimeout(() => {
        this.statsTimer = 0;
        this.computeStats();
      }, 200);
      return;
    }
    if (this.statsTimer) {
      clearTimeout(this.statsTimer);
      this.statsTimer = 0;
    }
    this.computeStats();
  }

  private computeStats(): void {
    const sheet = this.sheet;
    let count = 0;
    let nums = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    let cells = 0;
    for (const rg of this.grid.ranges) {
      cells += (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
      eachCell(sheet, rg, (_r, _c, cell) => {
        if (cell.v === null) return;
        if (sheet.hiddenRow(_r) || sheet.hiddenCol(_c)) return;
        count++;
        if (typeof cell.v === 'number') {
          nums++;
          sum += cell.v;
          if (cell.v < min) min = cell.v;
          if (cell.v > max) max = cell.v;
        }
      });
    }
    const f = (v: number) => (Number.isInteger(v) ? v.toLocaleString('en-US') : parseFloat(v.toPrecision(10)).toLocaleString('en-US', { maximumFractionDigits: 10 }));
    const parts: string[] = [];
    if (count > 1 || nums > 0) {
      if (nums > 0 && count > 1) parts.push(`Average: ${f(sum / nums)}`);
      if (count > 1) parts.push(`Count: ${count.toLocaleString('en-US')}`);
      if (nums > 1) {
        parts.push(`Min: ${f(min)}`, `Max: ${f(max)}`);
      }
      if (nums > 0 && count > 1) parts.push(`Sum: ${f(sum)}`);
    }
    if (cells > 1 && count === 0) parts.length = 0;
    this.el.stats.textContent = parts.join('    ');
  }

  updateQat(): void {
    const q = document.getElementById('qat')!;
    const u = q.querySelector('[data-qat=undo]') as HTMLButtonElement | null;
    const r = q.querySelector('[data-qat=redo]') as HTMLButtonElement | null;
    if (u) {
      u.disabled = !this.wb.canUndo();
      u.title = this.wb.canUndo() ? `Undo ${this.wb.undoLabel()}` : 'Undo';
    }
    if (r) {
      r.disabled = !this.wb.canRedo();
      r.title = this.wb.canRedo() ? `Redo ${this.wb.redoLabel()}` : 'Redo';
    }
  }

  // ------------------------------------------------------------------ window & title
  private async initWindowControls(): Promise<void> {
    document.getElementById('app-mark')!.innerHTML = appMark(18);
    document.getElementById('tb-search-ic')!.innerHTML = icon('search', 14);
    const min = document.getElementById('win-min')!;
    const max = document.getElementById('win-max')!;
    const close = document.getElementById('win-close')!;
    min.innerHTML = icon('minimize', 14);
    max.innerHTML = icon('maximize', 12);
    close.innerHTML = icon('close', 14);
    if (!fs.isTauri) {
      (document.querySelector('.tb-ctrl') as HTMLElement).style.display = 'none';
      return;
    }
    const { getCurrentWindow } = await import('@tauri-apps/api/window');
    const win = getCurrentWindow();
    min.addEventListener('click', () => void win.minimize());
    max.addEventListener('click', () => void win.toggleMaximize());
    close.addEventListener('click', () => void win.close());
    const updateMax = async () => {
      const m = await win.isMaximized();
      max.innerHTML = icon(m ? 'restore' : 'maximize', 12);
      max.title = m ? 'Restore Down' : 'Maximize';
    };
    void updateMax();
    void win.onResized(() => void updateMax());
    await win.onCloseRequested(async (ev) => {
      ev.preventDefault();
      if (await this.confirmDiscard()) {
        await fs.dataDelete('recovery.json');
        await win.destroy();
      }
    });
    // "Open with Sheetly"
    // recovery is checked first so a later autosave cannot overwrite recovery.json
    const p = await fs.cliFile();
    const recovered = await this.checkRecovery();
    if (p && (!recovered || (await this.confirmDiscard()))) await this.openPath(p, true);
  }

  updateTitle(): void {
    const t = `${this.doc.name}${this.doc.dirty ? '' : ''} - Sheetly`;
    this.el.title.textContent = t;
    document.title = t;
    if (fs.isTauri) {
      void import('@tauri-apps/api/window').then(({ getCurrentWindow }) => getCurrentWindow().setTitle(t));
    }
  }

  // ------------------------------------------------------------------ theme & zoom UI
  setTheme(dark: boolean, redraw = true): void {
    this.dark = dark;
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
    if (this.grid) {
      this.grid.theme = dark ? DARK : LIGHT;
    }
    try {
      localStorage.setItem('sheetly.theme', dark ? 'dark' : 'light');
    } catch {
      /* ignore */
    }
    if (redraw) {
      this.grid.invalidate();
      this.ribbon.update();
    }
  }

  private initStatusZoom(): void {
    const z = this.el.zoom;
    const minus = h('button', { type: 'button', title: 'Zoom Out', html: icon('minimize', 12) });
    const plus = h('button', { type: 'button', title: 'Zoom In', html: icon('plus', 12) });
    const slider = h('input', { type: 'range', min: '10', max: '400', step: '5', value: '100', title: 'Zoom' }) as HTMLInputElement;
    const pct = h('span', { class: 'pct', title: 'Zoom…' }, '100%');
    minus.addEventListener('click', () => this.grid.setZoom(Math.round(this.grid.zoom * 10 - 1) / 10));
    plus.addEventListener('click', () => this.grid.setZoom(Math.round(this.grid.zoom * 10 + 1) / 10));
    slider.addEventListener('input', () => this.grid.setZoom(parseInt(slider.value, 10) / 100));
    pct.addEventListener('click', () => void this.run('zoomDialog'));
    z.append(minus, slider, plus, pct);
    (z as any)._slider = slider;
    (z as any)._pct = pct;
  }
  private updateZoomUI(): void {
    const z = this.el.zoom as any;
    if (!z._slider) return;
    const v = Math.round(this.grid.zoom * 100);
    z._slider.value = String(v);
    z._pct.textContent = v + '%';
  }

  // ------------------------------------------------------------------ sheets
  private tabsStamp = '';
  renderTabsIfChanged(): void {
    const stamp = this.wb.sheets.map((s) => s.id + ':' + s.name + ':' + (s.tabColor ?? '')).join('|') + '#' + this.wb.sheets.indexOf(this.sheet);
    if (stamp !== this.tabsStamp) this.renderTabs();
  }
  renderTabs(): void {
    const box = this.el.sheetTabs;
    box.innerHTML = '';
    this.tabsStamp = this.wb.sheets.map((s) => s.id + ':' + s.name + ':' + (s.tabColor ?? '')).join('|') + '#' + this.wb.sheets.indexOf(this.sheet);
    this.wb.sheets.forEach((s, i) => {
      const tab = h('div', { class: 'stab' + (s === this.sheet ? ' active' : ''), draggable: 'true', title: s.name }, h('span', {}, s.name), s.tabColor ? h('span', { class: 'tabcolor', style: { background: s.tabColor } }) : null);
      tab.addEventListener('mousedown', (e) => {
        if (e.button === 0 && s !== this.sheet) this.setActiveSheet(s);
      });
      tab.addEventListener('click', () => this.grid.focusGrid());
      tab.addEventListener('dblclick', () => this.beginRenameTab(s, tab));
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        if (s !== this.sheet) this.setActiveSheet(s);
        this.tabMenu(e.clientX, e.clientY, s);
      });
      tab.addEventListener('dragstart', (e) => e.dataTransfer?.setData('text/sheetly-tab', String(i)));
      tab.addEventListener('dragover', (e) => {
        e.preventDefault();
        tab.classList.add('dragover');
      });
      tab.addEventListener('dragleave', () => tab.classList.remove('dragover'));
      tab.addEventListener('drop', (e) => {
        e.preventDefault();
        tab.classList.remove('dragover');
        const from = parseInt(e.dataTransfer?.getData('text/sheetly-tab') ?? '-1', 10);
        if (from >= 0 && from !== i) {
          moveSheet(this.wb, from, i > from ? i : i);
          this.renderTabs();
        }
      });
      box.append(tab);
    });
    const active = box.querySelector('.stab.active') as HTMLElement | null;
    active?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  private tabMenu(x: number, y: number, s: Sheet): void {
    showMenu(x, y, [
      { label: 'Insert…', icon: 'sheetAdd', cmd: () => this.run('addSheet') },
      { label: 'Delete', icon: 'deleteCells', cmd: () => this.run('deleteSheet') },
      { label: 'Rename', cmd: () => {
        const el = Array.from(this.el.sheetTabs.children)[this.wb.sheets.indexOf(s)] as HTMLElement;
        this.beginRenameTab(s, el);
      } },
      { label: 'Move or Copy…', submenu: [
        { label: 'Create a Copy', cmd: () => this.run('dupSheet') },
        { label: 'Move Left', cmd: () => this.run('moveSheet', -1) },
        { label: 'Move Right', cmd: () => this.run('moveSheet', 1) },
      ] },
      { sep: true },
      { label: 'Tab Color', cmd: () => this.run('tabColor') },
    ]);
  }

  private beginRenameTab(s: Sheet, tab: HTMLElement): void {
    const input = h('input', { type: 'text', value: s.name }) as HTMLInputElement;
    tab.innerHTML = '';
    tab.append(input);
    input.focus();
    input.select();
    let done = false;
    const finish = async (ok: boolean) => {
      if (done) return;
      done = true;
      if (ok && input.value !== s.name) {
        const err = this.validateSheetName(input.value, s);
        if (err) {
          await messageBox('Sheetly', err);
        } else renameSheet(this.wb, s, input.value);
      }
      this.renderTabs();
      this.grid.focusGrid();
    };
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') void finish(true);
      else if (e.key === 'Escape') void finish(false);
    });
    input.addEventListener('blur', () => void finish(true));
  }

  validateSheetName(name: string, except?: Sheet): string | null {
    const n = name.trim();
    if (!n) return 'The sheet name cannot be blank.';
    if (n.length > 31) return 'The sheet name must be 31 characters or fewer.';
    if (/[\\/?*[\]:]/.test(n)) return 'The sheet name cannot contain any of the following characters: \\ / ? * [ ] :';
    if (this.wb.sheets.some((s) => s !== except && s.name.toLowerCase() === n.toLowerCase())) return 'That name is already in use. Try a different name.';
    return null;
  }

  setActiveSheet(s: Sheet): void {
    if (s === this.sheet) return;
    this.grid.setSheet(s);
    this.renderTabs();
    this.refreshAll();
  }

  // ------------------------------------------------------------------ styles helpers
  activeStyle() {
    const { r, c } = this.grid.active;
    return this.wb.styleOf(this.sheet.get(r, c)?.s);
  }

  /** Apply a style patch to every selected range as one undo step. */
  formatSelection(patch: StylePatch, label = 'Format'): void {
    const rects = this.grid.ranges.map((r) => this.grid.expandForMerges(r));
    this.wb.transact(label, () => {
      for (const rg of rects) applyStyle(this.wb, this.sheet, rg, patch);
    });
    // wrapped rows grow automatically
    if (patch.wrap === true) {
      const rows = new Set<number>();
      for (const rg of rects) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 500); r++) if (!this.sheet.rows.sizes.has(r)) rows.add(r);
      if (rows.size) this.grid.autofitRows(Array.from(rows), false);
    }
  }

  // ------------------------------------------------------------------ clipboard
  makeClip(cut: boolean): { data: ClipData; text: string } | null {
    const rect = this.grid.expandForMerges(this.grid.range);
    if (rect.r2 - rect.r1 + 1 > 100000 || (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1) > 3_000_000) {
      // whole rows/columns: clip to the used area
      const u = this.sheet.usedRect();
      if (!u) return null;
      rect.r2 = Math.min(rect.r2, u.r2);
      rect.c2 = Math.min(rect.c2, u.c2);
    }
    const data = copyRange(this.wb, this.sheet, rect, cut);
    this.clip = { data, wb: this.wb, sheet: this.sheet, text: data.tsv };
    this.grid.invalidate();
    return { data, text: data.tsv };
  }

  private clipHtml(d: ClipData): string {
    const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    let html = '<table>';
    const lines = d.tsv.split('\r\n');
    for (let r = 0; r < d.rows; r++) {
      html += '<tr>';
      const cells = lines[r].split('\t');
      for (let c = 0; c < d.cols; c++) html += `<td>${esc(cells[c] ?? '')}</td>`;
      html += '</tr>';
    }
    return html + '</table>';
  }

  async copyToSystem(cut: boolean): Promise<void> {
    const c = this.makeClip(cut);
    if (!c) return;
    try {
      await navigator.clipboard.writeText(c.text);
    } catch {
      /* clipboard unavailable: internal clipboard still works */
    }
  }

  async pasteFromSystem(mode: PasteMode): Promise<void> {
    let text: string | null = null;
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = null;
    }
    this.doPaste(text, mode);
  }

  doPaste(text: string | null, mode: PasteMode): void {
    const sheet = this.sheet;
    const rect = this.grid.range;
    const norm = (s: string) => s.replace(/\r\n/g, '\n').replace(/\n$/, '');
    const internal = this.clip && (text === null || norm(text) === norm(this.clip.text));
    if (!internal && (text === null || text === '')) return;
    const r = rect.r1;
    const c = rect.c1;
    let dst: Rect | null = null;
    try {
      if (internal && this.clip) {
        const { data, wb: srcWb, sheet: srcSheet } = this.clip;
        if (data.cut && mode === 'all' && srcWb === this.wb) {
          this.wb.transact('Paste', () => {
            dst = moveRange(this.wb, srcSheet, data.rect, sheet, r, c);
          });
          this.clip = null;
        } else {
          this.wb.transact('Paste', () => {
            const single = data.rows === 1 && data.cols === 1 && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2) && mode !== 'transpose';
            if (single) {
              // cap the target: a huge selection only gets pasted into up to 1M cells
              const used = sheet.usedRect();
              const area = (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
              let tr = rect.r2;
              let tc = rect.c2;
              if (area > 1_000_000) {
                tr = Math.min(tr, Math.max(rect.r1, used ? used.r2 : rect.r1));
                tc = Math.min(tc, Math.max(rect.c1, used ? used.c2 : rect.c1));
                while ((tr - rect.r1 + 1) * (tc - rect.c1 + 1) > 1_000_000) {
                  if (tr - rect.r1 >= tc - rect.c1) tr = rect.r1 + Math.floor((tr - rect.r1) / 2);
                  else tc = rect.c1 + Math.floor((tc - rect.c1) / 2);
                }
              }
              for (let rr = rect.r1; rr <= tr; rr++)
                for (let cc = rect.c1; cc <= tc; cc++) pasteClip(this.wb, sheet, rr, cc, data, srcWb, mode);
              dst = { r1: rect.r1, c1: rect.c1, r2: tr, c2: tc };
            } else dst = pasteClip(this.wb, sheet, r, c, data, srcWb, mode);
          });
        }
      } else if (text !== null) {
        this.wb.transact('Paste', () => {
          dst = pasteText(this.wb, sheet, r, c, text);
        });
      }
    } catch (e) {
      void messageBox('Sheetly', 'The paste could not be completed: ' + (e as Error).message);
      return;
    }
    if (dst) this.grid.selectRect(this.grid.expandForMerges(dst));
    this.grid.invalidate();
  }

  // ------------------------------------------------------------------ commands
  async run(cmd: string, arg?: unknown): Promise<void> {
    try {
      await this.runInner(cmd, arg);
    } catch (err) {
      void messageBox('Sheetly', `Sorry, that command could not be completed.\n\n${(err as Error)?.message ?? String(err)}`);
    }
  }

  private async runInner(cmd: string, arg?: unknown): Promise<void> {
    const grid = this.grid;
    const wb = this.wb;
    const sheet = this.sheet;
    const rect = this.grid.expandForMerges(this.grid.range);
    const toggle = (key: 'bold' | 'italic' | 'underline' | 'strike') => {
      const on = !!this.activeStyle()[key];
      this.formatSelection({ [key]: on ? null : true } as StylePatch);
    };
    switch (cmd) {
      // ---- file ----
      case 'new': return void this.newWorkbook();
      case 'open': return void this.openDialog();
      case 'save': return void this.save();
      case 'saveAs': return void this.saveAs();
      case 'exportCsv': return void this.saveAs('csv');
      case 'exportPdf':
      case 'print': return void printSheet(this);
      case 'backstage': return this.backstage.show((arg as string) ?? 'info');
      // ---- edit ----
      case 'undo': {
        if (grid.editor.active) return;
        const e = wb.undo();
        if (e) this.afterHistory(e);
        return;
      }
      case 'redo': {
        if (grid.editor.active) return;
        const e = wb.redo();
        if (e) this.afterHistory(e);
        return;
      }
      case 'cut': return void this.copyToSystem(true);
      case 'copy': return void this.copyToSystem(false);
      case 'paste': return void this.pasteFromSystem((arg as PasteMode) ?? 'all');
      case 'pasteSpecial': return dlg.openPasteSpecial(this);
      case 'painter': {
        if (this.painter) {
          this.painter = null;
        } else {
          const { r, c } = grid.active;
          this.painter = { style: sheet.get(r, c)?.s ?? 0, sticky: !!arg };
        }
        this.ribbon.update();
        grid.area.style.cursor = this.painter ? 'copy' : 'cell';
        return;
      }
      case 'clear': {
        const what = (arg as 'all' | 'contents' | 'formats') ?? 'contents';
        wb.transact('Clear', () => {
          for (const rg of grid.ranges) clearRange(wb, sheet, this.grid.expandForMerges(rg), what);
          if (what !== 'contents') for (const rg of grid.ranges) sheet.merges = sheet.merges.filter((m) => !(m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2));
        }, what !== 'contents');
        return;
      }
      case 'fill': {
        const dir = arg as 'down' | 'right' | 'up' | 'left';
        wb.transact('Fill', () => {
          const rg = rect;
          if (dir === 'down' && rg.r2 > rg.r1) fillRange(wb, sheet, { ...rg, r2: rg.r1 }, 'down', rg.r2 - rg.r1, 'copy');
          else if (dir === 'right' && rg.c2 > rg.c1) fillRange(wb, sheet, { ...rg, c2: rg.c1 }, 'right', rg.c2 - rg.c1, 'copy');
          else if (dir === 'up' && rg.r2 > rg.r1) fillRange(wb, sheet, { ...rg, r1: rg.r2 }, 'up', rg.r2 - rg.r1, 'copy');
          else if (dir === 'left' && rg.c2 > rg.c1) fillRange(wb, sheet, { ...rg, c1: rg.c2 }, 'left', rg.c2 - rg.c1, 'copy');
          else if (dir === 'down' && rg.r1 > 0) fillRange(wb, sheet, { r1: rg.r1 - 1, r2: rg.r1 - 1, c1: rg.c1, c2: rg.c2 }, 'down', 1, 'copy');
          else if (dir === 'right' && rg.c1 > 0) fillRange(wb, sheet, { c1: rg.c1 - 1, c2: rg.c1 - 1, r1: rg.r1, r2: rg.r2 }, 'right', 1, 'copy');
        });
        return;
      }
      case 'fillSeries': {
        const dir = (rect.r2 > rect.r1 ? 'down' : 'right') as 'down' | 'right';
        wb.transact('Fill series', () => {
          if (dir === 'down') fillRange(wb, sheet, { ...rect, r2: rect.r1 + 1 > rect.r2 ? rect.r1 : rect.r1 + 1 }, 'down', rect.r2 - rect.r1 - 1);
        });
        return;
      }
      case 'find': return dlg.openFindReplace(this, false);
      case 'replace': return dlg.openFindReplace(this, true);
      case 'goto': return dlg.openGoTo(this);
      case 'selectAll': return grid.selectAll();
      case 'insertDate':
      case 'insertTime': {
        const v = cmd === 'insertDate' ? Math.floor(todaySerial()) : nowSerial() - Math.floor(nowSerial());
        const { r, c } = grid.active;
        wb.transact('Insert', () => {
          const old = sheet.get(r, c);
          const s = wb.patchStyle(old?.s, { fmt: cmd === 'insertDate' ? 'm/d/yyyy' : 'h:mm AM/PM' });
          wb.setCell(sheet, r, c, { v, s });
        });
        return;
      }
      // ---- formatting ----
      case 'bold': return toggle('bold');
      case 'italic': return toggle('italic');
      case 'underline': return toggle('underline');
      case 'strike': return toggle('strike');
      case 'font': return this.formatSelection({ font: (arg as string) === 'Calibri' ? null : (arg as string) });
      case 'size': return this.formatSelection({ size: arg === 11 ? null : (arg as number) });
      case 'sizeInc':
      case 'sizeDec': {
        const cur = this.activeStyle().size ?? 11;
        const steps = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
        let n = cur;
        if (cmd === 'sizeInc') n = steps.find((s) => s > cur) ?? cur + 2;
        else n = [...steps].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
        return this.formatSelection({ size: n === 11 ? null : n });
      }
      case 'fontColor': return this.formatSelection({ color: (arg as string | null) ?? null });
      case 'fillColor': return this.formatSelection({ fill: (arg as string | null) ?? null });
      case 'halign': {
        const cur = this.activeStyle().h;
        return this.formatSelection({ h: cur === arg ? null : (arg as 'left' | 'center' | 'right') });
      }
      case 'valign': {
        const cur = this.activeStyle().v ?? 'bottom';
        const a = arg as 'top' | 'middle' | 'bottom';
        return this.formatSelection({ v: cur === a && a !== 'bottom' ? null : a === 'bottom' ? null : a });
      }
      case 'wrap': return this.formatSelection({ wrap: this.activeStyle().wrap ? null : true });
      case 'indent': {
        const cur = this.activeStyle().indent ?? 0;
        const n = Math.max(0, cur + (arg as number));
        const h = this.activeStyle().h;
        return this.formatSelection({ indent: n || null, h: h ?? (n ? 'left' : null) });
      }
      case 'merge': return this.mergeCells((arg as string) ?? 'center');
      case 'borders': return this.applyBorders((arg as string) ?? 'all');
      case 'numFmt': return this.formatSelection({ fmt: arg === 'General' ? null : (arg as string) }, 'Number format');
      case 'currency': return this.formatSelection({ fmt: (arg as string) ?? FMT_CURRENCY });
      case 'accounting': return this.formatSelection({ fmt: FMT_ACCOUNTING });
      case 'percent': return this.formatSelection({ fmt: FMT_PERCENT });
      case 'comma': return this.formatSelection({ fmt: '#,##0.00' });
      case 'decInc':
      case 'decDec': {
        const cur = this.activeStyle().fmt ?? 'General';
        return this.formatSelection({ fmt: changeDecimals(cur, cmd === 'decInc' ? 1 : -1) }, 'Decimals');
      }
      case 'cellStyle': {
        const def = CELL_STYLES.find((s) => s.name === arg);
        if (def) this.formatSelection(def.patch, 'Cell Style');
        return;
      }
      case 'formatCells': return dlg.openFormatCells(this, (arg as string) ?? 'number');
      // ---- conditional formatting ----
      case 'cf': return dlg.openCF(this, arg as string);
      case 'cfClear': {
        const scope = (arg as string) ?? 'selection';
        wb.transact('Clear rules', () => {
          if (scope === 'sheet') sheet.cf = [];
          else sheet.cf = sheet.cf.filter((r) => !(r.rect.r1 >= rect.r1 && r.rect.r2 <= rect.r2 && r.rect.c1 >= rect.c1 && r.rect.c2 <= rect.c2));
        }, true);
        return;
      }
      case 'cfManage': return dlg.openCFManager(this);
      case 'cfScale': {
        const colors = arg as string[];
        wb.transact('Color scale', () => sheet.cf.push({ id: wb.nextRuleId++, rect: { ...rect }, type: 'scale', colors }), true);
        return;
      }
      case 'cfBar': {
        wb.transact('Data bars', () => sheet.cf.push({ id: wb.nextRuleId++, rect: { ...rect }, type: 'bar', barColor: arg as string }), true);
        return;
      }
      // ---- cells ----
      case 'insertDialog': return dlg.openInsertDelete(this, true);
      case 'deleteDialog': return dlg.openInsertDelete(this, false);
      case 'insertRows': {
        const n = rect.r2 - rect.r1 + 1;
        insertRows(wb, sheet, rect.r1, n);
        return;
      }
      case 'insertCols': {
        const n = rect.c2 - rect.c1 + 1;
        insertCols(wb, sheet, rect.c1, n);
        return;
      }
      case 'deleteRows': return deleteRows(wb, sheet, rect.r1, rect.r2 - rect.r1 + 1);
      case 'deleteCols': return deleteCols(wb, sheet, rect.c1, rect.c2 - rect.c1 + 1);
      case 'shiftCells': return this.shiftCells(arg as 'right' | 'down' | 'left' | 'up');
      case 'rowHeight': return dlg.openSizeDialog(this, 'row');
      case 'colWidth': return dlg.openSizeDialog(this, 'col');
      case 'autofitRows': {
        const rows: number[] = [];
        for (let r = rect.r1; r <= Math.min(rect.r2, rect.r1 + 5000); r++) rows.push(r);
        return grid.autofitRows(rows);
      }
      case 'autofitCols': {
        const cols: number[] = [];
        for (let c = rect.c1; c <= Math.min(rect.c2, rect.c1 + 1000); c++) cols.push(c);
        return grid.autofitCols(cols);
      }
      case 'defaultColWidth': return dlg.openSizeDialog(this, 'defcol');
      case 'hideRows':
      case 'hideCols':
      case 'unhideRows':
      case 'unhideCols': return this.hideUnhide(cmd);
      // ---- sheets ----
      case 'addSheet': {
        let s!: Sheet;
        wb.transact('Add sheet', () => {
          s = wb.addSheet(undefined, wb.sheets.indexOf(sheet) + 1);
        }, true);
        this.setActiveSheet(s);
        return;
      }
      case 'deleteSheet': {
        if (wb.sheets.length <= 1) return void messageBox('Sheetly', "A workbook must contain at least one visible worksheet.");
        const idx = wb.sheets.indexOf(sheet);
        if (sheet.usedRect()) {
          const r = await messageBox('Sheetly', `Delete "${sheet.name}"? The data on this sheet will be removed (you can undo this).`, ['Delete', 'Cancel']);
          if (r !== 'Delete') return;
        }
        deleteSheet(wb, sheet);
        grid.setSheet(wb.sheets[Math.min(idx, wb.sheets.length - 1)]);
        this.renderTabs();
        this.refreshAll();
        return;
      }
      case 'renameSheet': {
        const el = Array.from(this.el.sheetTabs.children)[wb.sheets.indexOf(sheet)] as HTMLElement;
        this.beginRenameTab(sheet, el);
        return;
      }
      case 'dupSheet': {
        const s = duplicateSheet(wb, sheet);
        this.setActiveSheet(s);
        return;
      }
      case 'moveSheet': {
        const i = wb.sheets.indexOf(sheet);
        const j = i + (arg as number);
        if (j >= 0 && j < wb.sheets.length) {
          moveSheet(wb, i, j);
          this.renderTabs();
        }
        return;
      }
      case 'tabColor': {
        const anchor = this.el.sheetTabs;
        showColorMenu(anchor, {
          noneLabel: 'No Color',
          onPick: (c) => wb.transact('Tab color', () => (sheet.tabColor = c ?? undefined), true),
        });
        return;
      }
      // ---- data ----
      case 'sort': return this.sortSelection(arg === 'asc');
      case 'sortDialog': return dlg.openSort(this);
      case 'filter': return this.toggleFilter();
      case 'filterClear': return sheet.filter && (clearFilter(wb, sheet), grid.invalidate());
      case 'filterReapply': return sheet.filter && (applyFilter(wb, sheet), grid.invalidate());
      // ---- formulas ----
      case 'autosum': return this.autosum((arg as string) ?? 'SUM');
      case 'insertFunction': return dlg.openInsertFunction(this);
      case 'insertFn': return this.insertFunctionByName(arg as string);
      case 'nameManager': return dlg.openNameManager(this);
      case 'defineName': return dlg.openDefineName(this);
      case 'showFormulas':
        this.showFormulas = !this.showFormulas;
        grid.invalidate();
        this.ribbon.update();
        return;
      case 'calcNow':
        wb.recalcAll();
        grid.invalidate();
        return;
      case 'calcMode':
        wb.autoCalc = arg === 'auto';
        if (wb.autoCalc) wb.recalcAll();
        this.ribbon.update();
        return;
      // ---- view ----
      case 'freeze': return this.freeze(arg as string);
      case 'zoomDialog': return dlg.openZoom(this);
      case 'zoom': return grid.setZoom((arg as number) / 100);
      case 'zoomSelection': return this.zoomToSelection();
      case 'gridlines':
        this.gridlines = !this.gridlines;
        grid.invalidate();
        this.ribbon.update();
        return;
      case 'headings':
        this.headings = !this.headings;
        grid.invalidate();
        this.ribbon.update();
        return;
      case 'formulaBar':
        this.showFormulaBar = !this.showFormulaBar;
        this.el.formulabar.style.display = this.showFormulaBar ? '' : 'none';
        this.ribbon.update();
        grid.resize();
        return;
      case 'theme': return this.setTheme(arg === undefined ? !this.dark : !!arg);
      case 'orientation':
        this.orientation = arg as 'portrait' | 'landscape';
        this.ribbon.update();
        return;
      case 'printGridlines':
        this.printGridlines = !this.printGridlines;
        this.ribbon.update();
        return;
      case 'stats': return dlg.openStats(this);
      case 'about': return dlg.openAbout();
    }
  }

  private afterHistory(e: { sheetId?: number; rect?: Rect }): void {
    if (e.sheetId !== undefined) {
      const s = this.wb.sheetById(e.sheetId);
      if (s && this.wb.sheets.includes(s) && s !== this.sheet) this.grid.setSheet(s);
    }
    if (!this.wb.sheets.includes(this.sheet)) this.grid.setSheet(this.wb.sheets[0]);
    if (e.rect && e.sheetId !== undefined) {
      const r = e.rect;
      if (r.r2 - r.r1 < 100000) {
        this.grid.selectRect(r);
        this.grid.scrollIntoView(r.r1, r.c1);
      }
    }
    this.renderTabs();
    this.refreshAll();
  }

  private shiftCells(dir: 'right' | 'down' | 'left' | 'up'): void {
    // Insert/delete cells shifting within the selected rows/columns only
    const wb = this.wb;
    const sheet = this.sheet;
    const rg = this.grid.range;
    const rows = rg.r2 - rg.r1 + 1;
    const cols = rg.c2 - rg.c1 + 1;
    const horiz = dir === 'right' || dir === 'left';
    const l1 = horiz ? rg.r1 : rg.c1;
    const l2 = horiz ? rg.r2 : rg.c2;
    const start = horiz ? rg.c1 : rg.r1;
    const n = horiz ? cols : rows;
    const limit = horiz ? MAX_COLS : MAX_ROWS;
    const forward = dir === 'right' || dir === 'down';
    // gather populated cells of the affected lines once (sparse)
    const lines = new Map<number, Map<number, Cell>>();
    for (const [k, cell] of sheet.cells) {
      const r = keyRow(k);
      const c = keyCol(k);
      const line = horiz ? r : c;
      const axis = horiz ? c : r;
      if (line < l1 || line > l2 || axis < start) continue;
      let m = lines.get(line);
      if (!m) lines.set(line, (m = new Map()));
      m.set(axis, cell);
    }
    wb.transact('Shift cells', () => {
      const put = (line: number, axis: number, cell: Cell | undefined) => {
        const r = horiz ? line : axis;
        const c = horiz ? axis : line;
        const nc = cell ? { v: cell.v, f: cell.f, s: cell.s } : undefined;
        if (!nc && !sheet.get(r, c)) return;
        wb.setCell(sheet, r, c, nc);
      };
      for (const [line, old] of lines) {
        const next = new Map<number, Cell>();
        for (const [axis, cell] of old) {
          if (forward) {
            if (axis + n < limit) next.set(axis + n, cell);
          } else if (axis >= start + n) next.set(axis - n, cell);
        }
        const touched = new Set<number>([...old.keys(), ...next.keys()]);
        for (const axis of touched) put(line, axis, next.get(axis));
      }
    });
  }

  private hideUnhide(cmd: string): void {
    const sheet = this.sheet;
    const rg = this.grid.range;
    this.wb.transact('Hide/Unhide', () => {
      if (cmd === 'hideRows') for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) sheet.rows.set(r, 0);
      else if (cmd === 'hideCols') for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 2000); c++) sheet.cols.set(c, 0);
      else if (cmd === 'unhideRows') {
        for (let r = Math.max(0, rg.r1 - 1); r <= Math.min(rg.r2 + 1, rg.r1 + 100000); r++) if (sheet.rows.sizes.get(r) === 0) sheet.rows.set(r, sheet.rows.def);
      } else for (let c = Math.max(0, rg.c1 - 1); c <= Math.min(rg.c2 + 1, rg.c1 + 2000); c++) if (sheet.cols.sizes.get(c) === 0) sheet.cols.set(c, sheet.cols.def);
    }, true);
    this.grid.invalidate();
  }

  mergeCells(kind: string): void {
    const sheet = this.sheet;
    const wb = this.wb;
    const rects = this.grid.ranges.map((r) => ({ ...r }));
    if (kind !== 'unmerge' && this.tooLarge(rects, 100_000, 'Merging')) return;
    wb.transact('Merge', () => {
      for (const rg of rects) {
        // remove merges overlapping the range
        const overlaps = (m: Rect) => m.r1 <= rg.r2 && m.r2 >= rg.r1 && m.c1 <= rg.c2 && m.c2 >= rg.c1;
        const hadMerge = sheet.merges.some(overlaps);
        sheet.merges = sheet.merges.filter((m) => !overlaps(m));
        if (kind === 'unmerge' || (kind === 'center' && hadMerge)) {
          if (kind === 'center') applyStyle(wb, sheet, rg, { h: null });
          continue;
        }
        if (rg.r1 === rg.r2 && rg.c1 === rg.c2) continue;
        const targets: Rect[] = kind === 'across' ? Array.from({ length: rg.r2 - rg.r1 + 1 }, (_, i) => ({ r1: rg.r1 + i, r2: rg.r1 + i, c1: rg.c1, c2: rg.c2 })) : [rg];
        for (const t of targets) {
          if (t.c1 === t.c2 && t.r1 === t.r2) continue;
          // keep the first non-empty value in the top-left cell like Excel keeps the upper-left value
          const doomed: [number, number, Cell][] = [];
          eachCell(sheet, t, (r, c, cell) => {
            if (r === t.r1 && c === t.c1) return;
            if (cell.v !== null || cell.f !== undefined) doomed.push([r, c, cell]);
          });
          for (const [r, c, cell] of doomed) wb.setCell(sheet, r, c, cell.s ? { v: null, s: cell.s } : undefined);
          sheet.merges.push({ ...t });
        }
        if (kind === 'center') applyStyle(wb, sheet, rg, { h: 'center' });
      }
    }, true);
    this.grid.invalidate();
    this.scheduleRefresh();
  }

  private applyBorders(kind: string): void {
    const wb = this.wb;
    const sheet = this.sheet;
    const rects = this.grid.ranges.map((r) => this.grid.expandForMerges(r));
    if (this.tooLarge(rects, 500_000, 'Borders')) return;
    const thin: Border = { style: 'thin', color: this.dark ? '#ffffff' : '#000000' };
    const med: Border = { style: 'medium', color: thin.color };
    const thick: Border = { style: 'thick', color: thin.color };
    const dbl: Border = { style: 'double', color: thin.color };
    wb.transact('Borders', () => {
      for (const rg of rects) {
        const R = rg.r2;
        const C = rg.c2;
        const patchEdges = (fn: (r: number, c: number) => StylePatch | null) => {
          for (let r = rg.r1; r <= R; r++)
            for (let c = rg.c1; c <= C; c++) {
              const p = fn(r, c);
              if (p) applyStyle(wb, sheet, { r1: r, c1: c, r2: r, c2: c }, p);
            }
        };
        switch (kind) {
          case 'none':
            patchEdges(() => ({ bl: null, br: null, bt: null, bb: null }));
            break;
          case 'all':
            patchEdges(() => ({ bl: thin, br: thin, bt: thin, bb: thin }));
            break;
          case 'inside':
            patchEdges((r, c) => ({ ...(c > rg.c1 ? { bl: thin } : {}), ...(c < rg.c2 ? { br: thin } : {}), ...(r > rg.r1 ? { bt: thin } : {}), ...(r < rg.r2 ? { bb: thin } : {}) }));
            break;
          case 'outside':
          case 'thickOutside': {
            const b = kind === 'outside' ? thin : thick;
            patchEdges((r, c) => ({ ...(c === rg.c1 ? { bl: b } : {}), ...(c === rg.c2 ? { br: b } : {}), ...(r === rg.r1 ? { bt: b } : {}), ...(r === rg.r2 ? { bb: b } : {}) }));
            break;
          }
          case 'bottom':
            patchEdges((r) => (r === rg.r2 ? { bb: thin } : null));
            break;
          case 'top':
            patchEdges((r) => (r === rg.r1 ? { bt: thin } : null));
            break;
          case 'left':
            patchEdges((_r, c) => (c === rg.c1 ? { bl: thin } : null));
            break;
          case 'right':
            patchEdges((_r, c) => (c === rg.c2 ? { br: thin } : null));
            break;
          case 'bottomDouble':
            patchEdges((r) => (r === rg.r2 ? { bb: dbl } : null));
            break;
          case 'thickBottom':
            patchEdges((r) => (r === rg.r2 ? { bb: thick } : null));
            break;
          case 'topBottom':
            patchEdges((r) => (r === rg.r1 ? { bt: thin } : r === rg.r2 ? { bb: thin } : null));
            break;
          case 'topThickBottom':
            patchEdges((r) => ({ ...(r === rg.r1 ? { bt: thin } : {}), ...(r === rg.r2 ? { bb: med } : {}) }));
            break;
        }
      }
    });
  }

  private sortSelection(asc: boolean): void {
    const sheet = this.sheet;
    let rg = this.grid.range;
    const { r, c } = this.grid.active;
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = currentRegion(sheet, r, c);
    if (rg.r1 === rg.r2) return;
    const hasHeader = detectHeader(this.wb, sheet, rg);
    const col = c >= rg.c1 && c <= rg.c2 ? c : rg.c1;
    sortRange(this.wb, sheet, rg, [{ col, asc }], hasHeader);
    this.grid.selectRect(rg);
  }

  sortWith(keys: SortKey[], hasHeader: boolean, rect: Rect): void {
    sortRange(this.wb, this.sheet, rect, keys, hasHeader);
    this.grid.selectRect(rect);
  }

  private toggleFilter(): void {
    const sheet = this.sheet;
    if (sheet.filter) {
      removeFilter(this.wb, sheet);
      this.grid.invalidate();
      return;
    }
    let rg = this.grid.range;
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = currentRegion(sheet, this.grid.active.r, this.grid.active.c);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !sheet.get(rg.r1, rg.c1)) {
      void messageBox('Sheetly', 'Select a cell inside your data first, then turn on the filter.');
      return;
    }
    this.wb.transact('Filter', () => {
      sheet.filter = { rect: { ...rg }, cols: new Map() };
    }, true);
    this.grid.invalidate();
  }

  private autosum(fn: string): void {
    const sheet = this.sheet;
    const wb = this.wb;
    const rg = this.grid.range;
    const has = (r: number, c: number) => typeof sheet.get(r, c)?.v === 'number';
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    if (!single) {
      // sum each column below the selection (or each row to the right when only one column is selected)
      wb.transact('AutoSum', () => {
        if (rg.c2 > rg.c1 || rg.r2 - rg.r1 === 0) {
          for (let c = rg.c1; c <= rg.c2; c++) {
            const target = rg.r2 + 1;
            if (target < MAX_ROWS) applyInput(wb, sheet, target, c, `=${fn}(${a1(rg.r1, c)}:${a1(rg.r2, c)})`);
          }
        } else {
          for (let r = rg.r1; r <= rg.r2; r++) applyInput(wb, sheet, r, rg.c2 + 1, `=${fn}(${a1(r, rg.c1)}:${a1(r, rg.c2)})`);
        }
      });
      return;
    }
    const { r, c } = this.grid.active;
    let formulaRange: string | null = null;
    if (r > 0 && has(r - 1, c)) {
      let r1 = r - 1;
      while (r1 > 0 && has(r1 - 1, c)) r1--;
      formulaRange = `${a1(r1, c)}:${a1(r - 1, c)}`;
    } else if (c > 0 && has(r, c - 1)) {
      let c1 = c - 1;
      while (c1 > 0 && has(r, c1 - 1)) c1--;
      formulaRange = `${a1(r, c1)}:${a1(r, c - 1)}`;
    }
    if (formulaRange) {
      wb.transact('AutoSum', () => applyInput(wb, sheet, r, c, `=${fn}(${formulaRange})`));
    } else {
      this.grid.beginEdit('enter', `=${fn}(`);
    }
  }

  insertFunctionByName(name: string): void {
    const ed = this.grid.editor;
    if (!ed.active) this.grid.beginEdit('edit', '');
    const t = ed.text;
    const ins = (t.startsWith('=') ? '' : '=') + name + '(';
    if (!t.startsWith('=')) ed.setText(ins, ins.length);
    else {
      const caret = ed.caret();
      ed.replaceRange(caret, caret, name + '(');
    }
  }

  private freeze(kind: string): void {
    const s = this.sheet;
    const { r, c } = this.grid.active;
    this.wb.transact('Freeze panes', () => {
      if (kind === 'none') {
        s.freezeRows = 0;
        s.freezeCols = 0;
      } else if (kind === 'top') {
        s.freezeRows = 1;
        s.freezeCols = 0;
      } else if (kind === 'col') {
        s.freezeRows = 0;
        s.freezeCols = 1;
      } else {
        s.freezeRows = r;
        s.freezeCols = c;
        s.scrollX = 0;
        s.scrollY = 0;
      }
    }, true);
    this.grid.clampScroll();
    this.grid.invalidate();
    this.ribbon.update();
  }

  private zoomToSelection(): void {
    const rg = this.grid.range;
    const g = this.grid;
    const s = this.sheet;
    const w = s.cols.top(rg.c2 + 1) - s.cols.top(rg.c1);
    const hgt = s.rows.top(rg.r2 + 1) - s.rows.top(rg.r1);
    const vw = g.area.clientWidth - 60;
    const vh = g.area.clientHeight - 40;
    const z = Math.min(vw / Math.max(w, 1), vh / Math.max(hgt, 1));
    g.setZoom(Math.max(0.1, Math.min(4, z)));
    g.scrollIntoView(rg.r1, rg.c1);
    s.scrollX = Math.max(0, s.cols.top(rg.c1) - s.cols.top(s.freezeCols));
    s.scrollY = Math.max(0, s.rows.top(rg.r1) - s.rows.top(s.freezeRows));
    g.clampScroll();
    g.invalidate();
  }

  // ------------------------------------------------------------------ navigation
  goToText(text: string): void {
    const t = text.trim();
    if (!t) return;
    // sheet-qualified?
    let sheet = this.sheet;
    let ref = t;
    const bang = t.lastIndexOf('!');
    if (bang > 0) {
      const sn = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
      const s = this.wb.sheetByName(sn);
      if (!s) {
        void messageBox('Sheetly', 'Reference is not valid.');
        return;
      }
      sheet = s;
      ref = t.slice(bang + 1);
    }
    const rect = parseRect(ref);
    if (rect) {
      if (sheet !== this.sheet) this.setActiveSheet(sheet);
      this.selectRectKeepActive(rect);
      return;
    }
    const def = this.wb.getName(t);
    if (def) {
      const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(def.refers);
      if (m) {
        const sn = m[1] ? this.wb.sheetByName(m[1]) : this.sheet;
        const r = parseRect(m[2]);
        if (sn && r) {
          if (sn !== this.sheet) this.setActiveSheet(sn);
          this.selectRectKeepActive(r);
          return;
        }
      }
    }
    // define a new name for the current selection
    if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(t) && !parseCell(t)) {
      const rg = this.grid.range;
      const q = /^[A-Za-z_][\w.]*$/.test(this.sheet.name) ? this.sheet.name : "'" + this.sheet.name.replace(/'/g, "''") + "'";
      this.wb.transact('Define name', () => this.wb.setName(t, `${q}!${rectAbs(rg)}`), true);
      this.wb.rebuildDeps();
      this.wb.requestFullRecalc();
      this.refreshAll();
      return;
    }
    void messageBox('Sheetly', 'Reference is not valid.');
  }

  selectRectKeepActive(rect: Rect): void {
    this.grid.selectRect(rect);
    this.grid.scrollIntoView(rect.r1, rect.c1);
  }

  // ------------------------------------------------------------------ files
  async confirmDiscard(): Promise<boolean> {
    if (!this.doc.dirty) return true;
    const r = await messageBox('Sheetly', `Do you want to save the changes you made to ${this.doc.name}?`, ['Save', "Don't Save", 'Cancel']);
    if (r === 'Cancel') return false;
    if (r === 'Save') return await this.save();
    return true;
  }

  private loadWorkbook(wb: Workbook, doc: Partial<typeof this.doc>, active = 0): void {
    this.grid.cancelEdit();
    this.wb = wb;
    this.cf = new CFEngine(wb);
    wb.onChange = () => this.onWorkbookChange();
    this.clip = null;
    this.painter = null;
    this.changeCount = 0;
    this.autosavedCount = 0;
    Object.assign(this.doc, { path: null, name: 'Book1', dirty: false, kind: 'xlsx' as DocKind }, doc);
    this.grid.setSheet(wb.sheets[Math.min(active, wb.sheets.length - 1)]);
    this.renderTabs();
    this.refreshAll();
    this.updateTitle();
    this.backstage.hide();
    this.grid.focusGrid();
  }

  async newWorkbook(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    this.untitledCount++;
    this.loadWorkbook(new Workbook(), { name: `Book${this.untitledCount}` });
    await fs.dataDelete('recovery.json');
  }

  async openDialog(): Promise<void> {
    if (!(await this.confirmDiscard())) return;
    const pick = await fs.pickOpen();
    if (!pick) return;
    if ('path' in pick) await this.openPath(pick.path, true);
    else await this.openFile(pick.file);
  }

  async openPath(path: string, skipConfirm = false): Promise<void> {
    if (!skipConfirm && !(await this.confirmDiscard())) return;
    try {
      const bytes = await fs.readBytes(path);
      const e = fs.ext(path);
      const { wb, kind, active } = this.parseFile(bytes, e);
      fs.pushRecent(path);
      this.loadWorkbook(wb, { path, name: fs.baseName(path).replace(/\.[^.]+$/, ''), kind, dirty: false }, active);
      if (kind !== 'xlsx' && kind !== 'sheetly') this.doc.path = e === 'csv' ? path : null;
      this.doc.name = fs.baseName(path);
      this.updateTitle();
    } catch (err) {
      fs.removeRecent(path);
      await messageBox('Sheetly', `Sorry, Sheetly couldn't open ${fs.baseName(path)}.\n\n${(err as Error).message ?? err}`);
    }
  }

  async openFile(file: File): Promise<void> {
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const { wb, kind, active } = this.parseFile(bytes, fs.ext(file.name));
      this.loadWorkbook(wb, { path: null, name: file.name, kind }, active);
    } catch (err) {
      await messageBox('Sheetly', `Sorry, Sheetly couldn't open ${file.name}.\n\n${(err as Error).message ?? err}`);
    }
  }

  private parseFile(bytes: Uint8Array, e: string): { wb: Workbook; kind: DocKind; active: number } {
    if (e === 'csv' || e === 'txt') return { wb: readCSV(new TextDecoder('utf-8').decode(bytes)), kind: 'csv', active: 0 };
    if (e === 'sheetly') {
      const { wb, active } = fromJSON(new TextDecoder().decode(bytes));
      return { wb, kind: 'sheetly', active };
    }
    return { wb: readXlsx(bytes), kind: 'xlsx', active: 0 };
  }

  async save(): Promise<boolean> {
    if (!this.doc.path) return this.saveAs();
    return this.writeTo(this.doc.path);
  }

  async saveAs(forceKind?: DocKind): Promise<boolean> {
    const base = this.doc.name.replace(/\.[^.]+$/, '');
    const kind = forceKind ?? this.doc.kind;
    const filters =
      kind === 'csv'
        ? [{ name: 'CSV (Comma delimited)', extensions: ['csv'] }, { name: 'Excel Workbook', extensions: ['xlsx'] }, { name: 'Sheetly Workbook', extensions: ['sheetly'] }]
        : kind === 'sheetly'
          ? [{ name: 'Sheetly Workbook', extensions: ['sheetly'] }, { name: 'Excel Workbook', extensions: ['xlsx'] }, { name: 'CSV (Comma delimited)', extensions: ['csv'] }]
          : [{ name: 'Excel Workbook', extensions: ['xlsx'] }, { name: 'CSV (Comma delimited)', extensions: ['csv'] }, { name: 'Sheetly Workbook', extensions: ['sheetly'] }];
    let path = await fs.pickSave(`${base}.${kind}`, filters);
    if (!path) return false;
    if (!/\.(xlsx|csv|sheetly)$/i.test(path)) path += '.' + filters[0].extensions[0];
    return this.writeTo(path);
  }

  private async writeTo(path: string): Promise<boolean> {
    try {
      if (this.grid.editor.active) this.grid.commitEdit('none');
      const e = fs.ext(path);
      if (e === 'csv') {
        if (this.wb.sheets.length > 1) {
          const r = await messageBox('Sheetly', 'The selected file type does not support workbooks that contain multiple sheets.\n\nOnly the active sheet will be saved. Do you want to continue?', ['Save', 'Cancel']);
          if (r !== 'Save') return false;
        }
        await fs.writeBytes(path, new TextEncoder().encode(writeCSV(this.wb, this.sheet)));
      } else if (e === 'sheetly') await fs.writeBytes(path, new TextEncoder().encode(toJSON(this.wb, this.wb.sheets.indexOf(this.sheet))));
      else await fs.writeBytes(path, writeXlsx(this.wb));
      this.doc.path = path;
      this.doc.name = fs.baseName(path);
      this.doc.kind = (e === 'csv' ? 'csv' : e === 'sheetly' ? 'sheetly' : 'xlsx') as DocKind;
      this.doc.dirty = false;
      this.autosavedCount = this.changeCount;
      fs.pushRecent(path);
      await fs.dataDelete('recovery.json');
      this.updateTitle();
      toast('Saved ' + this.doc.name);
      return true;
    } catch (err) {
      await messageBox('Sheetly', `Sheetly couldn't save the file.\n\n${(err as Error).message ?? err}`);
      return false;
    }
  }

  // ------------------------------------------------------------------ autosave
  private async autosave(): Promise<void> {
    if (!this.doc.dirty || this.changeCount === this.autosavedCount) return;
    const count = this.changeCount;
    try {
      const json = toJSON(this.wb, this.wb.sheets.indexOf(this.sheet));
      await fs.dataWrite('recovery.json', JSON.stringify({ path: this.doc.path, name: this.doc.name, kind: this.doc.kind, time: Date.now(), book: json }));
      this.autosavedCount = count; // only after the write succeeded
    } catch {
      /* best effort; retried on the next tick */
    }
  }

  /** Offers to recover an autosave. Returns true when the user recovered it. */
  private async checkRecovery(): Promise<boolean> {
    const raw = await fs.dataRead('recovery.json');
    if (!raw) return false;
    try {
      const rec = JSON.parse(raw) as { path: string | null; name: string; kind: DocKind; time: number; book: string };
      const when = new Date(rec.time).toLocaleString();
      const r = await messageBox('Document Recovery', `Sheetly found unsaved changes to "${rec.name}" from ${when}.\n\nDo you want to recover them?`, ['Recover', 'Discard']);
      if (r === 'Recover') {
        const { wb, active } = fromJSON(rec.book);
        this.loadWorkbook(wb, { path: rec.path, name: rec.name, kind: rec.kind, dirty: true }, active);
        return true;
      }
      await fs.dataDelete('recovery.json');
    } catch {
      // keep the unreadable file for inspection instead of deleting it
      try {
        await fs.dataWrite('recovery.bad.json', raw);
        await fs.dataDelete('recovery.json');
      } catch {
        /* ignore */
      }
    }
    return false;
  }

  // ------------------------------------------------------------------ search box (Tell me)
  commandIndex(): { label: string; run: () => void }[] {
    return this.ribbon.index;
  }
}

import { shiftFormula as shiftFormulaImpl } from '../core/parser';

export function rectAbs(r: Rect): string {
  const one = (rr: number, cc: number) => '$' + colName(cc) + '$' + (rr + 1);
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return '$' + colName(r.c1) + ':$' + colName(r.c2);
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return '$' + (r.r1 + 1) + ':$' + (r.r2 + 1);
  return r.r1 === r.r2 && r.c1 === r.c2 ? one(r.r1, r.c1) : one(r.r1, r.c1) + ':' + one(r.r2, r.c2);
}

export function detectHeader(wb: Workbook, sheet: Sheet, rg: Rect): boolean {
  if (rg.r2 - rg.r1 < 1) return false;
  let textTop = 0;
  let total = 0;
  let typeDiff = false;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const a = sheet.get(rg.r1, c);
    const b = sheet.get(rg.r1 + 1, c);
    if (a && a.v !== null) {
      total++;
      if (typeof a.v === 'string') textTop++;
    }
    if (a && b && typeof a.v === 'string' && typeof b.v !== 'string' && b.v !== null) typeDiff = true;
  }
  return total > 0 && textTop === total && (typeDiff || rg.c2 === rg.c1 ? typeDiff : true);
}

export { displayText, rectText, normRect, FUNCS, closeMenus, promptBox, openDialog, FErr };
