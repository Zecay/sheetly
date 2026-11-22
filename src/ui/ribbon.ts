// Ribbon, tab strip and quick access toolbar.
import { FUNCS } from '../core/eval';
import { FMT_ACCOUNTING, FMT_CURRENCY } from '../core/numfmt';
import { MenuItem, h, showColorMenu, showMenuBelow } from './dom';
import { icon } from './icons';
import type { App } from './app';
import { CELL_STYLES } from './app';

type Tip = [string, string, string?];

interface BtnOpts {
  icon?: string;
  label?: string;
  tip?: Tip;
  cmd?: () => void;
  large?: boolean;
  menu?: () => MenuItem[];
  split?: boolean; // main part runs cmd, caret opens menu
  on?: () => boolean;
  disabled?: () => boolean;
  bar?: () => string; // colour bar under the icon
  left?: boolean;
}

const FONTS = ['Calibri', 'Arial', 'Arial Black', 'Cambria', 'Candara', 'Comic Sans MS', 'Consolas', 'Courier New', 'Georgia', 'Impact', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

const CATEGORIES: Record<string, string[]> = {
  Logical: ['IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'SWITCH', 'TRUE', 'FALSE'],
  Text: ['CONCAT', 'CONCATENATE', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'UPPER', 'LOWER', 'PROPER', 'TRIM', 'CLEAN', 'EXACT', 'REPT', 'CHAR', 'CODE', 'VALUE', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'TEXT', 'NUMBERVALUE', 'T'],
  'Date & Time': ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'DATEVALUE', 'TIMEVALUE', 'DAYS', 'EDATE', 'EOMONTH', 'DATEDIF'],
  'Lookup & Reference': ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'LOOKUP', 'INDEX', 'MATCH', 'XMATCH', 'CHOOSE', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS'],
  'Math & Trig': ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'PRODUCT', 'ABS', 'SIGN', 'INT', 'TRUNC', 'MOD', 'POWER', 'SQRT', 'EXP', 'LN', 'LOG', 'LOG10', 'PI', 'RAND', 'RANDBETWEEN', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'CEILING', 'FLOOR', 'SIN', 'COS', 'TAN', 'ATAN', 'DEGREES', 'RADIANS'],
  Statistical: ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'MIN', 'MAX', 'MINIFS', 'MAXIFS', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MEDIAN', 'MODE', 'LARGE', 'SMALL', 'RANK', 'STDEV', 'STDEV.S', 'STDEV.P', 'VAR', 'VAR.S', 'VAR.P'],
  Information: ['ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL', 'ISBLANK', 'ISERROR', 'ISERR', 'ISNA', 'ISEVEN', 'ISODD', 'NA', 'N'],
};
export const FUNCTION_CATEGORIES = CATEGORIES;

export class Ribbon {
  index: { label: string; run: () => void }[] = [];
  private panels = new Map<string, HTMLElement>();
  private updaters: (() => void)[] = [];
  private current = 'home';
  private lastFill = '#ffff00';
  private lastFont = '#ff0000';
  private tabButtons = new Map<string, HTMLElement>();
  private tip: HTMLElement;
  private tipTimer = 0;

  constructor(private app: App) {
    this.tip = h('div', { id: 'tip', style: { display: 'none' } });
    document.body.append(this.tip);
    this.buildQat();
    this.buildTabs();
    this.buildSearch();
    this.show('home');
    this.bindTips();
  }

  // ------------------------------------------------------------------ chrome
  private buildQat(): void {
    const q = document.getElementById('qat')!;
    const mk = (ic: string, key: string, title: string, fn: () => void) => {
      const b = h('button', { type: 'button', title, 'data-qat': key, html: icon(ic, 16) });
      b.addEventListener('click', fn);
      q.append(b);
      return b;
    };
    mk('save', 'save', 'Save (Ctrl+S)', () => this.app.run('save'));
    mk('undo', 'undo', 'Undo (Ctrl+Z)', () => this.app.run('undo'));
    mk('redo', 'redo', 'Redo (Ctrl+Y)', () => this.app.run('redo'));
  }

  private buildTabs(): void {
    const tabs = this.app.el.tabs;
    const names: [string, string][] = [['home', 'Home'], ['insert', 'Insert'], ['layout', 'Page Layout'], ['formulas', 'Formulas'], ['data', 'Data'], ['review', 'Review'], ['view', 'View']];
    const file = h('button', { class: 'file', type: 'button' }, 'File');
    file.addEventListener('click', () => this.app.run('backstage', 'info'));
    tabs.append(file);
    for (const [id, label] of names) {
      const b = h('button', { type: 'button', role: 'tab' }, label);
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => this.show(id));
      tabs.append(b);
      this.tabButtons.set(id, b);
      this.panels.set(id, this.buildPanel(id));
    }
  }

  private buildSearch(): void {
    const input = document.getElementById('tb-search') as HTMLInputElement;
    input.addEventListener('input', () => {
      const q = input.value.trim().toLowerCase();
      if (!q) return;
      const items: MenuItem[] = this.index
        .filter((i) => i.label.toLowerCase().includes(q))
        .slice(0, 10)
        .map((i) => ({ label: i.label, cmd: () => i.run() }));
      if (!items.length) items.push({ label: 'No commands found', disabled: true });
      const r = input.getBoundingClientRect();
      showMenuBelow(input, items);
      void r;
      input.focus();
    });
    input.addEventListener('keydown', (e) => e.stopPropagation());
  }

  show(id: string): void {
    this.current = id;
    const panel = this.panels.get(id)!;
    const host = this.app.el.ribbon;
    host.innerHTML = '';
    host.append(panel);
    for (const [k, b] of this.tabButtons) b.classList.toggle('active', k === id);
    this.update();
  }

  update(): void {
    for (const u of this.updaters) u();
  }

  // ------------------------------------------------------------------ tooltips
  private bindTips(): void {
    const roots = [this.app.el.ribbon, document.getElementById('qat')!, document.getElementById('titlebar')!];
    for (const root of roots) {
      root.addEventListener('mouseover', (e) => {
        const el = (e.target as HTMLElement).closest('[data-tip]') as HTMLElement | null;
        clearTimeout(this.tipTimer);
        this.tip.style.display = 'none';
        if (!el) return;
        this.tipTimer = window.setTimeout(() => {
          const t = JSON.parse(el.dataset.tip!) as Tip;
          this.tip.innerHTML = `<b>${t[0]}${t[2] ? ` <span class="k">(${t[2]})</span>` : ''}</b>${t[1]}`;
          const r = el.getBoundingClientRect();
          this.tip.style.display = 'block';
          this.tip.style.left = Math.min(window.innerWidth - 270, Math.max(4, r.left)) + 'px';
          this.tip.style.top = r.bottom + 6 + 'px';
        }, 650);
      });
      root.addEventListener('mouseleave', () => {
        clearTimeout(this.tipTimer);
        this.tip.style.display = 'none';
      });
      root.addEventListener('mousedown', () => {
        clearTimeout(this.tipTimer);
        this.tip.style.display = 'none';
      });
    }
  }

  // ------------------------------------------------------------------ builders
  private reg(label: string | undefined, run: () => void): void {
    if (label) this.index.push({ label: label.replace(/…$/, ''), run });
  }

  private mkBtn(o: BtnOpts): HTMLElement {
    const large = !!o.large;
    const caretEl = () => h('span', { class: 'caret' });
    const build = (withLabel: boolean, withCaret: boolean, icoHtml: string) =>
      h('button', { class: 'rb-btn' + (large ? ' large' : ' small') + (o.left ? ' left' : '') + (o.bar ? ' has-bar' : ''), type: 'button' },
        h('span', { html: icoHtml, style: { display: 'flex' } }),
        withLabel && o.label ? h('span', { class: 'lbl' }, o.label) : null,
        withCaret ? caretEl() : null,
        o.bar ? h('span', { class: 'color-bar', style: { background: o.bar() } }) : null);
    const size = large ? 28 : 16;
    const ico = o.icon ? icon(o.icon, size) : '';
    const tip = o.tip ?? (o.label ? [o.label.replace(/…$/, ''), ''] : undefined);
    const attachTip = (el: HTMLElement) => {
      if (tip) el.dataset.tip = JSON.stringify(tip);
      el.addEventListener('mousedown', (e) => e.preventDefault()); // keep grid focus
    };
    let root: HTMLElement;
    const refresh: (() => void)[] = [];
    if (o.menu && o.split) {
      if (large) {
        root = h('div', { class: 'rb-split' });
        const top = build(false, false, ico);
        const bottom = h('button', { class: 'rb-btn small', type: 'button', style: { height: '20px', padding: '0 4px', marginTop: '-2px' } }, o.label ? h('span', { class: 'lbl' }, o.label) : null, caretEl());
        root.append(top, bottom);
        top.addEventListener('click', () => o.cmd?.());
        bottom.addEventListener('click', () => showMenuBelow(bottom, o.menu!()));
        attachTip(top);
        attachTip(bottom);
      } else {
        root = h('div', { class: 'rb-split h' });
        const main = build(!!o.label, false, ico);
        const arrow = h('button', { class: 'rb-btn small', type: 'button', style: { minWidth: '14px', padding: '0 2px' } }, caretEl());
        root.append(main, arrow);
        main.addEventListener('click', () => o.cmd?.());
        arrow.addEventListener('click', () => showMenuBelow(root, o.menu!()));
        attachTip(main);
        attachTip(arrow);
        refresh.push(() => {
          const on = o.on?.() ?? false;
          main.classList.toggle('on', on);
          const bar = main.querySelector('.color-bar') as HTMLElement | null;
          if (bar && o.bar) bar.style.background = o.bar();
        });
      }
    } else if (o.menu) {
      const b = build(true, true, ico);
      root = b;
      b.addEventListener('click', () => showMenuBelow(b, o.menu!()));
      attachTip(b);
    } else {
      const b = build(true, false, ico);
      root = b;
      b.addEventListener('click', () => o.cmd?.());
      attachTip(b);
      refresh.push(() => {
        if (o.on) b.classList.toggle('on', o.on());
        if (o.disabled) (b as HTMLButtonElement).disabled = o.disabled();
      });
    }
    if (o.cmd) this.reg(o.label, o.cmd);
    else if (o.menu && o.label) this.reg(o.label, () => showMenuBelow(root, o.menu!()));
    for (const f of refresh) this.updaters.push(f);
    return root;
  }

  private small(o: BtnOpts) {
    return this.mkBtn({ ...o, large: false });
  }
  private large(o: BtnOpts) {
    return this.mkBtn({ ...o, large: true });
  }

  private group(label: string, body: (HTMLElement | null)[], launcher?: () => void): HTMLElement {
    const g = h('div', { class: 'rb-group' }, h('div', { class: 'rb-body' }, body), h('div', { class: 'rb-label' }, label));
    if (launcher) {
      const l = h('button', { class: 'rb-launch', type: 'button', title: 'Format Cells' });
      l.addEventListener('click', launcher);
      g.append(l);
    }
    return g;
  }
  private col(...items: (HTMLElement | null)[]): HTMLElement {
    return h('div', { class: 'rb-col' }, items);
  }
  private row(...items: (HTMLElement | null)[]): HTMLElement {
    return h('div', { class: 'rb-row' }, items);
  }
  private sep(): HTMLElement {
    return h('div', { class: 'rb-sep' });
  }
  private check(label: string, get: () => boolean, set: () => void): HTMLElement {
    const inp = h('input', { type: 'checkbox' }) as HTMLInputElement;
    inp.addEventListener('change', () => {
      set();
      this.app.grid.focusGrid();
    });
    this.updaters.push(() => (inp.checked = get()));
    this.reg(label, set);
    return h('label', { class: 'rb-check' }, inp, label);
  }

  private combo(width: number, values: () => (string | number)[], get: () => string, apply: (v: string) => void, tip: Tip): HTMLElement {
    const input = h('input', { type: 'text', style: { width: width + 'px' }, spellcheck: 'false' }) as HTMLInputElement;
    const arrow = h('button', { type: 'button', tabindex: '-1', html: icon('chevron', 9) });
    const wrap = h('div', { class: 'rb-combo' }, input, arrow);
    wrap.dataset.tip = JSON.stringify(tip);
    arrow.addEventListener('mousedown', (e) => e.preventDefault());
    arrow.addEventListener('click', () =>
      showMenuBelow(wrap, values().map((v) => ({ label: String(v), checked: String(v) === input.value, cmd: () => apply(String(v)) }))),
    );
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') {
        apply(input.value);
        this.app.grid.focusGrid();
      } else if (e.key === 'Escape') {
        input.value = get();
        this.app.grid.focusGrid();
      }
    });
    input.addEventListener('focus', () => input.select());
    this.updaters.push(() => {
      if (document.activeElement !== input) input.value = get();
    });
    return wrap;
  }

  // ------------------------------------------------------------------ panels
  private buildPanel(id: string): HTMLElement {
    const p = h('div', { class: 'rb-panel' });
    const groups = (this as any)['panel_' + id]() as HTMLElement[];
    p.append(...groups);
    return p;
  }

  private colorSplit(kind: 'fill' | 'font'): HTMLElement {
    const isFill = kind === 'fill';
    const root = h('div', { class: 'rb-split h' });
    const main = h('button', { class: 'rb-btn small has-bar', type: 'button' }, h('span', { html: icon(isFill ? 'fill' : 'fontcolor', 16), style: { display: 'flex' } }), h('span', { class: 'color-bar' }));
    const arrow = h('button', { class: 'rb-btn small', type: 'button', style: { minWidth: '14px', padding: '0 2px' } }, h('span', { class: 'caret' }));
    const tip: Tip = isFill ? ['Fill Color', 'Color the background of the selected cells to make them stand out.'] : ['Font Color', 'Change the text color.'];
    for (const b of [main, arrow]) {
      b.dataset.tip = JSON.stringify(tip);
      b.addEventListener('mousedown', (e) => e.preventDefault());
    }
    root.append(main, arrow);
    const bar = main.querySelector('.color-bar') as HTMLElement;
    const setBar = () => (bar.style.background = isFill ? this.lastFill : this.lastFont);
    main.addEventListener('click', () => this.app.run(isFill ? 'fillColor' : 'fontColor', isFill ? this.lastFill : this.lastFont));
    arrow.addEventListener('click', () =>
      showColorMenu(root, {
        noneLabel: isFill ? 'No Fill' : 'Automatic',
        onPick: (c) => {
          if (c) {
            if (isFill) this.lastFill = c;
            else this.lastFont = c;
            setBar();
          }
          void this.app.run(isFill ? 'fillColor' : 'fontColor', c);
        },
      }),
    );
    this.updaters.push(setBar);
    this.reg(isFill ? 'Fill Color' : 'Font Color', () => this.app.run(isFill ? 'fillColor' : 'fontColor', isFill ? this.lastFill : this.lastFont));
    return root;
  }

  private bordersMenu(): MenuItem[] {
    const m = (label: string, kind: string, icon: string): MenuItem => ({ label, icon, cmd: () => this.app.run('borders', kind) });
    return [
      m('Bottom Border', 'bottom', 'bdBottom'), m('Top Border', 'top', 'bdTop'), m('Left Border', 'left', 'bdLeft'), m('Right Border', 'right', 'bdRight'),
      { sep: true },
      m('No Border', 'none', 'bdNone'), m('All Borders', 'all', 'bdAll'), m('Outside Borders', 'outside', 'bdOutside'), m('Thick Outside Borders', 'thickOutside', 'bdThickOutside'), m('Inside Borders', 'inside', 'bdInside'),
      { sep: true },
      m('Bottom Double Border', 'bottomDouble', 'bdBottomDouble'), m('Thick Bottom Border', 'thickBottom', 'bdThickBottom'), m('Top and Bottom Border', 'topBottom', 'bdTopBottom'), m('Top and Thick Bottom Border', 'topThickBottom', 'bdTopThickBottom'),
    ];
  }

  private numberFormatMenu(): MenuItem[] {
    const f = (label: string, code: string): MenuItem => ({ label, cmd: () => this.app.run('numFmt', code) });
    return [
      f('General', 'General'), f('Number', '0.00'), f('Currency', FMT_CURRENCY), f('Accounting', FMT_ACCOUNTING), f('Short Date', 'm/d/yyyy'),
      f('Long Date', 'dddd, mmmm d, yyyy'), f('Time', 'h:mm:ss AM/PM'), f('Percentage', '0.00%'), f('Fraction', '# ?/?'), f('Scientific', '0.00E+00'), f('Text', '@'),
      { sep: true },
      { label: 'More Number Formats…', cmd: () => this.app.run('formatCells', 'number') },
    ];
  }

  private formatName(): string {
    const f = this.app.activeStyle().fmt ?? 'General';
    const map: [string, string][] = [['General', 'General'], ['0.00', 'Number'], ['0', 'Number'], ['#,##0.00', 'Number'], ['#,##0', 'Number'], ['@', 'Text'], ['m/d/yyyy', 'Date'], ['0%', 'Percentage'], ['0.00%', 'Percentage'], ['0.00E+00', 'Scientific'], ['# ?/?', 'Fraction'], ['h:mm:ss AM/PM', 'Time']];
    for (const [k, v] of map) if (f === k) return v;
    if (/\$|"\$"/.test(f) && /_\(/.test(f)) return 'Accounting';
    if (/\$/.test(f)) return 'Currency';
    if (/%/.test(f)) return 'Percentage';
    if (/[ymd]/i.test(f.replace(/"[^"]*"/g, ''))) return /[hs]/i.test(f) && /[ymd]/.test(f) ? 'Custom' : /h|s/.test(f) ? 'Time' : 'Date';
    if (/E[+-]/.test(f)) return 'Scientific';
    return 'Custom';
  }

  private functionMenu(cat: string): MenuItem[] {
    return CATEGORIES[cat].filter((n) => FUNCS.has(n)).map((n) => ({ label: n, cmd: () => this.app.run('insertFn', n) }));
  }

  private stylesMenu(): MenuItem[] {
    const out: MenuItem[] = [];
    let group = '';
    for (const s of CELL_STYLES) {
      if (s.group !== group) {
        group = s.group;
        out.push({ heading: group });
      }
      out.push({ label: s.name, cmd: () => this.app.run('cellStyle', s.name) });
    }
    return out;
  }

  private cfMenu(): MenuItem[] {
    const c = (label: string, kind: string): MenuItem => ({ label, cmd: () => this.app.run('cf', kind) });
    return [
      { label: 'Highlight Cells Rules', submenu: [
        c('Greater Than…', 'gt'), c('Less Than…', 'lt'), c('Between…', 'between'), c('Equal To…', 'eq'), c('Text that Contains…', 'text'), c('Duplicate Values…', 'dup'),
      ] },
      { label: 'Top/Bottom Rules', submenu: [c('Top 10 Items…', 'top'), c('Bottom 10 Items…', 'bottom'), c('Above Average…', 'above'), c('Below Average…', 'below')] },
      { label: 'Data Bars', submenu: [
        { label: 'Blue Data Bar', cmd: () => this.app.run('cfBar', '#638ec6') },
        { label: 'Green Data Bar', cmd: () => this.app.run('cfBar', '#63c384') },
        { label: 'Red Data Bar', cmd: () => this.app.run('cfBar', '#ff555a') },
        { label: 'Orange Data Bar', cmd: () => this.app.run('cfBar', '#ffb628') },
      ] },
      { label: 'Color Scales', submenu: [
        { label: 'Green - Yellow - Red Color Scale', cmd: () => this.app.run('cfScale', ['#63be7b', '#ffeb84', '#f8696b']) },
        { label: 'Red - Yellow - Green Color Scale', cmd: () => this.app.run('cfScale', ['#f8696b', '#ffeb84', '#63be7b']) },
        { label: 'Green - White Color Scale', cmd: () => this.app.run('cfScale', ['#63be7b', '#ffffff']) },
        { label: 'White - Green Color Scale', cmd: () => this.app.run('cfScale', ['#ffffff', '#63be7b']) },
        { label: 'Blue - White - Red Color Scale', cmd: () => this.app.run('cfScale', ['#5a8ac6', '#fcfcff', '#f8696b']) },
        { label: 'Yellow - Red Color Scale', cmd: () => this.app.run('cfScale', ['#ffeb84', '#f8696b']) },
      ] },
      { sep: true },
      { label: 'New Rule…', cmd: () => this.app.run('cf', 'formula') },
      { label: 'Clear Rules', submenu: [
        { label: 'Clear Rules from Selected Cells', cmd: () => this.app.run('cfClear', 'selection') },
        { label: 'Clear Rules from Entire Sheet', cmd: () => this.app.run('cfClear', 'sheet') },
      ] },
      { label: 'Manage Rules…', cmd: () => this.app.run('cfManage') },
    ];
  }

  // ------------------------------------------------------------------ tab: Home
  panel_home(): HTMLElement[] {
    const a = this.app;
    const st = () => a.activeStyle();
    const clipboard = this.group('Clipboard', [
      this.large({
        icon: 'paste', label: 'Paste', tip: ['Paste', 'Add content from the Clipboard to your document.', 'Ctrl+V'], split: true,
        cmd: () => a.run('paste'),
        menu: () => [
          { label: 'Paste', icon: 'paste', cmd: () => a.run('paste') },
          { label: 'Values', cmd: () => a.run('paste', 'values') },
          { label: 'Formulas', cmd: () => a.run('paste', 'formulas') },
          { label: 'Formatting', cmd: () => a.run('paste', 'formats') },
          { label: 'Transpose', cmd: () => a.run('paste', 'transpose') },
          { sep: true },
          { label: 'Paste Special…', cmd: () => a.run('pasteSpecial') },
        ],
      }),
      this.col(
        this.small({ icon: 'cut', label: 'Cut', left: true, tip: ['Cut', 'Remove the selection and put it on the Clipboard.', 'Ctrl+X'], cmd: () => a.run('cut') }),
        this.small({ icon: 'copy', label: 'Copy', left: true, tip: ['Copy', 'Put a copy of the selection on the Clipboard.', 'Ctrl+C'], cmd: () => a.run('copy') }),
        this.small({ icon: 'painter', label: 'Format Painter', left: true, tip: ['Format Painter', 'Copy the look of one place and apply it to another. Double-click to apply to several places.'], cmd: () => a.run('painter'), on: () => a.painterActive() }),
      ),
    ]);
    const font = this.group('Font', [
      this.col(
        this.row(
          this.combo(104, () => FONTS, () => st().font ?? 'Calibri', (v) => a.run('font', v), ['Font', 'Pick a new font for your text.']),
          this.combo(34, () => SIZES, () => String(st().size ?? 11), (v) => { const n = parseFloat(v); if (n > 0 && n <= 409) void a.run('size', n); }, ['Font Size', 'Change the size of your text.']),
          this.small({ icon: 'plus', tip: ['Increase Font Size', 'Make your text a bit bigger.', 'Ctrl+Shift+>'], cmd: () => a.run('sizeInc') }),
          this.small({ icon: 'minimize', tip: ['Decrease Font Size', 'Make your text a bit smaller.', 'Ctrl+Shift+<'], cmd: () => a.run('sizeDec') }),
        ),
        this.row(
          this.small({ icon: 'bold', tip: ['Bold', 'Make your text bold.', 'Ctrl+B'], cmd: () => a.run('bold'), on: () => !!st().bold }),
          this.small({ icon: 'italic', tip: ['Italic', 'Italicize your text.', 'Ctrl+I'], cmd: () => a.run('italic'), on: () => !!st().italic }),
          this.small({ icon: 'underline', tip: ['Underline', 'Underline your text.', 'Ctrl+U'], cmd: () => a.run('underline'), on: () => !!st().underline }),
          this.small({ icon: 'strike', tip: ['Strikethrough', 'Cross something out by drawing a line through it.', 'Ctrl+5'], cmd: () => a.run('strike'), on: () => !!st().strike }),
          this.sep(),
          this.small({ icon: 'borders', tip: ['Borders', 'Apply borders to the selected cells.'], menu: () => this.bordersMenu() }),
          this.colorSplit('fill'),
          this.colorSplit('font'),
        ),
      ),
    ], () => a.run('formatCells', 'font'));
    const align = this.group('Alignment', [
      this.col(
        this.row(
          this.small({ icon: 'alignTop', tip: ['Top Align', 'Align text to the top of the cell.'], cmd: () => a.run('valign', 'top'), on: () => st().v === 'top' }),
          this.small({ icon: 'alignMiddle', tip: ['Middle Align', 'Align text so that it is centered between the top and bottom of the cell.'], cmd: () => a.run('valign', 'middle'), on: () => st().v === 'middle' }),
          this.small({ icon: 'alignBottom', tip: ['Bottom Align', 'Align text to the bottom of the cell.'], cmd: () => a.run('valign', 'bottom'), on: () => !st().v }),
          this.sep(),
          this.small({ icon: 'wrap', label: 'Wrap Text', tip: ['Wrap Text', 'Wrap extra-long text into multiple lines so you can see all of it.'], cmd: () => a.run('wrap'), on: () => !!st().wrap }),
        ),
        this.row(
          this.small({ icon: 'alignLeft', tip: ['Align Left', 'Align text to the left.'], cmd: () => a.run('halign', 'left'), on: () => st().h === 'left' }),
          this.small({ icon: 'alignCenter', tip: ['Center', 'Center your text.'], cmd: () => a.run('halign', 'center'), on: () => st().h === 'center' }),
          this.small({ icon: 'alignRight', tip: ['Align Right', 'Align text to the right.'], cmd: () => a.run('halign', 'right'), on: () => st().h === 'right' }),
          this.sep(),
          this.small({ icon: 'indentDec', tip: ['Decrease Indent', 'Move your content closer to the cell border.'], cmd: () => a.run('indent', -1) }),
          this.small({ icon: 'indentInc', tip: ['Increase Indent', 'Move your content farther from the cell border.'], cmd: () => a.run('indent', 1) }),
          this.small({
            icon: 'merge', label: 'Merge & Center', split: true, tip: ['Merge & Center', 'Join the selected cells into one larger cell and center the contents.'],
            cmd: () => a.run('merge', 'center'),
            menu: () => [
              { label: 'Merge & Center', cmd: () => a.run('merge', 'center') },
              { label: 'Merge Across', cmd: () => a.run('merge', 'across') },
              { label: 'Merge Cells', cmd: () => a.run('merge', 'merge') },
              { label: 'Unmerge Cells', cmd: () => a.run('merge', 'unmerge') },
            ],
          }),
        ),
      ),
    ], () => a.run('formatCells', 'alignment'));
    const number = this.group('Number', [
      this.col(
        this.row(this.numberCombo()),
        this.row(
          this.small({
            icon: 'currency', tip: ['Accounting Number Format', 'Choose an alternate currency format for the selected cell.'], split: true,
            cmd: () => a.run('currency', FMT_ACCOUNTING),
            menu: () => [
              { label: '$ English (United States)', cmd: () => a.run('currency', FMT_ACCOUNTING) },
              { label: '€ Euro (€ 123)', cmd: () => a.run('numFmt', '[$€-2] #,##0.00') },
              { label: '£ English (United Kingdom)', cmd: () => a.run('numFmt', '[$£-809]#,##0.00') },
              { label: '¥ Japanese (Japan)', cmd: () => a.run('numFmt', '[$¥-411]#,##0') },
              { sep: true },
              { label: 'More Accounting Formats…', cmd: () => a.run('formatCells', 'number') },
            ],
          }),
          this.small({ icon: 'percent', tip: ['Percent Style', 'Display the value of the cell as a percentage.', 'Ctrl+Shift+%'], cmd: () => a.run('percent') }),
          this.small({ icon: 'comma', tip: ['Comma Style', 'Display the value of the cell with a thousands separator.'], cmd: () => a.run('comma') }),
          this.small({ icon: 'decInc', tip: ['Increase Decimal', 'Show more precise values by showing more decimal places.'], cmd: () => a.run('decInc') }),
          this.small({ icon: 'decDec', tip: ['Decrease Decimal', 'Show less precise values by showing fewer decimal places.'], cmd: () => a.run('decDec') }),
        ),
      ),
    ], () => a.run('formatCells', 'number'));
    const styles = this.group('Styles', [
      this.large({ icon: 'condfmt', label: 'Conditional Formatting', tip: ['Conditional Formatting', 'Highlight interesting cells, emphasize unusual values and visualize data using data bars and color scales.'], menu: () => this.cfMenu() }),
      this.large({ icon: 'styles', label: 'Cell Styles', tip: ['Cell Styles', 'Quickly format a cell by choosing from pre-defined styles.'], menu: () => this.stylesMenu() }),
    ]);
    const cells = this.group('Cells', [
      this.large({
        icon: 'insertCells', label: 'Insert', tip: ['Insert', 'Insert cells, rows, columns or sheets.'],
        menu: () => [
          { label: 'Insert Cells…', icon: 'insertCells', cmd: () => a.run('insertDialog') },
          { label: 'Insert Sheet Rows', icon: 'rowsInsert', cmd: () => a.run('insertRows') },
          { label: 'Insert Sheet Columns', icon: 'colsInsert', cmd: () => a.run('insertCols') },
          { label: 'Insert Sheet', icon: 'sheetAdd', cmd: () => a.run('addSheet') },
        ],
      }),
      this.large({
        icon: 'deleteCells', label: 'Delete', tip: ['Delete', 'Delete cells, rows, columns or sheets.'],
        menu: () => [
          { label: 'Delete Cells…', icon: 'deleteCells', cmd: () => a.run('deleteDialog') },
          { label: 'Delete Sheet Rows', cmd: () => a.run('deleteRows') },
          { label: 'Delete Sheet Columns', cmd: () => a.run('deleteCols') },
          { label: 'Delete Sheet', cmd: () => a.run('deleteSheet') },
        ],
      }),
      this.large({
        icon: 'formatCells', label: 'Format', tip: ['Format', 'Change the row height or column width, organize sheets or protect/hide cells.'],
        menu: () => [
          { heading: 'Cell Size' },
          { label: 'Row Height…', cmd: () => a.run('rowHeight') },
          { label: 'AutoFit Row Height', cmd: () => a.run('autofitRows') },
          { label: 'Column Width…', cmd: () => a.run('colWidth') },
          { label: 'AutoFit Column Width', cmd: () => a.run('autofitCols') },
          { heading: 'Visibility' },
          { label: 'Hide Rows', cmd: () => a.run('hideRows') },
          { label: 'Hide Columns', cmd: () => a.run('hideCols') },
          { label: 'Unhide Rows', cmd: () => a.run('unhideRows') },
          { label: 'Unhide Columns', cmd: () => a.run('unhideCols') },
          { heading: 'Organize Sheets' },
          { label: 'Rename Sheet', cmd: () => a.run('renameSheet') },
          { label: 'Move Sheet Left', cmd: () => a.run('moveSheet', -1) },
          { label: 'Move Sheet Right', cmd: () => a.run('moveSheet', 1) },
          { label: 'Duplicate Sheet', cmd: () => a.run('dupSheet') },
          { label: 'Tab Color…', cmd: () => a.run('tabColor') },
          { sep: true },
          { label: 'Format Cells…', cmd: () => a.run('formatCells') },
        ],
      }),
    ]);
    const editing = this.group('Editing', [
      this.col(
        this.small({
          icon: 'autosum', label: 'AutoSum', split: true, tip: ['AutoSum', 'Automatically add a quick calculation to your worksheet, such as the sum or average.', 'Alt+='],
          cmd: () => a.run('autosum', 'SUM'),
          menu: () => ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((n) => ({ label: n === 'SUM' ? 'Sum' : n === 'AVERAGE' ? 'Average' : n === 'COUNT' ? 'Count Numbers' : n === 'MAX' ? 'Max' : 'Min', cmd: () => a.run('autosum', n) })),
        }),
        this.small({
          icon: 'fillDown', label: 'Fill', tip: ['Fill', 'Continue a pattern into one or more adjacent cells.'],
          menu: () => [
            { label: 'Down', shortcut: 'Ctrl+D', cmd: () => a.run('fill', 'down') },
            { label: 'Right', shortcut: 'Ctrl+R', cmd: () => a.run('fill', 'right') },
            { label: 'Up', cmd: () => a.run('fill', 'up') },
            { label: 'Left', cmd: () => a.run('fill', 'left') },
          ],
        }),
        this.small({
          icon: 'eraser', label: 'Clear', tip: ['Clear', 'Delete everything in the cell, or remove just the formatting, contents or comments.'],
          menu: () => [
            { label: 'Clear All', cmd: () => a.run('clear', 'all') },
            { label: 'Clear Formats', cmd: () => a.run('clear', 'formats') },
            { label: 'Clear Contents', cmd: () => a.run('clear', 'contents') },
          ],
        }),
      ),
      this.large({
        icon: 'sort', label: 'Sort & Filter', tip: ['Sort & Filter', 'Organize your data so it is easier to analyze.'],
        menu: () => [
          { label: 'Sort A to Z', cmd: () => a.run('sort', 'asc') },
          { label: 'Sort Z to A', cmd: () => a.run('sort', 'desc') },
          { label: 'Custom Sort…', cmd: () => a.run('sortDialog') },
          { sep: true },
          { label: 'Filter', icon: 'filter', checked: !!a.sheet.filter, cmd: () => a.run('filter') },
          { label: 'Clear', cmd: () => a.run('filterClear') },
          { label: 'Reapply', cmd: () => a.run('filterReapply') },
        ],
      }),
      this.large({
        icon: 'find', label: 'Find & Select', tip: ['Find & Select', 'Find and select specific text, formatting, or type of information within the workbook.'],
        menu: () => [
          { label: 'Find…', icon: 'find', shortcut: 'Ctrl+F', cmd: () => a.run('find') },
          { label: 'Replace…', shortcut: 'Ctrl+H', cmd: () => a.run('replace') },
          { label: 'Go To…', shortcut: 'Ctrl+G', cmd: () => a.run('goto') },
          { sep: true },
          { label: 'Select All', shortcut: 'Ctrl+A', cmd: () => a.run('selectAll') },
        ],
      }),
    ]);
    return [clipboard, font, align, number, styles, cells, editing];
  }

  private numberCombo(): HTMLElement {
    const btn = h('button', { class: 'rb-btn', type: 'button', style: { width: '128px', justifyContent: 'space-between', border: '1px solid var(--border)', background: 'var(--input-bg)', height: '22px' } }, h('span', {}, 'General'), h('span', { class: 'caret' }));
    btn.dataset.tip = JSON.stringify(['Number Format', 'Choose the format for the value in the cell, such as percentage, currency, date or time.']);
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    btn.addEventListener('click', () => showMenuBelow(btn, this.numberFormatMenu()));
    this.updaters.push(() => ((btn.firstChild as HTMLElement).textContent = this.formatName()));
    return btn;
  }

  // ------------------------------------------------------------------ tab: Insert
  panel_insert(): HTMLElement[] {
    const a = this.app;
    return [
      this.group('Sheets', [
        this.large({ icon: 'sheetAdd', label: 'New Sheet', tip: ['New Sheet', 'Add a new worksheet to this workbook.', 'Shift+F11'], cmd: () => a.run('addSheet') }),
        this.large({ icon: 'newDoc', label: 'Duplicate Sheet', tip: ['Duplicate Sheet', 'Make a copy of the current worksheet.'], cmd: () => a.run('dupSheet') }),
      ]),
      this.group('Rows & Columns', [
        this.large({ icon: 'rowsInsert', label: 'Insert Rows', tip: ['Insert Sheet Rows', 'Insert rows above the selection.'], cmd: () => a.run('insertRows') }),
        this.large({ icon: 'colsInsert', label: 'Insert Columns', tip: ['Insert Sheet Columns', 'Insert columns to the left of the selection.'], cmd: () => a.run('insertCols') }),
        this.large({ icon: 'insertCells', label: 'Insert Cells', tip: ['Insert Cells', 'Insert cells and shift existing ones right or down.', 'Ctrl+Shift++'], cmd: () => a.run('insertDialog') }),
      ]),
      this.group('Filters', [
        this.large({ icon: 'filter', label: 'Filter', tip: ['Filter', 'Turn on filtering so you can narrow down your data.', 'Ctrl+Shift+L'], cmd: () => a.run('filter'), on: () => !!a.sheet.filter }),
      ]),
      this.group('Date & Time', [
        this.large({ icon: 'calc', label: 'Insert Date', tip: ['Insert Date', "Insert today's date into the selected cell.", 'Ctrl+;'], cmd: () => a.run('insertDate') }),
        this.large({ icon: 'calc', label: 'Insert Time', tip: ['Insert Time', 'Insert the current time into the selected cell.', 'Ctrl+Shift+:'], cmd: () => a.run('insertTime') }),
      ]),
    ];
  }

  // ------------------------------------------------------------------ tab: Page Layout
  panel_layout(): HTMLElement[] {
    const a = this.app;
    return [
      this.group('Themes', [
        this.large({ icon: 'sun', label: 'Light', tip: ['Light Theme', 'Use the light color scheme.'], cmd: () => a.run('theme', false), on: () => !a.dark }),
        this.large({ icon: 'moon', label: 'Dark', tip: ['Dark Theme', 'Use the dark color scheme.'], cmd: () => a.run('theme', true), on: () => a.dark }),
      ]),
      this.group('Page Setup', [
        this.large({
          icon: 'print', label: 'Orientation', tip: ['Orientation', 'Switch between portrait and landscape layouts.'],
          menu: () => [
            { label: 'Portrait', checked: a.orientation === 'portrait', cmd: () => a.run('orientation', 'portrait') },
            { label: 'Landscape', checked: a.orientation === 'landscape', cmd: () => a.run('orientation', 'landscape') },
          ],
        }),
      ]),
      this.group('Sheet Options', [
        this.col(
          this.check('View Gridlines', () => a.gridlines, () => a.run('gridlines')),
          this.check('View Headings', () => a.headings, () => a.run('headings')),
          this.check('Print Gridlines', () => a.printGridlines, () => a.run('printGridlines')),
        ),
      ]),
      this.group('Print', [
        this.large({ icon: 'pdf', label: 'Print / PDF', tip: ['Print', 'Print the active sheet, or choose "Microsoft Print to PDF" to save it as a PDF.', 'Ctrl+P'], cmd: () => a.run('print') }),
      ]),
    ];
  }

  // ------------------------------------------------------------------ tab: Formulas
  panel_formulas(): HTMLElement[] {
    const a = this.app;
    const cat = (name: string, ic: string) => this.small({ icon: ic, label: name, tip: [name, `Browse ${name} functions.`], menu: () => this.functionMenu(name) });
    return [
      this.group('Function Library', [
        this.large({ icon: 'fx', label: 'Insert Function', tip: ['Insert Function', 'Edit the current cell by choosing functions and editing the arguments.', 'Shift+F3'], cmd: () => a.run('insertFunction') }),
        this.large({
          icon: 'autosum', label: 'AutoSum', tip: ['AutoSum', 'Automatically add a quick calculation to your worksheet.', 'Alt+='],
          menu: () => ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((n) => ({ label: n, cmd: () => a.run('autosum', n) })),
        }),
        this.col(cat('Logical', 'fx'), cat('Text', 'fx'), cat('Date & Time', 'fx')),
        this.col(cat('Lookup & Reference', 'fx'), cat('Math & Trig', 'fx'), cat('Statistical', 'fx')),
      ]),
      this.group('Defined Names', [
        this.large({ icon: 'names', label: 'Name Manager', tip: ['Name Manager', 'Create, edit, delete and find all the names used in the workbook.', 'Ctrl+F3'], cmd: () => a.run('nameManager') }),
        this.col(
          this.small({ icon: 'defineName', label: 'Define Name', left: true, tip: ['Define Name', 'Give a name to the selected cells.'], cmd: () => a.run('defineName') }),
          this.small({
            icon: 'fx', label: 'Use in Formula', left: true, tip: ['Use in Formula', 'Choose a name used in this workbook and insert it into the current formula.'],
            menu: () => {
              const names = Array.from(a.wb.names.values());
              return names.length ? names.map((n) => ({ label: n.name, cmd: () => a.run('insertFn', n.name) })) : [{ label: '(no names defined)', disabled: true }];
            },
          }),
        ),
      ]),
      this.group('Formula Auditing', [
        this.large({ icon: 'fx', label: 'Show Formulas', tip: ['Show Formulas', 'Display the formula in each cell instead of the resulting value.', 'Ctrl+`'], cmd: () => a.run('showFormulas'), on: () => a.showFormulas }),
      ]),
      this.group('Calculation', [
        this.large({ icon: 'calc', label: 'Calculate Now', tip: ['Calculate Now', 'Calculate the entire workbook now.', 'F9'], cmd: () => a.run('calcNow') }),
        this.large({
          icon: 'calc', label: 'Options', tip: ['Calculation Options', 'Specify when formulas are calculated.'],
          menu: () => [
            { label: 'Automatic', checked: a.wb.autoCalc, cmd: () => a.run('calcMode', 'auto') },
            { label: 'Manual', checked: !a.wb.autoCalc, cmd: () => a.run('calcMode', 'manual') },
          ],
        }),
      ]),
    ];
  }

  // ------------------------------------------------------------------ tab: Data
  panel_data(): HTMLElement[] {
    const a = this.app;
    return [
      this.group('Sort & Filter', [
        this.col(
          this.small({ icon: 'sort', label: 'Sort A to Z', left: true, tip: ['Sort A to Z', 'Sort the selection from lowest to highest.'], cmd: () => a.run('sort', 'asc') }),
          this.small({ icon: 'sort', label: 'Sort Z to A', left: true, tip: ['Sort Z to A', 'Sort the selection from highest to lowest.'], cmd: () => a.run('sort', 'desc') }),
        ),
        this.large({ icon: 'sort', label: 'Sort', tip: ['Sort', 'Open the Sort dialog box to sort by several columns at once.'], cmd: () => a.run('sortDialog') }),
        this.large({ icon: 'filter', label: 'Filter', tip: ['Filter', 'Turn on filtering for the selected cells.', 'Ctrl+Shift+L'], cmd: () => a.run('filter'), on: () => !!a.sheet.filter }),
        this.col(
          this.small({ icon: 'eraser', label: 'Clear', left: true, tip: ['Clear', 'Clear the filter and sort state for the current data range.'], cmd: () => a.run('filterClear') }),
          this.small({ icon: 'redo', label: 'Reapply', left: true, tip: ['Reapply', 'Reapply the filter in the current range.'], cmd: () => a.run('filterReapply') }),
        ),
      ]),
      this.group('Data Tools', [
        this.large({ icon: 'open', label: 'From CSV', tip: ['From CSV', 'Open a comma-separated values file.'], cmd: () => a.run('open') }),
        this.large({ icon: 'exportIcon', label: 'To CSV', tip: ['Save as CSV', 'Save the active sheet as a CSV file.'], cmd: () => a.run('exportCsv') }),
      ]),
    ];
  }

  // ------------------------------------------------------------------ tab: Review
  panel_review(): HTMLElement[] {
    const a = this.app;
    return [
      this.group('Workbook', [
        this.large({ icon: 'stats', label: 'Workbook Statistics', tip: ['Workbook Statistics', 'Show the number of sheets, cells and formulas in this workbook.'], cmd: () => a.run('stats') }),
      ]),
      this.group('Help', [
        this.large({ icon: 'info', label: 'About Sheetly', tip: ['About Sheetly', 'Version and license information.'], cmd: () => a.run('about') }),
      ]),
    ];
  }

  // ------------------------------------------------------------------ tab: View
  panel_view(): HTMLElement[] {
    const a = this.app;
    return [
      this.group('Workbook Views', [this.large({ icon: 'gridlines', label: 'Normal', tip: ['Normal', 'Normal worksheet view.'], on: () => true, cmd: () => {} })]),
      this.group('Show', [
        this.col(
          this.check('Formula Bar', () => a.showFormulaBar, () => a.run('formulaBar')),
          this.check('Gridlines', () => a.gridlines, () => a.run('gridlines')),
          this.check('Headings', () => a.headings, () => a.run('headings')),
        ),
      ]),
      this.group('Zoom', [
        this.large({ icon: 'zoom', label: 'Zoom', tip: ['Zoom', 'Open the Zoom dialog box to specify the zoom level.'], cmd: () => a.run('zoomDialog') }),
        this.large({ icon: 'zoom', label: '100%', tip: ['Zoom to 100%', 'Zoom the worksheet to 100% of the normal size.'], cmd: () => a.run('zoom', 100) }),
        this.large({ icon: 'zoom', label: 'Zoom to Selection', tip: ['Zoom to Selection', 'Zoom the worksheet so that the selected range fills the window.'], cmd: () => a.run('zoomSelection') }),
      ]),
      this.group('Window', [
        this.large({
          icon: 'freeze', label: 'Freeze Panes', tip: ['Freeze Panes', 'Keep rows and columns visible while the rest of the worksheet scrolls.'],
          menu: () => [
            { label: 'Freeze Panes', cmd: () => a.run('freeze', 'both') },
            { label: 'Freeze Top Row', cmd: () => a.run('freeze', 'top') },
            { label: 'Freeze First Column', cmd: () => a.run('freeze', 'col') },
            { sep: true },
            { label: 'Unfreeze Panes', cmd: () => a.run('freeze', 'none') },
          ],
        }),
      ]),
      this.group('Theme', [
        this.large({ icon: 'palette', label: 'Appearance', tip: ['Appearance', 'Change themes, colors, textures and animated backgrounds.'], cmd: () => a.run('appearance'), on: () => a.studio?.isOpen }),
        this.large({ icon: 'moon', label: 'Dark Mode', tip: ['Dark Mode', 'Switch between the light and dark color schemes.'], cmd: () => a.run('theme'), on: () => a.dark }),
      ]),
    ];
  }
}

export { FMT_CURRENCY };
