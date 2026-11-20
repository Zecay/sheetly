// Dialogs: Format Cells, Name Manager, Find & Replace, Sort, Conditional Formatting, filter panel and more.
import { MAX_COLS, MAX_ROWS, Rect, a1, colName, parseRect, rectText } from '../core/addr';
import { FUNCS } from '../core/eval';
import { Border, CFRule, CellStyle, StylePatch } from '../core/model';
import { formatValue } from '../core/numfmt';
import {
  FindOpts, SortKey, applyFilter, applyStyle, cellMatches, currentRegion, deleteCols, deleteRows, displayText, filterValueText, findAll, insertCols, insertRows, replaceInCell,
} from '../core/ops';
import { FErr } from '../core/values';
import { MenuItem, h, messageBox, openDialog, showColorMenu, showPopup, closeMenus, toast, shade } from './dom';
import { icon } from './icons';
import type { App } from './app';
import { detectHeader, rectAbs } from './app';
import { FUNCTION_CATEGORIES } from './ribbon';

const row = (label: string, ...ctl: (HTMLElement | string)[]) => h('div', { class: 'form-row' }, h('label', {}, label), ...ctl);

function select(options: [string, string][], value: string, onChange?: (v: string) => void): HTMLSelectElement {
  const s = h('select', {}, options.map(([v, l]) => h('option', { value: v, selected: v === value }, l))) as HTMLSelectElement;
  s.value = value;
  if (onChange) s.addEventListener('change', () => onChange(s.value));
  return s;
}

function tabs(names: [string, string][], panes: Record<string, HTMLElement>, initial: string): HTMLElement {
  const bar = h('div', { class: 'dtabs' });
  const body = h('div', { class: 'dpane' });
  const show = (id: string) => {
    body.innerHTML = '';
    body.append(panes[id]);
    Array.from(bar.children).forEach((b, i) => b.classList.toggle('on', names[i][0] === id));
  };
  for (const [id, label] of names) {
    const b = h('button', { type: 'button' }, label);
    b.addEventListener('click', () => show(id));
    bar.append(b);
  }
  show(initial);
  return h('div', {}, bar, body);
}

// ======================================================================== Format Cells

const DATE_TYPES: [string, string][] = [
  ['m/d/yyyy', 'm/d/yyyy'], ['dddd, mmmm d, yyyy', 'dddd, mmmm d, yyyy'], ['mmmm d, yyyy', 'mmmm d, yyyy'], ['d-mmm-yy', 'd-mmm-yy'], ['d-mmm', 'd-mmm'], ['mmm-yy', 'mmm-yy'],
  ['yyyy-mm-dd', 'yyyy-mm-dd'], ['dd/mm/yyyy', 'dd/mm/yyyy'], ['mm/dd/yy', 'mm/dd/yy'], ['m/d/yyyy h:mm', 'm/d/yyyy h:mm'], ['mmmm', 'mmmm'], ['dddd', 'dddd'],
];
const TIME_TYPES: [string, string][] = [
  ['h:mm AM/PM', 'h:mm AM/PM'], ['h:mm:ss AM/PM', 'h:mm:ss AM/PM'], ['h:mm', 'h:mm'], ['h:mm:ss', 'h:mm:ss'], ['[h]:mm:ss', '[h]:mm:ss'], ['mm:ss', 'mm:ss'],
];
const CURRENCIES: [string, string][] = [['$', '$'], ['€', '€'], ['£', '£'], ['¥', '¥'], ['', 'None']];

function categoryOf(fmt: string): string {
  if (!fmt || fmt === 'General') return 'General';
  if (fmt === '@') return 'Text';
  if (/E[+-]0/.test(fmt)) return 'Scientific';
  if (/\?\/\?/.test(fmt)) return 'Fraction';
  if (/_\(/.test(fmt) && /[$€£¥]/.test(fmt)) return 'Accounting';
  const bare = fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  if (/[$€£¥]/.test(fmt)) return 'Currency';
  if (/%/.test(bare)) return 'Percentage';
  if (/[ymd]/i.test(bare) && /[hs]/i.test(bare.replace(/[ymd]+/gi, ''))) return /[ymd]/i.test(bare.replace(/[hms:\s]|AM\/PM|\[h\]/gi, '')) ? 'Date' : 'Time';
  if (/[ymd]/i.test(bare)) return 'Date';
  if (/[hs]/i.test(bare)) return 'Time';
  if (/^[#,]*0(\.0+)?$/.test(fmt)) return 'Number';
  return 'Custom';
}

function numberCode(dec: number, sep: boolean): string {
  return (sep ? '#,##0' : '0') + (dec > 0 ? '.' + '0'.repeat(dec) : '');
}

function negVariants(base: string, red = true): [string, string][] {
  return [
    [base, '-1234.10'],
    ...(red ? ([[`${base};[Red]${base}`, '-1234.10 (red)']] as [string, string][]) : []),
    [`${base};(${base})`, '(1234.10)'],
    ...(red ? ([[`${base};[Red](${base})`, '(1234.10) (red)']] as [string, string][]) : []),
  ];
}

export function openFormatCells(app: App, startTab: string): void {
  const wb = app.wb;
  const st: CellStyle = { ...app.activeStyle() };
  const patch: StylePatch = {};
  const set = <K extends keyof CellStyle>(k: K, v: CellStyle[K] | null) => {
    (patch as any)[k] = v;
    if (v === null) delete (st as any)[k];
    else (st as any)[k] = v;
    updatePreview();
  };
  const { r, c } = app.activeCell;
  const sample = (() => {
    const v = app.sheet.get(r, c)?.v;
    return v === null || v === undefined ? 1234.5 : v;
  })();

  const preview = h('div', { class: 'preview-box' });
  const updatePreview = () => {
    const f = formatValue(typeof sample === 'object' ? 1234.5 : sample, st.fmt ?? 'General').text.split('').join(' ');
    preview.textContent = f || 'AaBbCcYyZz';
    Object.assign(preview.style, {
      fontFamily: st.font ? `"${st.font}", Calibri, sans-serif` : 'Calibri, Carlito, sans-serif',
      fontSize: (st.size ?? 11) * (4 / 3) + 'px',
      fontWeight: st.bold ? 'bold' : 'normal',
      fontStyle: st.italic ? 'italic' : 'normal',
      textDecoration: [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].join(' ') || 'none',
      color: st.color ?? '',
      background: st.fill ?? '',
    });
  };

  // ---------- Number ----------
  const numberPane = (() => {
    const cats = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Custom'];
    let cat = categoryOf(st.fmt ?? 'General');
    const list = h('div', { class: 'listbox', style: { height: '200px' } });
    const opts = h('div', {});
    const sampleBox = h('fieldset', {}, h('legend', {}, 'Sample'), h('div', { class: 'sample', style: { minHeight: '20px', fontFamily: 'Calibri, sans-serif' } }));
    const refresh = () => {
      sampleBox.querySelector('.sample')!.textContent = formatValue(typeof sample === 'object' ? 1234.5 : sample, st.fmt ?? 'General').text.split('').join(' ');
    };
    const dec0 = (() => {
      const m = /\.(0+)/.exec(st.fmt ?? '');
      return m ? m[1].length : 2;
    })();
    let dec = cat === 'Percentage' || cat === 'Scientific' || cat === 'Number' || cat === 'Currency' || cat === 'Accounting' ? dec0 : 2;
    let sep = /#,##0/.test(st.fmt ?? '');
    let sym = '$';
    for (const [s] of CURRENCIES) if (s && (st.fmt ?? '').includes(s)) sym = s;
    let neg = 0;
    const apply = (code: string) => {
      set('fmt', code === 'General' ? null : code);
      refresh();
    };
    const build = () => {
      opts.innerHTML = '';
      const decInput = h('input', { type: 'number', min: '0', max: '30', value: String(dec), style: { width: '60px' } }) as HTMLInputElement;
      const sepBox = h('input', { type: 'checkbox', checked: sep }) as HTMLInputElement;
      const symSel = select(CURRENCIES, sym);
      const decRow = row('Decimal places:', decInput);
      switch (cat) {
        case 'General':
          opts.append(h('div', { style: { color: 'var(--muted)' } }, 'General format cells have no specific number format.'));
          apply('General');
          break;
        case 'Number': {
          const negSel = select(negVariants(numberCode(dec, sep)).map(([v, l], i) => [String(i), l] as [string, string]), String(neg));
          const upd = () => {
            dec = Math.max(0, Math.min(30, parseInt(decInput.value, 10) || 0));
            sep = sepBox.checked;
            neg = parseInt(negSel.value, 10);
            apply(negVariants(numberCode(dec, sep))[neg][0]);
          };
          [decInput, sepBox, negSel].forEach((e) => e.addEventListener('input', upd));
          opts.append(decRow, h('div', { class: 'form-row' }, sepBox, h('label', { style: { minWidth: '0' } }, 'Use 1000 Separator (,)')), row('Negative numbers:', negSel));
          upd();
          break;
        }
        case 'Currency':
        case 'Accounting': {
          const negSel = select(negVariants('0').map(([, l], i) => [String(i), l] as [string, string]), String(neg));
          const upd = () => {
            dec = Math.max(0, Math.min(30, parseInt(decInput.value, 10) || 0));
            sym = symSel.value;
            neg = parseInt(negSel.value, 10);
            const num = '#,##0' + (dec ? '.' + '0'.repeat(dec) : '');
            const q = sym ? `"${sym}"` : '';
            if (cat === 'Currency') apply(negVariants(`${q}${num}`)[neg][0]);
            else apply(`_(${q}* ${num}_);_(${q}* (${num});_(${q}* "-"${'?'.repeat(Math.max(1, dec))}_);_(@_)`);
          };
          [decInput, symSel, negSel].forEach((e) => e.addEventListener('input', upd));
          opts.append(decRow, row('Symbol:', symSel));
          if (cat === 'Currency') opts.append(row('Negative numbers:', negSel));
          upd();
          break;
        }
        case 'Percentage':
        case 'Scientific': {
          const upd = () => {
            dec = Math.max(0, Math.min(30, parseInt(decInput.value, 10) || 0));
            apply((cat === 'Percentage' ? '0' : '0') + (dec ? '.' + '0'.repeat(dec) : '') + (cat === 'Percentage' ? '%' : 'E+00'));
          };
          decInput.addEventListener('input', upd);
          opts.append(decRow);
          upd();
          break;
        }
        case 'Date':
        case 'Time':
        case 'Fraction': {
          const types: [string, string][] = cat === 'Date' ? DATE_TYPES : cat === 'Time' ? TIME_TYPES : [['# ?/?', 'Up to one digit (1/4)'], ['# ??/??', 'Up to two digits (21/35)'], ['# ???/???', 'Up to three digits (312/943)'], ['?/?', 'As halves/quarters (1/4)']];
          const box = h('div', { class: 'listbox' });
          const cur = st.fmt ?? types[0][0];
          for (const [code] of types) {
            const label = cat === 'Fraction' ? types.find((t) => t[0] === code)![1] : formatValue(typeof sample === 'number' ? sample : 45000.75, code).text;
            const item = h('div', { class: 'li' + (code === cur ? ' sel' : '') }, label);
            item.addEventListener('click', () => {
              box.querySelectorAll('.li').forEach((e) => e.classList.remove('sel'));
              item.classList.add('sel');
              apply(code);
            });
            box.append(item);
          }
          opts.append(h('label', { class: 'lbl' }, 'Type:'), box);
          if (!types.some((t) => t[0] === st.fmt)) apply(types[0][0]);
          break;
        }
        case 'Text':
          opts.append(h('div', { style: { color: 'var(--muted)' } }, 'Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.'));
          apply('@');
          break;
        case 'Custom': {
          const inp = h('input', { type: 'text', class: 'txt', style: { width: '100%' }, value: st.fmt ?? 'General' }) as HTMLInputElement;
          const box = h('div', { class: 'listbox', style: { height: '120px', marginTop: '6px' } });
          const common = ['General', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '0.00E+00', '# ?/?', 'm/d/yyyy', 'd-mmm-yy', 'h:mm AM/PM', 'h:mm:ss', 'm/d/yyyy h:mm', '"$"#,##0.00', '#,##0;[Red]-#,##0', '@'];
          for (const code of common) {
            const item = h('div', { class: 'li' }, code);
            item.addEventListener('click', () => {
              inp.value = code;
              apply(code);
            });
            box.append(item);
          }
          inp.addEventListener('input', () => apply(inp.value || 'General'));
          opts.append(h('label', { class: 'lbl' }, 'Type:'), inp, box);
          break;
        }
      }
      refresh();
    };
    cats.forEach((name) => {
      const item = h('div', { class: 'li' + (name === cat ? ' sel' : '') }, name);
      item.addEventListener('click', () => {
        cat = name;
        list.querySelectorAll('.li').forEach((e) => e.classList.remove('sel'));
        item.classList.add('sel');
        build();
      });
      list.append(item);
    });
    // initial build must not overwrite an existing custom format
    const initialFmt = st.fmt;
    build();
    if (initialFmt === undefined) delete (patch as any).fmt;
    else if (initialFmt === (st.fmt ?? undefined)) delete (patch as any).fmt;
    refresh();
    return h('div', { class: 'cols2', style: { gridTemplateColumns: '150px 1fr' } }, h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Category:'), list), h('div', {}, sampleBox, opts));
  })();

  // ---------- Alignment ----------
  const selRect = app.grid.expandForMerges(app.grid.range);
  const mergedInit = app.grid.ranges.length === 1 && app.sheet.merges.some((m) => m.r1 === selRect.r1 && m.r2 === selRect.r2 && m.c1 === selRect.c1 && m.c2 === selRect.c2);
  let mergeWanted = mergedInit;
  const alignPane = (() => {
    const hSel = select([['', 'General'], ['left', 'Left (Indent)'], ['center', 'Center'], ['right', 'Right (Indent)'], ['justify', 'Justify']], st.h ?? '', (v) => set('h', (v || null) as any));
    const vSel = select([['top', 'Top'], ['middle', 'Center'], ['bottom', 'Bottom']], st.v ?? 'bottom', (v) => set('v', v === 'bottom' ? null : (v as any)));
    const indent = h('input', { type: 'number', min: '0', max: '15', value: String(st.indent ?? 0), style: { width: '60px' } }) as HTMLInputElement;
    indent.addEventListener('input', () => set('indent', parseInt(indent.value, 10) || null));
    const wrap = h('input', { type: 'checkbox', checked: !!st.wrap }) as HTMLInputElement;
    wrap.addEventListener('change', () => set('wrap', wrap.checked ? true : null));
    const mergeBox = h('input', { type: 'checkbox', checked: mergedInit }) as HTMLInputElement;
    mergeBox.addEventListener('change', () => (mergeWanted = mergeBox.checked));
    return h('div', {}, h('fieldset', {}, h('legend', {}, 'Text alignment'), row('Horizontal:', hSel), row('Indent:', indent), row('Vertical:', vSel)), h('fieldset', {}, h('legend', {}, 'Text control'), h('div', { class: 'form-row' }, wrap, h('label', { style: { minWidth: '0' } }, 'Wrap text')), h('div', { class: 'form-row' }, mergeBox, h('label', { style: { minWidth: '0' } }, 'Merge cells'))));
  })();

  // ---------- Font ----------
  const fontPane = (() => {
    const fonts = ['Calibri', 'Arial', 'Arial Black', 'Cambria', 'Candara', 'Comic Sans MS', 'Consolas', 'Courier New', 'Georgia', 'Impact', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
    const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
    const styles = ['Regular', 'Italic', 'Bold', 'Bold Italic'];
    const mkList = (items: (string | number)[], cur: string, fn: (v: string) => void) => {
      const box = h('div', { class: 'listbox', style: { height: '140px' } });
      for (const it of items) {
        const el = h('div', { class: 'li' + (String(it) === cur ? ' sel' : '') }, String(it));
        el.addEventListener('click', () => {
          box.querySelectorAll('.li').forEach((e) => e.classList.remove('sel'));
          el.classList.add('sel');
          fn(String(it));
        });
        box.append(el);
      }
      return box;
    };
    const curStyle = st.bold && st.italic ? 'Bold Italic' : st.bold ? 'Bold' : st.italic ? 'Italic' : 'Regular';
    const color = h('input', { type: 'color', value: st.color ?? '#000000' }) as HTMLInputElement;
    color.addEventListener('input', () => set('color', color.value));
    const auto = h('button', { class: 'btn small', type: 'button' }, 'Automatic');
    auto.addEventListener('click', () => set('color', null));
    const ul = h('input', { type: 'checkbox', checked: !!st.underline }) as HTMLInputElement;
    ul.addEventListener('change', () => set('underline', ul.checked ? true : null));
    const strike = h('input', { type: 'checkbox', checked: !!st.strike }) as HTMLInputElement;
    strike.addEventListener('change', () => set('strike', strike.checked ? true : null));
    return h('div', {},
      h('div', { class: 'cols2', style: { gridTemplateColumns: '1.3fr 1fr 0.7fr' } },
        h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Font:'), mkList(fonts, st.font ?? 'Calibri', (v) => set('font', v === 'Calibri' ? null : v))),
        h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Font style:'), mkList(styles, curStyle, (v) => { set('bold', v.includes('Bold') ? true : null); set('italic', v.includes('Italic') ? true : null); })),
        h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Size:'), mkList(sizes, String(st.size ?? 11), (v) => set('size', parseInt(v, 10) === 11 ? null : parseInt(v, 10)))),
      ),
      h('div', { class: 'form-row', style: { marginTop: '10px' } }, h('label', {}, 'Color:'), color, auto),
      h('div', { class: 'form-row' }, ul, h('label', { style: { minWidth: '90px' } }, 'Underline'), strike, h('label', {}, 'Strikethrough')),
      h('fieldset', { style: { marginTop: '8px' } }, h('legend', {}, 'Preview'), preview),
    );
  })();

  // ---------- Border ----------
  const bd = { top: !!st.bt, bottom: !!st.bb, left: !!st.bl, right: !!st.br, insideH: false, insideV: false };
  let bStyle: Border['style'] = st.bt?.style ?? st.bb?.style ?? st.bl?.style ?? st.br?.style ?? 'thin';
  let bColor = st.bt?.color ?? st.bb?.color ?? st.bl?.color ?? st.br?.color ?? '#000000';
  let bTouched = false;
  const borderPane = (() => {
    const box = h('div', { class: 'border-pad' });
    const mkEdge = (key: keyof typeof bd, label: string, style: Record<string, string>) => {
      const b = h('button', { type: 'button', class: 'btn small', style: { minWidth: '0', ...style } }, label);
      const refresh = () => (b.style.background = bd[key] ? 'var(--pressed)' : '');
      b.addEventListener('click', () => {
        bd[key] = !bd[key];
        bTouched = true;
        refresh();
      });
      refresh();
      return { b, refresh };
    };
    const top = mkEdge('top', 'Top', { gridColumn: '2', gridRow: '1' });
    const left = mkEdge('left', 'L', { gridColumn: '1', gridRow: '2' });
    const right = mkEdge('right', 'R', { gridColumn: '3', gridRow: '2' });
    const bottom = mkEdge('bottom', 'Bottom', { gridColumn: '2', gridRow: '3' });
    const mid = h('div', { class: 'cellbox', style: { gridColumn: '2', gridRow: '2' } });
    box.append(top.b, left.b, mid, right.b, bottom.b);
    const edges = [top, left, right, bottom];
    const presets = h('div', { class: 'chip-row' });
    const pre = (label: string, fn: () => void) => {
      const b = h('button', { class: 'btn small', type: 'button' }, label);
      b.addEventListener('click', () => {
        fn();
        bTouched = true;
        edges.forEach((e) => e.refresh());
      });
      presets.append(b);
    };
    pre('None', () => Object.assign(bd, { top: false, bottom: false, left: false, right: false, insideH: false, insideV: false }));
    pre('Outline', () => Object.assign(bd, { top: true, bottom: true, left: true, right: true }));
    pre('Inside', () => Object.assign(bd, { insideH: true, insideV: true }));
    const styleList = h('div', { class: 'listbox', style: { height: '150px' } });
    const styles: Border['style'][] = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair'];
    for (const s of styles) {
      const el = h('div', { class: 'li' + (s === bStyle ? ' sel' : '') }, s[0].toUpperCase() + s.slice(1));
      el.addEventListener('click', () => {
        styleList.querySelectorAll('.li').forEach((e) => e.classList.remove('sel'));
        el.classList.add('sel');
        bStyle = s;
        bTouched = true;
      });
      styleList.append(el);
    }
    const col = h('input', { type: 'color', value: bColor }) as HTMLInputElement;
    col.addEventListener('input', () => {
      bColor = col.value;
      bTouched = true;
    });
    return h('div', { class: 'cols2' },
      h('div', {}, h('fieldset', {}, h('legend', {}, 'Line'), h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Style:'), styleList, row('Color:', col))),
      h('div', {}, h('fieldset', {}, h('legend', {}, 'Presets'), presets), h('fieldset', {}, h('legend', {}, 'Border'), box, h('div', { style: { color: 'var(--muted)', textAlign: 'center' } }, 'Click the buttons to toggle each border.'))),
    );
  })();

  // ---------- Fill ----------
  const fillPane = (() => {
    const grid = h('div', { class: 'swatch-grid', style: { gridTemplateColumns: 'repeat(10, 24px)' } });
    const base = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
    const tints = [0, 0.8, 0.6, 0.4, -0.25, -0.5];
    for (const t of tints)
      for (const b of base) {
        const col = t === 0 ? b : shade(b, t);
        const sw = h('button', { class: 'swatch', type: 'button', style: { background: col, width: '24px', height: '24px' } });
        sw.addEventListener('click', () => {
          set('fill', col);
          custom.value = col;
        });
        grid.append(sw);
      }
    const none = h('button', { class: 'btn', type: 'button' }, 'No Color');
    none.addEventListener('click', () => set('fill', null));
    const custom = h('input', { type: 'color', value: st.fill ?? '#ffffff' }) as HTMLInputElement;
    custom.addEventListener('input', () => set('fill', custom.value));
    return h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Background Color:'), none, h('div', { style: { height: '8px' } }), grid, row('More colors:', custom), h('fieldset', { style: { marginTop: '10px' } }, h('legend', {}, 'Sample'), preview.cloneNode(false) as HTMLElement));
  })();
  void fillPane;

  const names: [string, string][] = [['number', 'Number'], ['alignment', 'Alignment'], ['font', 'Font'], ['border', 'Border'], ['fill', 'Fill']];
  const body = tabs(names, { number: numberPane, alignment: alignPane, font: fontPane, border: borderPane, fill: fillPane }, startTab);
  updatePreview();
  openDialog({
    title: 'Format Cells',
    content: body,
    width: 560,
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          const rects = app.grid.ranges.map((rg) => app.grid.expandForMerges(rg));
          const mergeKind = mergeWanted === mergedInit ? null : mergeWanted ? 'merge' : 'unmerge';
          // respect the merge size cap before changing anything
          if (mergeKind === 'merge' && app.tooLarge(app.grid.ranges, 100_000, 'Merging')) return;
          wb.transact('Format Cells', () => {
            if (Object.keys(patch).length) for (const rg of rects) applyStyle(wb, app.sheet, rg, patch);
            if (bTouched) applyBordersDialog(app, rects, bd, bStyle, bColor);
            if (mergeKind) app.mergeCells(mergeKind); // nested: stays one undo step
          });
          if (patch.wrap === true) app.grid.autofitRows(Array.from(new Set(rects.flatMap((rg) => Array.from({ length: Math.min(200, rg.r2 - rg.r1 + 1) }, (_, i) => rg.r1 + i)))), false);
          app.refreshAll();
        },
      },
      { label: 'Cancel' },
    ],
  });
}

function applyBordersDialog(app: App, rects: Rect[], bd: { top: boolean; bottom: boolean; left: boolean; right: boolean; insideH: boolean; insideV: boolean }, style: Border['style'], color: string): void {
  const b: Border = { style, color };
  const wb = app.wb;
  let total = 0;
  for (const rg of rects) total += (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
  if (total > 500_000) {
    void messageBox('Sheetly', 'Borders are limited to 500,000 cells at a time. Select a smaller range and try again.');
    return;
  }
  for (const rg of rects) {
    const R = rg.r2;
    const C = rg.c2;
    for (let r = rg.r1; r <= R; r++)
      for (let c = rg.c1; c <= C; c++) {
        const p: StylePatch = {};
        const top = r === rg.r1 ? bd.top : bd.insideH;
        const bottom = r === rg.r2 ? bd.bottom : bd.insideH;
        const left = c === rg.c1 ? bd.left : bd.insideV;
        const right = c === rg.c2 ? bd.right : bd.insideV;
        p.bt = top ? b : null;
        p.bb = bottom ? b : null;
        p.bl = left ? b : null;
        p.br = right ? b : null;
        applyStyle(wb, app.sheet, { r1: r, c1: c, r2: r, c2: c }, p);
      }
  }
}

// ======================================================================== names

export function openNameManager(app: App): void {
  const wb = app.wb;
  let selected: string | null = null;
  const tbody = h('tbody');
  const refInput = h('input', { type: 'text', class: 'txt', style: { width: '100%' } }) as HTMLInputElement;
  const render = () => {
    tbody.innerHTML = '';
    const names = Array.from(wb.names.values()).sort((a, b) => a.name.localeCompare(b.name));
    if (!names.length) tbody.append(h('tr', {}, h('td', { colspan: '2', style: { color: 'var(--muted)' } }, 'No names defined yet.')));
    for (const n of names) {
      const tr = h('tr', { class: selected === n.name.toLowerCase() ? 'sel' : '' }, h('td', {}, n.name), h('td', {}, n.refers));
      tr.addEventListener('click', () => {
        selected = n.name.toLowerCase();
        refInput.value = n.refers;
        render();
      });
      tbody.append(tr);
    }
  };
  const table = h('div', { class: 'listbox', style: { height: '220px', width: '100%' } }, h('table', { class: 'table-list' }, h('thead', {}, h('tr', {}, h('th', {}, 'Name'), h('th', {}, 'Refers To'))), tbody));
  const apply = () => {
    if (!selected) return;
    const def = wb.getName(selected);
    if (!def) return;
    wb.transact('Edit name', () => wb.setName(def.name, refInput.value.replace(/^=/, '')), true);
    wb.rebuildDeps();
    wb.requestFullRecalc();
    render();
  };
  const btnNew = h('button', { class: 'btn', type: 'button' }, 'New…');
  btnNew.addEventListener('click', () => openDefineName(app, () => render()));
  const btnDel = h('button', { class: 'btn', type: 'button' }, 'Delete');
  btnDel.addEventListener('click', () => {
    if (!selected) return;
    const def = wb.getName(selected);
    if (!def) return;
    wb.transact('Delete name', () => wb.deleteName(def.name), true);
    wb.rebuildDeps();
    wb.requestFullRecalc();
    selected = null;
    refInput.value = '';
    render();
  });
  const btnApply = h('button', { class: 'btn', type: 'button' }, 'Apply');
  btnApply.addEventListener('click', apply);
  const pick = h('button', { class: 'btn small', type: 'button', title: 'Use current selection' }, 'Use selection');
  pick.addEventListener('click', () => {
    const q = /^[A-Za-z_][\w.]*$/.test(app.sheet.name) ? app.sheet.name : `'${app.sheet.name.replace(/'/g, "''")}'`;
    refInput.value = `${q}!${rectAbs(app.grid.range)}`;
  });
  render();
  openDialog({
    title: 'Name Manager',
    width: 520,
    content: h('div', {}, h('div', { style: { display: 'flex', gap: '6px', marginBottom: '8px' } }, btnNew, btnDel), table, h('label', { class: 'lbl' }, 'Refers to:'), h('div', { style: { display: 'flex', gap: '6px' } }, refInput, pick, btnApply)),
    buttons: [{ label: 'Close', primary: true }],
  });
}

export function openDefineName(app: App, onDone?: () => void): void {
  const wb = app.wb;
  const name = h('input', { type: 'text', class: 'txt', style: { width: '100%' } }) as HTMLInputElement;
  const q = /^[A-Za-z_][\w.]*$/.test(app.sheet.name) ? app.sheet.name : `'${app.sheet.name.replace(/'/g, "''")}'`;
  const refers = h('input', { type: 'text', class: 'txt', style: { width: '100%' }, value: `${q}!${rectAbs(app.grid.range)}` }) as HTMLInputElement;
  const err = h('div', { class: 'err' });
  openDialog({
    title: 'New Name',
    width: 380,
    content: h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Name:'), name, h('label', { class: 'lbl' }, 'Refers to:'), refers, err),
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          const n = name.value.trim();
          if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(n)) {
            err.textContent = "The name that you entered is not valid.";
            return false;
          }
          if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[Rr]\d*[Cc]\d*$/.test(n)) {
            err.textContent = 'The name cannot look like a cell reference.';
            return false;
          }
          if (wb.getName(n)) {
            err.textContent = 'That name already exists.';
            return false;
          }
          wb.transact('Define name', () => wb.setName(n, refers.value.replace(/^=/, '')), true);
          wb.rebuildDeps();
          wb.requestFullRecalc();
          onDone?.();
          app.refreshAll();
        },
      },
      { label: 'Cancel' },
    ],
  });
}

// ======================================================================== find & replace

export function openFindReplace(app: App, withReplace: boolean): void {
  const fs = app.findState;
  const what = h('input', { type: 'text', class: 'txt', style: { width: '100%' }, value: fs.what }) as HTMLInputElement;
  const repl = h('input', { type: 'text', class: 'txt', style: { width: '100%' }, value: fs.replace }) as HTMLInputElement;
  const cb = (label: string, key: 'matchCase' | 'wholeCell') => {
    const i = h('input', { type: 'checkbox', checked: fs[key] }) as HTMLInputElement;
    i.addEventListener('change', () => (fs[key] = i.checked));
    return h('label', { class: 'rb-check', style: { padding: '0' } }, i, label);
  };
  const lookIn = select([['values', 'Values'], ['formulas', 'Formulas']], fs.inFormulas ? 'formulas' : 'values', (v) => (fs.inFormulas = v === 'formulas'));
  const within = select([['sheet', 'Sheet'], ['book', 'Workbook']], fs.allSheets ? 'book' : 'sheet', (v) => (fs.allSheets = v === 'book'));
  const order = select([['rows', 'By Rows'], ['cols', 'By Columns']], fs.byRows ? 'rows' : 'cols', (v) => (fs.byRows = v === 'rows'));
  const results = h('div', { class: 'listbox', style: { height: '120px', marginTop: '8px', display: 'none' } });
  const status = h('div', { style: { color: 'var(--muted)', minHeight: '16px', marginTop: '6px' } });
  const opts = (): FindOpts => ({ what: what.value, matchCase: fs.matchCase, wholeCell: fs.wholeCell, inFormulas: fs.inFormulas, allSheets: fs.allSheets, byRows: fs.byRows });
  const sync = () => {
    fs.what = what.value;
    fs.replace = repl.value;
  };
  const goTo = (f: { sheet: any; r: number; c: number }) => {
    if (f.sheet !== app.sheet) app.setActiveSheet(f.sheet);
    app.grid.select(f.r, f.c);
    app.grid.scrollIntoView(f.r, f.c);
  };
  const findNext = (): boolean => {
    sync();
    if (!what.value) return false;
    const all = findAll(app.wb, app.sheet, opts());
    if (!all.length) {
      status.textContent = "We couldn't find what you were searching for.";
      return false;
    }
    const cur = app.grid.active;
    const key = (f: { sheet: any; r: number; c: number }) => app.wb.sheets.indexOf(f.sheet) * 1e12 + (fs.byRows ? f.r * 20000 + f.c : f.c * 2000000 + f.r);
    const curKey = app.wb.sheets.indexOf(app.sheet) * 1e12 + (fs.byRows ? cur.r * 20000 + cur.c : cur.c * 2000000 + cur.r);
    const next = all.find((f) => key(f) > curKey) ?? all[0];
    goTo(next);
    status.textContent = '';
    return true;
  };
  const doReplace = () => {
    sync();
    const o = opts();
    const { r, c } = app.grid.active;
    const cell = app.sheet.get(r, c);
    if (cell && cellMatches(app.wb, cell, o)) {
      app.wb.transact('Replace', () => replaceInCell(app.wb, app.sheet, r, c, o, repl.value));
    }
    findNext();
  };
  const replaceAll = () => {
    sync();
    const o = opts();
    if (!o.what) return;
    const all = findAll(app.wb, app.sheet, o);
    let n = 0;
    app.wb.transact('Replace All', () => {
      for (const f of all) if (replaceInCell(app.wb, f.sheet, f.r, f.c, o, repl.value)) n++;
    });
    status.textContent = `${n} replacement${n === 1 ? '' : 's'} made.`;
  };
  const findAllClick = () => {
    sync();
    if (!what.value) return;
    const all = findAll(app.wb, app.sheet, opts());
    results.style.display = 'block';
    results.innerHTML = '';
    status.textContent = `${all.length} cell${all.length === 1 ? '' : 's'} found`;
    for (const f of all.slice(0, 2000)) {
      const cell = f.sheet.get(f.r, f.c);
      const li = h('div', { class: 'li' }, `${f.sheet.name}!${a1(f.r, f.c)}   ${displayText(app.wb, cell)}`);
      li.addEventListener('click', () => goTo(f));
      results.append(li);
    }
  };
  const btn = (label: string, fn: () => void, primary = false) => {
    const b = h('button', { class: 'btn' + (primary ? ' primary' : ''), type: 'button' }, label);
    b.addEventListener('click', fn);
    return b;
  };
  const body = h('div', {},
    h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Find what:'), what,
    withReplace ? h('label', { class: 'lbl' }, 'Replace with:') : null, withReplace ? repl : null,
    h('div', { class: 'cols2', style: { marginTop: '10px' } }, h('div', {}, cb('Match case', 'matchCase'), cb('Match entire cell contents', 'wholeCell')), h('div', {}, row('Within:', within), row('Search:', order), row('Look in:', lookIn))),
    h('div', { style: { display: 'flex', gap: '6px', flexWrap: 'wrap', marginTop: '10px' } },
      withReplace ? btn('Replace All', replaceAll) : null, withReplace ? btn('Replace', doReplace) : null, btn('Find All', findAllClick), btn('Find Next', findNext, true)),
    status, results);
  what.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      findNext();
    }
  });
  openDialog({ title: withReplace ? 'Find and Replace' : 'Find', content: body, width: 420, modeless: true, buttons: [{ label: 'Close' }] });
  void messageBox;
}

// ======================================================================== go to, zoom, paste special, insert/delete

export function openGoTo(app: App): void {
  const inp = h('input', { type: 'text', class: 'txt', style: { width: '100%' } }) as HTMLInputElement;
  const list = h('div', { class: 'listbox', style: { height: '120px' } });
  for (const n of app.wb.names.values()) {
    const li = h('div', { class: 'li' }, n.name);
    li.addEventListener('click', () => (inp.value = n.name));
    list.append(li);
  }
  openDialog({
    title: 'Go To',
    width: 340,
    content: h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Reference:'), inp, h('label', { class: 'lbl' }, 'Go to:'), list),
    buttons: [{ label: 'OK', primary: true, onClick: () => app.goToText(inp.value) }, { label: 'Cancel' }],
  });
}

export function openZoom(app: App): void {
  const cur = Math.round(app.grid.zoom * 100);
  const opts = [200, 100, 75, 50, 25];
  let value = cur;
  const custom = h('input', { type: 'number', min: '10', max: '400', value: String(cur), style: { width: '70px' } }) as HTMLInputElement;
  const radios = opts.map((o) => {
    const r = h('input', { type: 'radio', name: 'zoom', checked: o === cur }) as HTMLInputElement;
    r.addEventListener('change', () => (value = o));
    return h('label', { class: 'rb-check' }, r, o + '%');
  });
  const fit = h('input', { type: 'radio', name: 'zoom' }) as HTMLInputElement;
  const cust = h('input', { type: 'radio', name: 'zoom', checked: !opts.includes(cur) }) as HTMLInputElement;
  custom.addEventListener('input', () => {
    cust.checked = true;
    value = parseInt(custom.value, 10) || 100;
  });
  let fitSel = false;
  fit.addEventListener('change', () => (fitSel = true));
  openDialog({
    title: 'Zoom',
    width: 260,
    content: h('div', {}, h('fieldset', {}, h('legend', {}, 'Magnification'), ...radios, h('label', { class: 'rb-check' }, fit, 'Fit selection'), h('label', { class: 'rb-check' }, cust, 'Custom:', custom, '%'))),
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          if (fitSel) void app.run('zoomSelection');
          else app.grid.setZoom(Math.max(10, Math.min(400, value)) / 100);
        },
      },
      { label: 'Cancel' },
    ],
  });
}

export function openPasteSpecial(app: App): void {
  let mode = 'all';
  const items: [string, string][] = [['all', 'All'], ['formulas', 'Formulas'], ['values', 'Values'], ['formats', 'Formats'], ['transpose', 'Transpose']];
  const rows = items.map(([id, label]) => {
    const r = h('input', { type: 'radio', name: 'ps', checked: id === 'all' }) as HTMLInputElement;
    r.addEventListener('change', () => (mode = id));
    return h('label', { class: 'rb-check' }, r, label);
  });
  openDialog({
    title: 'Paste Special',
    width: 260,
    content: h('div', {}, h('fieldset', {}, h('legend', {}, 'Paste'), ...rows)),
    buttons: [{ label: 'OK', primary: true, onClick: () => void app.run('paste', mode) }, { label: 'Cancel' }],
  });
}

export function openInsertDelete(app: App, insert: boolean): void {
  let mode = insert ? 'down' : 'up';
  const items: [string, string][] = insert
    ? [['right', 'Shift cells right'], ['down', 'Shift cells down'], ['row', 'Entire row'], ['col', 'Entire column']]
    : [['left', 'Shift cells left'], ['up', 'Shift cells up'], ['row', 'Entire row'], ['col', 'Entire column']];
  const rows = items.map(([id, label]) => {
    const r = h('input', { type: 'radio', name: 'id', checked: id === mode }) as HTMLInputElement;
    r.addEventListener('change', () => (mode = id));
    return h('label', { class: 'rb-check' }, r, label);
  });
  openDialog({
    title: insert ? 'Insert' : 'Delete',
    width: 250,
    content: h('div', {}, ...rows),
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          if (mode === 'row') void app.run(insert ? 'insertRows' : 'deleteRows');
          else if (mode === 'col') void app.run(insert ? 'insertCols' : 'deleteCols');
          else void app.run('shiftCells', insert ? mode : mode === 'left' ? 'left' : 'up');
        },
      },
      { label: 'Cancel' },
    ],
  });
}

export function openSizeDialog(app: App, kind: 'row' | 'col' | 'defcol'): void {
  const sheet = app.sheet;
  const rg = app.grid.range;
  const isRow = kind === 'row';
  const curPx = isRow ? sheet.rows.size(app.grid.active.r) : sheet.cols.size(app.grid.active.c);
  const cur = isRow ? curPx * 0.75 : Math.max(0, (curPx - 5) / 7);
  const inp = h('input', { type: 'number', step: '0.01', min: '0', value: String(Math.round(cur * 100) / 100), class: 'txt', style: { width: '100px' } }) as HTMLInputElement;
  openDialog({
    title: isRow ? 'Row Height' : 'Column Width',
    width: 280,
    content: h('div', {}, row(isRow ? 'Row height (pt):' : 'Column width:', inp)),
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          const v = parseFloat(inp.value);
          if (!(v >= 0)) return false;
          app.wb.transact(isRow ? 'Row height' : 'Column width', () => {
            if (isRow) for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) sheet.rows.set(r, Math.round(v / 0.75));
            else for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 2000); c++) sheet.cols.set(c, v === 0 ? 0 : Math.round(v * 7 + 5));
          }, true);
          app.grid.invalidate();
        },
      },
      { label: 'Cancel' },
    ],
  });
  void kind;
}

// ======================================================================== sort

export function openSort(app: App): void {
  const sheet = app.sheet;
  let rect: Rect = app.grid.range;
  if (rect.r1 === rect.r2 && rect.c1 === rect.c2) rect = currentRegion(sheet, app.activeCell.r, app.activeCell.c);
  if (rect.r1 === rect.r2) {
    void messageBox('Sheetly', 'Select a range of cells that has more than one row to sort.');
    return;
  }
  let hasHeader = detectHeader(app.wb, sheet, rect);
  const levels: SortKey[] = [{ col: rect.c1 <= app.activeCell.c && app.activeCell.c <= rect.c2 ? app.activeCell.c : rect.c1, asc: true }];
  const list = h('div', {});
  const hdr = h('input', { type: 'checkbox', checked: hasHeader }) as HTMLInputElement;
  const colLabel = (c: number) => (hasHeader ? displayText(app.wb, sheet.get(rect.r1, c)) || `Column ${colName(c)}` : `Column ${colName(c)}`);
  const render = () => {
    list.innerHTML = '';
    levels.forEach((lv, i) => {
      const colSel = select(
        Array.from({ length: rect.c2 - rect.c1 + 1 }, (_, k) => [String(rect.c1 + k), colLabel(rect.c1 + k)] as [string, string]),
        String(lv.col), (v) => (lv.col = parseInt(v, 10)));
      const ordSel = select([['asc', 'A to Z / Smallest to Largest'], ['desc', 'Z to A / Largest to Smallest']], lv.asc ? 'asc' : 'desc', (v) => (lv.asc = v === 'asc'));
      list.append(row(i === 0 ? 'Sort by' : 'Then by', colSel, ordSel));
    });
  };
  hdr.addEventListener('change', () => {
    hasHeader = hdr.checked;
    render();
  });
  const add = h('button', { class: 'btn small', type: 'button' }, 'Add Level');
  add.addEventListener('click', () => {
    if (levels.length < 8) levels.push({ col: rect.c1, asc: true });
    render();
  });
  const del = h('button', { class: 'btn small', type: 'button' }, 'Delete Level');
  del.addEventListener('click', () => {
    if (levels.length > 1) levels.pop();
    render();
  });
  render();
  openDialog({
    title: 'Sort',
    width: 560,
    content: h('div', {}, h('div', { style: { display: 'flex', gap: '6px', marginBottom: '8px' } }, add, del, h('label', { class: 'rb-check', style: { marginLeft: 'auto' } }, hdr, 'My data has headers')), list),
    buttons: [{ label: 'OK', primary: true, onClick: () => app.sortWith(levels, hasHeader, rect) }, { label: 'Cancel' }],
  });
}

// ======================================================================== conditional formatting

const CF_PRESETS: [string, { color?: string; fill?: string; bold?: boolean }][] = [
  ['Light Red Fill with Dark Red Text', { color: '#9c0006', fill: '#ffc7ce' }],
  ['Yellow Fill with Dark Yellow Text', { color: '#9c5700', fill: '#ffeb9c' }],
  ['Green Fill with Dark Green Text', { color: '#006100', fill: '#c6efce' }],
  ['Light Red Fill', { fill: '#ffc7ce' }],
  ['Red Text', { color: '#9c0006' }],
  ['Bold Blue Text', { color: '#1f4e9c', bold: true }],
];

export function openCF(app: App, kind: string): void {
  const titles: Record<string, string> = {
    gt: 'Greater Than', lt: 'Less Than', between: 'Between', eq: 'Equal To', text: 'Text That Contains', dup: 'Duplicate Values', top: 'Top 10 Items', bottom: 'Bottom 10 Items',
    above: 'Above Average', below: 'Below Average', formula: 'New Formatting Rule',
  };
  const rect = app.grid.expandForMerges(app.grid.range);
  const active = app.sheet.get(app.activeCell.r, app.activeCell.c);
  const v1 = h('input', { type: 'text', class: 'txt', style: { width: '120px' }, value: typeof active?.v === 'number' ? String(active.v) : '' }) as HTMLInputElement;
  const v2 = h('input', { type: 'text', class: 'txt', style: { width: '120px' }, value: '' }) as HTMLInputElement;
  const n = h('input', { type: 'number', min: '1', max: '1000', value: '10', style: { width: '70px' } }) as HTMLInputElement;
  const formula = h('input', { type: 'text', class: 'txt', style: { width: '100%' }, value: '=' + a1(rect.r1, rect.c1) + '>0' }) as HTMLInputElement;
  const preset = select(CF_PRESETS.map(([l], i) => [String(i), l] as [string, string]), '0');
  const applies = h('input', { type: 'text', class: 'txt', style: { width: '160px' }, value: rectText(rect) }) as HTMLInputElement;
  const err = h('div', { class: 'err' });
  let description: HTMLElement;
  switch (kind) {
    case 'gt': case 'lt': case 'eq': description = row('Format cells that are ' + (kind === 'gt' ? 'GREATER THAN' : kind === 'lt' ? 'LESS THAN' : 'EQUAL TO') + ':', v1); break;
    case 'between': description = row('Format cells that are BETWEEN:', v1, h('span', {}, 'and'), v2); break;
    case 'text': description = row('Format cells that contain the text:', v1); break;
    case 'top': case 'bottom': description = row(`Format cells that rank in the ${kind === 'top' ? 'TOP' : 'BOTTOM'}:`, n, h('span', {}, 'items')); break;
    case 'formula': description = h('div', {}, h('label', { class: 'lbl', style: { marginTop: '0' } }, 'Format values where this formula is true:'), formula); break;
    default: description = h('div', {}, kind === 'dup' ? 'Format cells that contain DUPLICATE values.' : kind === 'above' ? 'Format cells that are ABOVE AVERAGE.' : 'Format cells that are BELOW AVERAGE.');
  }
  openDialog({
    title: titles[kind] ?? 'Conditional Formatting',
    width: 480,
    content: h('div', {}, description, row('with', preset), row('Applies to:', applies), err),
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          const rg = parseRect(applies.value.replace(/\$/g, ''));
          if (!rg) {
            err.textContent = 'The "Applies to" reference is not valid.';
            return false;
          }
          const style = CF_PRESETS[parseInt(preset.value, 10)][1];
          const base = { id: app.wb.nextRuleId++, rect: rg, style: { ...style } };
          let rule: CFRule;
          switch (kind) {
            case 'gt': rule = { ...base, type: 'cell', op: '>', v1: v1.value }; break;
            case 'lt': rule = { ...base, type: 'cell', op: '<', v1: v1.value }; break;
            case 'eq': rule = { ...base, type: 'cell', op: '=', v1: v1.value }; break;
            case 'between': rule = { ...base, type: 'cell', op: 'between', v1: v1.value, v2: v2.value }; break;
            case 'text': rule = { ...base, type: 'text', op: 'contains', v1: v1.value }; break;
            case 'dup': rule = { ...base, type: 'dup' }; break;
            case 'top': rule = { ...base, type: 'top', n: parseInt(n.value, 10) || 10 }; break;
            case 'bottom': rule = { ...base, type: 'bottom', n: parseInt(n.value, 10) || 10 }; break;
            case 'above': rule = { ...base, type: 'above' }; break;
            case 'below': rule = { ...base, type: 'below' }; break;
            default: rule = { ...base, type: 'formula', v1: formula.value };
          }
          app.wb.transact('Conditional formatting', () => app.sheet.cf.unshift(rule), true);
          app.grid.invalidate();
        },
      },
      { label: 'Cancel' },
    ],
  });
}

function ruleText(r: CFRule): string {
  switch (r.type) {
    case 'cell': return `Cell Value ${r.op === 'between' ? 'between ' + r.v1 + ' and ' + r.v2 : r.op + ' ' + r.v1}`;
    case 'text': return `Text contains "${r.v1}"`;
    case 'dup': return 'Duplicate Values';
    case 'unique': return 'Unique Values';
    case 'top': return `Top ${r.n ?? 10}`;
    case 'bottom': return `Bottom ${r.n ?? 10}`;
    case 'above': return 'Above Average';
    case 'below': return 'Below Average';
    case 'scale': return `${r.colors?.length ?? 2}-Color Scale`;
    case 'bar': return 'Data Bar';
    default: return 'Formula: ' + (r.v1 ?? '');
  }
}

export function openCFManager(app: App): void {
  const sheet = app.sheet;
  let sel = -1;
  const tbody = h('tbody');
  const render = () => {
    tbody.innerHTML = '';
    if (!sheet.cf.length) tbody.append(h('tr', {}, h('td', { colspan: '3', style: { color: 'var(--muted)' } }, 'No rules on this sheet.')));
    sheet.cf.forEach((r, i) => {
      const prev = h('span', { style: { display: 'inline-block', padding: '0 8px', background: r.style?.fill ?? (r.type === 'scale' ? `linear-gradient(90deg, ${r.colors?.join(',')})` : r.type === 'bar' ? r.barColor : ''), color: r.style?.color ?? '' } }, 'AaBbCc');
      const tr = h('tr', { class: i === sel ? 'sel' : '' }, h('td', {}, ruleText(r)), h('td', {}, prev), h('td', {}, rectText(r.rect)));
      tr.addEventListener('click', () => {
        sel = i;
        render();
      });
      tbody.append(tr);
    });
  };
  const mv = (d: number) => {
    if (sel < 0) return;
    const j = sel + d;
    if (j < 0 || j >= sheet.cf.length) return;
    app.wb.transact('Move rule', () => {
      const [x] = sheet.cf.splice(sel, 1);
      sheet.cf.splice(j, 0, x);
    }, true);
    sel = j;
    render();
    app.grid.invalidate();
  };
  const btn = (label: string, fn: () => void) => {
    const b = h('button', { class: 'btn small', type: 'button' }, label);
    b.addEventListener('click', fn);
    return b;
  };
  render();
  openDialog({
    title: 'Conditional Formatting Rules Manager',
    width: 580,
    content: h('div', {},
      h('div', { style: { display: 'flex', gap: '6px', marginBottom: '8px' } },
        btn('Delete Rule', () => {
          if (sel < 0) return;
          app.wb.transact('Delete rule', () => sheet.cf.splice(sel, 1), true);
          sel = -1;
          render();
          app.grid.invalidate();
        }),
        btn('Move Up', () => mv(-1)), btn('Move Down', () => mv(1))),
      h('div', { class: 'listbox', style: { height: '220px' } }, h('table', { class: 'table-list' }, h('thead', {}, h('tr', {}, h('th', {}, 'Rule (first match wins ties)'), h('th', {}, 'Format'), h('th', {}, 'Applies to'))), tbody))),
    buttons: [{ label: 'Close', primary: true }],
  });
}

// ======================================================================== insert function, stats, about

export function openInsertFunction(app: App): void {
  const search = h('input', { type: 'text', class: 'txt', style: { width: '100%' }, placeholder: 'Search for a function' }) as HTMLInputElement;
  const catSel = select([['all', 'All'], ...Object.keys(FUNCTION_CATEGORIES).map((k) => [k, k] as [string, string])], 'all');
  const list = h('div', { class: 'listbox', style: { height: '180px', marginTop: '6px' } });
  const desc = h('div', { style: { marginTop: '8px', minHeight: '48px' } });
  let selName = '';
  const render = () => {
    list.innerHTML = '';
    const q = search.value.trim().toUpperCase();
    const cat = catSel.value;
    const names = Array.from(FUNCS.keys())
      .filter((n) => (!q || n.includes(q)) && (cat === 'all' || FUNCTION_CATEGORIES[cat].includes(n)))
      .sort();
    for (const n of names) {
      const li = h('div', { class: 'li' + (n === selName ? ' sel' : '') }, n);
      li.addEventListener('click', () => {
        selName = n;
        const d = FUNCS.get(n)!;
        desc.innerHTML = '';
        desc.append(h('b', {}, `${n}(${d.sig})`), h('div', {}, d.desc));
        list.querySelectorAll('.li').forEach((e) => e.classList.remove('sel'));
        li.classList.add('sel');
      });
      li.addEventListener('dblclick', () => {
        selName = n;
        ok.click();
      });
      list.append(li);
    }
  };
  search.addEventListener('input', render);
  catSel.addEventListener('change', render);
  render();
  const ok = h('button', { class: 'btn primary', type: 'button' }, 'OK');
  const handle = openDialog({
    title: 'Insert Function',
    width: 440,
    content: h('div', {}, search, row('Category:', catSel), list, desc),
    buttons: [
      {
        label: 'OK',
        primary: true,
        onClick: () => {
          if (selName) setTimeout(() => void app.run('insertFn', selName), 0);
        },
      },
      { label: 'Cancel' },
    ],
  });
  void handle;
  void ok;
}

export function openStats(app: App): void {
  const lines: [string, string][] = [];
  let cells = 0;
  let formulas = 0;
  for (const s of app.wb.sheets) {
    let sc = 0;
    for (const c of s.cells.values()) {
      if (c.v !== null || c.f !== undefined) sc++;
      if (c.f !== undefined) formulas++;
    }
    cells += sc;
    lines.push([s.name, sc.toLocaleString('en-US') + ' cells']);
  }
  openDialog({
    title: 'Workbook Statistics',
    width: 360,
    content: h('div', {}, h('div', { class: 'form-row' }, h('b', {}, 'Sheets:'), String(app.wb.sheets.length)), h('div', { class: 'form-row' }, h('b', {}, 'Cells with data:'), cells.toLocaleString('en-US')), h('div', { class: 'form-row' }, h('b', {}, 'Formulas:'), formulas.toLocaleString('en-US')), h('div', { class: 'form-row' }, h('b', {}, 'Names:'), String(app.wb.names.size)), h('fieldset', {}, h('legend', {}, 'Per sheet'), ...lines.map(([a, b]) => h('div', { class: 'form-row' }, h('span', { style: { minWidth: '140px' } }, a), b)))),
    buttons: [{ label: 'Close', primary: true }],
  });
}

export function openAbout(): void {
  openDialog({
    title: 'About Sheetly',
    width: 380,
    content: h('div', {}, h('div', { style: { fontSize: '18px', fontWeight: '600', marginBottom: '6px' } }, 'Sheetly 0.1'), h('div', { class: 'msg-text' }, 'A small, fast, offline spreadsheet for Windows.\n\nSheetly never connects to the internet and collects no data. It is made for personal, non-commercial use.\n\nBuilt with Tauri 2, Rust and TypeScript.')),
    buttons: [{ label: 'OK', primary: true }],
  });
}

// ======================================================================== filter dropdown

export function showFilterPanel(app: App, col: number, anchor: DOMRect): void {
  const sheet = app.sheet;
  const f = sheet.filter;
  if (!f) return;
  const header = displayText(app.wb, sheet.get(f.rect.r1, col)) || `Column ${colName(col)}`;
  const counts = new Map<string, number>();
  for (let r = f.rect.r1 + 1; r <= f.rect.r2; r++) {
    const t = filterValueText(app.wb, sheet.get(r, col));
    counts.set(t, (counts.get(t) ?? 0) + 1);
  }
  const values = Array.from(counts.keys()).sort((a, b) => {
    const na = parseFloat(a);
    const nb = parseFloat(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0;
  });
  const allowed = new Set<string>(f.cols.get(col) ?? values);
  const wrap = h('div', { style: { width: '230px' } });
  const mk = (label: string, icn: string | null, fn: () => void, disabled = false) => {
    const it = h('div', { class: 'menu-item' + (disabled ? ' disabled' : '') }, h('span', { class: 'mi-icon', html: icn ? icon(icn, 14) : '' }), h('span', { class: 'mi-label' }, label));
    if (!disabled) it.addEventListener('click', () => {
      closeMenus();
      fn();
    });
    return it;
  };
  const sortBy = (asc: boolean) => app.sortWith([{ col, asc }], true, f.rect);
  wrap.append(
    mk('Sort A to Z', 'sort', () => sortBy(true)),
    mk('Sort Z to A', 'sort', () => sortBy(false)),
    h('div', { class: 'menu-sep' }),
    mk(`Clear Filter From "${header}"`, null, () => {
      f.cols.delete(col);
      applyFilter(app.wb, sheet);
      app.grid.invalidate();
    }, !f.cols.has(col)),
    h('div', { class: 'menu-sep' }),
  );
  const search = h('input', { type: 'text', class: 'txt', placeholder: 'Search', style: { width: 'calc(100% - 8px)', margin: '2px 4px' } }) as HTMLInputElement;
  const list = h('div', { style: { maxHeight: '200px', overflowY: 'auto', margin: '2px 4px', border: '1px solid var(--border-soft)' } });
  const allBox = h('input', { type: 'checkbox' }) as HTMLInputElement;
  const boxes: [string, HTMLInputElement][] = [];
  const renderList = () => {
    list.innerHTML = '';
    boxes.length = 0;
    const q = search.value.trim().toLowerCase();
    let shown = 0;
    let hidden = 0;
    for (const v of values) {
      if (q && !v.toLowerCase().includes(q)) continue;
      if (shown >= 1000) {
        hidden++;
        continue;
      }
      shown++;
      const cb = h('input', { type: 'checkbox', checked: allowed.has(v) }) as HTMLInputElement;
      cb.addEventListener('change', () => {
        if (cb.checked) allowed.add(v);
        else allowed.delete(v);
        syncAll();
      });
      boxes.push([v, cb]);
      list.append(h('label', { class: 'rb-check', style: { padding: '1px 4px' } }, cb, v === '' ? '(Blanks)' : v));
    }
    if (hidden) list.append(h('div', { style: { padding: '2px 6px', color: 'var(--muted)' } }, `${hidden.toLocaleString('en-US')} more not shown. Use search to narrow the list.`));
    syncAll();
  };
  const syncAll = () => {
    allBox.checked = boxes.length > 0 && boxes.every(([, b]) => b.checked);
    allBox.indeterminate = !allBox.checked && boxes.some(([, b]) => b.checked);
  };
  allBox.addEventListener('change', () => {
    for (const [v, b] of boxes) {
      b.checked = allBox.checked;
      if (allBox.checked) allowed.add(v);
      else allowed.delete(v);
    }
  });
  search.addEventListener('input', renderList);
  search.addEventListener('keydown', (e) => e.stopPropagation());
  renderList();
  const ok = h('button', { class: 'btn primary small', type: 'button' }, 'OK');
  ok.addEventListener('click', () => {
    closeMenus();
    if (values.every((v) => allowed.has(v))) f.cols.delete(col);
    else f.cols.set(col, new Set(allowed));
    applyFilter(app.wb, sheet);
    app.grid.invalidate();
  });
  const cancel = h('button', { class: 'btn small', type: 'button' }, 'Cancel');
  cancel.addEventListener('click', () => closeMenus());
  wrap.append(search, h('label', { class: 'rb-check', style: { padding: '1px 8px' } }, allBox, '(Select All)'), list, h('div', { style: { display: 'flex', gap: '6px', justifyContent: 'flex-end', padding: '6px 6px 2px' } }, ok, cancel));
  const a = h('div', { style: { position: 'fixed', left: anchor.right - 230 + 'px', top: anchor.bottom + 'px', width: '0', height: '0' } });
  document.body.append(a);
  showPopup(a, wrap);
  setTimeout(() => a.remove(), 0);
  void toast;
  void FErr;
  void deleteRows;
  void deleteCols;
  void insertRows;
  void insertCols;
  void MAX_COLS;
  void MAX_ROWS;
  void showColorMenu;
  void ({} as MenuItem);
}
