// Print / save as PDF: renders the used range as an HTML table and calls the system print dialog.
import { colName } from '../core/addr';
import { displayText } from '../core/ops';
import { FErr } from '../core/values';
import { toast } from './dom';
import type { App } from './app';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function printSheet(app: App): void {
  const sheet = app.sheet;
  const wb = app.wb;
  const used = sheet.usedRect();
  if (!used) {
    toast('There is nothing to print on this sheet.');
    return;
  }
  const c2 = Math.min(used.c2, 200);
  // cap printed rows (20k) and total cells (300k) so the print table stays manageable
  const r2 = Math.min(used.r2, 19999, Math.max(0, Math.floor(300000 / (c2 + 1)) - 1));
  if (r2 < used.r2) toast(`Only the first ${(r2 + 1).toLocaleString('en-US')} rows are printed.`);
  const merged = new Map<string, { rs: number; cs: number } | 'skip'>();
  for (const m of sheet.merges) {
    if (m.r1 > r2 || m.c1 > c2) continue;
    for (let r = m.r1; r <= Math.min(m.r2, r2); r++) for (let c = m.c1; c <= Math.min(m.c2, c2); c++) merged.set(r + ',' + c, 'skip');
    merged.set(m.r1 + ',' + m.c1, { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
  }
  const showHdr = app.printGridlines;
  let html = '<table>';
  html += '<colgroup>';
  if (showHdr) html += '<col style="width:34px">';
  for (let c = 0; c <= c2; c++) if (!sheet.hiddenCol(c)) html += `<col style="width:${sheet.cols.size(c)}px">`;
  html += '</colgroup>';
  if (showHdr) {
    html += '<tr><th></th>';
    for (let c = 0; c <= c2; c++) if (!sheet.hiddenCol(c)) html += `<th>${colName(c)}</th>`;
    html += '</tr>';
  }
  for (let r = 0; r <= r2; r++) {
    if (sheet.hiddenRow(r)) continue;
    html += `<tr style="height:${sheet.rows.size(r)}px">`;
    if (showHdr) html += `<th>${r + 1}</th>`;
    for (let c = 0; c <= c2; c++) {
      if (sheet.hiddenCol(c)) continue;
      const m = merged.get(r + ',' + c);
      if (m === 'skip') continue;
      const cell = sheet.get(r, c);
      const st = wb.styleOf(cell?.s);
      const css: string[] = [];
      if (st.font) css.push(`font-family:"${st.font}"`);
      if (st.size) css.push(`font-size:${st.size}pt`);
      if (st.bold) css.push('font-weight:bold');
      if (st.italic) css.push('font-style:italic');
      const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
      if (deco) css.push(`text-decoration:${deco}`);
      if (st.color) css.push(`color:${st.color}`);
      if (st.fill) css.push(`background:${st.fill};-webkit-print-color-adjust:exact;print-color-adjust:exact`);
      const v = cell?.v;
      const ha = st.h ?? (typeof v === 'number' ? 'right' : typeof v === 'boolean' || v instanceof FErr ? 'center' : 'left');
      css.push(`text-align:${ha}`);
      css.push(`vertical-align:${st.v === 'top' ? 'top' : st.v === 'middle' ? 'middle' : 'bottom'}`);
      if (st.wrap) css.push('white-space:pre-wrap');
      for (const [k, side] of [['bl', 'left'], ['br', 'right'], ['bt', 'top'], ['bb', 'bottom']] as const) {
        const b = st[k];
        if (b) css.push(`border-${side}:${b.style === 'thick' ? 3 : b.style === 'medium' ? 2 : 1}px ${b.style === 'double' ? 'double' : b.style === 'dashed' ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid'} ${b.color}`);
      }
      const span = m ? ` rowspan="${m.rs}" colspan="${m.cs}"` : '';
      html += `<td${span} style="${css.join(';')}">${esc(displayText(wb, cell)).replace(/\n/g, '<br>')}</td>`;
    }
    html += '</tr>';
  }
  html += '</table>';
  const area = document.getElementById('print-area')!;
  area.className = app.printGridlines ? 'grid' : '';
  area.innerHTML = html;
  let style = document.getElementById('page-style') as HTMLStyleElement | null;
  if (!style) {
    style = document.createElement('style');
    style.id = 'page-style';
    document.head.append(style);
  }
  style.textContent = `@page { size: ${app.orientation}; margin: 0.6in; }`;
  const cleanup = () => {
    area.innerHTML = '';
    window.removeEventListener('afterprint', cleanup);
    clearTimeout(safety);
  };
  window.addEventListener('afterprint', cleanup);
  // afterprint may never fire; free the table anyway
  const safety = window.setTimeout(cleanup, 120000);
  setTimeout(() => window.print(), 50);
}
