// CSV import / export.
import { MAX_COLS, MAX_ROWS } from './addr';
import { Sheet, Workbook } from './model';
import { applyInput, displayText } from './ops';

export function detectDelimiter(text: string): string {
  const head = text.slice(0, 4000).split(/\r?\n/).slice(0, 5).join('\n');
  const counts: [string, number][] = [',', ';', '\t', '|'].map((d) => [d, head.split(d).length - 1]);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ',';
}

export function parseCSV(text: string, delim?: string): string[][] {
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const d = delim ?? detectDelimiter(text);
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let inQ = false;
  let started = false;
  const n = text.length;
  for (let i = 0; i < n; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i++;
        } else inQ = false;
      } else cur += ch;
      continue;
    }
    if (ch === '"' && cur === '') {
      inQ = true;
      started = true;
    } else if (ch === d) {
      row.push(cur);
      cur = '';
      started = true;
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur);
      rows.push(row);
      row = [];
      cur = '';
      started = false;
    } else {
      cur += ch;
      started = true;
    }
  }
  if (inQ) throw new Error('The CSV file ends inside a quoted value (missing closing quote).');
  if (started || cur !== '' || row.length) {
    row.push(cur);
    rows.push(row);
  }
  return rows;
}

export function readCSV(text: string): Workbook {
  const wb = new Workbook();
  const sheet = wb.sheets[0];
  const rows = parseCSV(text);
  if (rows.length > MAX_ROWS) throw new Error(`The CSV file has ${rows.length.toLocaleString()} rows; Sheetly supports at most ${MAX_ROWS.toLocaleString()}.`);
  if (rows.some((r) => r.length > MAX_COLS)) throw new Error(`The CSV file has more than ${MAX_COLS.toLocaleString()} columns.`);
  wb.transact('import', () => {
    for (let r = 0; r < rows.length; r++)
      for (let c = 0; c < rows[r].length; c++) {
        const t = rows[r][c];
        if (t === '') continue;
        // keep identifiers, long digit strings and "=..." as text so nothing is mangled
        if (t[0] === '=' || /^[+-]?0\d/.test(t) || /^[+-]?\d{16,}$/.test(t)) {
          wb.setValue(sheet, r, c, t);
          continue;
        }
        try {
          applyInput(wb, sheet, r, c, t);
        } catch {
          wb.setValue(sheet, r, c, t);
        }
      }
  });
  wb.clearHistory();
  return wb;
}

export function writeCSV(wb: Workbook, sheet: Sheet): string {
  const used = sheet.usedRect();
  if (!used) return '';
  const lines: string[] = [];
  for (let r = 0; r <= used.r2; r++) {
    const parts: string[] = [];
    for (let c = 0; c <= used.c2; c++) {
      const s = displayText(wb, sheet.get(r, c));
      parts.push(/[",\r\n]/.test(s) || /^\s|\s$/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
    }
    lines.push(parts.join(','));
  }
  return '﻿' + lines.join('\r\n') + '\r\n';
}
