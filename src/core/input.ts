// Turning typed text into cell content, and cell content back into editable text.
import { FErr, E, Value, errFromCode, ymdToSerial } from './values';
import { parseNumberText, parseDateText, numToStr } from './eval';
import { isDateFormat, formatValue } from './numfmt';
import type { Cell, CellStyle } from './model';

export interface Parsed {
  v: Value;
  f?: string;
  fmt?: string; // number format to apply when the cell is currently General
}

export function parseInput(text: string, existingFmt?: string): Parsed {
  if (text === '') return { v: null };
  if (existingFmt === '@') return { v: text.startsWith("'") ? text : text };
  if (text.startsWith("'")) return { v: text.slice(1) };
  if (text.startsWith('=') && text.length > 1) return { v: null, f: text.slice(1) };
  const t = text.trim();
  if (t === '') return { v: text };
  const up = t.toUpperCase();
  if (up === 'TRUE') return { v: true };
  if (up === 'FALSE') return { v: false };
  if (t.startsWith('#')) {
    const e = errFromCode(t);
    if (e) return { v: e };
  }
  // percent
  let m = /^([-+]?)(\d[\d,]*\.?\d*|\.\d+)\s*%$/.exec(t);
  if (m) {
    const n = parseNumberText(t);
    if (n !== null) {
      const dec = (m[2].split('.')[1] ?? '').length;
      return { v: n, fmt: dec ? '0.' + '0'.repeat(dec) + '%' : '0%' };
    }
  }
  // currency
  m = /^([-+]?)\$\s*([\d,]+\.?\d*|\.\d+)$/.exec(t) || /^\(\$\s*([\d,]+\.?\d*)\)$/.exec(t);
  if (m) {
    const n = parseNumberText(t.replace(/^\(\$\s*/, '(').replace(/^([-+]?)\$\s*/, '$1'));
    if (n !== null) {
      const dec = ((m[2] ?? m[1]).split('.')[1] ?? '').length;
      return { v: n, fmt: dec > 0 ? '"$"#,##0.' + '0'.repeat(dec) : '"$"#,##0' };
    }
  }
  const n = parseNumberText(t);
  if (n !== null) {
    if (t.includes(',')) {
      const dec = (t.split('.')[1] ?? '').length;
      return { v: n, fmt: dec ? '#,##0.' + '0'.repeat(dec) : '#,##0' };
    }
    return { v: n };
  }
  // dates and times
  const d = parseDateText(t);
  if (d !== null) {
    const hasDate = /[\/-]|\b[A-Za-z]{3,}\b/.test(t) && !/^\d{1,2}:\d{2}/.test(t);
    const hasTime = /\d:\d{2}/.test(t);
    let fmt: string;
    if (hasDate && hasTime) fmt = 'm/d/yyyy h:mm';
    else if (hasDate) fmt = /^\d{4}-/.test(t) ? 'yyyy-mm-dd' : /[A-Za-z]/.test(t) ? 'd-mmm-yy' : 'm/d/yyyy';
    else if (/AM|PM|A|P$/i.test(t)) fmt = 'h:mm AM/PM';
    else fmt = /\d:\d{2}:\d{2}/.test(t) ? 'h:mm:ss' : 'h:mm';
    return { v: d, fmt };
  }
  // month/day without year
  m = /^(\d{1,2})\/(\d{1,2})$/.exec(t);
  if (m) {
    const mo = parseInt(m[1], 10);
    const dd = parseInt(m[2], 10);
    if (mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31) {
      return { v: ymdToSerial(new Date().getFullYear(), mo, dd), fmt: 'd-mmm' };
    }
  }
  return { v: text };
}

/** Text shown in the formula bar / cell editor. */
export function editText(cell: Cell | undefined, style?: CellStyle): string {
  if (!cell) return '';
  if (cell.f !== undefined) return '=' + cell.f;
  const v = cell.v;
  if (v === null) return '';
  if (v instanceof FErr) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    const fmt = style?.fmt;
    if (fmt && isDateFormat(fmt)) {
      const hasTime = /[hs]/i.test(fmt.replace(/"[^"]*"/g, ''));
      const hasDate = /[ymd]/i.test(fmt.replace(/"[^"]*"/g, ''));
      if (hasDate && hasTime) return formatValue(v, 'm/d/yyyy h:mm:ss').text;
      if (hasDate) return formatValue(v, 'm/d/yyyy').text;
      return formatValue(v, 'h:mm:ss').text;
    }
    if (fmt && /%/.test(fmt) && !isDateFormat(fmt)) {
      return numToStr(parseFloat((v * 100).toPrecision(15))) + '%';
    }
    return numToStr(v);
  }
  // text: protect strings that would be re-parsed as something else
  if (v.startsWith('=') || v.startsWith("'")) return "'" + v;
  const p = parseInput(v);
  if (typeof p.v !== 'string' || p.f !== undefined) return "'" + v;
  return v;
}

export { E };
