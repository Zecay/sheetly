// .xlsx reader and writer (values, formulas, basic formatting, widths, merges, names, simple conditional formats).
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { MAX_COLS, MAX_ROWS, Rect, a1, cellKey, colIndex, colName, keyCol, keyRow, parseCell, parseRect, rectText } from './addr';
import { BUILTIN_FORMATS } from './numfmt';
import { Border, CFRule, Cell, CellStyle, Sheet, Workbook } from './model';
import { XNode, child, esc, escText, kids, parseXml } from './xml';
import { E, FErr, Value, errFromCode } from './values';
import { shiftFormula, tokenize } from './parser';

// ---------------- colours ----------------

const THEME_DEFAULT = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

function tint(hex: string, t: number): string {
  const c = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = c.map((v) => {
    const x = t < 0 ? v * (1 + t) : v + (255 - v) * t;
    return Math.max(0, Math.min(255, Math.round(x)));
  });
  return out.map((v) => v.toString(16).padStart(2, '0')).join('');
}

function readColor(n: XNode | undefined, theme: string[]): string | undefined {
  if (!n) return undefined;
  const a = n.attrs;
  let hex: string | undefined;
  if (a.rgb) hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb;
  else if (a.theme !== undefined) hex = theme[parseInt(a.theme, 10)] ?? THEME_DEFAULT[parseInt(a.theme, 10)];
  else if (a.indexed !== undefined) hex = INDEXED[parseInt(a.indexed, 10)];
  else if (a.auto) return undefined;
  if (!hex) return undefined;
  if (a.tint) hex = tint(hex, parseFloat(a.tint));
  return '#' + hex.toLowerCase();
}

// ---------------- reading ----------------

function textOf(si: XNode): string {
  // <si><t>..</t></si> or rich text <r><t>..</t></r>
  let s = '';
  for (const c of si.children) {
    if (c.tag === 't') s += c.text;
    else if (c.tag === 'r') s += kids(c, 't').map((t) => t.text).join('');
  }
  return s;
}

const BORDER_MAP: Record<string, Border['style']> = {
  thin: 'thin', medium: 'medium', thick: 'thick', dashed: 'dashed', dotted: 'dotted', double: 'double', hair: 'hair',
  mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed', dashDotDot: 'dashed', mediumDashDotDot: 'dashed', slantDashDot: 'dashed',
};

function parseFont(f: XNode, theme: string[]): CellStyle {
  const st: CellStyle = {};
  const name = child(f, 'name')?.attrs.val;
  if (name && name !== 'Calibri') st.font = name;
  const sz = child(f, 'sz')?.attrs.val;
  if (sz && parseFloat(sz) !== 11) st.size = parseFloat(sz);
  const b = child(f, 'b');
  if (b && b.attrs.val !== '0' && b.attrs.val !== 'false') st.bold = true;
  const i = child(f, 'i');
  if (i && i.attrs.val !== '0' && i.attrs.val !== 'false') st.italic = true;
  const u = child(f, 'u');
  if (u && u.attrs.val !== 'none') st.underline = true;
  const s = child(f, 'strike');
  if (s && s.attrs.val !== '0' && s.attrs.val !== 'false') st.strike = true;
  const col = readColor(child(f, 'color'), theme);
  if (col && col !== '#000000') st.color = col;
  return st;
}

function parseBorderSide(n: XNode | undefined, theme: string[]): Border | undefined {
  if (!n || !n.attrs.style) return undefined;
  const style = BORDER_MAP[n.attrs.style];
  if (!style) return undefined;
  return { style, color: readColor(child(n, 'color'), theme) ?? '#000000' };
}

export interface XlsxResult {
  wb: Workbook;
}

const MAX_ENTRY_BYTES = 100 * 1024 * 1024;
const MAX_TOTAL_BYTES = 200 * 1024 * 1024;

export function readXlsx(data: Uint8Array): Workbook {
  let totalSize = 0;
  const files = unzipSync(data, {
    filter: (f) => {
      if (f.originalSize > MAX_ENTRY_BYTES) throw new Error('The .xlsx file contains an entry that is too large to open safely.');
      totalSize += f.originalSize;
      if (totalSize > MAX_TOTAL_BYTES) throw new Error('The .xlsx file expands to too much data to open safely.');
      return true;
    },
  });
  const get = (name: string): string | undefined => {
    const f = files[name] ?? files[name.replace(/^\//, '')];
    return f ? strFromU8(f) : undefined;
  };
  const wbXml = get('xl/workbook.xml');
  if (!wbXml) throw new Error('Not a valid .xlsx file (xl/workbook.xml is missing).');
  const wbNode = parseXml(wbXml);
  const rels = new Map<string, string>();
  const relXml = get('xl/_rels/workbook.xml.rels');
  if (relXml) for (const r of kids(parseXml(relXml), 'Relationship')) rels.set(r.attrs.Id, r.attrs.Target);

  // theme colours
  let theme = THEME_DEFAULT.slice();
  const themeXml = get('xl/theme/theme1.xml');
  if (themeXml) {
    const scheme = child(child(parseXml(themeXml), 'themeElements'), 'clrScheme');
    if (scheme) {
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      theme = order.map((k, i) => {
        const n = child(scheme, k);
        const c = n?.children[0];
        const v = c?.attrs.lastClr ?? c?.attrs.val;
        return v && /^[0-9A-Fa-f]{6}$/.test(v) ? v : THEME_DEFAULT[i];
      });
    }
  }

  // shared strings
  const sst: string[] = [];
  const sstXml = get('xl/sharedStrings.xml');
  (sst as any).missing = !sstXml;
  if (sstXml) for (const si of kids(parseXml(sstXml), 'si')) sst.push(textOf(si));

  const wb = new Workbook();
  wb.sheets = [];

  // styles
  const xfStyle: number[] = [];
  const dxfs: { color?: string; fill?: string; bold?: boolean; italic?: boolean }[] = [];
  const stylesXml = get('xl/styles.xml');
  if (stylesXml) {
    const sn = parseXml(stylesXml);
    const numFmts = new Map<number, string>();
    for (const nf of kids(child(sn, 'numFmts'), 'numFmt')) numFmts.set(parseInt(nf.attrs.numFmtId, 10), nf.attrs.formatCode);
    const fonts = kids(child(sn, 'fonts'), 'font').map((f) => parseFont(f, theme));
    const fills = kids(child(sn, 'fills'), 'fill').map((f) => {
      const pf = child(f, 'patternFill');
      if (!pf || pf.attrs.patternType !== 'solid') return undefined;
      return readColor(child(pf, 'fgColor'), theme);
    });
    const borders = kids(child(sn, 'borders'), 'border').map((b) => ({
      bl: parseBorderSide(child(b, 'left'), theme) ?? parseBorderSide(child(b, 'start'), theme),
      br: parseBorderSide(child(b, 'right'), theme) ?? parseBorderSide(child(b, 'end'), theme),
      bt: parseBorderSide(child(b, 'top'), theme),
      bb: parseBorderSide(child(b, 'bottom'), theme),
    }));
    for (const xf of kids(child(sn, 'cellXfs'), 'xf')) {
      const st: CellStyle = {};
      const fid = parseInt(xf.attrs.fontId ?? '0', 10);
      Object.assign(st, fonts[fid] ?? {});
      const fill = fills[parseInt(xf.attrs.fillId ?? '0', 10)];
      if (fill) st.fill = fill;
      const bd = borders[parseInt(xf.attrs.borderId ?? '0', 10)];
      if (bd) {
        if (bd.bl) st.bl = bd.bl;
        if (bd.br) st.br = bd.br;
        if (bd.bt) st.bt = bd.bt;
        if (bd.bb) st.bb = bd.bb;
      }
      const nid = parseInt(xf.attrs.numFmtId ?? '0', 10);
      const code = numFmts.get(nid) ?? BUILTIN_FORMATS[nid];
      if (code && code !== 'General') st.fmt = code;
      const al = child(xf, 'alignment');
      if (al) {
        const h = al.attrs.horizontal;
        if (h === 'left' || h === 'center' || h === 'right' || h === 'justify') st.h = h;
        else if (h === 'centerContinuous') st.h = 'center';
        const v = al.attrs.vertical;
        if (v === 'top') st.v = 'top';
        else if (v === 'center') st.v = 'middle';
        if (al.attrs.wrapText === '1' || al.attrs.wrapText === 'true') st.wrap = true;
        if (al.attrs.indent) st.indent = parseInt(al.attrs.indent, 10);
      }
      xfStyle.push(wb.styleIndex(st));
    }
    for (const d of kids(child(sn, 'dxfs'), 'dxf')) {
      const font = child(d, 'font');
      const fill = child(child(d, 'fill'), 'patternFill');
      dxfs.push({
        color: readColor(child(font, 'color'), theme),
        bold: font && child(font, 'b') ? true : undefined,
        italic: font && child(font, 'i') ? true : undefined,
        fill: readColor(child(fill, 'bgColor'), theme) ?? readColor(child(fill, 'fgColor'), theme),
      });
    }
  }

  // sheets
  const sheetNodes = kids(child(wbNode, 'sheets'), 'sheet');
  const parsedSheets: { sheet: Sheet; xml: XNode }[] = [];
  for (const sn of sheetNodes) {
    const rid = sn.attrs['r:id'] ?? sn.attrs.id;
    let target = rels.get(rid) ?? '';
    if (target.startsWith('/')) target = target.slice(1);
    else if (!target.startsWith('xl/')) target = 'xl/' + target;
    const xml = get(target);
    const sheet = wb.addSheet(wb.uniqueSheetName(sn.attrs.name || 'Sheet'));
    if (xml) parsedSheets.push({ sheet, xml: parseXml(xml) });
  }
  if (!wb.sheets.length) wb.addSheet('Sheet1');

  for (const { sheet, xml } of parsedSheets) readSheet(wb, sheet, xml, sst, xfStyle, dxfs, theme);

  // defined names
  for (const dn of kids(child(wbNode, 'definedNames'), 'definedName')) {
    const name = dn.attrs.name;
    if (!name || name.startsWith('_xlnm') || dn.attrs.hidden === '1') continue;
    if (!wb.getName(name)) wb.setName(name, stripXlfn(dn.text.replace(/^=/, '')));
  }
  wb.rebuildDeps();
  wb.recalcAll();
  wb.clearHistory();
  return wb;
}

function stripXlfn(f: string): string {
  return f.replace(/_xlfn\.|_xlws\./g, '');
}

function readSheet(wb: Workbook, sheet: Sheet, root: XNode, sst: string[], xfStyle: number[], dxfs: any[], theme: string[]): void {
  const tab = child(child(root, 'sheetPr'), 'tabColor');
  const tc = readColor(tab, theme);
  if (tc) sheet.tabColor = tc;
  // views
  const view = child(child(root, 'sheetViews'), 'sheetView');
  if (view) {
    if (view.attrs.showGridLines === '0') sheet.showGrid = false;
    const pane = child(view, 'pane');
    if (pane && (pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit')) {
      sheet.freezeCols = Math.floor(parseFloat(pane.attrs.xSplit ?? '0'));
      sheet.freezeRows = Math.floor(parseFloat(pane.attrs.ySplit ?? '0'));
    }
  }
  const fmtPr = child(root, 'sheetFormatPr');
  const defH = fmtPr?.attrs.defaultRowHeight ? parseFloat(fmtPr.attrs.defaultRowHeight) * (4 / 3) : 20;
  if (Math.abs(defH - 20) > 0.5) sheet.rows.def = Math.round(defH);
  // columns
  for (const col of kids(child(root, 'cols'), 'col')) {
    const min = parseInt(col.attrs.min, 10) - 1;
    const max = Math.min(parseInt(col.attrs.max, 10) - 1, MAX_COLS - 1);
    const hidden = col.attrs.hidden === '1' || col.attrs.hidden === 'true';
    const w = col.attrs.width ? Math.round(parseFloat(col.attrs.width) * 7 + 5) : undefined;
    if (max - min > 2000) continue;
    for (let c = min; c <= max; c++) {
      if (hidden) sheet.cols.set(c, 0);
      else if (w !== undefined) sheet.cols.set(c, w);
    }
  }
  // cells
  const shared = new Map<string, { f: string; r: number; c: number }>();
  const sd = child(root, 'sheetData');
  let maxR = -1;
  for (const row of kids(sd, 'row')) {
    const rIdx = row.attrs.r ? parseInt(row.attrs.r, 10) - 1 : maxR + 1;
    maxR = rIdx;
    if (rIdx >= MAX_ROWS || rIdx < 0) continue;
    if (row.attrs.hidden === '1' || row.attrs.hidden === 'true') sheet.rows.set(rIdx, 0);
    else if (row.attrs.ht && (row.attrs.customHeight === '1' || row.attrs.customHeight === 'true')) sheet.rows.set(rIdx, Math.round(parseFloat(row.attrs.ht) * (4 / 3)));
    let cIdx = -1;
    for (const c of kids(row, 'c')) {
      let r = rIdx;
      let col: number;
      if (c.attrs.r) {
        const p = parseCell(c.attrs.r);
        if (!p) continue;
        r = p.r;
        col = p.c;
      } else col = cIdx + 1;
      cIdx = col;
      if (r >= MAX_ROWS || r < 0 || col >= MAX_COLS || col < 0) continue;
      const s = c.attrs.s ? xfStyle[parseInt(c.attrs.s, 10)] : undefined;
      const t = c.attrs.t;
      const fNode = child(c, 'f');
      const vNode = child(c, 'v');
      let v: Value = null;
      const vt = vNode?.text ?? '';
      if (t === 's') {
        if ((sst as any).missing) throw new Error('Not a valid .xlsx file (shared strings part is missing).');
        v = sst[parseInt(vt, 10)] ?? '';
      }
      else if (t === 'str') v = vt;
      else if (t === 'inlineStr') v = textOf(child(c, 'is') ?? { tag: 'is', attrs: {}, children: [], text: '' });
      else if (t === 'b') v = vt === '1';
      else if (t === 'e') v = errFromCode(vt) ?? E.VALUE;
      else if (vNode && vt !== '') {
        const num = parseFloat(vt);
        v = Number.isNaN(num) ? vt : num;
      }
      let f: string | undefined;
      if (fNode) {
        const ft = fNode.attrs.t;
        if (ft === 'shared') {
          const si = fNode.attrs.si;
          if (fNode.text) {
            f = stripXlfn(fNode.text);
            shared.set(si, { f, r, c: col });
          } else {
            const m = shared.get(si);
            if (m) f = shiftFormula(m.f, r - m.r, col - m.c);
          }
        } else if (fNode.text) f = stripXlfn(fNode.text);
      }
      if (v === null && f === undefined && s === undefined) continue;
      const cell: Cell = { v, s: s || undefined };
      if (f !== undefined) cell.f = f;
      sheet.cells.set(cellKey(r, col), cell);
      if (r > sheet.maxRow) sheet.maxRow = r;
      if (col > sheet.maxCol) sheet.maxCol = col;
    }
  }
  // merges
  for (const m of kids(child(root, 'mergeCells'), 'mergeCell')) {
    const r = parseRect(m.attrs.ref ?? '');
    if (r && (r.r2 > r.r1 || r.c2 > r.c1)) sheet.merges.push(r);
  }
  // auto filter
  const af = child(root, 'autoFilter');
  if (af?.attrs.ref) {
    const r = parseRect(af.attrs.ref);
    if (r) sheet.filter = { rect: r, cols: new Map() };
  }
  // conditional formatting
  for (const cf of kids(root, 'conditionalFormatting')) {
    const rect = parseRect((cf.attrs.sqref ?? '').split(' ')[0]);
    if (!rect) continue;
    for (const rule of kids(cf, 'cfRule')) {
      const base: CFRule = { id: wb.nextRuleId++, rect, type: 'cell' };
      const dx = rule.attrs.dxfId !== undefined ? dxfs[parseInt(rule.attrs.dxfId, 10)] : undefined;
      if (dx) base.style = { color: dx.color, fill: dx.fill, bold: dx.bold, italic: dx.italic };
      const fs = kids(rule, 'formula').map((x) => stripXlfn(x.text));
      switch (rule.attrs.type) {
        case 'cellIs': {
          const opMap: Record<string, string> = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=', equal: '=', notEqual: '<>', between: 'between', notBetween: 'notbetween' };
          base.type = 'cell';
          base.op = opMap[rule.attrs.operator ?? ''] ?? '>';
          base.v1 = fs[0];
          base.v2 = fs[1];
          break;
        }
        case 'expression':
          base.type = 'formula';
          base.v1 = fs[0];
          break;
        case 'containsText':
          base.type = 'text';
          base.op = 'contains';
          base.v1 = rule.attrs.text ?? '';
          break;
        case 'duplicateValues':
          base.type = 'dup';
          break;
        case 'uniqueValues':
          base.type = 'unique';
          break;
        case 'colorScale': {
          const cs = child(rule, 'colorScale');
          const colors = kids(cs, 'color').map((c) => readColor(c, theme) ?? '#ffffff');
          if (colors.length < 2) continue;
          base.type = 'scale';
          base.colors = colors;
          break;
        }
        case 'dataBar': {
          const db = child(rule, 'dataBar');
          base.type = 'bar';
          base.barColor = readColor(child(db, 'color'), theme) ?? '#638ec6';
          break;
        }
        default:
          continue;
      }
      sheet.cf.push(base);
    }
  }
  sheet.cols.touch();
  sheet.rows.touch();
}

// ---------------- writing ----------------

const XLFN = new Set(['IFS', 'CONCAT', 'TEXTJOIN', 'MAXIFS', 'MINIFS', 'SWITCH', 'XLOOKUP', 'XMATCH', 'IFNA', 'DAYS', 'NUMBERVALUE', 'XOR', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P']);

function addXlfn(f: string): string {
  let out = '';
  let pos = 0;
  for (const t of tokenize(f)) {
    if (t.kind === 'func' && XLFN.has(t.text.toUpperCase())) {
      out += f.slice(pos, t.start) + '_xlfn.' + t.text;
      pos = t.end;
    }
  }
  return out + f.slice(pos);
}

const REV_BUILTIN = new Map<string, number>();
for (const [id, code] of Object.entries(BUILTIN_FORMATS)) if (!REV_BUILTIN.has(code)) REV_BUILTIN.set(code, parseInt(id, 10));

function hex6(c: string | undefined): string | undefined {
  if (!c) return undefined;
  const m = /^#?([0-9a-fA-F]{6})$/.exec(c.trim());
  return m ? 'FF' + m[1].toUpperCase() : undefined;
}

function num(n: number): string {
  return Number.isInteger(n) ? String(n) : String(parseFloat(n.toPrecision(17)));
}

export function writeXlsx(wb: Workbook): Uint8Array {
  // ----- styles -----
  const fonts: string[] = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'];
  const fontIdx = new Map<string, number>([[fonts[0], 0]]);
  const fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const fillIdx = new Map<string, number>();
  const borders: string[] = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
  const borderIdx = new Map<string, number>([[borders[0], 0]]);
  const numFmts: { id: number; code: string }[] = [];
  const numFmtIdx = new Map<string, number>();
  const xfs: string[] = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const xfMap = new Map<number, number>([[0, 0]]);

  const intern = (arr: string[], map: Map<string, number>, xml: string): number => {
    let i = map.get(xml);
    if (i === undefined) {
      i = arr.length;
      arr.push(xml);
      map.set(xml, i);
    }
    return i;
  };

  const styleXf = (si: number): number => {
    const got = xfMap.get(si);
    if (got !== undefined) return got;
    const st = wb.styleOf(si);
    // font
    let fx = '<font>';
    if (st.bold) fx += '<b/>';
    if (st.italic) fx += '<i/>';
    if (st.strike) fx += '<strike/>';
    if (st.underline) fx += '<u/>';
    fx += `<sz val="${st.size ?? 11}"/>`;
    const fc = hex6(st.color);
    fx += fc ? `<color rgb="${fc}"/>` : '<color theme="1"/>';
    fx += `<name val="${esc(st.font ?? 'Calibri')}"/><family val="2"/>`;
    if (!st.font) fx += '<scheme val="minor"/>';
    fx += '</font>';
    const fontId = intern(fonts, fontIdx, fx);
    // fill
    let fillId = 0;
    const fcol = hex6(st.fill);
    if (fcol) fillId = intern(fills, fillIdx, `<fill><patternFill patternType="solid"><fgColor rgb="${fcol}"/><bgColor indexed="64"/></patternFill></fill>`);
    // border
    const side = (tag: string, b?: Border) => (b ? `<${tag} style="${b.style}"><color rgb="${hex6(b.color) ?? 'FF000000'}"/></${tag}>` : `<${tag}/>`);
    const bx = `<border>${side('left', st.bl)}${side('right', st.br)}${side('top', st.bt)}${side('bottom', st.bb)}<diagonal/></border>`;
    const borderId = intern(borders, borderIdx, bx);
    // number format
    let nid = 0;
    if (st.fmt && st.fmt !== 'General') {
      const b = REV_BUILTIN.get(st.fmt);
      if (b !== undefined) nid = b;
      else {
        let id = numFmtIdx.get(st.fmt);
        if (id === undefined) {
          id = 164 + numFmts.length;
          numFmts.push({ id, code: st.fmt });
          numFmtIdx.set(st.fmt, id);
        }
        nid = id;
      }
    }
    let xf = `<xf numFmtId="${nid}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"`;
    if (nid) xf += ' applyNumberFormat="1"';
    if (fontId) xf += ' applyFont="1"';
    if (fillId) xf += ' applyFill="1"';
    if (borderId) xf += ' applyBorder="1"';
    if (st.h || st.v || st.wrap || st.indent) {
      xf += ' applyAlignment="1"><alignment';
      if (st.h) xf += ` horizontal="${st.h}"`;
      if (st.v) xf += ` vertical="${st.v === 'middle' ? 'center' : st.v}"`;
      if (st.wrap) xf += ' wrapText="1"';
      if (st.indent) xf += ` indent="${st.indent}"`;
      xf += '/></xf>';
    } else xf += '/>';
    const idx = xfs.length;
    xfs.push(xf);
    xfMap.set(si, idx);
    return idx;
  };

  // ----- shared strings -----
  const sstList: string[] = [];
  const sstIdx = new Map<string, number>();
  const sstOf = (s: string) => {
    let i = sstIdx.get(s);
    if (i === undefined) {
      i = sstList.length;
      sstList.push(s);
      sstIdx.set(s, i);
    }
    return i;
  };

  // ----- dxfs (conditional formats) -----
  const dxfs: string[] = [];
  const dxfIdx = new Map<string, number>();
  const dxfOf = (st: NonNullable<CFRule['style']>): number => {
    let x = '<dxf>';
    if (st.bold || st.italic || st.color) {
      x += '<font>' + (st.bold ? '<b/>' : '') + (st.italic ? '<i/>' : '') + (st.color ? `<color rgb="${hex6(st.color)}"/>` : '') + '</font>';
    }
    if (st.fill) x += `<fill><patternFill><bgColor rgb="${hex6(st.fill)}"/></patternFill></fill>`;
    x += '</dxf>';
    return intern(dxfs, dxfIdx, x);
  };

  // ----- sheets -----
  const sheetXmls: string[] = [];
  wb.sheets.forEach((sheet) => {
    const rect = sheet.usedRect();
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    if (sheet.tabColor) x += `<sheetPr><tabColor rgb="${hex6(sheet.tabColor)}"/></sheetPr>`;
    x += `<dimension ref="${rect ? rectText(rect) : 'A1'}"/>`;
    x += '<sheetViews><sheetView workbookViewId="0"' + (sheet.showGrid ? '' : ' showGridLines="0"') + (sheet === wb.sheets[0] ? ' tabSelected="1"' : '') + '>';
    if (sheet.freezeRows || sheet.freezeCols) {
      const tl = a1(sheet.freezeRows, sheet.freezeCols);
      const pane = sheet.freezeRows && sheet.freezeCols ? 'bottomRight' : sheet.freezeRows ? 'bottomLeft' : 'topRight';
      x += `<pane${sheet.freezeCols ? ` xSplit="${sheet.freezeCols}"` : ''}${sheet.freezeRows ? ` ySplit="${sheet.freezeRows}"` : ''} topLeftCell="${tl}" activePane="${pane}" state="frozen"/>`;
    }
    x += '</sheetView></sheetViews>';
    x += `<sheetFormatPr defaultRowHeight="${num(sheet.rows.def * 0.75)}"/>`;
    // columns: group consecutive equal widths
    const colIdxs = Array.from(sheet.cols.sizes.keys()).sort((a, b) => a - b);
    if (colIdxs.length) {
      x += '<cols>';
      let i = 0;
      while (i < colIdxs.length) {
        let j = i;
        while (j + 1 < colIdxs.length && colIdxs[j + 1] === colIdxs[j] + 1 && sheet.cols.sizes.get(colIdxs[j + 1]) === sheet.cols.sizes.get(colIdxs[i])) j++;
        const px = sheet.cols.sizes.get(colIdxs[i])!;
        if (px === 0) x += `<col min="${colIdxs[i] + 1}" max="${colIdxs[j] + 1}" width="9.140625" hidden="1" customWidth="1"/>`;
        else x += `<col min="${colIdxs[i] + 1}" max="${colIdxs[j] + 1}" width="${num(Math.round(((px - 5) / 7) * 1e6) / 1e6)}" customWidth="1"/>`;
        i = j + 1;
      }
      x += '</cols>';
    }
    // rows
    const keys = Array.from(sheet.cells.keys()).sort((a, b) => a - b);
    const rowSet = new Set<number>(sheet.rows.sizes.keys());
    for (const k of keys) rowSet.add(keyRow(k));
    const rowList = Array.from(rowSet).sort((a, b) => a - b);
    x += '<sheetData>';
    let ki = 0;
    for (const r of rowList) {
      const sz = sheet.rows.sizes.get(r);
      let attrs = ` r="${r + 1}"`;
      if (sz === 0) attrs += ' hidden="1"';
      else if (sz !== undefined) attrs += ` ht="${num(sz * 0.75)}" customHeight="1"`;
      let cx = '';
      while (ki < keys.length && keyRow(keys[ki]) < r) ki++;
      while (ki < keys.length && keyRow(keys[ki]) === r) {
        const k = keys[ki++];
        const cell = sheet.cells.get(k)!;
        const c = keyCol(k);
        const sAttr = cell.s ? ` s="${styleXf(cell.s)}"` : '';
        const ref = a1(r, c);
        const v = typeof cell.v === 'number' && !Number.isFinite(cell.v) ? E.NUM : cell.v;
        if (cell.f !== undefined) {
          const fx = `<f>${escText(addXlfn(cell.f))}</f>`;
          if (typeof v === 'number') cx += `<c r="${ref}"${sAttr}>${fx}<v>${num(v)}</v></c>`;
          else if (typeof v === 'string') cx += `<c r="${ref}"${sAttr} t="str">${fx}<v>${escText(v)}</v></c>`;
          else if (typeof v === 'boolean') cx += `<c r="${ref}"${sAttr} t="b">${fx}<v>${v ? 1 : 0}</v></c>`;
          else if (v instanceof FErr) cx += `<c r="${ref}"${sAttr} t="e">${fx}<v>${esc(v.code === '#CIRC!' ? '#REF!' : v.code)}</v></c>`;
          else cx += `<c r="${ref}"${sAttr}>${fx}</c>`;
        } else if (v === null) cx += `<c r="${ref}"${sAttr}/>`;
        else if (typeof v === 'number') cx += `<c r="${ref}"${sAttr}><v>${num(v)}</v></c>`;
        else if (typeof v === 'string') cx += `<c r="${ref}"${sAttr} t="s"><v>${sstOf(v)}</v></c>`;
        else if (typeof v === 'boolean') cx += `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
        else cx += `<c r="${ref}"${sAttr} t="e"><v>${esc(v.code)}</v></c>`;
      }
      x += cx ? `<row${attrs}>${cx}</row>` : `<row${attrs}/>`;
    }
    x += '</sheetData>';
    if (sheet.filter) x += `<autoFilter ref="${rectText(sheet.filter.rect)}"/>`;
    if (sheet.merges.length) x += `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${rectText(m)}"/>`).join('')}</mergeCells>`;
    let prio = 1;
    for (const rule of sheet.cf) {
      const sq = rectText(rule.rect);
      const st = rule.style ? ` dxfId="${dxfOf(rule.style)}"` : '';
      let body = '';
      switch (rule.type) {
        case 'cell': {
          const opMap: Record<string, string> = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '<>': 'notEqual', between: 'between', notbetween: 'notBetween' };
          body = `<cfRule type="cellIs"${st} priority="${prio++}" operator="${opMap[rule.op ?? '>']}"><formula>${escText(rule.v1 ?? '0')}</formula>${rule.op === 'between' || rule.op === 'notbetween' ? `<formula>${escText(rule.v2 ?? '0')}</formula>` : ''}</cfRule>`;
          break;
        }
        case 'formula':
          body = `<cfRule type="expression"${st} priority="${prio++}"><formula>${escText(rule.v1 ?? 'FALSE')}</formula></cfRule>`;
          break;
        case 'text': {
          const first = a1(rule.rect.r1, rule.rect.c1);
          const t = rule.v1 ?? '';
          body = `<cfRule type="containsText"${st} priority="${prio++}" operator="containsText" text="${esc(t)}"><formula>NOT(ISERROR(SEARCH("${esc(t.replace(/"/g, '""'))}",${first})))</formula></cfRule>`;
          break;
        }
        case 'dup':
          body = `<cfRule type="duplicateValues"${st} priority="${prio++}"/>`;
          break;
        case 'unique':
          body = `<cfRule type="uniqueValues"${st} priority="${prio++}"/>`;
          break;
        case 'scale': {
          const cols = rule.colors ?? ['#f8696b', '#63be7b'];
          const kinds = cols.length === 3 ? ['<cfvo type="min"/>', '<cfvo type="percentile" val="50"/>', '<cfvo type="max"/>'] : ['<cfvo type="min"/>', '<cfvo type="max"/>'];
          body = `<cfRule type="colorScale" priority="${prio++}"><colorScale>${kinds.join('')}${cols.map((c) => `<color rgb="${hex6(c)}"/>`).join('')}</colorScale></cfRule>`;
          break;
        }
        case 'bar':
          body = `<cfRule type="dataBar" priority="${prio++}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${hex6(rule.barColor ?? '#638ec6')}"/></dataBar></cfRule>`;
          break;
        default:
          break;
      }
      if (body) x += `<conditionalFormatting sqref="${sq}">${body}</conditionalFormatting>`;
    }
    x += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
    x += '</worksheet>';
    sheetXmls.push(x);
  });

  // styles.xml
  let styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">';
  if (numFmts.length) styles += `<numFmts count="${numFmts.length}">${numFmts.map((n) => `<numFmt numFmtId="${n.id}" formatCode="${esc(n.code)}"/>`).join('')}</numFmts>`;
  styles += `<fonts count="${fonts.length}">${fonts.join('')}</fonts>`;
  styles += `<fills count="${fills.length}">${fills.join('')}</fills>`;
  styles += `<borders count="${borders.length}">${borders.join('')}</borders>`;
  styles += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
  styles += `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`;
  styles += '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
  styles += `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`;
  styles += '</styleSheet>';

  const sstXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sstList.length + '" uniqueCount="' + sstList.length + '">' +
    sstList.map((s) => `<si><t xml:space="preserve">${escText(s)}</t></si>`).join('') + '</sst>';

  let wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>';
  wb.sheets.forEach((s, i) => {
    wbx += `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
  });
  wbx += '</sheets>';
  if (wb.names.size) wbx += '<definedNames>' + Array.from(wb.names.values()).map((n) => `<definedName name="${esc(n.name)}">${escText(addXlfn(n.refers))}</definedName>`).join('') + '</definedNames>';
  wbx += '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';

  const n = wb.sheets.length;
  let rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  for (let i = 0; i < n; i++) rels += `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`;
  rels += `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
  rels += `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;

  let ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
  for (let i = 0; i < n; i++) ct += `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
  ct += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>';

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
    'xl/workbook.xml': strToU8(wbx),
    'xl/_rels/workbook.xml.rels': strToU8(rels),
    'xl/styles.xml': strToU8(styles),
    'xl/sharedStrings.xml': strToU8(sstXml),
  };
  sheetXmls.forEach((x, i) => (files[`xl/worksheets/sheet${i + 1}.xml`] = strToU8(x)));
  return zipSync(files, { level: 6 });
}

export { colName, colIndex, MAX_ROWS };
export type { Rect };
