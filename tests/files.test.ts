import { describe, expect, it } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { Workbook } from '../src/core/model';
import { applyInput, applyStyle } from '../src/core/ops';
import { readXlsx, writeXlsx } from '../src/core/xlsx';
import { parseCSV, readCSV, writeCSV } from '../src/core/csv';
import { fromJSON, toJSON } from '../src/core/serialize';
import { parseXml } from '../src/core/xml';
import { FErr } from '../src/core/values';

function sample(): Workbook {
  const wb = new Workbook();
  const s = wb.sheets[0];
  wb.transact('x', () => {
    applyInput(wb, s, 0, 0, 'Name');
    applyInput(wb, s, 0, 1, 'Amount');
    applyInput(wb, s, 1, 0, 'Ann & "Bob" <x>');
    applyInput(wb, s, 1, 1, '1234.5');
    applyInput(wb, s, 2, 0, 'Total');
    applyInput(wb, s, 2, 1, '=SUM(B2:B2)*2');
    applyInput(wb, s, 3, 0, '3/15/2023');
    applyInput(wb, s, 3, 1, '=TEXTJOIN(",",TRUE,A1:A2)');
    applyInput(wb, s, 4, 0, '=1/0');
    applyInput(wb, s, 4, 1, 'TRUE');
    applyStyle(wb, s, { r1: 0, c1: 0, r2: 0, c2: 1 }, { bold: true, fill: '#ffee00', color: '#ff0000' });
    applyStyle(wb, s, { r1: 1, c1: 1, r2: 1, c2: 1 }, { fmt: '"$"#,##0.00', h: 'right', bb: { style: 'thin', color: '#000000' } });
    s.cols.set(0, 120);
    s.rows.set(1, 40);
    s.merges.push({ r1: 6, c1: 0, r2: 6, c2: 2 });
    s.freezeRows = 1;
    s.tabColor = '#00aa55';
    s.cf.push({ id: 1, rect: { r1: 1, c1: 1, r2: 4, c2: 1 }, type: 'cell', op: '>', v1: '100', style: { fill: '#ffc7ce', color: '#9c0006' } });
    const s2 = wb.addSheet('Second Sheet');
    applyInput(wb, s2, 0, 0, '=Sheet1!B2+1');
  });
  wb.setName('Rate', 'Sheet1!$B$2');
  return wb;
}

describe('xlsx round trip', () => {
  it('keeps values, formulas, formats and structure', () => {
    const wb = sample();
    const bytes = writeXlsx(wb);
    expect(bytes.length).toBeGreaterThan(500);
    const back = readXlsx(bytes);
    const s = back.sheets[0];
    expect(back.sheets.map((x) => x.name)).toEqual(['Sheet1', 'Second Sheet']);
    expect(s.get(1, 0)?.v).toBe('Ann & "Bob" <x>');
    expect(s.get(1, 1)?.v).toBe(1234.5);
    expect(s.get(2, 1)?.f).toBe('SUM(B2:B2)*2');
    expect(s.get(2, 1)?.v).toBe(2469);
    expect(s.get(3, 1)?.f).toBe('TEXTJOIN(",",TRUE,A1:A2)');
    expect(s.get(3, 0)?.v).toBe(45000);
    expect(back.styleOf(s.get(3, 0)?.s).fmt).toBe('m/d/yyyy');
    expect(s.get(4, 0)?.v instanceof FErr).toBe(true);
    expect(s.get(4, 1)?.v).toBe(true);
    const hs = back.styleOf(s.get(0, 0)?.s);
    expect(hs.bold).toBe(true);
    expect(hs.fill).toBe('#ffee00');
    expect(hs.color).toBe('#ff0000');
    const cs = back.styleOf(s.get(1, 1)?.s);
    expect(cs.fmt).toBe('"$"#,##0.00');
    expect(cs.h).toBe('right');
    expect(cs.bb?.style).toBe('thin');
    expect(s.cols.size(0)).toBe(120);
    expect(s.rows.size(1)).toBe(40);
    expect(s.merges).toEqual([{ r1: 6, c1: 0, r2: 6, c2: 2 }]);
    expect(s.freezeRows).toBe(1);
    expect(s.tabColor).toBe('#00aa55');
    expect(s.cf.length).toBe(1);
    expect(s.cf[0].op).toBe('>');
    expect(s.cf[0].style?.fill).toBe('#ffc7ce');
    expect(back.sheets[1].get(0, 0)?.v).toBe(1235.5);
    expect(back.getName('rate')?.refers).toBe('Sheet1!$B$2');
  });
  it('reads shared formulas, shared strings and rich text', () => {
    const sheet = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>
      <row r="3"><c r="A3" t="s"><v>0</v></c><c r="B3"><f t="shared" si="0"/></c></row></sheetData></worksheet>`;
    const files = {
      '[Content_Types].xml': strToU8('<Types/>'),
      'xl/workbook.xml': strToU8('<workbook xmlns:r="x"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'),
      'xl/_rels/workbook.xml.rels': strToU8('<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'),
      'xl/sharedStrings.xml': strToU8('<sst><si><r><t>Hel</t></r><r><t>lo</t></r></si></sst>'),
      'xl/worksheets/sheet1.xml': strToU8(sheet),
    };
    const wb = readXlsx(zipSync(files));
    const s = wb.sheets[0];
    expect(s.get(1, 1)?.f).toBe('A2*2');
    expect(s.get(2, 1)?.f).toBe('A3*2');
    expect(s.get(2, 0)?.v).toBe('Hello');
    expect(s.get(1, 1)?.v).toBe(4);
  });
  it('rejects non-xlsx data', () => {
    expect(() => readXlsx(new Uint8Array([1, 2, 3]))).toThrow();
  });
});

describe('csv', () => {
  it('parses quotes, delimiters and newlines', () => {
    expect(parseCSV('a,b\n1,"x,y"\n"q""r",\n')).toEqual([['a', 'b'], ['1', 'x,y'], ['q"r', '']]);
    expect(parseCSV('a;b\r\n1;2')).toEqual([['a', 'b'], ['1', '2']]);
    expect(parseCSV('"multi\nline",2')).toEqual([['multi\nline', '2']]);
  });
  it('imports typed values and exports display text', () => {
    const wb = readCSV('n,d,p\n5,3/15/2023,50%\n"a,b",=1+1,x\n');
    const s = wb.sheets[0];
    expect(s.get(1, 0)?.v).toBe(5);
    expect(s.get(1, 1)?.v).toBe(45000);
    expect(s.get(1, 2)?.v).toBe(0.5);
    expect(s.get(2, 1)?.v).toBe('=1+1');
    const out = writeCSV(wb, s);
    expect(out).toBe('﻿n,d,p\r\n5,3/15/2023,50%\r\n"a,b",=1+1,x\r\n');
  });
});

describe('native format', () => {
  it('round trips everything including errors and styles', () => {
    const wb = sample();
    const { wb: back } = fromJSON(toJSON(wb, 1));
    expect(back.sheets.length).toBe(2);
    expect(back.sheets[0].get(2, 1)?.v).toBe(2469);
    expect(back.sheets[0].get(4, 0)?.v instanceof FErr).toBe(true);
    expect(back.styleOf(back.sheets[0].get(0, 0)?.s).bold).toBe(true);
    expect(back.sheets[0].cf.length).toBe(1);
    expect(back.getName('rate')?.refers).toBe('Sheet1!$B$2');
  });
});

describe('xml', () => {
  it('parses entities, CDATA and self-closing tags', () => {
    const n = parseXml('<?xml version="1.0"?><a x="1&amp;2"><b/><c>t&lt;&#65;<![CDATA[<z>]]></c></a>');
    expect(n.attrs.x).toBe('1&2');
    expect(n.children[1].text).toBe('t<A<z>');
  });
});
