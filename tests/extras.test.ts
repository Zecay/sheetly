import { describe, expect, it } from 'vitest';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { CFEngine } from '../src/core/cf';
import { CFRule, Sheet, Workbook } from '../src/core/model';
import { applyInput, insertRows, moveExtras, moveRange, sortRange } from '../src/core/ops';
import { readXlsx, writeXlsx } from '../src/core/xlsx';

const R = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

function sorted3(): { wb: Workbook; s: Sheet } {
  const wb = new Workbook();
  const s = wb.sheets[0];
  [30, 10, 20].forEach((v, i) => applyInput(wb, s, i, 0, String(v)));
  return { wb, s };
}
const rule = (wb: Workbook, x: Partial<CFRule>): CFRule => ({ id: wb.nextRuleId++, rect: R(0, 0, 9, 0), type: 'cell', style: { fill: '#ff0000' }, ...x });
const hits = (wb: Workbook, s: Sheet, n = 10) => {
  const cf = new CFEngine(wb);
  const out: number[] = [];
  for (let r = 0; r < n; r++) if (cf.evaluate(s, r, 0)?.fill) out.push(r);
  return out;
};
const sheetXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']);

describe('notes, links and validation follow cells', () => {
  it('follow a sort, and undo restores them', () => {
    const { wb, s } = sorted3();
    s.comments.push({ r: 0, c: 0, author: 'a', text: 'thirty' });
    s.hyperlinks.push({ rect: R(1, 0), target: 'https://example.com/ten' });
    s.validations.push({ rects: [R(2, 0)], attrs: { type: 'whole' }, f1: '1' });
    sortRange(wb, s, R(0, 0, 2, 0), [{ col: 0, asc: true }], false);
    expect(s.get(2, 0)?.v).toBe(30);
    expect(s.comments[0]).toMatchObject({ r: 2, c: 0 });
    expect(s.hyperlinks[0].rect).toEqual(R(0, 0));
    expect(s.validations[0].rects).toEqual([R(1, 0)]);
    wb.undo();
    expect(s.get(0, 0)?.v).toBe(30);
    expect(s.comments[0].r).toBe(0);
    expect(s.hyperlinks[0].rect).toEqual(R(1, 0));
    expect(s.validations[0].rects).toEqual([R(2, 0)]);
    wb.redo();
    expect(s.comments[0].r).toBe(2);
  });

  it('a rule covering the whole sorted column is left as one range', () => {
    const { wb, s } = sorted3();
    s.validations.push({ rects: [R(0, 0, 99, 0)], attrs: { type: 'whole' }, f1: '1' });
    sortRange(wb, s, R(0, 0, 2, 0), [{ col: 0, asc: true }], false);
    expect(s.validations[0].rects).toEqual([R(0, 0, 99, 0)]);
  });

  it('keeps the header row in place', () => {
    const { wb, s } = sorted3();
    s.comments.push({ r: 0, c: 0, author: 'a', text: 'header' });
    sortRange(wb, s, R(0, 0, 2, 0), [{ col: 0, asc: true }], true);
    expect(s.comments[0].r).toBe(0);
  });

  it('follow cut and paste, replace what was at the destination, and undo', () => {
    const { wb, s } = sorted3();
    s.comments.push({ r: 0, c: 0, author: 'a', text: 'moves' }, { r: 5, c: 3, author: 'a', text: 'replaced' }, { r: 8, c: 8, author: 'a', text: 'stays' });
    s.hyperlinks.push({ rect: R(1, 0), location: 'Sheet1!A1' });
    s.validations.push({ rects: [R(0, 0, 3, 0)], attrs: { type: 'whole' }, f1: '1' });
    moveRange(wb, s, R(0, 0, 1, 0), s, 5, 3);
    expect(s.comments.map((n) => n.text).sort()).toEqual(['moves', 'stays']);
    expect(s.comments.find((n) => n.text === 'moves')).toMatchObject({ r: 5, c: 3 });
    expect(s.hyperlinks[0].rect).toEqual(R(6, 3));
    expect(s.hyperlinks[0].location).toBe('Sheet1!D6');
    expect(s.validations.flatMap((v) => v.rects)).toEqual(expect.arrayContaining([R(2, 0, 3, 0), R(5, 3, 6, 3)]));
    wb.undo();
    expect(s.comments).toHaveLength(3);
    expect(s.hyperlinks[0].rect).toEqual(R(1, 0));
    expect(s.validations[0].rects).toEqual([R(0, 0, 3, 0)]);
  });

  it('move to another sheet with formulas still pointing at the original sheet', () => {
    const { wb, s } = sorted3();
    const t = wb.addSheet('Other');
    applyInput(wb, s, 4, 0, '=A1+A2');
    applyInput(wb, s, 6, 0, '=A5');
    s.comments.push({ r: 4, c: 0, author: 'a', text: 'n' });
    moveRange(wb, s, R(4, 0), t, 0, 1);
    expect(s.comments).toHaveLength(0);
    expect(t.comments[0]).toMatchObject({ r: 0, c: 1 });
    expect(t.get(0, 1)?.v).toBe(40);
    expect(s.get(6, 0)?.f).toBe('Other!B1');
    expect(s.get(6, 0)?.v).toBe(40);
  });

  it('shift moves notes and removes those in deleted cells', () => {
    const { wb, s } = sorted3();
    s.comments.push({ r: 0, c: 0, author: 'a', text: 'gone' }, { r: 2, c: 0, author: 'a', text: 'up' });
    moveExtras(wb, s, R(1, 0, 1048575, 0), s, -1, 0);
    expect(s.comments).toHaveLength(1);
    expect(s.comments[0]).toMatchObject({ text: 'up', r: 1 });
  });
});

describe('conditional formats on the sheet', () => {
  it('moves whole rules and splits per-cell rules, keeping relative formulas', () => {
    const { wb, s } = sorted3();
    s.cf.push(rule(wb, { type: 'formula', rect: R(0, 0, 2, 0), v1: 'A1>15' }));
    moveRange(wb, s, R(2, 0), s, 2, 3);
    const moved = s.cf.find((x) => x.rect.c1 === 3)!;
    expect(moved.rect).toEqual(R(2, 3));
    expect(moved.v1).toBe('D3>15');
    expect(s.cf.find((x) => x.rect.c1 === 0)!.rect).toEqual(R(0, 0, 1, 0));
  });

  it('never splits a ranking rule', () => {
    const { wb, s } = sorted3();
    s.cf.push(rule(wb, { type: 'top', n: 1, rect: R(0, 0, 2, 0) }));
    moveRange(wb, s, R(2, 0), s, 2, 3);
    expect(s.cf).toHaveLength(1);
    expect(s.cf[0].rect).toEqual(R(0, 0, 2, 0));
  });

  it('formula references follow inserted rows', () => {
    const { wb, s } = sorted3();
    s.cf.push(rule(wb, { type: 'cell', op: '>', v1: '=$C$5', rect: R(0, 0, 2, 0) }));
    insertRows(wb, s, 0, 1);
    expect(s.cf[0].v1).toBe('=$C$6');
  });

  it('a cell-value formula is evaluated relative to each cell', () => {
    const wb = new Workbook();
    const s = wb.sheets[0];
    [5, 5].forEach((v, r) => applyInput(wb, s, r, 0, String(v)));
    applyInput(wb, s, 0, 1, '1');
    applyInput(wb, s, 1, 1, '9');
    s.cf.push(rule(wb, { op: '>', v1: '=B1', rect: R(0, 0, 1, 0) }));
    expect(hits(wb, s, 2)).toEqual([0]);
  });

  it('top percent, equal-to-average, standard deviation and stop-if-true', () => {
    const { wb, s } = sorted3();
    s.cf = [rule(wb, { type: 'top', n: 34, percent: true })];
    expect(hits(wb, s)).toEqual([0]);
    s.cf = [rule(wb, { type: 'above', equal: true })];
    expect(hits(wb, s)).toEqual([0, 2]);
    s.cf = [rule(wb, { type: 'above', stdDev: 2 })];
    expect(hits(wb, s)).toEqual([]);
    s.cf = [rule(wb, { type: 'above', style: { color: '#000000' }, stop: true }), rule(wb, { type: 'top', n: 3 })];
    expect(hits(wb, s)).toEqual([1, 2]); // row 0 matched the first rule, which stops the fill rule
  });
});

describe('conditional formats in .xlsx', () => {
  it('every rule type survives save and open and gives the same highlights', () => {
    const { wb, s } = sorted3();
    const rules: Partial<CFRule>[] = [
      { type: 'top', n: 2 }, { type: 'bottom', n: 1 }, { type: 'top', n: 25, percent: true }, { type: 'above' }, { type: 'below', equal: true }, { type: 'above', stdDev: 1 },
      { type: 'text', op: 'contains', v1: '1' }, { type: 'text', op: 'notcontains', v1: '1' }, { type: 'text', op: 'begins', v1: '2' }, { type: 'text', op: 'ends', v1: '0' },
      { type: 'cell', op: 'between', v1: '15', v2: '25' }, { type: 'formula', v1: '=A1>15' }, { type: 'dup' }, { type: 'unique' }, { type: 'above', stop: true },
    ];
    for (const x of rules) {
      s.cf = [rule(wb, x)];
      const want = hits(wb, s);
      const back = readXlsx(writeXlsx(wb));
      expect(back.sheets[0].cf).toHaveLength(1);
      expect(back.sheets[0].cf[0]).toMatchObject({ type: x.type, ...(x.n ? { n: x.n } : {}), ...(x.percent ? { percent: true } : {}), ...(x.stdDev ? { stdDev: x.stdDev } : {}), ...(x.stop ? { stop: true } : {}), ...(x.type === 'text' ? { op: x.op, v1: x.v1 } : {}) });
      expect(hits(back, back.sheets[0])).toEqual(want);
    }
  });

  it('writes operands the way Excel requires', () => {
    const { wb, s } = sorted3();
    s.cf = [rule(wb, { op: '=', v1: 'abc' }), rule(wb, { op: '>', v1: '=$B$1' }), rule(wb, { type: 'formula', v1: '=A1>0' }), rule(wb, { type: 'formula', v1: '=(((' })];
    const x = sheetXml(writeXlsx(wb));
    expect(x).toContain('<formula>&quot;abc&quot;</formula>');
    expect(x).toContain('<formula>$B$1</formula>');
    expect(x).toContain('<formula>A1&gt;0</formula>');
    expect(x).not.toContain('<formula>=');
    expect(x.match(/<cfRule/g)).toHaveLength(3);
  });

  it('reads priority, every range, reference operands and formula-based rule types', () => {
    const { wb, s } = sorted3();
    applyInput(wb, s, 0, 2, '25');
    const files = unzipSync(writeXlsx(wb));
    const cf =
      '<conditionalFormatting sqref="A1:A1 A3:A3"><cfRule type="expression" priority="2"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="A2"><cfRule type="cellIs" priority="1" operator="lessThan"><formula>$C$1</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="B1"><cfRule type="containsBlanks" priority="3"><formula>LEN(TRIM(B1))=0</formula></cfRule></conditionalFormatting>';
    files['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(files['xl/worksheets/sheet1.xml']).replace('<pageMargins', cf + '<pageMargins'));
    const back = readXlsx(zipSync(files)).sheets[0];
    expect(back.cf.map((x) => x.rect)).toEqual([R(1, 0), R(0, 0), R(2, 0), R(0, 1)]);
    expect(back.cf[0].v1).toBe('=$C$1');
    expect(back.cf[2].v1).toBe('A3>0');
    expect(back.cf[3].type).toBe('formula');
  });
});
