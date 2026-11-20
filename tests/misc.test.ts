import { describe, expect, it } from 'vitest';
import { Axis, Workbook } from '../src/core/model';
import { applyInput, applyStyle } from '../src/core/ops';
import { CFEngine } from '../src/core/cf';
import { colIndex, colName, parseCell, parseRect, rectText } from '../src/core/addr';
import { MAX_COLS, MAX_ROWS } from '../src/core/addr';
import { refSpans, cycleReference, callAt } from '../src/ui/editor';

describe('addresses', () => {
  it('converts column names both ways up to XFD', () => {
    expect(colName(0)).toBe('A');
    expect(colName(25)).toBe('Z');
    expect(colName(26)).toBe('AA');
    expect(colName(16383)).toBe('XFD');
    expect(colIndex('XFD')).toBe(16383);
    expect(colIndex('aa')).toBe(26);
    expect(MAX_COLS).toBe(16384);
    expect(MAX_ROWS).toBe(1048576);
  });
  it('parses cells and ranges', () => {
    expect(parseCell('B3')).toEqual({ r: 2, c: 1 });
    expect(parseCell('XFE1')).toBeNull();
    expect(parseCell('A1048577')).toBeNull();
    expect(parseRect('$A$1:C3')).toEqual({ r1: 0, c1: 0, r2: 2, c2: 2 });
    expect(parseRect('B:C')).toEqual({ r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 2 });
    expect(rectText({ r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 2 })).toBe('B:C');
  });
});

describe('row/column axis with custom sizes', () => {
  it('maps indexes to positions and back across a million rows', () => {
    const a = new Axis(20);
    a.set(5, 40);
    a.set(10, 0); // hidden
    a.set(1000000, 100);
    expect(a.top(0)).toBe(0);
    expect(a.top(5)).toBe(100);
    expect(a.top(6)).toBe(140);
    expect(a.top(11)).toBe(220);
    expect(a.at(0, MAX_ROWS)).toBe(0);
    expect(a.at(105, MAX_ROWS)).toBe(5);
    expect(a.at(141, MAX_ROWS)).toBe(6);
    // position of the hidden row start belongs to the next visible row
    expect(a.at(a.top(10), MAX_ROWS)).toBe(11);
    expect(a.total(MAX_ROWS)).toBe(MAX_ROWS * 20 + 80);
    const p = a.top(999999);
    expect(a.at(p + 1, MAX_ROWS)).toBe(999999);
  });
});

describe('conditional formatting engine', () => {
  const make = () => {
    const wb = new Workbook();
    const s = wb.sheets[0];
    wb.transact('d', () => {
      for (let i = 0; i < 5; i++) applyInput(wb, s, i, 0, String((i + 1) * 10));
      applyInput(wb, s, 5, 0, 'apple');
      applyInput(wb, s, 6, 0, 'Apple');
    });
    return { wb, s, cf: new CFEngine(wb) };
  };
  it('cellIs rules', () => {
    const { wb, s, cf } = make();
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 6, c2: 0 }, type: 'cell', op: '>', v1: '25', style: { fill: '#ff0000' } });
    expect(cf.evaluate(s, 0, 0)).toBeNull();
    expect(cf.evaluate(s, 2, 0)?.fill).toBe('#ff0000');
    expect(cf.evaluate(s, 5, 0)).toBeNull();
  });
  it('between, top, duplicates and text', () => {
    const { wb, s, cf } = make();
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 6, c2: 0 }, type: 'cell', op: 'between', v1: '15', v2: '35', style: { fill: '#00ff00' } });
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 4, c2: 0 }, type: 'top', n: 2, style: { color: '#0000ff' } });
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 5, c1: 0, r2: 6, c2: 0 }, type: 'dup', style: { bold: true } });
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 5, c1: 0, r2: 6, c2: 0 }, type: 'text', op: 'contains', v1: 'PPL', style: { italic: true } });
    expect(cf.evaluate(s, 1, 0)?.fill).toBe('#00ff00');
    expect(cf.evaluate(s, 4, 0)?.color).toBe('#0000ff');
    expect(cf.evaluate(s, 3, 0)?.color).toBe('#0000ff');
    expect(cf.evaluate(s, 2, 0)?.color).toBeUndefined();
    expect(cf.evaluate(s, 5, 0)?.bold).toBe(true);
    expect(cf.evaluate(s, 6, 0)?.italic).toBe(true);
  });
  it('color scales and data bars', () => {
    const { wb, s, cf } = make();
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 4, c2: 0 }, type: 'scale', colors: ['#000000', '#ffffff'] });
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 4, c2: 0 }, type: 'bar', barColor: '#123456' });
    expect(cf.evaluate(s, 0, 0)?.fill).toBe('#000000');
    expect(cf.evaluate(s, 4, 0)?.fill).toBe('#ffffff');
    expect(cf.evaluate(s, 2, 0)?.fill).toBe('#808080');
    expect(cf.evaluate(s, 4, 0)?.bar?.pct).toBe(1);
    expect(cf.evaluate(s, 0, 0)?.bar?.pct).toBeCloseTo(0.2, 5);
  });
  it('formula rules use relative references', () => {
    const { wb, s, cf } = make();
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 4, c2: 0 }, type: 'formula', v1: '=MOD(ROW(),2)=0', style: { fill: '#eeeeee' } });
    expect(cf.evaluate(s, 0, 0)).toBeNull();
    expect(cf.evaluate(s, 1, 0)?.fill).toBe('#eeeeee');
  });
  it('recomputes statistics after data changes', () => {
    const { wb, s, cf } = make();
    s.cf.push({ id: wb.nextRuleId++, rect: { r1: 0, c1: 0, r2: 4, c2: 0 }, type: 'scale', colors: ['#000000', '#ffffff'] });
    expect(cf.evaluate(s, 4, 0)?.fill).toBe('#ffffff');
    wb.transact('e', () => applyInput(wb, s, 4, 0, '1000'));
    expect(cf.evaluate(s, 3, 0)?.fill).not.toBe('#ffffff');
  });
});

describe('workbook history for formatting and names', () => {
  it('undoes style changes', () => {
    const wb = new Workbook();
    const s = wb.sheets[0];
    wb.transact('v', () => applyInput(wb, s, 0, 0, 'x'));
    wb.transact('bold', () => applyStyle(wb, s, { r1: 0, c1: 0, r2: 0, c2: 0 }, { bold: true }));
    expect(wb.styleOf(s.get(0, 0)?.s).bold).toBe(true);
    wb.undo();
    expect(wb.styleOf(s.get(0, 0)?.s).bold).toBeUndefined();
    wb.redo();
    expect(wb.styleOf(s.get(0, 0)?.s).bold).toBe(true);
  });
  it('undoes column width and merge changes (meta transactions)', () => {
    const wb = new Workbook();
    const s = wb.sheets[0];
    wb.transact('w', () => {
      s.cols.set(0, 200);
      s.merges.push({ r1: 0, c1: 0, r2: 1, c2: 1 });
    }, true);
    expect(s.cols.size(0)).toBe(200);
    wb.undo();
    expect(s.cols.size(0)).toBe(64);
    expect(s.merges.length).toBe(0);
    wb.redo();
    expect(s.merges.length).toBe(1);
  });
});

describe('formula editor helpers', () => {
  it('assigns one colour per distinct reference', () => {
    const spans = refSpans('=A1+B2*A1+Sheet2!C3');
    expect(spans.map((x) => x.text)).toEqual(['A1', 'B2', 'A1', 'Sheet2!C3']);
    expect(spans[0].color).toBe(spans[2].color);
    expect(spans[0].color).not.toBe(spans[1].color);
  });
  it('cycles absolute and relative references with F4', () => {
    let t = '=A1+1';
    const steps: string[] = [];
    for (let i = 0; i < 4; i++) {
      const r = cycleReference(t, 3)!;
      t = r.text;
      steps.push(t);
    }
    expect(steps).toEqual(['=$A$1+1', '=A$1+1', '=$A1+1', '=A1+1']);
  });
  it('finds the function call and argument at the caret', () => {
    expect(callAt('=SUM(A1,', 8)).toEqual({ name: 'SUM', idx: 1 });
    expect(callAt('=IF(A1>1,SUM(B1,', 16)).toEqual({ name: 'SUM', idx: 1 });
    expect(callAt('=IF(A1>1,SUM(B1,B2),', 21)).toEqual({ name: 'IF', idx: 2 });
    expect(callAt('=1+2', 4)).toBeNull();
  });
});
