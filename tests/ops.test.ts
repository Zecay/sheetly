import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/model';
import {
  applyInput, copyRange, deleteCols, deleteRows, fillRange, insertCols, insertRows, moveRange, parseTSV, pasteClip, pasteText,
  renameSheet, sortRange, deleteSheet, findAll,
} from '../src/core/ops';
import { parseCell } from '../src/core/addr';
import { FErr, Value } from '../src/core/values';
import { changeDecimals, formatValue } from '../src/core/numfmt';
import { parseInput, editText } from '../src/core/input';

function mk(cells: Record<string, string>) {
  const wb = new Workbook();
  wb.transact('init', () => {
    for (const [k, v] of Object.entries(cells)) {
      const p = parseCell(k)!;
      applyInput(wb, wb.sheets[0], p.r, p.c, v);
    }
  });
  return wb;
}
const g = (wb: Workbook, ref: string, si = 0): Value => {
  const p = parseCell(ref)!;
  const v = wb.getValue(wb.sheets[si], p.r, p.c);
  return v instanceof FErr ? v.code : v;
};
const f = (wb: Workbook, ref: string, si = 0) => {
  const p = parseCell(ref)!;
  return wb.sheets[si].get(p.r, p.c)?.f;
};

describe('number formats', () => {
  const fmt = (v: number | string, code: string) => formatValue(v, code).text;
  it('general', () => {
    expect(fmt(1, 'General')).toBe('1');
    expect(fmt(3.14159265358979, 'General')).toBe('3.141592654');
    expect(fmt(123456.789012345, 'General')).toBe('123456.789');
    expect(fmt(0.000123456789, 'General')).toBe('0.000123457');
    expect(fmt(12345678901234, 'General')).toBe('1.23457E+13');
    expect(fmt(-0.5, 'General')).toBe('-0.5');
  });
  it('fixed decimals and grouping', () => {
    expect(fmt(1234.5, '0.00')).toBe('1234.50');
    expect(fmt(1234.5, '#,##0')).toBe('1,235');
    expect(fmt(1234567.891, '#,##0.00')).toBe('1,234,567.89');
    expect(fmt(0.5, '#.00')).toBe('.50');
    expect(fmt(5, '000')).toBe('005');
    expect(fmt(-1234.5, '#,##0.00')).toBe('-1,234.50');
    expect(fmt(2.5, '0.#')).toBe('2.5');
    expect(fmt(2, '0.0#')).toBe('2.0');
  });
  it('percent, currency, scientific', () => {
    expect(fmt(0.256, '0%')).toBe('26%');
    expect(fmt(0.256, '0.0%')).toBe('25.6%');
    expect(fmt(1234.5, '"$"#,##0.00')).toBe('$1,234.50');
    expect(fmt(-1234.5, '"$"#,##0.00')).toBe('-$1,234.50');
    expect(fmt(-1234.5, '$#,##0.00;($#,##0.00)')).toBe('($1,234.50)');
    expect(fmt(12345.6789, '0.00E+00')).toBe('1.23E+04');
    expect(fmt(0.00012, '0.00E+00')).toBe('1.20E-04');
  });
  it('sections, conditions and text', () => {
    expect(fmt(5, '0;-0;"zero"')).toBe('5');
    expect(fmt(0, '0;-0;"zero"')).toBe('zero');
    expect(fmt(-5, '0;(0)')).toBe('(5)');
    expect(formatValue(-5, '0;[Red]0').color).toBeTruthy();
    expect(fmt(150, '[>100]"big";"small"')).toBe('big');
    expect(fmt(50, '[>100]"big";"small"')).toBe('small');
    expect(fmt('hi', '@')).toBe('hi');
    expect(fmt('hi', '"<"@">"')).toBe('<hi>');
  });
  it('dates and times', () => {
    expect(fmt(45000, 'm/d/yyyy')).toBe('3/15/2023');
    expect(fmt(45000, 'yyyy-mm-dd')).toBe('2023-03-15');
    expect(fmt(45000, 'd-mmm-yy')).toBe('15-Mar-23');
    expect(fmt(45000, 'dddd, mmmm d, yyyy')).toBe('Wednesday, March 15, 2023');
    expect(fmt(45000.5, 'h:mm AM/PM')).toBe('12:00 PM');
    expect(fmt(45000.75, 'h:mm:ss')).toBe('18:00:00');
    expect(fmt(0.0625, 'hh:mm')).toBe('01:30');
    expect(fmt(1.5, '[h]:mm')).toBe('36:00');
    expect(fmt(45000.25, 'm/d/yyyy h:mm')).toBe('3/15/2023 6:00');
  });
  it('accounting keeps a fill marker', () => {
    const t = fmt(1234.5, '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)');
    expect(t.replace(//g, '|')).toBe(' $|1,234.50 ');
  });
  it('changeDecimals', () => {
    expect(changeDecimals('General', 1)).toBe('0.0');
    expect(changeDecimals('0.00', 1)).toBe('0.000');
    expect(changeDecimals('0.00', -1)).toBe('0.0');
    expect(changeDecimals('0.0', -1)).toBe('0');
    expect(changeDecimals('#,##0.00', -1)).toBe('#,##0.0');
  });
});

describe('typed input', () => {
  it('detects types', () => {
    expect(parseInput('42').v).toBe(42);
    expect(parseInput('TRUE').v).toBe(true);
    expect(parseInput("'123").v).toBe('123');
    expect(parseInput('=A1+1').f).toBe('A1+1');
    expect(parseInput('hello').v).toBe('hello');
    expect(parseInput('12.5%')).toMatchObject({ v: 0.125, fmt: '0.0%' });
    expect(parseInput('$1,234.50')).toMatchObject({ v: 1234.5, fmt: '"$"#,##0.00' });
    expect(parseInput('1,234')).toMatchObject({ v: 1234, fmt: '#,##0' });
    expect(parseInput('3/15/2023')).toMatchObject({ v: 45000, fmt: 'm/d/yyyy' });
    expect(parseInput('1:30 PM')).toMatchObject({ fmt: 'h:mm AM/PM' });
  });
  it('round-trips editor text', () => {
    const wb = mk({ A1: '12.5%', B1: '3/15/2023', C1: "'007", D1: '=1+1' });
    const s = wb.sheets[0];
    expect(editText(s.get(0, 0), wb.styleOf(s.get(0, 0)!.s))).toBe('12.5%');
    expect(editText(s.get(0, 1), wb.styleOf(s.get(0, 1)!.s))).toBe('3/15/2023');
    expect(editText(s.get(0, 2))).toBe("'007");
    expect(editText(s.get(0, 3))).toBe('=1+1');
  });
});

describe('insert / delete', () => {
  it('insert rows shifts cells and references', () => {
    const wb = mk({ A1: '1', A2: '2', A3: '=SUM(A1:A2)', B1: '=A2', C1: '=$A$2' });
    insertRows(wb, wb.sheets[0], 1, 1);
    expect(g(wb, 'A3')).toBe(2);
    expect(g(wb, 'A4')).toBe(3);
    expect(f(wb, 'A4')).toBe('SUM(A1:A3)');
    expect(f(wb, 'B1')).toBe('A3');
    expect(f(wb, 'C1')).toBe('$A$3');
    expect(g(wb, 'B1')).toBe(2);
  });
  it('delete rows shrinks ranges and produces #REF!', () => {
    const wb = mk({ A1: '1', A2: '2', A3: '3', B1: '=SUM(A1:A3)', B2: '=A2', B3: '=A3' });
    deleteRows(wb, wb.sheets[0], 1, 1);
    expect(f(wb, 'B1')).toBe('SUM(A1:A2)');
    expect(g(wb, 'B1')).toBe(4);
    expect(g(wb, 'B2')).toBe(3);
    expect(f(wb, 'B2')).toBe('A2');
    const wb2 = mk({ A1: '1', A2: '2', B1: '=A2' });
    deleteRows(wb2, wb2.sheets[0], 1, 1);
    expect(f(wb2, 'B1')).toBe('#REF!');
    expect(g(wb2, 'B1')).toBe('#REF!');
  });
  it('insert / delete columns', () => {
    const wb = mk({ A1: '1', B1: '2', C1: '=A1+B1' });
    insertCols(wb, wb.sheets[0], 1, 2);
    expect(f(wb, 'E1')).toBe('A1+D1');
    expect(g(wb, 'E1')).toBe(3);
    deleteCols(wb, wb.sheets[0], 1, 2);
    expect(f(wb, 'C1')).toBe('A1+B1');
    expect(g(wb, 'C1')).toBe(3);
  });
  it('updates references from other sheets and names', () => {
    const wb = mk({ A1: '1', A2: '2' });
    const s2 = wb.addSheet('Other');
    wb.transact('x', () => applyInput(wb, s2, 0, 0, '=Sheet1!A2*10'));
    wb.setName('N', 'Sheet1!$A$2');
    insertRows(wb, wb.sheets[0], 0, 1);
    expect(s2.get(0, 0)?.f).toBe('Sheet1!A3*10');
    expect(wb.getName('n')?.refers).toBe('Sheet1!$A$3');
    expect(wb.getValue(s2, 0, 0)).toBe(20);
  });
  it('undo and redo of a structural change', () => {
    const wb = mk({ A1: '1', A2: '=A1+1' });
    insertRows(wb, wb.sheets[0], 0, 1);
    expect(g(wb, 'A3')).toBe(2);
    wb.undo();
    expect(g(wb, 'A2')).toBe(2);
    expect(f(wb, 'A2')).toBe('A1+1');
    wb.redo();
    expect(f(wb, 'A3')).toBe('A2+1');
    expect(g(wb, 'A3')).toBe(2);
  });
  it('sheet rename and delete fix formulas', () => {
    const wb = mk({ A1: '5' });
    const s2 = wb.addSheet('Two');
    wb.transact('x', () => applyInput(wb, s2, 0, 0, '=Sheet1!A1+1'));
    renameSheet(wb, wb.sheets[0], 'My Data');
    expect(s2.get(0, 0)?.f).toBe("'My Data'!A1+1");
    expect(wb.getValue(s2, 0, 0)).toBe(6);
    deleteSheet(wb, wb.sheets[0]);
    expect(wb.getValue(wb.sheets[0], 0, 0) instanceof FErr).toBe(true);
  });
});

describe('undo / redo', () => {
  it('undoes and redoes cell edits with recalculation', () => {
    const wb = mk({ A1: '1', B1: '=A1*2' });
    wb.transact('edit', () => applyInput(wb, wb.sheets[0], 0, 0, '10'));
    expect(g(wb, 'B1')).toBe(20);
    wb.undo();
    expect(g(wb, 'A1')).toBe(1);
    expect(g(wb, 'B1')).toBe(2);
    wb.redo();
    expect(g(wb, 'B1')).toBe(20);
  });
  it('keeps at least 100 steps', () => {
    const wb = new Workbook();
    for (let i = 0; i < 150; i++) wb.transact('e' + i, () => applyInput(wb, wb.sheets[0], 0, 0, String(i)));
    let n = 0;
    while (wb.canUndo()) {
      wb.undo();
      n++;
    }
    expect(n).toBeGreaterThanOrEqual(100);
  });
  it('a new edit clears the redo stack', () => {
    const wb = mk({ A1: '1' });
    wb.transact('a', () => applyInput(wb, wb.sheets[0], 0, 0, '2'));
    wb.undo();
    wb.transact('b', () => applyInput(wb, wb.sheets[0], 0, 0, '3'));
    expect(wb.canRedo()).toBe(false);
  });
});

describe('fill', () => {
  const fillDown = (cells: Record<string, string>, src: string, n: number) => {
    const wb = mk(cells);
    const a = src.split(':');
    const p1 = parseCell(a[0])!;
    const p2 = parseCell(a[1] ?? a[0])!;
    wb.transact('fill', () => fillRange(wb, wb.sheets[0], { r1: p1.r, c1: p1.c, r2: p2.r, c2: p2.c }, 'down', n));
    return wb;
  };
  it('extends number series', () => {
    const wb = fillDown({ A1: '1', A2: '2' }, 'A1:A2', 3);
    expect([g(wb, 'A3'), g(wb, 'A4'), g(wb, 'A5')]).toEqual([3, 4, 5]);
    const wb2 = fillDown({ A1: '10', A2: '20' }, 'A1:A2', 2);
    expect([g(wb2, 'A3'), g(wb2, 'A4')]).toEqual([30, 40]);
  });
  it('copies a single number', () => {
    const wb = fillDown({ A1: '7' }, 'A1', 2);
    expect([g(wb, 'A2'), g(wb, 'A3')]).toEqual([7, 7]);
  });
  it('extends day and month names', () => {
    const wb = fillDown({ A1: 'Mon' }, 'A1', 3);
    expect([g(wb, 'A2'), g(wb, 'A3'), g(wb, 'A4')]).toEqual(['Tue', 'Wed', 'Thu']);
    const wb2 = fillDown({ A1: 'Nov' }, 'A1', 3);
    expect([g(wb2, 'A2'), g(wb2, 'A3'), g(wb2, 'A4')]).toEqual(['Dec', 'Jan', 'Feb']);
    const wb3 = fillDown({ A1: 'January' }, 'A1', 1);
    expect(g(wb3, 'A2')).toBe('February');
  });
  it('increments text with a trailing number', () => {
    const wb = fillDown({ A1: 'Item 1' }, 'A1', 2);
    expect([g(wb, 'A2'), g(wb, 'A3')]).toEqual(['Item 2', 'Item 3']);
  });
  it('shifts relative references in formulas', () => {
    const wb = fillDown({ A1: '1', A2: '2', A3: '3', B1: '=A1*2+$A$1' }, 'B1', 2);
    expect(f(wb, 'B2')).toBe('A2*2+$A$1');
    expect(f(wb, 'B3')).toBe('A3*2+$A$1');
    expect(g(wb, 'B3')).toBe(7);
  });
  it('fills right and up', () => {
    const wb = mk({ A1: '1', B1: '2' });
    wb.transact('f', () => fillRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 0, c2: 1 }, 'right', 2));
    expect([g(wb, 'C1'), g(wb, 'D1')]).toEqual([3, 4]);
    const wb2 = mk({ A5: '5', A6: '6' });
    wb2.transact('f', () => fillRange(wb2, wb2.sheets[0], { r1: 4, c1: 0, r2: 5, c2: 0 }, 'up', 2));
    expect([g(wb2, 'A4'), g(wb2, 'A3')]).toEqual([4, 3]);
  });
});

describe('sort', () => {
  it('sorts rows ascending and descending, blanks last', () => {
    const wb = mk({ A1: 'b', B1: '2', A2: 'a', B2: '3', A3: 'c', B3: '1', A4: '', B4: '9' });
    sortRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 0, asc: true }], false);
    expect([g(wb, 'A1'), g(wb, 'A2'), g(wb, 'A3')]).toEqual(['a', 'b', 'c']);
    expect(g(wb, 'B1')).toBe(3);
    expect(g(wb, 'B4')).toBe(9);
    sortRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 3, c2: 1 }, [{ col: 1, asc: false }], false);
    expect([g(wb, 'B1'), g(wb, 'B2'), g(wb, 'B3'), g(wb, 'B4')]).toEqual([9, 3, 2, 1]);
  });
  it('keeps the header row and adjusts same-row formulas', () => {
    const wb = mk({ A1: 'n', A2: '3', B2: '=A2*10', A3: '1', B3: '=A3*10' });
    sortRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 2, c2: 1 }, [{ col: 0, asc: true }], true);
    expect(g(wb, 'A1')).toBe('n');
    expect(g(wb, 'A2')).toBe(1);
    expect(g(wb, 'B2')).toBe(10);
    expect(f(wb, 'B2')).toBe('A2*10');
  });
  it('sorts numbers before text', () => {
    const wb = mk({ A1: 'x', A2: '5', A3: '2' });
    sortRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 2, c2: 0 }, [{ col: 0, asc: true }], false);
    expect([g(wb, 'A1'), g(wb, 'A2'), g(wb, 'A3')]).toEqual([2, 5, 'x']);
  });
});

describe('clipboard', () => {
  it('parses TSV with quotes and newlines', () => {
    expect(parseTSV('a\tb\r\n1\t2\r\n')).toEqual([['a', 'b'], ['1', '2']]);
    expect(parseTSV('"x\ny"\t"q""r"\n')).toEqual([['x\ny', 'q"r']]);
    expect(parseTSV('a\t\tb')).toEqual([['a', '', 'b']]);
  });
  it('copies as TSV and pastes with shifted formulas', () => {
    const wb = mk({ A1: '1', B1: '=A1+1', A2: 'x' });
    const clip = copyRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 1, c2: 1 });
    expect(clip.tsv).toBe('1\t2\r\nx\t\r\n');
    wb.transact('p', () => pasteClip(wb, wb.sheets[0], 4, 0, clip, wb));
    expect(f(wb, 'B5')).toBe('A5+1');
    expect(g(wb, 'B5')).toBe(2);
    expect(g(wb, 'A6')).toBe('x');
  });
  it('paste values only', () => {
    const wb = mk({ A1: '=1+2' });
    const clip = copyRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 0, c2: 0 });
    wb.transact('p', () => pasteClip(wb, wb.sheets[0], 3, 3, clip, wb, 'values'));
    expect(g(wb, 'D4')).toBe(3);
    expect(f(wb, 'D4')).toBeUndefined();
  });
  it('pastes text from another app', () => {
    const wb = new Workbook();
    wb.transact('p', () => pasteText(wb, wb.sheets[0], 1, 1, 'a\t1\r\nb\t=C2*2\r\n'));
    expect(g(wb, 'B2')).toBe('a');
    expect(g(wb, 'C2')).toBe(1);
    expect(g(wb, 'C3')).toBe(2);
  });
  it('cut and paste updates references to moved cells', () => {
    const wb = mk({ A1: '5', B1: '=A1*2', C1: '=A1' });
    wb.transact('mv', () => moveRange(wb, wb.sheets[0], { r1: 0, c1: 0, r2: 0, c2: 0 }, wb.sheets[0], 4, 0));
    expect(g(wb, 'A5')).toBe(5);
    expect(g(wb, 'A1')).toBe(null);
    expect(f(wb, 'B1')).toBe('A5*2');
    expect(g(wb, 'B1')).toBe(10);
  });
});

describe('find', () => {
  it('finds by value, case and whole cell', () => {
    const wb = mk({ A1: 'Apple', A2: 'pineapple', B1: '=1+1', B2: '2' });
    const o = { what: 'apple', matchCase: false, wholeCell: false, inFormulas: false, allSheets: false, byRows: true };
    expect(findAll(wb, wb.sheets[0], o).map((x) => [x.r, x.c])).toEqual([[0, 0], [1, 0]]);
    expect(findAll(wb, wb.sheets[0], { ...o, matchCase: true }).length).toBe(1);
    expect(findAll(wb, wb.sheets[0], { ...o, wholeCell: true }).length).toBe(1);
    expect(findAll(wb, wb.sheets[0], { ...o, what: '2' }).length).toBe(2);
    expect(findAll(wb, wb.sheets[0], { ...o, what: '=1', inFormulas: true }).length).toBe(1);
  });
});
