// Stress 5a: weird inputs. Huge strings, deep nesting, long expressions, circular refs, full-sheet ranges, #REF! after delete.
import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/core/model';
import { FormulaError, deleteCols, deleteRows } from '../../src/core/ops';
import { FErr } from '../../src/core/values';
import { get, put } from './helpers';

function sheet() {
  const wb = new Workbook();
  return { wb, sh: wb.sheets[0] };
}

describe('huge strings', () => {
  it('1 MB text cell and a 1 MB string literal inside a formula', () => {
    const { wb, sh } = sheet();
    const big = 'x'.repeat(1_000_000);
    const t0 = performance.now();
    put(wb, sh, 'A1', big);
    put(wb, sh, 'A2', '=LEN(A1)');
    put(wb, sh, 'A3', `="${big}"`);
    put(wb, sh, 'A4', '=LEN(A3)');
    const ms = performance.now() - t0;
    console.log(`[stress] 1 MB string cell + formula literal: ${ms.toFixed(0)} ms`);
    expect(get(wb, sh, 'A2')).toBe(1_000_000);
    expect(get(wb, sh, 'A4')).toBe(1_000_000);
    expect(ms).toBeLessThan(5000);
  }, 60_000);
});

describe('deep nesting and long expressions', () => {
  it('300 nested parentheses evaluate to 1', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'A1', '=' + '('.repeat(300) + '1' + ')'.repeat(300));
    expect(get(wb, sh, 'A1')).toBe(1);
  });

  it('5000 nested parentheses: either accepted or a FormulaError, never a raw RangeError', () => {
    const { wb, sh } = sheet();
    let caught: unknown = null;
    try {
      put(wb, sh, 'A1', '=' + '('.repeat(5000) + '1' + ')'.repeat(5000));
    } catch (e) {
      caught = e;
    }
    if (caught !== null) expect(caught, 'raw exception from parser').toBeInstanceOf(FormulaError);
  });

  it('5000 unary minus signs: either accepted or a FormulaError, never a raw RangeError', () => {
    const { wb, sh } = sheet();
    let caught: unknown = null;
    try {
      put(wb, sh, 'A1', '=' + '-'.repeat(5000) + '1');
    } catch (e) {
      caught = e;
    }
    if (caught !== null) expect(caught, 'raw exception from parser').toBeInstanceOf(FormulaError);
  });

  it('64 nested ABS() calls (Excel limit) evaluate to 1', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'A1', '=' + 'ABS('.repeat(64) + '1' + ')'.repeat(64));
    expect(get(wb, sh, 'A1')).toBe(1);
  });

  it('a 300-term sum (=1+1+...+1) evaluates to 300', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'A1', '=' + Array(300).fill('1').join('+'));
    expect(get(wb, sh, 'A1')).toBe(300);
  });
});

describe('circular references', () => {
  it('self loop, 2-cell loop, range loop, and dependents all give #CIRC!; editing inside a loop does not hang', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'A1', '=A1');
    put(wb, sh, 'D1', '=A1+1');
    put(wb, sh, 'A2', '=B2');
    put(wb, sh, 'B2', '=A2');
    put(wb, sh, 'C2', '=A2+1');
    put(wb, sh, 'A3', '=SUM(A3:A5)');
    put(wb, sh, 'A4', '1');
    put(wb, sh, 'A5', '2');
    for (const ref of ['A1', 'D1', 'A2', 'B2', 'C2', 'A3']) expect(get(wb, sh, ref), ref).toBe('#CIRC!');

    const t0 = performance.now();
    for (let i = 0; i < 200; i++) put(wb, sh, 'A5', String(i));
    expect(performance.now() - t0).toBeLessThan(5000);
    expect(get(wb, sh, 'A3')).toBe('#CIRC!');

    // break the loops and check recovery
    put(wb, sh, 'A1', '5');
    expect(get(wb, sh, 'D1')).toBe(6);
    put(wb, sh, 'B2', '3');
    expect(get(wb, sh, 'A2')).toBe(3);
    expect(get(wb, sh, 'C2')).toBe(4);
  });
});

// Formulas live on a second sheet: a formula inside A1:XFD1048576 on the same sheet is circular by design.
function twoSheets() {
  const wb = new Workbook();
  const data = wb.sheets[0]; // Sheet1
  const calc = wb.addSheet('Calc');
  return { wb, data, calc };
}

describe('full-sheet ranges (A1:XFD1048576)', () => {
  it('SUM, COUNTIF and COUNT(A:A) over the whole sheet on a sparse sheet are correct', () => {
    const { wb, data, calc } = twoSheets();
    put(wb, data, 'A1', '1');
    put(wb, data, 'A3', '2');
    put(wb, data, 'Z500', '4');
    const t0 = performance.now();
    put(wb, calc, 'C1', '=SUM(Sheet1!A1:XFD1048576)');
    put(wb, calc, 'D1', '=COUNTIF(Sheet1!A1:XFD1048576,">0")');
    put(wb, calc, 'E1', '=COUNT(Sheet1!A:A)');
    const ms = performance.now() - t0;
    expect(get(wb, calc, 'C1')).toBe(7);
    expect(get(wb, calc, 'D1')).toBe(3);
    expect(get(wb, calc, 'E1')).toBe(2);
    expect(ms).toBeLessThan(1000);
  });

  it('COUNTIF over A1:XFD1048576 with one cell at XFD1000 (far corner) stays fast', () => {
    const { wb, data, calc } = twoSheets();
    put(wb, data, 'A1', '1');
    put(wb, data, 'XFD1000', '1');
    const t0 = performance.now();
    put(wb, calc, 'D1', '=COUNTIF(Sheet1!A1:XFD1048576,">0")');
    const ms = performance.now() - t0;
    console.log(`[stress] COUNTIF over full sheet with cells at A1 and XFD1000: ${ms.toFixed(0)} ms`);
    expect(get(wb, calc, 'D1')).toBe(2);
    expect(ms).toBeLessThan(500);
  }, 60_000);

  it('COUNTIF with cells at A1 and XFD1048576 (sparse scan)', () => {
    const { wb, data, calc } = twoSheets();
    put(wb, data, 'A1', '1');
    put(wb, data, 'XFD1048576', '1');
    put(wb, calc, 'D1', '=COUNTIF(Sheet1!A1:XFD1048576,">0")');
    expect(get(wb, calc, 'D1')).toBe(2);
  });

  it('elementwise =Sheet1!A1:XFD1048576*2 (toMatrixFull allocates up to 1e6 x 16384)', () => {
    const { wb, data, calc } = twoSheets();
    put(wb, data, 'A1', '1');
    put(wb, calc, 'B1', '=Sheet1!A1:XFD1048576*2');
    expect(get(wb, calc, 'B1')).toBeDefined();
  });
});

describe('#REF! after delete', () => {
  it('deleting the row a formula points at turns the reference into #REF!', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'A1', '5');
    put(wb, sh, 'A2', '=A1*2');
    deleteRows(wb, sh, 0, 1);
    const v = get(wb, sh, 'A1');
    expect(v).toBe('#REF!');
    expect(sh.get(0, 0)?.f).toContain('#REF!');
  });

  it('deleting the column a formula points at turns the reference into #REF!', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'C1', '3');
    put(wb, sh, 'B1', '=C1+1');
    deleteCols(wb, sh, 2, 1);
    expect(get(wb, sh, 'B1')).toBe('#REF!');
  });

  it('a range loses rows and shrinks; a range that is fully deleted becomes #REF!', () => {
    const { wb, sh } = sheet();
    for (let i = 1; i <= 5; i++) put(wb, sh, `A${i}`, String(i));
    put(wb, sh, 'B1', '=SUM(A1:A5)');
    put(wb, sh, 'C1', '=SUM(A2:A3)');
    deleteRows(wb, sh, 2, 2); // removes rows 3 and 4 (values 3 and 4)
    expect(get(wb, sh, 'B1')).toBe(1 + 2 + 5);
    deleteRows(wb, sh, 1, 2); // removes what was A2:A3 (only the cell that still holds 2 and 5 remain in range)
    expect(get(wb, sh, 'C1')).toBe('#REF!');
  });

  it('dependents of a #REF! cell recalc without throwing', () => {
    const { wb, sh } = sheet();
    put(wb, sh, 'A1', '1');
    put(wb, sh, 'A2', '=A1+1');
    put(wb, sh, 'A3', '=A2+1');
    deleteRows(wb, sh, 1, 1);
    // old row 3 (=A2+1) now points at the deleted row, so it becomes #REF!
    expect(get(wb, sh, 'A2')).toBe('#REF!');
    put(wb, sh, 'A1', '10');
    expect(get(wb, sh, 'A1')).toBe(10);
  });
});
