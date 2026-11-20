// Stress 4: insert/delete rows and columns with formulas, many times. No throws, and incremental recalc must match a full recalc.
import { describe, expect, it } from 'vitest';
import { Sheet, Workbook } from '../../src/core/model';
import { deleteCols, deleteRows, insertCols, insertRows } from '../../src/core/ops';
import { dumpFormulaValues, dumpSheet, put, rng } from './helpers';

function grid(wb: Workbook, sh: Sheet, rows = 60): void {
  wb.transact('grid', () => {
    for (let r = 0; r < rows; r++) {
      put(wb, sh, `A${r + 1}`, String(r + 1));
      put(wb, sh, `B${r + 1}`, `=A${r + 1}*2`);
      put(wb, sh, `C${r + 1}`, `=SUM(A1:A${r + 1})`);
      put(wb, sh, `D${r + 1}`, `=B${r + 1}+C${r + 1}`);
    }
  });
}

describe('insert/delete rows and columns with formulas', () => {
  it('300 random insert/delete/edit ops: no throws, and incremental values match a full recalc', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    grid(wb, sh);
    const rand = rng(99);
    const failures: string[] = [];
    const cols = ['A', 'B', 'C', 'D', 'E'];
    for (let k = 0; k < 300; k++) {
      const op = rand();
      try {
        if (op < 0.2) insertRows(wb, sh, Math.floor(rand() * (sh.maxRow + 2)), 1 + Math.floor(rand() * 3));
        else if (op < 0.4) deleteRows(wb, sh, Math.floor(rand() * (sh.maxRow + 1)), 1 + Math.floor(rand() * 2));
        else if (op < 0.5) insertCols(wb, sh, Math.floor(rand() * 6), 1);
        else if (op < 0.6) deleteCols(wb, sh, Math.floor(rand() * 6), 1);
        else {
          const r = 1 + Math.floor(rand() * 70);
          const c = cols[Math.floor(rand() * cols.length)];
          const text = rand() < 0.5 ? `=${c}${r}+1` : rand() < 0.5 ? `=SUM(A1:A${r})` : String(k);
          put(wb, sh, `${c}${r}`, text);
        }
      } catch (e) {
        failures.push(`op ${k}: threw ${(e as Error).name}: ${(e as Error).message}`);
      }
      if (k % 10 === 9) {
        const before = dumpFormulaValues(sh);
        wb.recalcAll();
        const after = dumpFormulaValues(sh);
        if (before !== after) failures.push(`op ${k}: incremental values differ from full recalc`);
      }
    }
    expect(failures.slice(0, 10)).toEqual([]);
  }, 120_000);

  it('insertRows(at) then deleteRows(at) restores every formula and value (100 positions)', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    grid(wb, sh);
    const before = dumpSheet(sh);
    const rand = rng(5);
    for (let k = 0; k < 100; k++) {
      const at = Math.floor(rand() * 60);
      insertRows(wb, sh, at, 1);
      deleteRows(wb, sh, at, 1);
      expect(dumpSheet(sh), `round trip at row ${at}, iteration ${k}`).toBe(before);
    }
  }, 120_000);

  it('insertCols(at) then deleteCols(at) restores every formula and value (100 positions)', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    grid(wb, sh);
    const before = dumpSheet(sh);
    const rand = rng(6);
    for (let k = 0; k < 100; k++) {
      const at = Math.floor(rand() * 4);
      insertCols(wb, sh, at, 1);
      deleteCols(wb, sh, at, 1);
      expect(dumpSheet(sh), `round trip at col ${at}, iteration ${k}`).toBe(before);
    }
  }, 120_000);
});
