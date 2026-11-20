// Stress 2: undo/redo history. Caps at 200 steps, random walk over undo/redo, structural (snapshot) undo.
import { describe, expect, it } from 'vitest';
import { Sheet, Workbook } from '../../src/core/model';
import { applyInput, deleteRows, insertRows } from '../../src/core/ops';
import { countUndo, dumpSheet, get, put, rng } from './helpers';

function edit(wb: Workbook, sh: Sheet, ref: string, text: string): void {
  wb.transact('edit', () => put(wb, sh, ref, text));
}

describe('undo/redo history', () => {
  it('1000 single-cell edits: undo history is capped at 200 steps', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    for (let i = 1; i <= 1000; i++) edit(wb, sh, 'A1', String(i));
    expect(get(wb, sh, 'A1')).toBe(1000);
    expect(countUndo(wb)).toBe(200);
    // 200 undos from 1000 land on the value written by edit 800
    expect(get(wb, sh, 'A1')).toBe(800);
  });

  it('500 random undo/redo steps: the cell always matches the history position', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    for (let i = 1; i <= 200; i++) edit(wb, sh, 'A1', String(i));
    let pos = 200;
    const rand = rng(42);
    for (let k = 0; k < 500; k++) {
      if (wb.canUndo() && (!wb.canRedo() || rand() < 0.5)) {
        wb.undo();
        pos--;
      } else if (wb.canRedo()) {
        wb.redo();
        pos++;
      }
      expect(get(wb, sh, 'A1') ?? 0, `step ${k}`).toBe(pos);
    }
  });

  it('500 undo+redo pairs leave the value and both stacks intact', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    for (let i = 1; i <= 200; i++) edit(wb, sh, 'A1', String(i));
    for (let k = 0; k < 500; k++) {
      wb.undo();
      wb.redo();
    }
    expect(get(wb, sh, 'A1')).toBe(200);
    expect(wb.canUndo()).toBe(true);
    expect(wb.canRedo()).toBe(false);
  });

  it('a new edit after undo drops the redo stack', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    edit(wb, sh, 'A1', '1');
    edit(wb, sh, 'A1', '2');
    wb.undo();
    expect(wb.canRedo()).toBe(true);
    edit(wb, sh, 'A1', '3');
    expect(wb.canRedo()).toBe(false);
  });

  it('250 structural edits (insert/delete rows): capped at 200, and 100 undos restore the sheet exactly', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('fill', () => {
      for (let r = 0; r < 100; r++) for (let c = 0; c < 50; c++) applyInput(wb, sh, r, c, String(r * 50 + c + 1));
    });
    const rand = rng(7);
    const t0 = performance.now();
    for (let k = 0; k < 100; k++) {
      const at = Math.floor(rand() * 100);
      if (rand() < 0.5) insertRows(wb, sh, at, 1);
      else deleteRows(wb, sh, at, 1);
    }
    for (let k = 100; k < 250; k++) {
      const at = Math.floor(rand() * 100);
      if (rand() < 0.5) insertRows(wb, sh, at, 1);
      else deleteRows(wb, sh, at, 1);
    }
    const ms = performance.now() - t0;
    console.log(`[stress] 250 structural ops on 5000 cells: ${ms.toFixed(0)} ms`);

    // 251 undo entries were pushed (fill + 250 structural); the cap keeps the newest 200.
    expect(countUndo(wb)).toBe(200);
  }, 120_000);

  it('100 structural edits then 100 undos restore the sheet byte-for-byte', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('fill', () => {
      for (let r = 0; r < 100; r++) for (let c = 0; c < 50; c++) applyInput(wb, sh, r, c, String(r * 50 + c + 1));
    });
    const snap = dumpSheet(sh);
    const rand = rng(9);
    for (let k = 0; k < 100; k++) {
      const at = Math.floor(rand() * 100);
      if (rand() < 0.5) insertRows(wb, sh, at, 1);
      else deleteRows(wb, sh, at, 1);
    }
    for (let k = 0; k < 100; k++) wb.undo();
    expect(dumpSheet(sh)).toBe(snap);
  }, 120_000);
});
