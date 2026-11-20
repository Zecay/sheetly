// Stress 5b: bulk operations. Sort 100k rows, and xlsx / csv round trips of 50k cells.
import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/core/model';
import { applyInput, sortRange } from '../../src/core/ops';
import { colName } from '../../src/core/addr';
import { readCSV, writeCSV } from '../../src/core/csv';
import { readXlsx, writeXlsx } from '../../src/core/xlsx';
import { FErr, Value } from '../../src/core/values';
import { rng } from './helpers';

function sameValue(a: Value, b: Value): boolean {
  if (a instanceof FErr || b instanceof FErr) return a instanceof FErr && b instanceof FErr && a.code === b.code;
  return a === b;
}

describe('bulk operations', () => {
  it('sort 100k rows by a number column; formula column follows its row', () => {
    const N = 100_000;
    const rand = rng(11);
    const wb = new Workbook();
    const sh = wb.sheets[0];
    const t0 = performance.now();
    wb.transact('build', () => {
      for (let r = 0; r < N; r++) {
        applyInput(wb, sh, r, 0, String(Math.floor(rand() * 1_000_000)));
        applyInput(wb, sh, r, 1, `row ${r}`);
        applyInput(wb, sh, r, 2, `=A${r + 1}*2`);
      }
    });
    const buildMs = performance.now() - t0;

    const t1 = performance.now();
    sortRange(wb, sh, { r1: 0, c1: 0, r2: N - 1, c2: 2 }, [{ col: 0, asc: true }], false);
    const sortMs = performance.now() - t1;
    console.log(`[stress] 100k rows: build ${buildMs.toFixed(0)} ms, sort ${sortMs.toFixed(0)} ms`);

    let prev = -Infinity;
    let badOrder = 0;
    let badFormula = 0;
    const labels = new Set<unknown>();
    for (let r = 0; r < N; r++) {
      const a = wb.getValue(sh, r, 0) as number;
      if (a < prev) badOrder++;
      prev = a;
      if (wb.getValue(sh, r, 2) !== a * 2) badFormula++;
      if (sh.get(r, 2)?.f !== `A${r + 1}*2`) badFormula++;
      labels.add(wb.getValue(sh, r, 1));
    }
    expect(badOrder).toBe(0);
    expect(badFormula).toBe(0);
    expect(labels.size).toBe(N);
    expect(sortMs).toBeLessThan(20_000);
  }, 180_000);

  function buildMixed(): Workbook {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    const R = 500;
    const C = 100;
    wb.transact('build', () => {
      for (let r = 0; r < R; r++)
        for (let c = 0; c < C; c++) {
          const k = c % 4;
          let text: string;
          if (k === 0) text = String(r * 100 + c + 0.5);
          else if (k === 1) text = r % 10 === 0 ? `a,"b" ${r}` : `item ${r}-${c}`;
          else if (k === 2) text = `=${colName(c - 2)}${r + 1}*2`;
          else text = r % 2 === 0 ? 'TRUE' : 'FALSE';
          applyInput(wb, sh, r, c, text);
        }
    });
    return wb;
  }

  it('xlsx round trip of 50k cells (values, formulas, booleans, quoted text)', () => {
    const orig = buildMixed();
    const sh = orig.sheets[0];
    const t0 = performance.now();
    const bytes = writeXlsx(orig);
    const t1 = performance.now();
    const back = readXlsx(bytes);
    const t2 = performance.now();
    const bs = back.sheets[0];
    let valBad = 0;
    let fBad = 0;
    const samples: string[] = [];
    for (let r = 0; r < 500; r++)
      for (let c = 0; c < 100; c++) {
        const a = orig.getValue(sh, r, c);
        const b = back.getValue(bs, r, c);
        if (!sameValue(a, b)) {
          valBad++;
          if (samples.length < 5) samples.push(`r${r}c${c}: ${String(a)} vs ${String(b)}`);
        }
        if (sh.get(r, c)?.f !== bs.get(r, c)?.f) fBad++;
      }
    console.log(`[stress] xlsx 50k: write ${(t1 - t0).toFixed(0)} ms, read ${(t2 - t1).toFixed(0)} ms, ${bytes.length} bytes`);
    expect(samples).toEqual([]);
    expect(valBad).toBe(0);
    expect(fBad).toBe(0);
  }, 180_000);

  it('csv round trip of 50k cells (values; formulas come back as values)', () => {
    const orig = buildMixed();
    const sh = orig.sheets[0];
    const t0 = performance.now();
    const text = writeCSV(orig, sh);
    const t1 = performance.now();
    const back = readCSV(text);
    const t2 = performance.now();
    const bs = back.sheets[0];
    let valBad = 0;
    const samples: string[] = [];
    for (let r = 0; r < 500; r++)
      for (let c = 0; c < 100; c++) {
        const a = orig.getValue(sh, r, c);
        const b = back.getValue(bs, r, c);
        if (!sameValue(a, b)) {
          valBad++;
          if (samples.length < 5) samples.push(`r${r}c${c}: ${String(a)} vs ${String(b)}`);
        }
      }
    console.log(`[stress] csv 50k: write ${(t1 - t0).toFixed(0)} ms, read ${(t2 - t1).toFixed(0)} ms, ${text.length} chars`);
    expect(samples).toEqual([]);
    expect(valBad).toBe(0);
  }, 180_000);
});
