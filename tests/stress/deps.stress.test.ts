// Stress 3: dependency graph cleanup. Formulas overwritten by values or removed must leave no entries behind.
import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/core/model';
import { clearRange, deleteRows } from '../../src/core/ops';
import { get, put, rng } from './helpers';

// Internal dependency maps (private fields), read only.
function stats(wb: Workbook) {
  const w = wb as unknown as {
    deps: Map<number, unknown>;
    pointDeps: Map<number, Set<number>>;
    rangeBuckets: Map<number, Set<unknown>>;
    tallRanges: Map<number, Set<unknown>>;
    volatileSet: Set<number>;
  };
  let pointEdges = 0;
  for (const s of w.pointDeps.values()) pointEdges += s.size;
  let bucketEntries = 0;
  for (const s of w.rangeBuckets.values()) bucketEntries += s.size;
  let tallEntries = 0;
  for (const s of w.tallRanges.values()) tallEntries += s.size;
  return {
    deps: w.deps.size,
    pointKeys: w.pointDeps.size,
    pointEdges,
    bucketKeys: w.rangeBuckets.size,
    bucketEntries,
    tallEntries,
    volatile: w.volatileSet.size,
  };
}

const ZERO = { deps: 0, pointKeys: 0, pointEdges: 0, bucketKeys: 0, bucketEntries: 0, tallEntries: 0, volatile: 0 };

describe('dependency graph cleanup', () => {
  it('overwriting formulas with values removes every dependency entry', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('build', () => {
      for (let i = 1; i <= 1000; i++) {
        put(wb, sh, `A${i}`, String(i));
        put(wb, sh, `B${i}`, `=A${i}*2`);
      }
      put(wb, sh, 'C1', '=SUM(A1:A1000)'); // row-bucketed range
      put(wb, sh, 'D1', '=SUM(A1:A100000)'); // tall range (more than 64 buckets)
    });
    expect(get(wb, sh, 'C1')).toBe(500500);
    let s = stats(wb);
    expect(s.deps).toBe(1002);
    expect(s.pointKeys).toBe(1000);
    expect(s.tallEntries).toBe(1);

    wb.transact('overwrite B', () => {
      for (let i = 1; i <= 1000; i++) put(wb, sh, `B${i}`, String(i));
    });
    s = stats(wb);
    expect(s.deps).toBe(2);
    expect(s.pointKeys).toBe(0);
    expect(s.pointEdges).toBe(0);

    wb.transact('overwrite C1 and D1', () => {
      put(wb, sh, 'C1', 'x');
      put(wb, sh, 'D1', 'x');
    });
    expect(stats(wb)).toEqual(ZERO);
  });

  it('clearing a block of formulas (clearRange all) leaves nothing behind', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('build', () => {
      for (let i = 1; i <= 1000; i++) {
        put(wb, sh, `A${i}`, String(i));
        put(wb, sh, `B${i}`, `=A${i}+1`);
      }
    });
    wb.transact('clear B', () => clearRange(wb, sh, { r1: 0, c1: 1, r2: 999, c2: 1 }, 'all'));
    const s = stats(wb);
    expect(s.deps).toBe(0);
    expect(s.pointKeys).toBe(0);
  });

  it('deleting rows drops the dependents that lived in them', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('build', () => {
      for (let i = 1; i <= 1000; i++) {
        put(wb, sh, `A${i}`, String(i));
        put(wb, sh, `B${i}`, `=A${i}*2`);
      }
    });
    deleteRows(wb, sh, 0, 500); // removes A1:A500 and B1:B500
    const s = stats(wb);
    expect(s.deps).toBe(500);
    expect(s.pointKeys).toBe(500);
  });

  it('2000 formula overwrites on one cell do not accumulate entries', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('values', () => {
      for (let i = 1; i <= 50; i++) put(wb, sh, `A${i}`, String(i));
    });
    const rand = rng(3);
    for (let k = 0; k < 2000; k++) {
      const r = 1 + Math.floor(rand() * 50);
      if (k % 3 === 0) put(wb, sh, 'Z1', `=A${r}*2`);
      else if (k % 3 === 1) put(wb, sh, 'Z1', `=SUM(A1:A${r})`);
      else put(wb, sh, 'Z1', String(k));
      if (k % 500 === 0) {
        const s = stats(wb);
        expect(s.deps, `after ${k} overwrites`).toBeLessThanOrEqual(1);
        expect(s.pointKeys, `after ${k} overwrites`).toBeLessThanOrEqual(1);
        expect(s.bucketEntries, `after ${k} overwrites`).toBeLessThanOrEqual(1);
      }
    }
    put(wb, sh, 'Z1', '1');
    expect(stats(wb)).toEqual(ZERO);
  });

  it('5000 different precedents, one formula cell: point index holds only the current one', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    wb.transact('values', () => {
      for (let i = 1; i <= 5000; i++) put(wb, sh, `A${i}`, String(i));
    });
    for (let i = 1; i <= 5000; i++) put(wb, sh, 'Z1', `=A${i}`);
    const s = stats(wb);
    expect(s.deps).toBe(1);
    expect(s.pointKeys).toBe(1);
    expect(s.pointEdges).toBe(1);
    expect(get(wb, sh, 'Z1')).toBe(5000);
  });

  it('an overwritten cell no longer follows its old precedent', () => {
    const wb = new Workbook();
    const sh = wb.sheets[0];
    put(wb, sh, 'A1', '1');
    put(wb, sh, 'B1', '=A1*10');
    put(wb, sh, 'B1', 'hello');
    put(wb, sh, 'A1', '5');
    expect(get(wb, sh, 'B1')).toBe('hello');
    put(wb, sh, 'B1', '=A1+1');
    expect(get(wb, sh, 'B1')).toBe(6);
    put(wb, sh, 'A1', '7');
    expect(get(wb, sh, 'B1')).toBe(8);
  });
});
