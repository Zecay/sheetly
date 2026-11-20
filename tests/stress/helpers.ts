// Shared helpers for the stress tests. Test-only: nothing here touches src/.
import { Sheet, Workbook } from '../../src/core/model';
import { applyInput } from '../../src/core/ops';
import { parseCell } from '../../src/core/addr';
import { FErr, Value } from '../../src/core/values';

/** Seeded PRNG (mulberry32) so any failure can be reproduced. */
export function rng(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Type text into a cell, the same path as the UI (applyInput). */
export function put(wb: Workbook, sh: Sheet, ref: string, text: string): void {
  const p = parseCell(ref);
  if (!p) throw new Error('bad ref ' + ref);
  applyInput(wb, sh, p.r, p.c, text);
}

/** Read a cell. Errors come back as their code string, e.g. '#REF!'. */
export function get(wb: Workbook, sh: Sheet, ref: string): Value {
  const p = parseCell(ref);
  if (!p) throw new Error('bad ref ' + ref);
  const v = wb.getValue(sh, p.r, p.c);
  return v instanceof FErr ? v.code : v;
}

export function formulaAt(sh: Sheet, ref: string): string | undefined {
  const p = parseCell(ref);
  if (!p) throw new Error('bad ref ' + ref);
  return sh.get(p.r, p.c)?.f;
}

/** Drains the undo stack and returns how many steps it held. */
export function countUndo(wb: Workbook): number {
  let n = 0;
  while (wb.canUndo()) {
    wb.undo();
    n++;
  }
  return n;
}

/** Stable text dump of every cell (value and formula), for equality checks. */
export function dumpSheet(sh: Sheet): string {
  const out: string[] = [];
  for (const [k, c] of sh.cells) {
    const v = c.v instanceof FErr ? c.v.code : String(c.v);
    out.push(`${k}=${v}|${c.f ?? ''}`);
  }
  return out.sort().join('\n');
}

/** Dump of formula cells only (value per cell), for incremental-vs-full recalc checks. */
export function dumpFormulaValues(sh: Sheet): string {
  const out: string[] = [];
  for (const [k, c] of sh.cells) {
    if (c.f === undefined) continue;
    const v = c.v instanceof FErr ? c.v.code : String(c.v);
    out.push(`${k}=${v}`);
  }
  return out.sort().join('\n');
}
