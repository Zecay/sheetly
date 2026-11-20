// Cell values, error values and date serial helpers.

export class FErr {
  constructor(public code: string) {}
  toString(): string {
    return this.code;
  }
}

export const E = {
  DIV0: new FErr('#DIV/0!'),
  REF: new FErr('#REF!'),
  NAME: new FErr('#NAME?'),
  VALUE: new FErr('#VALUE!'),
  NA: new FErr('#N/A'),
  NUM: new FErr('#NUM!'),
  NULL: new FErr('#NULL!'),
  CIRC: new FErr('#CIRC!'),
};

const ERR_BY_CODE = new Map<string, FErr>(Object.values(E).map((e) => [e.code, e]));
export function errFromCode(code: string): FErr | undefined {
  return ERR_BY_CODE.get(code.toUpperCase());
}

export type Value = number | string | boolean | FErr | null;

export function isErr(v: unknown): v is FErr {
  return v instanceof FErr;
}

// ---------- dates (1900 system, including the Lotus leap-year bug) ----------

const MS_DAY = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

/** Real calendar date -> Excel serial (1 = 1900-01-01; 60 = fake 1900-02-29). */
export function ymdToSerial(y: number, m: number, d: number): number {
  const ms = Date.UTC(y, m - 1, d);
  const days = Math.round((ms - Date.UTC(1899, 11, 31)) / MS_DAY); // 1900-01-01 -> 1
  return days >= 60 ? days + 1 : days;
}

export function serialToYmd(serial: number): { y: number; m: number; d: number } {
  let n = Math.floor(serial);
  if (n === 60) return { y: 1900, m: 2, d: 29 };
  if (n > 60) n -= 1;
  const dt = new Date(Date.UTC(1899, 11, 31) + n * MS_DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}

export function weekdayOf(serial: number): number {
  // Excel: serial 1 (1900-01-01) is a Sunday in the 1900 system -> 0 = Sunday
  const n = Math.floor(serial);
  return ((n - 1) % 7 + 7) % 7;
}

export function serialToTime(serial: number): { h: number; mi: number; s: number; ms: number } {
  let frac = serial - Math.floor(serial);
  // round to milliseconds to avoid 59.9999 artefacts
  let total = Math.round(frac * 86400000);
  if (total >= 86400000) total = 86399999;
  const h = Math.floor(total / 3600000);
  const mi = Math.floor((total % 3600000) / 60000);
  const s = Math.floor((total % 60000) / 1000);
  const ms = total % 1000;
  return { h, mi, s, ms };
}

export function nowSerial(): number {
  const d = new Date();
  const local = d.getTime() - d.getTimezoneOffset() * 60000;
  return (local - EPOCH) / MS_DAY + 0; // fractional part is the local time
}

export function todaySerial(): number {
  return Math.floor(nowSerial());
}

export const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
export const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
