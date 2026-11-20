// File access: Tauri commands when running as the desktop app, browser fallbacks otherwise (dev preview).
import { invoke } from '@tauri-apps/api/core';

export const isTauri = typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;

export interface FileFilter {
  name: string;
  extensions: string[];
}

export const OPEN_FILTERS: FileFilter[] = [
  { name: 'Spreadsheets', extensions: ['xlsx', 'csv', 'sheetly'] },
  { name: 'Excel Workbook', extensions: ['xlsx'] },
  { name: 'CSV (Comma delimited)', extensions: ['csv'] },
  { name: 'Sheetly Workbook', extensions: ['sheetly'] },
];

export function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}
export function ext(path: string): string {
  const m = /\.([^.\\/]+)$/.exec(path);
  return m ? m[1].toLowerCase() : '';
}

export async function pickOpen(): Promise<{ path: string } | { file: File } | null> {
  if (isTauri) {
    const { open } = await import('@tauri-apps/plugin-dialog');
    const p = await open({ multiple: false, filters: OPEN_FILTERS });
    if (!p) return null;
    return { path: typeof p === 'string' ? p : (p as any).path };
  }
  return new Promise((resolve) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = '.xlsx,.csv,.sheetly';
    inp.onchange = () => resolve(inp.files && inp.files[0] ? { file: inp.files[0] } : null);
    inp.click();
  });
}

export async function pickSave(defaultName: string, filters: FileFilter[]): Promise<string | null> {
  if (isTauri) {
    const { save } = await import('@tauri-apps/plugin-dialog');
    const p = await save({ defaultPath: defaultName, filters });
    return p ?? null;
  }
  return defaultName;
}

export async function readBytes(path: string): Promise<Uint8Array> {
  const buf = (await invoke('read_file', { path })) as ArrayBuffer;
  return new Uint8Array(buf);
}

export async function writeBytes(path: string, bytes: Uint8Array): Promise<void> {
  if (isTauri) {
    await invoke('write_file', bytes, { headers: { path: encodeURIComponent(path) } });
    return;
  }
  const blob = new Blob([bytes as BlobPart]);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = baseName(path);
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

export async function fileExists(path: string): Promise<boolean> {
  if (!isTauri) return false;
  try {
    return (await invoke('path_exists', { path })) as boolean;
  } catch {
    return false;
  }
}

export async function cliFile(): Promise<string | null> {
  if (!isTauri) return null;
  try {
    return ((await invoke('cli_file')) as string | null) ?? null;
  } catch {
    return null;
  }
}

// ---- app data (autosave recovery) ----
export async function dataWrite(name: string, text: string): Promise<void> {
  if (!isTauri) {
    try {
      localStorage.setItem('sheetly.data.' + name, text);
    } catch {
      /* quota */
    }
    return;
  }
  await invoke('data_write', new TextEncoder().encode(text), { headers: { name } });
}
export async function dataRead(name: string): Promise<string | null> {
  if (!isTauri) return localStorage.getItem('sheetly.data.' + name);
  try {
    if (!((await invoke('data_exists', { name })) as boolean)) return null;
    const buf = (await invoke('data_read', { name })) as ArrayBuffer;
    return new TextDecoder().decode(new Uint8Array(buf));
  } catch {
    return null;
  }
}
export async function dataDelete(name: string): Promise<void> {
  if (!isTauri) {
    localStorage.removeItem('sheetly.data.' + name);
    return;
  }
  try {
    await invoke('data_delete', { name });
  } catch {
    /* ignore */
  }
}

// ---- recent files ----
export interface Recent {
  path: string;
  name: string;
  time: number;
}
export function loadRecent(): Recent[] {
  try {
    const v: unknown = JSON.parse(localStorage.getItem('sheetly.recent') ?? '[]');
    if (!Array.isArray(v)) return [];
    return v.filter((r): r is Recent => !!r && typeof r.path === 'string').map((r) => ({ path: r.path, name: typeof r.name === 'string' ? r.name : baseName(r.path), time: typeof r.time === 'number' ? r.time : 0 }));
  } catch {
    return [];
  }
}
export function pushRecent(path: string): void {
  try {
    const list = loadRecent().filter((r) => r.path !== path);
    list.unshift({ path, name: baseName(path), time: Date.now() });
    localStorage.setItem('sheetly.recent', JSON.stringify(list.slice(0, 12)));
  } catch {
    /* ignore */
  }
}
export function removeRecent(path: string): void {
  try {
    localStorage.setItem('sheetly.recent', JSON.stringify(loadRecent().filter((r) => r.path !== path)));
  } catch {
    /* ignore */
  }
}
