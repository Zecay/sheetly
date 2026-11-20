// Small DOM helpers: element builder, popup menus, modal dialogs, colour picker.
import { icon } from './icons';

type Child = Node | string | null | undefined | false;

export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, any> = {},
  ...children: (Child | Child[])[]
): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else if (k === 'html') el.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'checked' || k === 'disabled' || k === 'selected' || k === 'value' || k === 'readOnly') (el as any)[k] = v;
    else el.setAttribute(k, v === true ? '' : String(v));
  }
  for (const c of children.flat()) {
    if (c === null || c === undefined || c === false) continue;
    el.append(typeof c === 'string' ? document.createTextNode(c) : c);
  }
  return el;
}

export function clear(el: HTMLElement): void {
  while (el.firstChild) el.removeChild(el.firstChild);
}

// ---------------- popup menus ----------------

export interface MenuItem {
  label?: string;
  icon?: string;
  cmd?: () => void;
  sep?: boolean;
  disabled?: boolean;
  checked?: boolean;
  shortcut?: string;
  submenu?: MenuItem[];
  custom?: HTMLElement;
  heading?: string;
}

interface OpenMenu {
  el: HTMLElement;
  close: () => void;
}
let openMenus: OpenMenu[] = [];

export function closeMenus(): void {
  for (const m of openMenus) m.el.remove();
  openMenus = [];
  document.removeEventListener('mousedown', outsideHandler, true);
  document.removeEventListener('keydown', escHandler, true);
}
function outsideHandler(e: MouseEvent) {
  if (openMenus.some((m) => m.el.contains(e.target as Node))) return;
  closeMenus();
}
function escHandler(e: KeyboardEvent) {
  if (e.key === 'Escape' && openMenus.length) {
    e.stopPropagation();
    closeMenus();
  }
}

function buildMenu(items: MenuItem[], depth: number): HTMLElement {
  const el = h('div', { class: 'menu', role: 'menu' });
  el.addEventListener('contextmenu', (e) => e.preventDefault());
  for (const it of items) {
    if (it.sep) {
      el.append(h('div', { class: 'menu-sep' }));
      continue;
    }
    if (it.heading) {
      el.append(h('div', { class: 'menu-heading' }, it.heading));
      continue;
    }
    if (it.custom) {
      el.append(it.custom);
      continue;
    }
    const row = h(
      'div',
      { class: 'menu-item' + (it.disabled ? ' disabled' : ''), role: 'menuitem' },
      h('span', { class: 'mi-icon', html: it.checked ? icon('check', 14) : it.icon ? icon(it.icon, 16) : '' }),
      h('span', { class: 'mi-label' }, it.label ?? ''),
      h('span', { class: 'mi-key' }, it.shortcut ?? ''),
      it.submenu ? h('span', { class: 'mi-arrow', html: icon('right', 10) }) : null,
    );
    if (it.submenu) {
      let sub: HTMLElement | null = null;
      row.addEventListener('mouseenter', () => {
        openMenus = openMenus.filter((m) => {
          if (m.el !== sub && (m.el as any)._depth > depth) {
            m.el.remove();
            return false;
          }
          return true;
        });
        const r = row.getBoundingClientRect();
        sub = buildMenu(it.submenu!, depth + 1);
        (sub as any)._depth = depth + 1;
        document.body.append(sub);
        placeAt(sub, r.right - 2, r.top - 4, r.left);
        openMenus.push({ el: sub, close: () => sub?.remove() });
      });
    } else {
      row.addEventListener('mouseenter', () => {
        openMenus = openMenus.filter((m) => {
          if ((m.el as any)._depth > depth) {
            m.el.remove();
            return false;
          }
          return true;
        });
      });
    }
    if (!it.disabled && !it.submenu)
      row.addEventListener('click', (e) => {
        e.stopPropagation();
        closeMenus();
        it.cmd?.();
      });
    el.append(row);
  }
  return el;
}

function placeAt(el: HTMLElement, x: number, y: number, altRight?: number) {
  el.style.left = '0px';
  el.style.top = '0px';
  const w = el.offsetWidth;
  const hgt = el.offsetHeight;
  let left = x;
  if (left + w > window.innerWidth - 4) left = altRight !== undefined ? altRight - w + 2 : window.innerWidth - w - 4;
  let top = y;
  if (top + hgt > window.innerHeight - 4) top = Math.max(4, window.innerHeight - hgt - 4);
  el.style.left = Math.max(2, left) + 'px';
  el.style.top = Math.max(2, top) + 'px';
}

export function showMenu(x: number, y: number, items: MenuItem[]): void {
  closeMenus();
  const el = buildMenu(items, 0);
  (el as any)._depth = 0;
  document.body.append(el);
  placeAt(el, x, y);
  openMenus.push({ el, close: () => el.remove() });
  document.addEventListener('mousedown', outsideHandler, true);
  document.addEventListener('keydown', escHandler, true);
}

export function showMenuBelow(anchor: HTMLElement, items: MenuItem[]): void {
  const r = anchor.getBoundingClientRect();
  showMenu(r.left, r.bottom, items);
}

export function showPopup(anchor: HTMLElement, content: HTMLElement): void {
  closeMenus();
  const el = h('div', { class: 'menu popup' }, content);
  (el as any)._depth = 0;
  document.body.append(el);
  const r = anchor.getBoundingClientRect();
  placeAt(el, r.left, r.bottom);
  openMenus.push({ el, close: () => el.remove() });
  document.addEventListener('mousedown', outsideHandler, true);
  document.addEventListener('keydown', escHandler, true);
}

// ---------------- colour picker ----------------

const THEME_COLS = ['#ffffff', '#000000', '#e7e6e6', '#44546a', '#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
const TINTS: number[][] = [
  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  [-0.05, 0.5, -0.1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8],
  [-0.15, 0.35, -0.25, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6],
  [-0.25, 0.25, -0.5, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4],
  [-0.35, 0.15, -0.75, -0.25, -0.25, -0.25, -0.25, -0.25, -0.25, -0.25],
  [-0.5, 0.05, -0.9, -0.5, -0.5, -0.5, -0.5, -0.5, -0.5, -0.5],
];
const STANDARD = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];

export function shade(hex: string, t: number): string {
  const n = parseInt(hex.slice(1), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => {
    const x = t < 0 ? v * (1 + t) : v + (255 - v) * t;
    return Math.max(0, Math.min(255, Math.round(x)));
  });
  return '#' + ch.map((v) => v.toString(16).padStart(2, '0')).join('');
}

export function showColorMenu(anchor: HTMLElement, opts: { noneLabel: string; onPick: (c: string | null) => void }): void {
  const wrap = h('div', { class: 'colors' });
  const none = h('div', { class: 'menu-item color-none' }, h('span', { class: 'mi-icon', html: icon('close', 14) }), h('span', { class: 'mi-label' }, opts.noneLabel));
  none.addEventListener('click', () => {
    closeMenus();
    opts.onPick(null);
  });
  wrap.append(none);
  const sw = (c: string, title: string) => {
    const b = h('button', { class: 'swatch', style: { background: c }, title, type: 'button' });
    b.addEventListener('click', () => {
      closeMenus();
      opts.onPick(c);
    });
    return b;
  };
  wrap.append(h('div', { class: 'menu-heading' }, 'Theme Colors'));
  const grid = h('div', { class: 'swatch-grid' });
  for (const row of TINTS) {
    row.forEach((t, i) => grid.append(sw(t === 0 ? THEME_COLS[i] : shade(THEME_COLS[i], t), 'Theme color')));
  }
  wrap.append(grid);
  wrap.append(h('div', { class: 'menu-heading' }, 'Standard Colors'));
  const std = h('div', { class: 'swatch-grid one' });
  for (const c of STANDARD) std.append(sw(c, 'Standard color'));
  wrap.append(std);
  const more = h('div', { class: 'menu-item' }, h('span', { class: 'mi-icon' }), h('span', { class: 'mi-label' }, 'More Colors…'));
  const inp = h('input', { type: 'color', style: { position: 'absolute', opacity: '0', pointerEvents: 'none', width: '0', height: '0' } });
  inp.addEventListener('input', () => {
    opts.onPick(inp.value);
  });
  inp.addEventListener('change', () => closeMenus());
  more.addEventListener('click', () => inp.click());
  wrap.append(h('div', { class: 'menu-sep' }), more, inp);
  showPopup(anchor, wrap);
}

// ---------------- dialogs ----------------

export interface DialogButton {
  label: string;
  primary?: boolean;
  /** return false to keep the dialog open */
  onClick?: () => boolean | void | Promise<boolean | void>;
}

export interface DialogHandle {
  el: HTMLElement;
  close: () => void;
}

export function openDialog(opts: { title: string; content: HTMLElement; buttons?: DialogButton[]; width?: number; className?: string; modeless?: boolean; onClose?: () => void }): DialogHandle {
  closeMenus();
  const backdrop = h('div', { class: 'dlg-backdrop' + (opts.modeless ? ' modeless' : '') });
  const body = h('div', { class: 'dlg-body' }, opts.content);
  const foot = h('div', { class: 'dlg-foot' });
  const closeBtn = h('button', { class: 'dlg-x', type: 'button', title: 'Close', html: icon('close', 12) });
  const dlg = h('div', { class: 'dlg ' + (opts.className ?? ''), role: 'dialog', 'aria-modal': 'true', style: opts.width ? { width: opts.width + 'px' } : {} }, h('div', { class: 'dlg-title' }, h('span', {}, opts.title), closeBtn), body, foot);
  backdrop.append(dlg);
  document.body.append(backdrop);
  const prevFocus = document.activeElement as HTMLElement | null;
  let closed = false;
  const close = () => {
    if (closed) return;
    closed = true;
    backdrop.remove();
    document.removeEventListener('keydown', keyHandler, true);
    opts.onClose?.();
    prevFocus?.focus?.();
  };
  const buttons = opts.buttons ?? [{ label: 'OK', primary: true }];
  let primaryBtn: HTMLButtonElement | null = null;
  for (const b of buttons) {
    const btn = h('button', { class: 'btn' + (b.primary ? ' primary' : ''), type: 'button' }, b.label);
    btn.addEventListener('click', async () => {
      let r: unknown;
      try {
        r = await b.onClick?.();
      } catch (err) {
        void messageBox('Sheetly', 'Something went wrong.\n\n' + ((err as Error)?.message ?? String(err)));
        return;
      }
      if (r !== false) close();
    });
    if (b.primary) primaryBtn = btn;
    foot.append(btn);
  }
  const keyHandler = (e: KeyboardEvent) => {
    if (!backdrop.isConnected) return;
    // only the topmost dialog reacts
    const all = document.querySelectorAll('.dlg-backdrop');
    if (all[all.length - 1] !== backdrop) return;
    if (e.key === 'Escape') {
      e.stopPropagation();
      e.preventDefault();
      close();
    } else if (e.key === 'Enter' && primaryBtn && !(e.target instanceof HTMLTextAreaElement) && !(e.target instanceof HTMLButtonElement) && !(e.target instanceof HTMLSelectElement)) {
      e.preventDefault();
      primaryBtn.click();
    }
  };
  document.addEventListener('keydown', keyHandler, true);
  closeBtn.addEventListener('click', close);
  backdrop.addEventListener('mousedown', (e) => {
    if (e.target === backdrop) dlg.classList.add('shake'), setTimeout(() => dlg.classList.remove('shake'), 300);
  });
  // focus first input or primary button
  setTimeout(() => {
    const first = dlg.querySelector<HTMLElement>('input:not([type=checkbox]):not([type=radio]), textarea, select');
    (first ?? primaryBtn)?.focus();
    if (first instanceof HTMLInputElement && first.type === 'text') first.select();
  }, 0);
  return { el: dlg, close };
}

export function messageBox(title: string, message: string, buttons: string[] = ['OK']): Promise<string> {
  return new Promise((resolve) => {
    let result = buttons[buttons.length - 1];
    const content = h('div', { class: 'msg' }, h('div', { class: 'msg-text' }, message));
    const handle = openDialog({
      title,
      content,
      width: 380,
      buttons: buttons.map((b, i) => ({
        label: b,
        primary: i === 0,
        onClick: () => {
          result = b;
        },
      })),
      onClose: () => resolve(result),
    });
    void handle;
  });
}

export function promptBox(title: string, label: string, value: string, validate?: (v: string) => string | null): Promise<string | null> {
  return new Promise((resolve) => {
    let out: string | null = null;
    const input = h('input', { type: 'text', value, class: 'txt', style: { width: '100%' } });
    const err = h('div', { class: 'err' });
    const content = h('div', {}, h('label', { class: 'lbl' }, label), input, err);
    openDialog({
      title,
      content,
      width: 360,
      buttons: [
        {
          label: 'OK',
          primary: true,
          onClick: () => {
            const m = validate?.(input.value) ?? null;
            if (m) {
              err.textContent = m;
              return false;
            }
            out = input.value;
          },
        },
        { label: 'Cancel' },
      ],
      onClose: () => resolve(out),
    });
  });
}

let toastTimer = 0;
export function toast(msg: string): void {
  let t = document.getElementById('toast');
  if (!t) {
    t = h('div', { id: 'toast' });
    document.body.append(t);
  }
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => t!.classList.remove('show'), 2600);
}
