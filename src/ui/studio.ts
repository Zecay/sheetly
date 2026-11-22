// The Appearance panel: docked beside the sheet so every change is seen live on the real app.
import { clear, h, promptBox, toast } from './dom';
import { icon } from './icons';
import { ColorKey, FONTS, Grad, Look, PRESETS, clone, hasBackdrop, isDark, isHex, loadSaved, mix, palette, preset, readable, revealBackdrop, storeImage, storeSaved } from './theme';
import { SCENES, TEXTURES, previewScene, textureUrl } from './wallpaper';
import type { App } from './app';

type Tab = 'themes' | 'colors' | 'background' | 'animation' | 'style';
const TABS: [Tab, string][] = [['themes', 'Themes'], ['colors', 'Colors'], ['background', 'Background'], ['animation', 'Motion'], ['style', 'Style']];

const COLOR_GROUPS: [string, [ColorKey, string][]][] = [
  ['Top bar', [['bar', 'Background'], ['barText', 'Text and icons']]],
  ['Toolbars', [['accent', 'Accent'], ['ribbon', 'Ribbon'], ['chrome', 'Formula and tab bar'], ['status', 'Status bar'], ['hover', 'Hover']]],
  ['Sheet', [['sheet', 'Background'], ['cellText', 'Text'], ['grid', 'Gridlines'], ['hdr', 'Row and column headers'], ['hdrText', 'Header text'], ['sel', 'Selection']]],
  ['Menus and text', [['panel', 'Menus and dialogs'], ['text', 'Text'], ['muted', 'Secondary text'], ['border', 'Borders']]],
];
const GRADIENTS: [string, string][] = [
  ['#ff9a6b', '#d8347f'], ['#1fa2ff', '#12d8a5'], ['#7f53ff', '#ef5fa7'], ['#f6d365', '#fda085'], ['#0f2027', '#2c5364'], ['#141e30', '#35577d'],
  ['#2b1055', '#7597de'], ['#42275a', '#b6657f'], ['#134e5e', '#71b280'], ['#232526', '#5a6066'], ['#fdfbfb', '#dfe7ee'], ['#ffecd2', '#fcb69f'],
];
const QUICK = ['#217346', '#2f6fd0', '#0a84c6', '#14a98a', '#7048e8', '#c678dd', '#d6336c', '#f2643d', '#e0a100', '#607080'];

export class Studio {
  private el = h('aside', { id: 'studio', 'aria-label': 'Appearance' });
  private body = h('div', { class: 'st-body' });
  private tabsEl = h('div', { class: 'st-tabs', role: 'tablist' });
  private undoBtn: HTMLButtonElement;
  private tab: Tab = 'themes';
  private history: Look[] = [];
  private lastEdit = 0;

  constructor(private app: App) {
    const close = h('button', { class: 'st-icon', type: 'button', title: 'Close (Esc)', html: icon('close', 12) });
    close.addEventListener('click', () => this.hide());
    this.undoBtn = h('button', { class: 'btn small', type: 'button', title: 'Undo the last appearance change' }, 'Undo');
    this.undoBtn.addEventListener('click', () => this.undo());
    const reset = h('button', { class: 'btn small', type: 'button', title: 'Go back to the default Sheetly look' }, 'Reset');
    reset.addEventListener('click', () => this.pick(preset(isDark(this.look) ? 'dark' : 'classic')));
    const save = h('button', { class: 'btn small primary', type: 'button' }, 'Save as theme');
    save.addEventListener('click', () => void this.saveTheme());
    this.el.append(
      h('div', { class: 'st-head' }, h('span', { class: 'st-title' }, 'Appearance'), close),
      this.tabsEl,
      this.body,
      h('div', { class: 'st-foot' }, this.undoBtn, reset, h('span', { style: { flex: '1' } }), save),
    );
    // Typing and arrow keys in this panel must not reach the sheet.
    this.el.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') this.hide();
    });
    this.el.hidden = true;
    document.body.append(this.el);
  }

  private get look(): Look {
    return this.app.look;
  }
  get isOpen(): boolean {
    return !this.el.hidden;
  }

  show(tab?: Tab): void {
    if (tab) this.tab = tab;
    this.el.hidden = false;
    document.documentElement.dataset.studio = 'on';
    this.render();
    this.app.ribbon.update();
  }
  hide(): void {
    this.el.hidden = true;
    delete document.documentElement.dataset.studio;
    this.app.ribbon.update();
    this.app.grid.focusGrid();
  }
  toggle(): void {
    if (this.isOpen) this.hide();
    else this.show();
  }
  /** Called when the look changes from outside the panel (ribbon light/dark). */
  refresh(): void {
    if (this.isOpen) this.render();
  }

  // ---------------- state ----------------
  private remember(force = false): void {
    const now = Date.now();
    // A drag on a slider or colour picker is one undo step.
    if (force || now - this.lastEdit > 700) {
      this.history.push(clone(this.look));
      if (this.history.length > 40) this.history.shift();
    }
    this.lastEdit = now;
    this.undoBtn.disabled = !this.history.length;
  }
  private edit(fn: (l: Look) => void, rerender = false): void {
    this.remember();
    const l = clone(this.look);
    fn(l);
    if (l.id !== 'custom') {
      l.id = 'custom';
      l.name = 'Custom';
    }
    this.app.setLook(l);
    if (rerender) this.render();
  }
  private pick(l: Look): void {
    this.remember(true);
    this.app.setLook(clone(l));
    this.render();
  }
  private undo(): void {
    const l = this.history.pop();
    if (l) this.app.setLook(l);
    this.lastEdit = 0;
    this.render();
  }
  private async saveTheme(): Promise<void> {
    const name = await promptBox('Save theme', 'Theme name', 'My theme', (v) => (v.trim() ? null : 'Enter a name.'));
    if (!name) return;
    const l = clone(this.look);
    l.id = 'my-' + Date.now().toString(36);
    l.name = name.trim().slice(0, 40);
    storeSaved([l, ...loadSaved()].slice(0, 40));
    this.app.setLook(l);
    this.tab = 'themes';
    this.render();
    toast('Theme saved');
  }

  // ---------------- rendering ----------------
  private render(): void {
    const top = this.body.scrollTop;
    clear(this.tabsEl);
    for (const [id, label] of TABS) {
      const b = h('button', { type: 'button', role: 'tab', class: this.tab === id ? 'on' : '' }, label);
      b.addEventListener('click', () => {
        const same = this.tab === id;
        this.tab = id;
        this.render();
        if (!same) this.body.scrollTop = 0;
      });
      this.tabsEl.append(b);
    }
    clear(this.body);
    this[`page_${this.tab}` as const]();
    this.body.scrollTop = top;
    this.undoBtn.disabled = !this.history.length;
  }

  private section(title: string, ...kids: (HTMLElement | null)[]): void {
    this.body.append(h('section', { class: 'st-sec' }, h('h3', {}, title), kids));
  }
  private row(label: string, control: HTMLElement): HTMLElement {
    return h('label', { class: 'st-row' }, h('span', {}, label), control);
  }
  private slider(label: string, value: number, min: number, max: number, unit: string, set: (l: Look, v: number) => void, step = 1): HTMLElement {
    const out = h('output', {}, value + unit);
    const input = h('input', { type: 'range', min: String(min), max: String(max), step: String(step), value: String(value) }) as HTMLInputElement;
    input.addEventListener('input', () => {
      out.textContent = input.value + unit;
      this.edit((l) => set(l, +input.value));
    });
    return h('label', { class: 'st-slider' }, h('span', {}, label), input, out);
  }
  private toggle_(label: string, value: boolean, set: (l: Look, v: boolean) => void, rerender = false): HTMLElement {
    const input = h('input', { type: 'checkbox', checked: value }) as HTMLInputElement;
    input.addEventListener('change', () => this.edit((l) => set(l, input.checked), rerender));
    return h('label', { class: 'st-toggle' }, h('span', {}, label), input, h('i', {}));
  }
  private color(value: string, set: (l: Look, v: string) => void): HTMLElement {
    const pick = h('input', { type: 'color', value }) as HTMLInputElement;
    const hex = h('input', { type: 'text', value, maxLength: 7, spellcheck: 'false', 'aria-label': 'Hex color' }) as HTMLInputElement;
    pick.addEventListener('input', () => {
      hex.value = pick.value;
      this.edit((l) => set(l, pick.value));
    });
    hex.addEventListener('input', () => {
      const v = hex.value.trim().toLowerCase();
      if (!isHex(v)) return;
      pick.value = v;
      this.edit((l) => set(l, v));
    });
    return h('span', { class: 'st-color' }, pick, hex);
  }
  private seg<T extends string>(items: [T, string][], value: T, set: (v: T) => void): HTMLElement {
    const el = h('div', { class: 'st-seg' });
    for (const [v, label] of items) {
      const b = h('button', { type: 'button', class: v === value ? 'on' : '' }, label);
      b.addEventListener('click', () => set(v));
      el.append(b);
    }
    return el;
  }

  private card(l: Look, onDelete?: () => void): HTMLElement {
    const c = l.c;
    const bar = l.barGrad.on ? `linear-gradient(${l.barGrad.angle}deg, ${c.bar}, ${l.barGrad.to})` : c.bar;
    const wall = l.wall.mode === 'gradient' ? `linear-gradient(${l.wall.angle}deg, ${l.wall.a}, ${l.wall.b})` : c.sheet;
    const sheet = h('div', { class: 'tc-sheet', style: { background: hasBackdrop(l) ? wall : c.sheet } });
    if (l.anim.id !== 'none') {
      const cv = h('canvas', { width: 136, height: 58 });
      previewScene(cv, l.anim.id, l);
      sheet.append(cv);
    }
    sheet.append(h('div', { class: 'tc-cells', style: { borderColor: c.grid, background: hasBackdrop(l) ? 'transparent' : c.sheet } }, h('b', { style: { background: c.sel } })));
    const el = h('button', { type: 'button', class: 'theme-card' + (this.look.id === l.id ? ' on' : ''), title: l.name },
      h('div', { class: 'tc-mock', style: { borderColor: c.border } }, h('div', { class: 'tc-bar', style: { background: bar } }), h('div', { class: 'tc-rib', style: { background: c.ribbon } }), sheet),
      h('span', { class: 'tc-name' }, l.name),
    );
    el.addEventListener('click', () => this.pick(l));
    if (onDelete) {
      const x = h('span', { class: 'tc-del', title: 'Delete theme', html: icon('close', 9) });
      x.addEventListener('click', (e) => {
        e.stopPropagation();
        onDelete();
      });
      el.append(x);
    }
    return el;
  }

  // ---------------- pages ----------------
  private page_themes(): void {
    const quick = h('div', { class: 'st-quick' });
    const dark = isDark(this.look);
    for (const q of QUICK) {
      const b = h('button', { type: 'button', class: 'st-dot', style: { background: q }, title: 'Build a theme from this color' });
      b.addEventListener('click', () => this.fromColor(q, dark));
      quick.append(b);
    }
    const own = h('input', { type: 'color', value: this.look.c.accent, title: 'Any color', class: 'st-dot-pick' }) as HTMLInputElement;
    own.addEventListener('input', () => this.fromColor(own.value, isDark(this.look), false));
    quick.append(own);
    const shuffle = h('button', { class: 'btn small', type: 'button', title: 'Generate a random color theme' }, 'Surprise me');
    shuffle.addEventListener('click', () => {
      const hue = Math.random() * 360;
      this.fromColor(hslHex(hue, 62, 46), Math.random() < 0.5, true, hslHex(hue + 35 + Math.random() * 50, 68, 42));
    });
    const mode = this.seg([['light', 'Light'], ['dark', 'Dark']], dark ? 'dark' : 'light', (v) => this.app.setTheme(v === 'dark'));
    this.section('Make it yours in one click', quick, h('div', { class: 'st-line' }, mode, shuffle));
    const saved = loadSaved();
    if (saved.length) {
      this.body.append(h('section', { class: 'st-sec' }, h('h3', {}, 'My themes'), h('div', { class: 'theme-grid' }, saved.map((l) => this.card(l, () => {
        storeSaved(loadSaved().filter((x) => x.id !== l.id));
        this.render();
      })))));
    }
    for (const g of PRESETS) this.body.append(h('section', { class: 'st-sec' }, h('h3', {}, g.title), h('div', { class: 'theme-grid' }, g.items.map((l) => this.card(l)))));
  }

  private fromColor(accent: string, dark: boolean, rerender = true, barTo?: string): void {
    this.edit((l) => {
      l.c = palette(accent, dark, 0.09);
      l.barGrad = barTo ? { on: true, to: barTo, angle: 90 } : { ...l.barGrad, to: mix(l.c.bar, '#000000', 0.25) };
      l.ribbonGrad.to = mix(l.c.ribbon, accent, 0.12);
    }, rerender);
  }

  private gradControls(g: Grad, key: 'barGrad' | 'ribbonGrad'): HTMLElement {
    const box = h('div', { class: 'st-sub' });
    box.append(this.row('End color', this.color(g.to, (l, v) => { l[key].to = v; })), this.slider('Direction', g.angle, 0, 360, '°', (l, v) => { l[key].angle = v; }, 5));
    return box;
  }

  private page_colors(): void {
    for (const [title, items] of COLOR_GROUPS) {
      const rows = items.map(([k, label]) => this.row(label, this.color(this.look.c[k], (l, v) => {
        l.c[k] = v;
        if (k === 'bar' && !l.barGrad.on) l.c.barText = readable(v);
      })));
      if (title === 'Top bar') {
        rows.push(this.toggle_('Gradient', this.look.barGrad.on, (l, v) => { l.barGrad.on = v; }, true));
        if (this.look.barGrad.on) rows.push(this.gradControls(this.look.barGrad, 'barGrad'));
      }
      if (title === 'Toolbars') {
        rows.push(this.toggle_('Ribbon gradient', this.look.ribbonGrad.on, (l, v) => { l.ribbonGrad.on = v; }, true));
        if (this.look.ribbonGrad.on) rows.push(this.gradControls(this.look.ribbonGrad, 'ribbonGrad'));
      }
      this.section(title, ...rows);
    }
  }

  private page_background(): void {
    const w = this.look.wall;
    const mode = this.seg([['none', 'None'], ['gradient', 'Gradient'], ['image', 'Picture']], w.mode, (v) => {
      if (v === 'image') return this.choosePicture();
      this.edit((l) => {
        l.wall.mode = v;
        if (v !== 'none') revealBackdrop(l);
      }, true);
    });
    const kids: HTMLElement[] = [mode];
    if (w.mode === 'gradient') {
      const swatches = h('div', { class: 'st-grads' });
      for (const [a, b] of GRADIENTS) {
        const s = h('button', { type: 'button', style: { background: `linear-gradient(135deg, ${a}, ${b})` }, title: 'Use this gradient' });
        s.addEventListener('click', () => this.edit((l) => { l.wall.a = a; l.wall.b = b; }, true));
        swatches.append(s);
      }
      kids.push(swatches, this.row('Start color', this.color(w.a, (l, v) => { l.wall.a = v; })), this.row('End color', this.color(w.b, (l, v) => { l.wall.b = v; })), this.toggle_('Radial', w.radial, (l, v) => { l.wall.radial = v; }, true));
      if (!w.radial) kids.push(this.slider('Direction', w.angle, 0, 360, '°', (l, v) => { l.wall.angle = v; }, 5));
    } else if (w.mode === 'image') {
      const again = h('button', { class: 'btn small', type: 'button' }, 'Choose another picture');
      again.addEventListener('click', () => this.choosePicture());
      kids.push(again);
    }
    this.section('Wallpaper', ...kids);

    const grid = h('div', { class: 'tex-grid' });
    const dark = isDark(this.look);
    const none = h('button', { type: 'button', class: 'tex' + (this.look.texture.id === 'none' ? ' on' : '') }, h('span', { class: 'tex-img' }), h('small', {}, 'None'));
    none.addEventListener('click', () => this.edit((l) => { l.texture.id = 'none'; }, true));
    grid.append(none);
    for (const [id, name] of TEXTURES) {
      const img = h('span', { class: 'tex-img', style: { backgroundImage: `url("${textureUrl(id, dark)}")` } });
      const b = h('button', { type: 'button', class: 'tex' + (this.look.texture.id === id ? ' on' : '') }, img, h('small', {}, name));
      b.addEventListener('click', () => this.edit((l) => {
        l.texture.id = id;
        revealBackdrop(l);
      }, true));
      grid.append(b);
    }
    const t = this.look.texture;
    this.section('Texture', grid, t.id === 'none' ? null : this.slider('Strength', t.strength, 5, 100, '%', (l, v) => { l.texture.strength = v; }), t.id === 'none' ? null : this.slider('Size', t.scale, 50, 250, '%', (l, v) => { l.texture.scale = v; }, 10));
    this.seeThrough();
  }

  private seeThrough(): void {
    if (!hasBackdrop(this.look)) return;
    this.section('See-through',
      this.slider('Sheet', this.look.sheetAlpha, 15, 100, '%', (l, v) => { l.sheetAlpha = v; }),
      this.slider('Toolbars', this.look.uiAlpha, 30, 100, '%', (l, v) => { l.uiAlpha = v; }),
      h('p', { class: 'st-note' }, 'Lower values show more background. Keep the sheet higher if text gets hard to read.'),
    );
  }

  private choosePicture(): void {
    const input = h('input', { type: 'file', accept: 'image/png,image/jpeg,image/webp' }) as HTMLInputElement;
    input.addEventListener('change', async () => {
      const file = input.files?.[0];
      if (!file) return;
      try {
        const bmp = await createImageBitmap(file);
        const k = Math.min(1, 1600 / Math.max(bmp.width, bmp.height));
        const cv = h('canvas', { width: Math.round(bmp.width * k), height: Math.round(bmp.height * k) });
        cv.getContext('2d')!.drawImage(bmp, 0, 0, cv.width, cv.height);
        bmp.close();
        if (!storeImage(cv.toDataURL('image/jpeg', 0.78))) return toast('That picture is too large to keep.');
        this.edit((l) => {
          l.wall.mode = 'image';
          revealBackdrop(l);
        }, true);
        this.app.applyLookImage();
      } catch {
        toast('Sheetly could not read that picture.');
      }
    });
    input.click();
  }

  private page_animation(): void {
    const grid = h('div', { class: 'anim-grid' });
    const none = h('button', { type: 'button', class: 'anim' + (this.look.anim.id === 'none' ? ' on' : '') }, h('span', { class: 'anim-none' }, 'Off'), h('small', {}, 'No animation'));
    none.addEventListener('click', () => this.edit((l) => { l.anim.id = 'none'; }, true));
    grid.append(none);
    for (const [id, name] of SCENES) {
      const cv = h('canvas', { width: 148, height: 84 });
      previewScene(cv, id, this.look);
      const b = h('button', { type: 'button', class: 'anim' + (this.look.anim.id === id ? ' on' : '') }, cv, h('small', {}, name));
      b.addEventListener('click', () => this.edit((l) => {
        l.anim.id = id;
        revealBackdrop(l);
      }, true));
      grid.append(b);
    }
    this.section('Animated background', grid);
    const a = this.look.anim;
    if (a.id !== 'none') {
      this.section('Tune it',
        this.slider('Speed', a.speed, 25, 250, '%', (l, v) => { l.anim.speed = v; }, 5),
        this.slider('Amount', a.amount, 25, 150, '%', (l, v) => { l.anim.amount = v; }, 5),
        h('div', { class: 'st-row' }, h('span', {}, 'Performance'), this.seg([['eco', 'Saver'], ['balanced', 'Balanced'], ['smooth', 'Smooth']], a.quality, (v) => this.edit((l) => { l.anim.quality = v; }, true))),
        this.toggle_('Pause when Sheetly is not focused', a.pauseUnfocused, (l, v) => { l.anim.pauseUnfocused = v; }),
        this.app.wallpaper.reducedMotion ? h('p', { class: 'st-note' }, 'Windows animation effects are off, so the background stays still.') : null,
      );
      this.seeThrough();
    }
    this.section('Interface', this.toggle_('Animate menus, dialogs and hover', this.look.motion, (l, v) => { l.motion = v; }));
  }

  private page_style(): void {
    const font = h('select', {}, FONTS.map(([v, n]) => h('option', { value: v, selected: v === this.look.font }, n))) as HTMLSelectElement;
    font.addEventListener('change', () => this.edit((l) => { l.font = font.value; }));
    this.section('Text', this.row('Interface font', font), this.slider('Text size', this.look.fontSize, 11, 14, 'px', (l, v) => { l.fontSize = v; }));
    this.section('Shape', this.slider('Rounded corners', this.look.radius, 0, 12, 'px', (l, v) => { l.radius = v; }));
    this.body.append(h('p', { class: 'st-note' }, 'Appearance belongs to this copy of Sheetly. It is never saved into your spreadsheets.'));
  }
}

function hslHex(hue: number, s: number, l: number): string {
  const S = s / 100;
  const L = l / 100;
  const f = (n: number) => {
    const k = (n + ((hue % 360) + 360) % 360 / 30) % 12;
    const v = L - S * Math.min(L, 1 - L) * Math.max(-1, Math.min(k - 3, 9 - k, 1));
    return Math.round(v * 255).toString(16).padStart(2, '0');
  };
  return `#${f(0)}${f(8)}${f(4)}`;
}
