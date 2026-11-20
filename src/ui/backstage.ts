// The File menu ("backstage").
import { h } from './dom';
import { icon } from './icons';
import { loadRecent, removeRecent } from './fsapi';
import type { App } from './app';

export class Backstage {
  private el = document.getElementById('backstage')!;
  private page = 'info';
  private keyHandler = (e: KeyboardEvent) => {
    if (e.key === 'Escape' && !this.el.hidden && !document.querySelector('.dlg-backdrop')) {
      e.stopPropagation();
      this.hide();
    }
  };

  constructor(private app: App) {
    document.addEventListener('keydown', this.keyHandler, true);
  }

  show(page = 'info'): void {
    this.page = page;
    this.el.hidden = false;
    this.render();
  }

  hide(): void {
    this.el.hidden = true;
    this.app.grid.focusGrid();
  }

  private render(): void {
    const a = this.app;
    const items: [string, string, string][] = [
      ['info', 'Info', 'info'],
      ['new', 'New', 'newDoc'],
      ['open', 'Open', 'open'],
      ['save', 'Save', 'save'],
      ['saveAs', 'Save As', 'saveAs'],
      ['export', 'Export', 'exportIcon'],
      ['print', 'Print', 'print'],
      ['options', 'Options', 'window'],
    ];
    const side = h('div', { class: 'bs-side' });
    const back = h('button', { class: 'back', type: 'button', title: 'Back', html: icon('left', 20) });
    back.addEventListener('click', () => this.hide());
    side.append(back);
    for (const [id, label, ic] of items) {
      const b = h('button', { type: 'button', class: this.page === id ? 'on' : '' }, h('span', { html: icon(ic, 18) }), label);
      b.addEventListener('click', () => {
        if (id === 'save') {
          this.hide();
          void a.run('save');
        } else if (id === 'saveAs' ) {
          this.page = id;
          this.render();
        } else {
          this.page = id;
          this.render();
        }
      });
      side.append(b);
    }
    const close = h('button', { type: 'button', style: { marginTop: 'auto', marginBottom: '8px' } }, h('span', { html: icon('close', 18) }), 'Close');
    close.addEventListener('click', () => {
      this.hide();
      void a.newWorkbook();
    });
    side.append(close);
    const main = h('div', { class: 'bs-main' });
    this.fillMain(main);
    this.el.innerHTML = '';
    this.el.append(side, main);
  }

  private card(ic: string, title: string, desc: string, fn: () => void): HTMLElement {
    const c = h('div', { class: 'bs-card' }, h('span', { html: icon(ic, 32) }), h('div', {}, h('div', { class: 't' }, title), h('div', { class: 'd' }, desc)));
    c.addEventListener('click', fn);
    return c;
  }

  private fillMain(main: HTMLElement): void {
    const a = this.app;
    const doRun = (cmd: string, arg?: unknown) => {
      this.hide();
      void a.run(cmd, arg);
    };
    switch (this.page) {
      case 'info': {
        const wbStats = (() => {
          let cells = 0;
          let formulas = 0;
          for (const s of a.wb.sheets) for (const c of s.cells.values()) {
            if (c.v !== null || c.f !== undefined) cells++;
            if (c.f !== undefined) formulas++;
          }
          return { cells, formulas };
        })();
        main.append(
          h('h1', {}, a.doc.name),
          h('div', { class: 'bs-recent', style: { marginTop: '0' } },
            h('div', { class: 'row' }, h('span', {}, 'Location'), h('span', { class: 'path' }, a.doc.path ?? 'Not saved yet')),
            h('div', { class: 'row' }, h('span', {}, 'Sheets'), h('span', {}, String(a.wb.sheets.length))),
            h('div', { class: 'row' }, h('span', {}, 'Cells with data'), h('span', {}, wbStats.cells.toLocaleString('en-US'))),
            h('div', { class: 'row' }, h('span', {}, 'Formulas'), h('span', {}, wbStats.formulas.toLocaleString('en-US'))),
            h('div', { class: 'row' }, h('span', {}, 'Status'), h('span', {}, a.doc.dirty ? 'Has unsaved changes' : 'Saved')),
          ),
        );
        break;
      }
      case 'new':
        main.append(h('h1', {}, 'New'), this.card('newDoc', 'Blank workbook', 'Start with an empty workbook.', () => { this.hide(); void a.newWorkbook(); }));
        break;
      case 'open': {
        main.append(h('h1', {}, 'Open'), this.card('open', 'Browse…', 'Open an .xlsx, .csv or .sheetly file from your computer.', () => doRun('open')));
        const recent = loadRecent();
        const box = h('div', { class: 'bs-recent' }, h('div', { class: 'menu-heading', style: { padding: '0 0 6px' } }, 'Recent'));
        if (!recent.length) box.append(h('div', { class: 'row' }, h('span', { style: { color: 'var(--muted)' } }, 'No recent files')));
        for (const r of recent) {
          const row = h('div', { class: 'row' }, h('div', {}, h('div', {}, r.name), h('div', { class: 'path' }, r.path)), h('span', { style: { color: 'var(--muted)' } }, new Date(r.time).toLocaleDateString()));
          row.addEventListener('click', async () => {
            this.hide();
            await a.openPath(r.path);
          });
          row.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            removeRecent(r.path);
            this.render();
          });
          box.append(row);
        }
        main.append(box);
        break;
      }
      case 'saveAs':
        main.append(
          h('h1', {}, 'Save As'),
          this.card('save', 'Excel Workbook (.xlsx)', 'Values, formulas, formatting, sheet names and column widths.', () => { this.hide(); void a.saveAs('xlsx'); }),
          this.card('save', 'CSV (Comma delimited) (.csv)', 'The active sheet as plain text.', () => { this.hide(); void a.saveAs('csv'); }),
          this.card('save', 'Sheetly Workbook (.sheetly)', 'Keeps everything Sheetly supports.', () => { this.hide(); void a.saveAs('sheetly'); }),
        );
        break;
      case 'export':
        main.append(
          h('h1', {}, 'Export'),
          this.card('exportIcon', 'Export as CSV', 'Save the active sheet as a .csv file.', () => { this.hide(); void a.saveAs('csv'); }),
          this.card('pdf', 'Create PDF', 'Choose "Microsoft Print to PDF" in the print dialog.', () => doRun('exportPdf')),
        );
        break;
      case 'print':
        main.append(h('h1', {}, 'Print'), this.card('print', 'Print the active sheet', 'Opens the print dialog. Choose "Microsoft Print to PDF" to save a PDF.', () => doRun('print')));
        break;
      case 'options': {
        main.append(
          h('h1', {}, 'Options'),
          this.card(a.dark ? 'sun' : 'moon', a.dark ? 'Switch to the light theme' : 'Switch to the dark theme', 'Change the color scheme of Sheetly.', () => { a.run('theme'); this.render(); }),
          this.card('info', 'About Sheetly', 'Version and license information.', () => { this.hide(); void a.run('about'); }),
        );
        break;
      }
    }
  }
}
