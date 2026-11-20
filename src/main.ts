import './style.css';
import { App } from './ui/app';

window.addEventListener('DOMContentLoaded', () => {
  // Disable the browser/WebView context menu and default drag behaviour; Sheetly has its own menus.
  window.addEventListener('contextmenu', (e) => {
    const t = e.target as HTMLElement;
    if (!(t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement)) e.preventDefault();
  });
  new App();
});
