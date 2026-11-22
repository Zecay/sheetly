// Original line icons (24x24 grid). `a` marks accent-coloured parts via CSS class.
const P = (d: string, cls = '') => `<path d="${d}"${cls ? ` class="${cls}"` : ''}/>`;
const R = (x: number, y: number, w: number, h: number, cls = '', rx = 0) => `<rect x="${x}" y="${y}" width="${w}" height="${h}"${rx ? ` rx="${rx}"` : ''}${cls ? ` class="${cls}"` : ''}/>`;
const T = (x: number, y: number, s: string, size = 12, cls = '') => `<text x="${x}" y="${y}" font-size="${size}" font-family="Segoe UI, Arial" font-weight="600"${cls ? ` class="${cls}"` : ''} stroke="none" fill="currentColor">${s}</text>`;

// Border pictograms: faint dotted guide square (+ optional cross) with the relevant edges drawn bold (`bk`) or heavy (`bh`).
const BX = { t: 'M4 4h16', b: 'M4 20h16', l: 'M4 4v16', r: 'M20 4v16', h: 'M4 12h16', v: 'M12 4v16' };
const bIcon = (guide: string[], bold: string[], heavy: string[] = []) =>
  `<g class="bf">${P('M4 4h16v16H4z')}${guide.map((k) => P(BX[k as keyof typeof BX])).join('')}</g>` +
  bold.map((k) => P(BX[k as keyof typeof BX], 'bk')).join('') +
  heavy.map((k) => P(BX[k as keyof typeof BX], 'bh')).join('');

const defs: Record<string, string> = {
  bdBottom: bIcon([], ['b']),
  bdTop: bIcon([], ['t']),
  bdLeft: bIcon([], ['l']),
  bdRight: bIcon([], ['r']),
  bdNone: bIcon(['h', 'v'], []),
  bdAll: bIcon([], ['t', 'b', 'l', 'r', 'h', 'v']),
  bdOutside: bIcon(['h', 'v'], ['t', 'b', 'l', 'r']),
  bdThickOutside: bIcon(['h', 'v'], [], ['t', 'b', 'l', 'r']),
  bdInside: bIcon([], ['h', 'v']),
  bdBottomDouble: bIcon([], [], []) + P('M4 17h16', 'bk') + P('M4 21h16', 'bk'),
  bdThickBottom: bIcon([], [], ['b']),
  bdTopBottom: bIcon([], ['t', 'b']),
  bdTopThickBottom: bIcon([], ['t'], ['b']),
  save: P('M5 3h12l3 3v15H5z') + P('M8 3v5h8V3') + R(8, 14, 8, 6),
  undo: P('M9 8H15a5 5 0 010 10H8') + P('M9 4L5 8l4 4'),
  redo: P('M15 8H9a5 5 0 000 10h7') + P('M15 4l4 4-4 4'),
  cut: P('M6 3l9 12M18 3L9 15') + `<circle cx="7" cy="18" r="3"/><circle cx="17" cy="18" r="3"/>`,
  copy: R(8, 8, 12, 13, '', 1) + P('M16 8V4H5v13h3'),
  paste: R(5, 5, 14, 16, '', 1) + R(9, 3, 6, 4, 'f') + P('M9 12h6M9 16h6'),
  painter: R(4, 4, 14, 6, '', 1) + P('M18 7h2v6H11v3') + R(9, 16, 4, 5, 'a'),
  bold: P('M7 4h6a4 4 0 010 8H7zM7 12h7a4 4 0 010 8H7z'),
  italic: P('M10 4h8M6 20h8M14 4L10 20'),
  underline: P('M7 4v7a5 5 0 0010 0V4M5 21h14'),
  strike: P('M4 12h16') + P('M17 7c-1-2-3-3-5.5-3C9 4 7 5.500 7 7.500c0 1.500 1 2.500 3 3M14 13c2 .800 3 1.800 3 3.500 0 2-2 3.500-5 3.500-2.500 0-4.500-1-5.500-3'),
  borders: R(4, 4, 16, 16) + P('M12 4v16M4 12h16', 'dim'),
  fill: P('M5 13l7-8 7 8-6 6z') + P('M19 17c1 1.500 1.500 2.500 1 3.500', '') + R(3, 21, 18, 2, 'a'),
  fontcolor: P('M6 17l6-13 6 13M8.500 13h7') + R(4, 20, 16, 3, 'a'),
  alignTop: P('M4 4h16M12 20V8M8 12l4-4 4 4'),
  alignMiddle: P('M4 12h16M12 4v5M12 20v-5M9 6l3 3 3-3M9 18l3-3 3 3'),
  alignBottom: P('M4 20h16M12 4v12M8 12l4 4 4-4'),
  alignLeft: P('M4 5h16M4 10h10M4 15h16M4 20h10'),
  alignCenter: P('M4 5h16M7 10h10M4 15h16M7 20h10'),
  alignRight: P('M4 5h16M10 10h10M4 15h16M10 20h10'),
  wrap: P('M4 6h16M4 12h13a3 3 0 010 6h-4') + P('M15 15l-2 3 2 3') + P('M4 18h5'),
  merge: R(3, 6, 18, 12, '', 1) + P('M3 12h5M21 12h-5M6 9l-3 3 3 3M18 9l3 3-3 3'),
  indentInc: P('M4 5h16M12 10h8M12 15h8M4 20h16M4 9l4 3-4 3z'),
  indentDec: P('M4 5h16M12 10h8M12 15h8M4 20h16M8 9L4 12l4 3z'),
  currency: `<circle cx="12" cy="12" r="8"/>` + P('M14.500 9.500c-.500-1-1.500-1.500-2.500-1.500-1.500 0-2.500.800-2.500 2s1 1.600 2.500 2 2.500.900 2.500 2-1 2-2.500 2c-1.200 0-2.200-.600-2.600-1.600M12 6.500V8M12 16v1.500'),
  percent: P('M18 5L6 19') + `<circle cx="7.500" cy="7.500" r="2.500"/><circle cx="16.500" cy="16.500" r="2.500"/>`,
  comma: T(7, 14, '000', 9) + P('M10 19c1 0 1 2-1 3'),
  decInc: T(2, 13, '.0', 11) + P('M15 8l-4 4 4 4', 'a') + P('M11 12h10', 'a') + T(15, 21, '.00', 7),
  decDec: T(2, 13, '.00', 9) + P('M20 8l4 4-4 4', 'dim') + P('M24 12H14', 'dim'),
  condfmt: R(3, 4, 18, 16, '', 1) + P('M3 9h18M3 14h18M9 4v16') + R(10, 10, 5, 3, 'a') + R(3.500, 15, 5, 4, 'b'),
  styles: R(3, 4, 8, 6, 'g', 1) + R(13, 4, 8, 6, 'r', 1) + R(3, 14, 8, 6, 'y', 1) + R(13, 14, 8, 6, '', 1),
  insertCells: R(4, 8, 12, 12, '', 1) + P('M16 4v8M12 8h8', 'a'),
  deleteCells: R(4, 8, 12, 12, '', 1) + P('M12 6l8 8M20 6l-8 8', 'r'),
  formatCells: R(3, 5, 18, 14, '', 1) + P('M3 10h18M9 10v9'),
  autosum: P('M17 5H7l6 7-6 7h10', 'a'),
  fillDown: P('M12 4v14M7 13l5 5 5-5') + P('M5 21h14'),
  eraser: P('M4 16l9-11 7 6-6 9H8z') + P('M9 21h11'),
  sort: P('M6 4v16M3 17l3 3 3-3M14 6h7M14 11h5M14 16h3'),
  filter: P('M3 5h18l-7 8v6l-4 2v-8z'),
  find: `<circle cx="10" cy="10" r="6"/>` + P('M15 15l6 6'),
  search: `<circle cx="10.500" cy="10.500" r="5.500"/>` + P('M15 15l5 5'),
  fx: T(3, 17, 'fx', 14),
  names: R(3, 6, 18, 12, '', 1) + P('M7 10h4M7 14h10'),
  defineName: P('M4 7l3-3h13v10l-3 3H4z') + `<circle cx="8" cy="9" r="1.200"/>`,
  freeze: R(3, 4, 18, 16) + P('M3 10h18M10 4v16') + R(3, 4, 7, 6, 'a'),
  zoom: `<circle cx="10" cy="10" r="6"/>` + P('M15 15l6 6M7.500 10h5M10 7.500v5'),
  gridlines: R(3, 4, 18, 16) + P('M3 10h18M3 15h18M9 4v16M15 4v16'),
  sheetAdd: R(4, 5, 16, 14, '', 1) + P('M12 9v6M9 12h6', 'a'),
  rowsInsert: R(4, 4, 16, 6) + R(4, 14, 16, 6) + P('M12 10.500v3M10.500 12h3', 'a'),
  colsInsert: R(4, 4, 6, 16) + R(14, 4, 6, 16) + P('M10.500 12h3M12 10.500v3', 'a'),
  print: P('M7 9V3h10v6') + R(4, 9, 16, 8, '', 1) + R(7, 14, 10, 7),
  pdf: P('M6 3h9l4 4v14H6z') + P('M15 3v4h4') + T(7.500, 18, 'PDF', 6.500),
  newDoc: P('M6 3h9l4 4v14H6z') + P('M15 3v4h4M12 11v6M9 14h6', 'a'),
  open: P('M3 7h7l2 2h9v10H3z') + P('M3 11h18'),
  saveAs: P('M5 3h11l4 4v14H5z') + P('M8 3v5h7V3') + P('M15 14l3 3 3-3', 'a'),
  exportIcon: P('M12 3v12M7 8l5-5 5 5') + P('M5 15v5h14v-5'),
  info: `<circle cx="12" cy="12" r="9"/>` + P('M12 11v6M12 7.500v.500'),
  close: P('M5 5l14 14M19 5L5 19'),
  minimize: P('M5 12h14'),
  maximize: R(5, 5, 14, 14),
  restore: R(5, 8, 11, 11) + P('M8 8V5h11v11h-3'),
  moon: P('M20 14A8 8 0 019 3a8 8 0 1011 11z'),
  sun: `<circle cx="12" cy="12" r="4"/>` + P('M12 2v3M12 19v3M2 12h3M19 12h3M5 5l2 2M17 17l2 2M19 5l-2 2M7 17l-2 2'),
  chevron: P('M6 9l6 6 6-6'),
  plus: P('M12 5v14M5 12h14'),
  left: P('M15 5l-7 7 7 7'),
  right: P('M9 5l7 7-7 7'),
  check: P('M5 12l5 5 9-10'),
  table: R(3, 4, 18, 16) + P('M3 9h18M9 4v16') + R(3, 4, 18, 5, 'a'),
  unfreeze: R(3, 4, 18, 16) + P('M3 10h18M10 4v16'),
  header: R(3, 4, 18, 16) + P('M3 9h18') + P('M7 13h10M7 16h6', 'dim'),
  calc: R(5, 3, 14, 18, '', 1) + R(8, 6, 8, 3, 'a') + P('M8 13h2M14 13h2M8 17h2M14 17h2'),
  stats: R(4, 12, 4, 8) + R(10, 6, 4, 14) + R(16, 9, 4, 11),
  paintbrush: P('M14 4l6 6-8 8H6v-6z') + P('M4 20h4'),
  trash: P('M5 7h14M9 7V4h6v3M7 7l1 13h8l1-13'),
  window: R(4, 5, 16, 14) + P('M4 9h16'),
  palette: P('M12 3a9 9 0 100 18c1.4 0 2-.9 2-1.8 0-1.3-1-1.7-1-2.7 0-.9.7-1.5 1.7-1.5H17a4 4 0 004-4c0-4.4-4-8-9-8z') + P('M7.5 11.5h.01M10 7.5h.01M15 7.5h.01', 'a'),
};

export function icon(name: string, size = 16): string {
  const body = defs[name];
  if (!body) return '';
  return `<svg class="ic" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;
}

/** The app mark: a green grid with an "S". */
export function appMark(size = 16): string {
  return `<svg width="${size}" height="${size}" viewBox="0 0 24 24" aria-hidden="true"><rect x="1" y="1" width="22" height="22" rx="4" fill="#fff" fill-opacity=".95"/><rect x="1" y="1" width="22" height="22" rx="4" fill="none" stroke="#14592f" stroke-width="1"/><path d="M8 1v22M16 1v22M1 8h22M1 16h22" stroke="#217346" stroke-opacity=".28"/><path d="M16.200 7.800C15.500 6.800 14.200 6.200 12.600 6.200 10.500 6.200 9 7.200 9 8.800c0 1.600 1.200 2.200 3.200 2.800 2.300.600 3.600 1.300 3.600 3 0 1.600-1.400 2.800-3.600 2.800-1.500 0-2.800-.600-3.700-1.700" fill="none" stroke="#1d7a45" stroke-width="2" stroke-linecap="round"/></svg>`;
}
