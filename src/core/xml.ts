// Tiny XML reader (enough for OOXML) and escaping helpers. Works in the browser and in Node.

export interface XNode {
  tag: string; // local name (namespace prefix removed)
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// eslint-disable-next-line no-control-regex
const BAD_XML = /[\x00-\x08\x0b\x0c\x0e-\x1f]/g;
export function escText(s: string): string {
  return esc(s).replace(BAD_XML, '');
}

function unesc(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_m, e: string) => {
    switch (e) {
      case 'lt': return '<';
      case 'gt': return '>';
      case 'amp': return '&';
      case 'quot': return '"';
      case 'apos': return "'";
    }
    const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return Number.isInteger(code) && code >= 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : '�';
  });
}

function local(name: string): string {
  const i = name.indexOf(':');
  return i < 0 ? name : name.slice(i + 1);
}

export function parseXml(src: string): XNode {
  const root: XNode = { tag: '#root', attrs: {}, children: [], text: '' };
  const stack: XNode[] = [root];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) break;
    if (lt > i) {
      const top = stack[stack.length - 1];
      top.text += unesc(src.slice(i, lt));
    }
    if (src.startsWith('<!--', lt)) {
      const e = src.indexOf('-->', lt + 4);
      i = e < 0 ? n : e + 3;
      continue;
    }
    if (src.startsWith('<![CDATA[', lt)) {
      const e = src.indexOf(']]>', lt + 9);
      const top = stack[stack.length - 1];
      top.text += src.slice(lt + 9, e < 0 ? n : e);
      i = e < 0 ? n : e + 3;
      continue;
    }
    if (src[lt + 1] === '?' || src[lt + 1] === '!') {
      const e = src.indexOf('>', lt);
      i = e < 0 ? n : e + 1;
      continue;
    }
    if (src[lt + 1] === '/') {
      const e = src.indexOf('>', lt);
      stack.pop();
      if (!stack.length) stack.push(root);
      i = e < 0 ? n : e + 1;
      continue;
    }
    // open tag: find matching '>' while respecting quotes
    let j = lt + 1;
    let q = '';
    while (j < n) {
      const ch = src[j];
      if (q) {
        if (ch === q) q = '';
      } else if (ch === '"' || ch === "'") q = ch;
      else if (ch === '>') break;
      j++;
    }
    let body = src.slice(lt + 1, j);
    const selfClose = body.endsWith('/');
    if (selfClose) body = body.slice(0, -1);
    const m = /^([^\s/>]+)/.exec(body)!;
    const node: XNode = { tag: local(m[1]), attrs: {}, children: [], text: '' };
    const attrRe = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
    let am: RegExpExecArray | null;
    const rest = body.slice(m[1].length);
    while ((am = attrRe.exec(rest))) node.attrs[am[1]] = unesc(am[3] ?? am[4] ?? '');
    stack[stack.length - 1].children.push(node);
    if (!selfClose) stack.push(node);
    i = j + 1;
  }
  return root.children[0] ?? root;
}

export function child(n: XNode | undefined, tag: string): XNode | undefined {
  return n?.children.find((c) => c.tag === tag);
}
export function kids(n: XNode | undefined, tag: string): XNode[] {
  return n ? n.children.filter((c) => c.tag === tag) : [];
}
