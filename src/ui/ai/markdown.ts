// A small, safe Markdown renderer for the AI assistant's replies (builds DOM nodes; never parses
// HTML). Cell references like Data!B7 or A2:C9 become links that select those cells.
import { h } from '../dom';

export interface RefLinks {
  /** True when `ref` points at real cells (so "e.g!A1" isn't linked). */
  valid(ref: string): boolean;
  open(ref: string): void;
}

const CELL = String.raw`\$?[A-Z]{1,3}\$?[1-9]\d{0,6}`;
const QUALIFIED = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!${CELL}(?::${CELL})?`;
const REF_RE = new RegExp(String.raw`(${QUALIFIED})|\b(${CELL}:${CELL})\b`, 'g');
const INLINE_RE = /(`[^`]+`)|(\*\*[^*]+\*\*)|(\*[^*\s][^*]*\*)/g;

function refLink(ref: string, links: RefLinks, code = false): HTMLElement {
  return h('a', { class: `ai-ref${code ? ' code' : ''}`, href: '#', title: `Go to ${ref}`, onclick: (e: Event) => {
    e.preventDefault();
    links.open(ref);
  } }, ref);
}

/** Plain text with cell references turned into links (no other formatting). */
export function withRefs(text: string, links: RefLinks | null): (Node | string)[] {
  if (!links) return [text];
  const out: (Node | string)[] = [];
  let last = 0;
  for (const m of text.matchAll(REF_RE)) {
    const ref = m[0];
    if (!links.valid(ref)) continue;
    if (m.index! > last) out.push(text.slice(last, m.index));
    out.push(refLink(ref, links));
    last = m.index! + ref.length;
  }
  if (last < text.length) out.push(text.slice(last));
  return out;
}

export function inline(text: string, links: RefLinks | null): (Node | string)[] {
  const out: (Node | string)[] = [];
  let last = 0;
  for (const m of text.matchAll(INLINE_RE)) {
    if (m.index! > last) out.push(...withRefs(text.slice(last, m.index), links));
    const t = m[0];
    if (m[1]) {
      const code = t.slice(1, -1);
      out.push(links && new RegExp(`^(?:${QUALIFIED}|${CELL}(?::${CELL})?)$`).test(code) && links.valid(code) ? refLink(code, links, true) : h('code', null, code));
    } else if (m[2]) out.push(h('strong', null, ...inline(t.slice(2, -2), links)));
    else out.push(h('em', null, ...inline(t.slice(1, -1), links)));
    last = m.index! + t.length;
  }
  if (last < text.length) out.push(...withRefs(text.slice(last), links));
  return out;
}

function cells(row: string): string[] {
  return row.trim().replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
}

/** Renders Markdown text (paragraphs, lists, headings, code, tables) into a container. */
export function renderMarkdown(text: string, links: RefLinks | null): HTMLElement {
  const root = h('div', { class: 'ai-md' });
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) {
      i++;
      continue;
    }
    // Code block
    if (/^\s*```/.test(line)) {
      const body: string[] = [];
      i++;
      while (i < lines.length && !/^\s*```/.test(lines[i])) body.push(lines[i++]);
      i++;
      root.append(h('pre', null, h('code', null, body.join('\n'))));
      continue;
    }
    // Heading
    const hm = /^(#{1,4})\s+(.*)$/.exec(line);
    if (hm) {
      root.append(h('div', { class: `ai-h ai-h${hm[1].length}` }, ...inline(hm[2], links)));
      i++;
      continue;
    }
    // Table
    if (/^\s*\|/.test(line) && i + 1 < lines.length && /^\s*\|?\s*:?-{2,}/.test(lines[i + 1])) {
      const head = cells(line);
      i += 2;
      const rows: string[][] = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) rows.push(cells(lines[i++]));
      root.append(
        h('div', { class: 'ai-table-wrap' }, h('table', { class: 'ai-table' }, h('thead', null, h('tr', null, ...head.map((c) => h('th', null, ...inline(c, links))))), h('tbody', null, ...rows.map((r) => h('tr', null, ...r.map((c) => h('td', null, ...inline(c, links)))))))),
      );
      continue;
    }
    // Lists
    const listRe = /^(\s*)([-*•]|\d+[.)])\s+(.*)$/;
    if (listRe.test(line)) {
      const ordered = /^\s*\d/.test(line);
      const list = h(ordered ? 'ol' : 'ul', null);
      while (i < lines.length && listRe.test(lines[i])) {
        const m = listRe.exec(lines[i])!;
        const item = h('li', null, ...inline(m[3], links));
        i++;
        // Continuation lines (indented, not a new item)
        while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !listRe.test(lines[i])) {
          item.append(h('br'), ...inline(lines[i].trim(), links));
          i++;
        }
        if (m[1].length >= 2 && list.lastElementChild) {
          // Nested item: attach to the previous one
          let sub = list.lastElementChild.querySelector(':scope > ul, :scope > ol');
          if (!sub) {
            sub = h(ordered ? 'ol' : 'ul', null);
            list.lastElementChild.append(sub);
          }
          sub.append(item);
        } else list.append(item);
      }
      root.append(list);
      continue;
    }
    // Paragraph
    const para = h('p', null, ...inline(line, links));
    i++;
    while (i < lines.length && lines[i].trim() && !/^\s*```/.test(lines[i]) && !/^#{1,4}\s/.test(lines[i]) && !listRe.test(lines[i]) && !/^\s*\|/.test(lines[i])) {
      para.append(h('br'), ...inline(lines[i], links));
      i++;
    }
    root.append(para);
  }
  return root;
}
