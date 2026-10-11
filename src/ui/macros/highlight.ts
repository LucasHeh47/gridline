// Syntax coloring for VBA code in the macro editor.

const KEYWORDS = new Set(
  (
    'and as boolean byref byte byval call case const currency date decimal declare dim do double each else elseif empty end enum ' +
    'eqv erase error exit explicit false for friend function get global gosub goto if imp implements in integer is let like long ' +
    'longlong longptr loop me mod new next not nothing null object on option optional or paramarray preserve private property ' +
    'public redim rem resume return select set single static step stop string sub then to true type typeof until variant wend ' +
    'while with withevents xor base compare text binary attribute'
  ).split(' '),
);

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** One line of code as HTML with <span class="hl-…"> tokens. */
export function highlightLine(line: string): string {
  let out = '';
  let i = 0;
  const n = line.length;
  while (i < n) {
    const ch = line[i];
    if (ch === "'") {
      out += `<span class="hl-com">${esc(line.slice(i))}</span>`;
      break;
    }
    if (ch === '"') {
      let j = i + 1;
      while (j < n) {
        if (line[j] === '"') {
          if (line[j + 1] === '"') {
            j += 2;
            continue;
          }
          j++;
          break;
        }
        j++;
      }
      out += `<span class="hl-str">${esc(line.slice(i, j))}</span>`;
      i = j;
      continue;
    }
    if (/[0-9]/.test(ch) || (ch === '&' && /[hHoO]/.test(line[i + 1] ?? ''))) {
      const m = /^(&[hHoO][0-9a-fA-F]+&?|\d*\.?\d+(?:[eE][+-]?\d+)?[%&!#@]?)/.exec(line.slice(i));
      if (m) {
        out += `<span class="hl-num">${esc(m[0])}</span>`;
        i += m[0].length;
        continue;
      }
    }
    if (ch === '#') {
      const end = line.indexOf('#', i + 1);
      if (end > i && /\d/.test(line.slice(i, end))) {
        out += `<span class="hl-num">${esc(line.slice(i, end + 1))}</span>`;
        i = end + 1;
        continue;
      }
    }
    if (/[A-Za-z_]/.test(ch)) {
      const m = /^[A-Za-z_][A-Za-z0-9_]*[$%&!#@]?/.exec(line.slice(i))!;
      const w = m[0];
      const lower = w.toLowerCase();
      if (lower === 'rem' && !line.slice(0, i).trim()) {
        out += `<span class="hl-com">${esc(line.slice(i))}</span>`;
        break;
      }
      const prev = line.slice(0, i).trimEnd();
      if (KEYWORDS.has(lower) && !prev.endsWith('.')) out += `<span class="hl-kw">${esc(w)}</span>`;
      else if (/^(sub|function|property)$/i.test(prev.split(/\s+/).pop() ?? '')) out += `<span class="hl-fn">${esc(w)}</span>`;
      else out += esc(w);
      i += w.length;
      continue;
    }
    out += esc(ch);
    i++;
  }
  return out;
}

export function highlightCode(code: string): string {
  return code
    .split('\n')
    .map((l) => highlightLine(l.replace(/\r$/, '')))
    .join('\n');
}
