// Reads and writes vbaProject.bin, the part of an .xlsm file that holds VBA code ([MS-OVBA]).
// The file is an OLE compound file: a PROJECT text stream, and a VBA storage with a "dir" stream
// (module list) and one compressed stream per module.
//
// Writing: when the code is unchanged the original bytes are kept. Otherwise the module streams are
// rebuilt with source only and the _VBA_PROJECT cache is marked stale (version 0xFFFF), which makes
// Excel recompile the code when it opens the file.
import * as CFB from 'cfb';

export type VbaModuleKind = 'standard' | 'document' | 'class' | 'form';

export interface VbaModuleSource {
  name: string;
  kind: VbaModuleKind;
  /** Full stored source, including "Attribute VB_…" lines. */
  source: string;
}

export interface VbaProjectData {
  name: string;
  codePage: number;
  modules: VbaModuleSource[];
}

// ---------- compression ([MS-OVBA] 2.4.1) ----------

function copyTokenBits(difference: number): number {
  let bits = 0;
  while (1 << bits < difference) bits++;
  return Math.max(bits, 4);
}

export function decompress(data: Uint8Array, start = 0): Uint8Array {
  if (data[start] !== 1) throw new Error('VBA data is damaged (bad compression signature).');
  const out: number[] = [];
  let pos = start + 1;
  while (pos < data.length - 1) {
    const header = data[pos] | (data[pos + 1] << 8);
    const size = (header & 0x0fff) + 3;
    const compressed = (header & 0x8000) !== 0;
    const chunkEnd = Math.min(pos + size, data.length);
    pos += 2;
    const chunkStart = out.length;
    if (!compressed) {
      for (let i = 0; i < 4096 && pos < data.length; i++) out.push(data[pos++]);
      continue;
    }
    while (pos < chunkEnd) {
      const flags = data[pos++];
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        if (((flags >> bit) & 1) === 0) {
          out.push(data[pos++]);
          continue;
        }
        const token = data[pos] | (data[pos + 1] << 8);
        pos += 2;
        const bits = copyTokenBits(out.length - chunkStart);
        const lengthMask = 0xffff >> bits;
        const length = (token & lengthMask) + 3;
        const offset = (token >> (16 - bits)) + 1;
        const from = out.length - offset;
        if (from < 0) throw new Error('VBA data is damaged (bad copy token).');
        for (let i = 0; i < length; i++) out.push(out[from + i]);
      }
    }
    pos = chunkEnd;
  }
  return Uint8Array.from(out);
}

export function compress(data: Uint8Array): Uint8Array {
  const out: number[] = [1];
  for (let start = 0; start < data.length; start += 4096) {
    const end = Math.min(start + 4096, data.length);
    const chunk = compressChunk(data, start, end);
    if (chunk.length > 4096) {
      // Incompressible: store raw (always 4096 bytes, padded)
      out.push(0xff, 0x3f);
      for (let i = start; i < start + 4096; i++) out.push(i < end ? data[i] : 0);
    } else {
      const header = 0xb000 | (chunk.length + 2 - 3);
      out.push(header & 0xff, header >> 8);
      for (const b of chunk) out.push(b);
    }
  }
  return Uint8Array.from(out);
}

function compressChunk(data: Uint8Array, start: number, end: number): number[] {
  const out: number[] = [];
  // Positions of each 3-byte sequence seen so far in this chunk, newest last
  const seen = new Map<number, number[]>();
  const key = (i: number) => (data[i] << 16) | (data[i + 1] << 8) | data[i + 2];
  const remember = (i: number) => {
    if (i + 2 >= end) return;
    const k = key(i);
    let list = seen.get(k);
    if (!list) seen.set(k, (list = []));
    list.push(i);
  };
  let pos = start;
  while (pos < end) {
    const flagAt = out.length;
    out.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && pos < end; bit++) {
      let bestLen = 0;
      let bestOff = 0;
      const difference = pos - start;
      if (difference > 0 && pos + 2 < end) {
        const bits = copyTokenBits(difference);
        const maxLen = Math.min((0xffff >> bits) + 3, end - pos);
        const candidates = seen.get(key(pos));
        if (candidates) {
          for (let c = candidates.length - 1, tries = 0; c >= 0 && tries < 128; c--, tries++) {
            const from = candidates[c];
            let len = 0;
            while (len < maxLen && data[from + len] === data[pos + len]) len++;
            if (len > bestLen) {
              bestLen = len;
              bestOff = pos - from;
              if (len === maxLen) break;
            }
          }
        }
        if (bestLen >= 3) {
          const token = ((bestOff - 1) << (16 - bits)) | (bestLen - 3);
          out.push(token & 0xff, token >> 8);
          flags |= 1 << bit;
          for (let i = 0; i < bestLen; i++) remember(pos + i);
          pos += bestLen;
          continue;
        }
      }
      out.push(data[pos]);
      remember(pos);
      pos++;
    }
    out[flagAt] = flags;
  }
  return out;
}

// ---------- text in the project's code page ----------

function decoderLabel(cp: number): string {
  if (cp >= 1250 && cp <= 1258) return `windows-${cp}`;
  return { 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 65001: 'utf-8', 10000: 'macintosh' }[cp] ?? 'windows-1252';
}

export function decodeText(bytes: Uint8Array, cp: number): string {
  try {
    return new TextDecoder(decoderLabel(cp)).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

const encoders = new Map<number, Map<string, number>>();

/** Encodes text in a single-byte Windows code page (characters it can't hold become "?"). */
export function encodeText(text: string, cp: number): Uint8Array {
  if (cp === 65001) return new TextEncoder().encode(text);
  let table = encoders.get(cp);
  if (!table) {
    table = new Map();
    if (cp >= 1250 && cp <= 1258) {
      const dec = new TextDecoder(decoderLabel(cp));
      for (let b = 0x80; b <= 0xff; b++) {
        const ch = dec.decode(Uint8Array.of(b));
        if (ch !== '\ufffd' && !table.has(ch)) table.set(ch, b);
      }
    }
    encoders.set(cp, table);
  }
  const out = new Uint8Array(text.length);
  let n = 0;
  for (const ch of text) {
    const code = ch.codePointAt(0)!;
    out[n++] = code < 0x80 ? code : (table.get(ch) ?? 0x3f);
  }
  return out.slice(0, n);
}

function utf16(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    out[i * 2] = c & 0xff;
    out[i * 2 + 1] = c >> 8;
  }
  return out;
}

// ---------- dir stream ([MS-OVBA] 2.3.4.2) ----------

interface DirRecord {
  id: number;
  data: Uint8Array;
  /** Byte position of the record in the decompressed stream. */
  at: number;
}

function readRecords(dir: Uint8Array): DirRecord[] {
  const out: DirRecord[] = [];
  const view = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const at = pos;
    const id = view.getUint16(pos, true);
    let size = view.getUint32(pos + 2, true);
    pos += 6;
    if (id === 0x0009) size = 6; // PROJECTVERSION: the size field is a fixed "4" but 6 bytes follow
    out.push({ id, data: dir.subarray(pos, pos + size), at });
    pos += size;
    if (id === 0x0010) break;
  }
  return out;
}

interface DirModule {
  name: string;
  stream: string;
  offset: number;
  procedural: boolean;
}

interface DirInfo {
  codePage: number;
  projectName: string;
  modules: DirModule[];
  /** Start of the PROJECTMODULES record: everything before it (information and references) can be reused. */
  modulesAt: number;
}

function parseDir(dir: Uint8Array): DirInfo {
  const records = readRecords(dir);
  let codePage = 1252;
  let projectName = 'VBAProject';
  let modulesAt = -1;
  const modules: DirModule[] = [];
  let cur: Partial<DirModule> & { unicodeName?: string; unicodeStream?: string } = {};
  // Text fields are decoded after the code page is known
  const raw: { id: number; data: Uint8Array }[] = [];
  for (const r of records) {
    if (r.id === 0x0003) codePage = r.data[0] | (r.data[1] << 8);
    if (r.id === 0x000f && modulesAt < 0) modulesAt = r.at;
    raw.push(r);
  }
  const text = (d: Uint8Array) => decodeText(d, codePage);
  const uni = (d: Uint8Array) => {
    let s = '';
    for (let i = 0; i + 1 < d.length; i += 2) s += String.fromCharCode(d[i] | (d[i + 1] << 8));
    return s;
  };
  for (const r of raw) {
    switch (r.id) {
      case 0x0004:
        projectName = text(r.data);
        break;
      case 0x0019:
        cur = { name: text(r.data), procedural: true, offset: 0 };
        break;
      case 0x0047:
        cur.unicodeName = uni(r.data);
        break;
      case 0x001a:
        cur.stream = text(r.data);
        break;
      case 0x0032:
        cur.unicodeStream = uni(r.data);
        break;
      case 0x0031:
        cur.offset = new DataView(r.data.buffer, r.data.byteOffset, 4).getUint32(0, true);
        break;
      case 0x0021:
        cur.procedural = true;
        break;
      case 0x0022:
        cur.procedural = false;
        break;
      case 0x002b:
        if (cur.name !== undefined) {
          modules.push({
            name: cur.unicodeName || cur.name,
            stream: cur.unicodeStream || cur.stream || cur.name,
            offset: cur.offset ?? 0,
            procedural: cur.procedural ?? true,
          });
        }
        cur = {};
        break;
    }
  }
  return { codePage, projectName, modules, modulesAt };
}

function record(id: number, data: Uint8Array | number[]): number[] {
  const d = Array.from(data);
  return [id & 0xff, id >> 8, d.length & 0xff, (d.length >> 8) & 0xff, (d.length >> 16) & 0xff, d.length >>> 24, ...d];
}

const u16 = (n: number) => [n & 0xff, (n >> 8) & 0xff];
const u32 = (n: number) => [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff];

/** PROJECTINFORMATION and PROJECTREFERENCES for a brand-new project (references stdole like Excel does). */
function newProjectInfo(name: string, codePage: number): number[] {
  const enc = (s: string) => encodeText(s, codePage);
  const libid = '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation';
  const lib = enc(libid);
  return [
    ...record(0x0001, u32(1)), // SysKind: 32-bit Windows
    ...record(0x0002, u32(0x0409)), // LCID
    ...record(0x0014, u32(0x0409)), // LCIDINVOKE
    ...record(0x0003, u16(codePage)),
    ...record(0x0004, enc(name)),
    ...record(0x0005, []),
    ...record(0x0040, []), // doc string (unicode)
    ...record(0x0006, []),
    ...record(0x003d, []), // help file 2
    ...record(0x0007, u32(0)),
    ...record(0x0008, u32(0)),
    0x09, 0x00, ...u32(4), ...u32(1), ...u16(0), // PROJECTVERSION 1.0
    ...record(0x000c, []),
    ...record(0x003c, []), // constants (unicode)
    // Reference to OLE Automation (stdole)
    ...record(0x0016, enc('stdole')),
    ...record(0x003e, utf16('stdole')),
    ...record(0x000d, [...u32(lib.length), ...lib, ...u32(0), ...u16(0)]),
  ];
}

function moduleRecords(m: { name: string; kind: VbaModuleKind }, codePage: number): number[] {
  const name = encodeText(m.name, codePage);
  return [
    ...record(0x0019, name),
    ...record(0x0047, utf16(m.name)),
    ...record(0x001a, name),
    ...record(0x0032, utf16(m.name)),
    ...record(0x001c, []),
    ...record(0x0048, []),
    ...record(0x0031, u32(0)), // source starts at the beginning of the stream (no compiled code)
    ...record(0x001e, u32(0)),
    ...record(0x002c, u16(0xffff)),
    ...record(m.kind === 'standard' ? 0x0021 : 0x0022, []),
    ...record(0x002b, []),
  ];
}

// ---------- PROJECT stream ([MS-OVBA] 2.3.1) ----------

/** Data encryption used for the CMG, DPB and GC lines ([MS-OVBA] 2.4.3). */
export function encryptProjectData(data: number[], projectId: string, seed = Math.floor(Math.random() * 256)): string {
  let projKey = 0;
  for (const ch of projectId) projKey = (projKey + ch.charCodeAt(0)) & 0xff;
  const version = 2;
  const out = [seed, seed ^ version, seed ^ projKey];
  let unencrypted1 = projKey;
  let encrypted1 = out[2];
  let encrypted2 = out[1];
  const push = (b: number) => {
    const enc = b ^ ((encrypted2 + unencrypted1) & 0xff);
    out.push(enc);
    encrypted2 = encrypted1;
    encrypted1 = enc;
    unencrypted1 = b;
  };
  const ignored = (seed & 6) >> 1;
  for (let i = 0; i < ignored; i++) push(0);
  for (const b of u32(data.length)) push(b);
  for (const b of data) push(b);
  return out.map((b) => b.toString(16).toUpperCase().padStart(2, '0')).join('');
}

export function decryptProjectData(hex: string): number[] {
  const bytes = hex.match(/../g)!.map((h) => parseInt(h, 16));
  const seed = bytes[0];
  const projKey = seed ^ bytes[2];
  let unencrypted1 = projKey;
  let encrypted1 = bytes[2];
  let encrypted2 = bytes[1];
  let pos = 3;
  const next = () => {
    const enc = bytes[pos++];
    const b = enc ^ ((encrypted2 + unencrypted1) & 0xff);
    encrypted2 = encrypted1;
    encrypted1 = enc;
    unencrypted1 = b;
    return b;
  };
  for (let i = 0; i < (seed & 6) >> 1; i++) next();
  const len = next() | (next() << 8) | (next() << 16) | (next() << 24);
  const out: number[] = [];
  for (let i = 0; i < len; i++) out.push(next());
  return out;
}

const MODULE_LINE = /^(Document|Module|Class|BaseClass)=/i;

function projectLine(m: { name: string; kind: VbaModuleKind }): string {
  switch (m.kind) {
    case 'document':
      return `Document=${m.name}/&H00000000`;
    case 'class':
      return `Class=${m.name}`;
    case 'form':
      return `BaseClass=${m.name}`;
    default:
      return `Module=${m.name}`;
  }
}

function newProjectStream(name: string, modules: { name: string; kind: VbaModuleKind }[]): string {
  const id = `{${guid()}}`;
  return [
    `ID="${id}"`,
    ...modules.map(projectLine),
    `Name="${name}"`,
    'HelpContextID="0"',
    'VersionCompatible32="393222000"',
    `CMG="${encryptProjectData(u32(0), id)}"`, // not protected
    `DPB="${encryptProjectData([0], id)}"`, // no password
    `GC="${encryptProjectData([0xff], id)}"`, // visible
    '',
    '[Host Extender Info]',
    '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
    '',
    '[Workspace]',
    ...modules.map((m) => `${m.name}=0, 0, 0, 0, C`),
    '',
  ].join('\r\n');
}

/** The original PROJECT stream with its module lines and [Workspace] replaced. */
function updateProjectStream(original: string, modules: { name: string; kind: VbaModuleKind }[]): string {
  const lines = original.split(/\r?\n/);
  const out: string[] = [];
  let inserted = false;
  let section = '';
  for (const line of lines) {
    if (/^\[.*\]$/.test(line.trim())) {
      section = line.trim().toLowerCase();
      out.push(line);
      if (section === '[workspace]') out.push(...modules.map((m) => `${m.name}=0, 0, 0, 0, C`));
      continue;
    }
    if (section === '[workspace]') continue;
    if (!section && MODULE_LINE.test(line)) {
      if (!inserted) {
        out.push(...modules.map(projectLine));
        inserted = true;
      }
      continue;
    }
    if (!section && !inserted && /^Name=/i.test(line)) {
      out.push(...modules.map(projectLine));
      inserted = true;
    }
    out.push(line);
  }
  return out.join('\r\n');
}

function moduleKindsFromProject(text: string): Map<string, VbaModuleKind> {
  const kinds = new Map<string, VbaModuleKind>();
  for (const line of text.split(/\r?\n/)) {
    if (/^\[/.test(line)) break;
    const m = /^(Document|Module|Class|BaseClass)=([^/\r\n]+)/i.exec(line);
    if (!m) continue;
    const k = m[1].toLowerCase();
    kinds.set(m[2].trim().toLowerCase(), k === 'document' ? 'document' : k === 'class' ? 'class' : k === 'baseclass' ? 'form' : 'standard');
  }
  return kinds;
}

/** A random GUID in upper case (works without a secure context, unlike crypto.randomUUID). */
export function guid(): string {
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const x = Array.from(b, (v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20)}`;
}

// ---------- read / write ----------

function streamBytes(cfb: CFB.CFB$Container, path: string): Uint8Array | null {
  const e = CFB.find(cfb, path);
  if (!e || !e.content) return null;
  return e.content instanceof Uint8Array ? e.content : Uint8Array.from(e.content as number[]);
}

export function readVbaProject(bin: Uint8Array): VbaProjectData {
  const cfb = CFB.parse(bin);
  const dirBytes = streamBytes(cfb, '/VBA/dir');
  if (!dirBytes) throw new Error('This file has a VBA project without a module list.');
  const info = parseDir(decompress(dirBytes));
  const projectText = decodeText(streamBytes(cfb, '/PROJECT') ?? new Uint8Array(), info.codePage);
  const kinds = moduleKindsFromProject(projectText);
  const modules: VbaModuleSource[] = [];
  for (const m of info.modules) {
    const bytes = streamBytes(cfb, `/VBA/${m.stream}`);
    let source = '';
    if (bytes && m.offset < bytes.length) {
      try {
        source = decodeText(decompress(bytes, m.offset), info.codePage);
      } catch {
        source = '';
      }
    }
    modules.push({ name: m.name, kind: kinds.get(m.name.toLowerCase()) ?? (m.procedural ? 'standard' : 'class'), source });
  }
  return { name: info.projectName, codePage: info.codePage, modules };
}

/** Builds vbaProject.bin. With `original`, keeps its references, protection lines and forms. */
export function writeVbaProject(project: { name: string; modules: VbaModuleSource[] }, original?: Uint8Array | null): Uint8Array {
  let cfb: CFB.CFB$Container;
  let codePage = 1252;
  let info: number[];
  let projectText: string;
  if (original) {
    cfb = CFB.parse(original);
    const dirBytes = streamBytes(cfb, '/VBA/dir');
    const dir = dirBytes ? decompress(dirBytes) : null;
    const parsed = dir ? parseDir(dir) : null;
    codePage = parsed?.codePage ?? 1252;
    info = dir && parsed && parsed.modulesAt > 0 ? Array.from(dir.subarray(0, parsed.modulesAt)) : newProjectInfo(project.name, codePage);
    // Old module streams and compiled-code caches are replaced
    for (const m of parsed?.modules ?? []) CFB.utils.cfb_del(cfb, `/VBA/${m.stream}`);
    for (const p of [...cfb.FullPaths]) if (/\/VBA\/__SRP_\d+$/i.test(p)) CFB.utils.cfb_del(cfb, p);
    const oldProject = streamBytes(cfb, '/PROJECT');
    projectText = oldProject ? updateProjectStream(decodeText(oldProject, codePage), project.modules) : newProjectStream(project.name, project.modules);
  } else {
    cfb = CFB.utils.cfb_new();
    info = newProjectInfo(project.name, codePage);
    projectText = newProjectStream(project.name, project.modules);
  }

  const dir = [...info, ...record(0x000f, u16(project.modules.length)), ...record(0x0013, u16(0xffff))];
  for (const m of project.modules) dir.push(...moduleRecords(m, codePage));
  dir.push(...record(0x0010, []));

  const add = (path: string, data: Uint8Array) => CFB.utils.cfb_add(cfb, path, data);
  add('/VBA/dir', compress(Uint8Array.from(dir)));
  // Version 0xFFFF: the compiled-code cache is stale, so Excel recompiles from source
  add('/VBA/_VBA_PROJECT', Uint8Array.of(0xcc, 0x61, 0xff, 0xff, 0x00, 0x00, 0x00));
  for (const m of project.modules) add(`/VBA/${m.name}`, compress(encodeText(normalizeNewlines(m.source), codePage)));
  add('/PROJECT', encodeText(projectText, codePage));
  const wm: number[] = [];
  for (const m of project.modules) wm.push(...encodeText(m.name, codePage), 0, ...utf16(m.name), 0, 0);
  wm.push(0, 0);
  add('/PROJECTwm', Uint8Array.from(wm));
  const out = CFB.write(cfb, { type: 'array', fileType: 'cfb' }) as number[] | Uint8Array;
  return out instanceof Uint8Array ? out : Uint8Array.from(out);
}

function normalizeNewlines(s: string): string {
  const t = s.replace(/\r?\n/g, '\r\n');
  return t.endsWith('\r\n') || !t ? t : t + '\r\n';
}

// ---------- the hidden "Attribute" lines ----------

const BASES: Record<string, string> = {
  thisworkbook: '0{00020819-0000-0000-C000-000000000046}',
  sheet: '0{00020820-0000-0000-C000-000000000046}',
  class: '0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}',
};

/** Splits stored source into what the editor shows and the hidden attribute lines. */
export function splitAttributes(source: string): { code: string; header: string[]; procAttrs: Record<string, string[]> } {
  const lines = source.replace(/\r\n?/g, '\n').split('\n');
  const header: string[] = [];
  let i = 0;
  // Exported-class header (VERSION 1.0 CLASS / BEGIN … END) may appear in some files
  if (/^VERSION\s+[\d.]+\s+CLASS/i.test(lines[0] ?? '')) {
    while (i < lines.length && !/^END\s*$/i.test(lines[i])) header.push(lines[i++]);
    if (i < lines.length) header.push(lines[i++]);
  }
  while (i < lines.length && /^Attribute\s+VB_/i.test(lines[i])) header.push(lines[i++]);
  const procAttrs: Record<string, string[]> = {};
  const body: string[] = [];
  for (; i < lines.length; i++) {
    const m = /^\s*Attribute\s+([A-Za-z_][\w]*)\.(VB_\w+.*)$/i.exec(lines[i]);
    if (m) (procAttrs[m[1].toLowerCase()] ??= []).push(lines[i].trim());
    else body.push(lines[i]);
  }
  while (body.length && body[body.length - 1] === '') body.pop();
  return { code: body.join('\r\n'), header, procAttrs };
}

/** Rebuilds stored source: attribute header, code, and procedure attributes after each Sub line. */
export function joinAttributes(
  name: string,
  kind: VbaModuleKind,
  code: string,
  saved?: { header?: string[]; procAttrs?: Record<string, string[]> },
  shortcuts: Record<string, string> = {},
): string {
  let header = saved?.header?.length ? [...saved.header] : defaultHeader(name, kind);
  header = header.map((l) => (/^Attribute\s+VB_Name\s*=/i.test(l) ? `Attribute VB_Name = "${name}"` : l));
  const out = [...header];
  const procAttrs = saved?.procAttrs ?? {};
  for (const line of code.replace(/\r\n?/g, '\n').split('\n')) {
    out.push(line);
    const m = /^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function)\s+([A-Za-z_]\w*)/i.exec(line);
    if (!m) continue;
    const key = m[1].toLowerCase();
    const attrs = (procAttrs[key] ?? []).filter((a) => !/VB_ProcData\.VB_Invoke_Func/i.test(a));
    const sc = shortcuts[key];
    if (sc) attrs.push(`Attribute ${m[1]}.VB_ProcData.VB_Invoke_Func = "${sc}\\n14"`);
    out.push(...attrs);
  }
  return out.join('\r\n');
}

function defaultHeader(name: string, kind: VbaModuleKind): string[] {
  if (kind === 'standard') return [`Attribute VB_Name = "${name}"`];
  const base = kind === 'document' ? (name.toLowerCase() === 'thisworkbook' ? BASES.thisworkbook : BASES.sheet) : BASES.class;
  const doc = kind === 'document';
  return [
    `Attribute VB_Name = "${name}"`,
    `Attribute VB_Base = "${base}"`,
    'Attribute VB_GlobalNameSpace = False',
    'Attribute VB_Creatable = False',
    `Attribute VB_PredeclaredId = ${doc ? 'True' : 'False'}`,
    `Attribute VB_Exposed = ${doc ? 'True' : 'False'}`,
    'Attribute VB_TemplateDerived = False',
    `Attribute VB_Customizable = ${doc ? 'True' : 'False'}`,
  ];
}

/** Excel shortcut keys stored as VB_Invoke_Func attributes: { main: 'm' } (uppercase = Ctrl+Shift). */
export function invokeShortcuts(procAttrs: Record<string, string[]>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [proc, attrs] of Object.entries(procAttrs)) {
    for (const a of attrs) {
      const m = /VB_Invoke_Func\s*=\s*"([A-Za-z])\\n\d+"/i.exec(a);
      if (m) out[proc] = m[1];
    }
  }
  return out;
}
