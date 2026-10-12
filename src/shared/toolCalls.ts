// Local models sometimes write their tool calls as text instead of sending them as structured
// calls — for example Qwen3-Coder's XML (<function=name><parameter=x>…</parameter></function>),
// <tool_call>{json}</tool_call>, Llama's {"name": …, "parameters": …} or a ```json block. Servers
// like Ollama and LM Studio don't always convert these, so the app would just show the markup.
// This turns them into real calls, and drops <think>…</think> reasoning from the visible text.
import type { AiToolDef } from './api';

export interface TextToolCall {
  name: string;
  input: Record<string, unknown>;
}

const MARKER = /<tool_call>|<function=|<\|python_tag\|>|```|\{\s*"name"\s*:/;
const THINK_OPEN = '<think>';
const THINK_CLOSE = '</think>';
const PREFIXES = ['<tool_call>', '<function=', '<|python_tag|>', '```', THINK_OPEN, THINK_CLOSE];

/** Index from which the end of `s` might be the start of a marker split across chunks, or -1. */
function partialStart(s: string): number {
  for (let i = Math.max(0, s.length - 16); i < s.length; i++) {
    const rest = s.slice(i);
    if (PREFIXES.some((p) => p.startsWith(rest))) return i;
    if (/^\{\s*(?:"[a-z]{0,4}"?\s*)?$/.test(rest)) return i;
  }
  return -1;
}

/**
 * Filters a streamed reply: text passes through as it arrives, except anything from the first
 * tool-call marker on, which is held back and parsed when the reply ends.
 */
export class ToolCallStream {
  private pending = '';
  private held = '';
  private holding = false;
  private thinking = false;
  /** Drop blank space right after reasoning, even when it arrives in later pieces. */
  private trimStart = false;

  constructor(private tools: AiToolDef[]) {}

  /** Returns the text that can be shown now. */
  push(chunk: string): string {
    let out = this.take(chunk);
    if (this.trimStart && out) {
      out = out.replace(/^\s+/, '');
      if (out) this.trimStart = false;
    }
    return out;
  }

  private take(chunk: string): string {
    if (this.holding) {
      this.held += chunk;
      return '';
    }
    let buf = this.pending + chunk;
    this.pending = '';
    let out = '';
    for (;;) {
      if (this.thinking) {
        const end = buf.indexOf(THINK_CLOSE);
        if (end < 0) {
          // Keep only what could be the start of the closing tag
          const p = partialStart(buf);
          this.pending = p >= 0 ? buf.slice(p) : '';
          return out;
        }
        this.thinking = false;
        this.trimStart = !out.trim();
        buf = buf.slice(end + THINK_CLOSE.length);
        continue;
      }
      const t = buf.indexOf(THINK_OPEN);
      const m = MARKER.exec(buf);
      if (t >= 0 && (!m || t < m.index)) {
        out += buf.slice(0, t);
        this.thinking = true;
        buf = buf.slice(t + THINK_OPEN.length);
        continue;
      }
      if (m) {
        out += buf.slice(0, m.index);
        this.holding = true;
        this.held = buf.slice(m.index);
        return out;
      }
      const p = partialStart(buf);
      if (p >= 0) {
        this.pending = buf.slice(p);
        return out + buf.slice(0, p);
      }
      return out + buf;
    }
  }

  /** Call when the reply has ended: the remaining text to show and any tool calls found. */
  finish(): { text: string; calls: TextToolCall[] } {
    if (this.thinking) return { text: '', calls: [] };
    if (!this.holding) {
      const text = this.trimStart ? this.pending.replace(/^\s+/, '') : this.pending;
      this.pending = '';
      return { text, calls: [] };
    }
    const held = this.held + this.pending;
    this.held = '';
    this.pending = '';
    return parseToolCalls(held, this.tools);
  }
}

// ---------- parsing ----------

type Schema = { type?: string | string[]; properties?: Record<string, Schema> };

/** Reads a JSON object or array starting at `start` (skipping spaces); returns it and where it ends. */
function readJson(text: string, start: number): { value: unknown; end: number } | null {
  let i = start;
  while (i < text.length && /\s/.test(text[i])) i++;
  const open = text[i];
  if (open !== '{' && open !== '[') return null;
  let depth = 0;
  let inStr = false;
  for (let j = i; j < text.length; j++) {
    const ch = text[j];
    if (inStr) {
      if (ch === '\\') j++;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === '{' || ch === '[') depth++;
    else if (ch === '}' || ch === ']') {
      depth--;
      if (depth === 0) {
        try {
          return { value: JSON.parse(text.slice(i, j + 1)), end: j + 1 };
        } catch {
          return null;
        }
      }
    }
  }
  return null;
}

/** A parameter written as text, converted to what the tool's schema expects. */
function coerce(raw: string, schema: Schema | undefined): unknown {
  const types = schema?.type === undefined ? [] : Array.isArray(schema.type) ? schema.type : [schema.type];
  const v = raw;
  const t = v.trim();
  const wantsStructured = types.length === 0 || types.some((x) => x === 'array' || x === 'object' || x === 'number' || x === 'integer' || x === 'boolean' || x === 'null');
  if (types.includes('string') && !types.some((x) => x === 'array' || x === 'object')) {
    // A quoted string ("…") is JSON for the same text
    if (/^".*"$/s.test(t)) {
      try {
        const s = JSON.parse(t);
        if (typeof s === 'string') return s;
      } catch {
        // keep as written
      }
    }
    if (types.length === 1 || !/^-?\d+(\.\d+)?$/.test(t)) return v;
  }
  if (wantsStructured) {
    try {
      return JSON.parse(t);
    } catch {
      if (types.includes('number') || types.includes('integer')) {
        const n = Number(t);
        if (t !== '' && Number.isFinite(n)) return n;
      }
      if (types.includes('boolean') && /^(true|false)$/i.test(t)) return t.toLowerCase() === 'true';
    }
  }
  return v;
}

/** A call written as JSON: {"name": …, "arguments"|"parameters"|"input": …} or {"function": {…}}. */
function fromJson(value: unknown, names: Set<string>): TextToolCall[] | null {
  const list = Array.isArray(value) ? value : [value];
  const out: TextToolCall[] = [];
  for (const item of list) {
    if (!item || typeof item !== 'object') return null;
    const o = (item as { function?: unknown }).function && typeof (item as { function: unknown }).function === 'object' ? ((item as { function: object }).function as Record<string, unknown>) : (item as Record<string, unknown>);
    const name = typeof o.name === 'string' ? o.name : null;
    if (!name || !names.has(name)) return null;
    let args = o.arguments ?? o.parameters ?? o.input ?? {};
    if (typeof args === 'string') {
      try {
        args = JSON.parse(args);
      } catch {
        return null;
      }
    }
    if (!args || typeof args !== 'object' || Array.isArray(args)) return null;
    out.push({ name, input: args as Record<string, unknown> });
  }
  return out.length ? out : null;
}

/** Finds tool calls written as text. Returns the text without them, and the calls in order. */
export function parseToolCalls(text: string, tools: AiToolDef[]): { text: string; calls: TextToolCall[] } {
  const names = new Set(tools.map((t) => t.name));
  const schemas = new Map(tools.map((t) => [t.name, t.input_schema as Schema]));
  const spans: { start: number; end: number; calls: TextToolCall[] }[] = [];
  const free = (a: number, b: number) => !spans.some((s) => a < s.end && b > s.start);

  // Qwen3-Coder: [<tool_call>] <function=name> <parameter=p>value</parameter>… </function> [</tool_call>]
  const fnRe = /(?:<tool_call>\s*)?<function=([\w.-]+)>([\s\S]*?)(?:<\/function>(?:\s*<\/tool_call>)?|(?=<function=)|(?=<tool_call>)|$)/g;
  for (const m of text.matchAll(fnRe)) {
    const name = m[1];
    if (!names.has(name)) continue;
    const props = schemas.get(name)?.properties ?? {};
    const input: Record<string, unknown> = {};
    const paramRe = /<parameter=([\w.-]+)>\r?\n?([\s\S]*?)\r?\n?(?:<\/parameter>|(?=<parameter=)|$)/g;
    for (const p of m[2].matchAll(paramRe)) input[p[1]] = coerce(p[2], props[p[1]]);
    spans.push({ start: m.index!, end: m.index! + m[0].length, calls: [{ name, input }] });
  }

  // <tool_call>{json}</tool_call>, <|python_tag|>{json}, ```json {…} ```, and bare {"name": …}
  const starts: { re: RegExp; close?: RegExp }[] = [
    { re: /<tool_call>/g, close: /^\s*<\/tool_call>/ },
    { re: /<\|python_tag\|>/g, close: /^\s*(?:<\|eom_id\|>|<\|eot_id\|>)?/ },
    { re: /```(?:json|tool_call|tool)?[ \t]*\r?\n?/g, close: /^\s*```/ },
    { re: /(?=\{\s*"name"\s*:)/g },
  ];
  for (const { re, close } of starts) {
    for (const m of text.matchAll(re)) {
      const at = m.index! + m[0].length;
      const j = readJson(text, at);
      if (!j) continue;
      const calls = fromJson(j.value, names);
      if (!calls) continue;
      let end = j.end;
      if (close) {
        const c = close.exec(text.slice(end));
        if (re.source.startsWith('```') && !c) continue;
        if (c) end += c[0].length;
      }
      if (free(m.index!, end)) spans.push({ start: m.index!, end, calls });
    }
  }

  if (!spans.length) return { text, calls: [] };
  spans.sort((a, b) => a.start - b.start);
  let rest = '';
  let last = 0;
  for (const s of spans) {
    rest += text.slice(last, s.start);
    last = s.end;
  }
  rest += text.slice(last);
  rest = rest.replace(/<\/?tool_call>|<\|python_tag\|>|<\|eom_id\|>|<\|eot_id\|>/g, '').replace(/\n{3,}/g, '\n\n').trim();
  return { text: rest, calls: spans.flatMap((s) => s.calls) };
}
