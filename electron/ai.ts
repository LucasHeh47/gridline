// The AI assistant's connection to a model. Runs in the main process so the API key never reaches
// the page: the key is stored encrypted with the operating system's keychain (safeStorage), and
// requests stream back to the window as events.
// Keeping the key secret:
//  • the page can save or remove the key and ask whether one is saved, but never read it back;
//  • it is only ever sent to api.anthropic.com over HTTPS, and redirects are refused so the header
//    can't follow a response anywhere else;
//  • messages passed back to the page are scrubbed of anything that looks like a key;
//  • only the app's own page may use these channels.
//  • Claude: Anthropic's Messages API with the user's own key.
//  • Local: any OpenAI-compatible server on the PC (Ollama, LM Studio).
import { app, ipcMain, net, safeStorage, type IpcMainEvent, type IpcMainInvokeEvent, type WebContents } from 'electron';
import fs from 'node:fs/promises';
import path from 'node:path';
import type { AiBlock, AiEvent, AiMessage, AiRequest, AiToolDef } from '../src/shared/api';
import { redactSecrets, validKeyFormat } from '../src/shared/secrets';
import { ToolCallStream } from '../src/shared/toolCalls';

const keyPath = () => path.join(app.getPath('userData'), 'ai-key.bin');
const running = new Map<string, AbortController>();

/**
 * Where Claude requests go. Development runs (not the installed app) may point them at a local test
 * server with GRIDLINE_AI_TEST_URL=http://127.0.0.1:8124; installed copies always use Anthropic.
 */
function claudeBase(): string {
  const t = process.env.GRIDLINE_AI_TEST_URL;
  if (!app.isPackaged && t && /^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?\/?$/.test(t)) return t.replace(/\/$/, '');
  return 'https://api.anthropic.com';
}

let devUrl: string | undefined;

/** True for messages from the app's own page (not another site or an embedded frame). */
function fromApp(e: IpcMainEvent | IpcMainInvokeEvent): boolean {
  const frame = e.senderFrame;
  if (!frame || frame !== e.sender.mainFrame) return false;
  if (devUrl) return frame.url.startsWith(devUrl);
  try {
    const u = new URL(frame.url);
    return u.protocol === 'file:' && decodeURIComponent(u.pathname).replace(/\\/g, '/').endsWith('/dist/index.html');
  } catch {
    return false;
  }
}

async function readKey(): Promise<string | null> {
  try {
    const data = await fs.readFile(keyPath());
    return safeStorage.decryptString(data);
  } catch {
    return null;
  }
}

function send(target: WebContents, id: string, e: AiEvent, key?: string | null): void {
  if (target.isDestroyed()) return;
  // Error text can include what a server echoed back: never let a key through
  if (e.type === 'error') e = { type: 'error', message: redactSecrets(e.message, key) };
  target.send('ai:event', id, e);
}

/** Reads a server-sent-events stream, calling `onEvent` with each event's name and data. */
async function readSse(res: Response, onEvent: (event: string, data: string) => void): Promise<void> {
  const reader = res.body!.getReader();
  const dec = new TextDecoder();
  let buf = '';
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += dec.decode(value, { stream: true });
    let i: number;
    while ((i = buf.search(/\r?\n\r?\n/)) >= 0) {
      const chunk = buf.slice(0, i);
      buf = buf.slice(i).replace(/^\r?\n\r?\n/, '');
      let event = 'message';
      const data: string[] = [];
      for (const line of chunk.split(/\r?\n/)) {
        if (line.startsWith('event:')) event = line.slice(6).trim();
        else if (line.startsWith('data:')) data.push(line.slice(5).replace(/^ /, ''));
      }
      if (data.length) onEvent(event, data.join('\n'));
    }
  }
}

async function errorText(res: Response, provider: string): Promise<string> {
  let detail = '';
  try {
    const body = await res.text();
    const j = JSON.parse(body) as { error?: { message?: string } | string; message?: string };
    detail = typeof j.error === 'string' ? j.error : (j.error?.message ?? j.message ?? '');
  } catch {
    // not JSON
  }
  if (provider === 'claude') {
    if (res.status === 401) return 'Your Claude API key was rejected. Check it in Settings → AI assistant.';
    if (res.status === 403) return `Your Claude API key isn't allowed to do this.${detail ? ' ' + detail : ''}`;
    if (res.status === 429) return 'Too many requests to Claude right now (or the account is out of credit). Wait a moment and try again.';
    if (res.status === 529 || res.status === 503) return 'Claude is overloaded right now. Try again in a minute.';
  }
  return `The AI service returned an error (${res.status})${detail ? ': ' + detail : '.'}`;
}

// ---------- Claude ----------

async function chatClaude(target: WebContents, id: string, req: AiRequest, signal: AbortSignal): Promise<void> {
  const key = await readKey();
  if (!key) {
    send(target, id, { type: 'error', message: 'Add your Claude API key in Settings → AI assistant first.' });
    return;
  }
  const tools = req.tools.map((t, i) => (i === req.tools.length - 1 ? { ...t, cache_control: { type: 'ephemeral' } } : t));
  const body = {
    model: req.model,
    max_tokens: req.maxTokens ?? 8192,
    stream: true,
    // The instructions and tool list are the same for every message, so cache them (cheaper and faster)
    system: [{ type: 'text', text: req.system, cache_control: { type: 'ephemeral' } }],
    tools,
    messages: req.messages,
  };
  let res: Response;
  try {
    res = await net.fetch(`${claudeBase()}/v1/messages`, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify(body),
      // Never follow a redirect: the key header would go along to the new address
      redirect: 'error',
      signal,
    });
  } catch (err) {
    if (signal.aborted) throw err;
    send(target, id, { type: 'error', message: `Couldn't reach Claude: ${(err as Error).message}` }, key);
    return;
  }
  if (!res.ok) {
    send(target, id, { type: 'error', message: await errorText(res, 'claude') }, key);
    return;
  }
  const blocks = new Map<number, { type: string; id?: string; name?: string; json: string }>();
  let stopReason = 'end_turn';
  const usage = { input: 0, output: 0 };
  await readSse(res, (event, data) => {
    let j: Record<string, any>;
    try {
      j = JSON.parse(data);
    } catch {
      return;
    }
    switch (event) {
      case 'message_start':
        usage.input = j.message?.usage?.input_tokens ?? 0;
        break;
      case 'content_block_start':
        blocks.set(j.index, { type: j.content_block?.type, id: j.content_block?.id, name: j.content_block?.name, json: '' });
        break;
      case 'content_block_delta': {
        const b = blocks.get(j.index);
        if (j.delta?.type === 'text_delta') send(target, id, { type: 'text', text: j.delta.text });
        else if (j.delta?.type === 'input_json_delta' && b) b.json += j.delta.partial_json ?? '';
        break;
      }
      case 'content_block_stop': {
        const b = blocks.get(j.index);
        if (b?.type === 'tool_use') {
          let input: unknown = {};
          try {
            input = b.json ? JSON.parse(b.json) : {};
          } catch {
            input = {};
          }
          send(target, id, { type: 'tool_use', id: b.id!, name: b.name!, input });
        }
        break;
      }
      case 'message_delta':
        if (j.delta?.stop_reason) stopReason = j.delta.stop_reason;
        usage.output = j.usage?.output_tokens ?? usage.output;
        break;
      case 'error':
        send(target, id, { type: 'error', message: j.error?.message ?? 'The AI service reported an error.' }, key);
        break;
    }
  });
  send(target, id, { type: 'done', stopReason, usage });
}

// ---------- local (OpenAI-compatible) ----------

function toOpenAi(system: string, messages: AiMessage[]): unknown[] {
  const out: unknown[] = [{ role: 'system', content: system }];
  for (const m of messages) {
    const texts = m.content.filter((b): b is Extract<AiBlock, { type: 'text' }> => b.type === 'text').map((b) => b.text);
    if (m.role === 'assistant') {
      const calls = m.content
        .filter((b): b is Extract<AiBlock, { type: 'tool_use' }> => b.type === 'tool_use')
        .map((b) => ({ id: b.id, type: 'function', function: { name: b.name, arguments: JSON.stringify(b.input ?? {}) } }));
      out.push({ role: 'assistant', content: texts.join('\n') || null, ...(calls.length ? { tool_calls: calls } : {}) });
    } else {
      for (const b of m.content) if (b.type === 'tool_result') out.push({ role: 'tool', tool_call_id: b.tool_use_id, content: b.content });
      if (texts.length) out.push({ role: 'user', content: texts.join('\n') });
    }
  }
  return out;
}

function baseUrlOf(req: { baseUrl?: string }): string {
  return (req.baseUrl || 'http://localhost:11434/v1').replace(/\/+$/, '');
}

async function chatLocal(target: WebContents, id: string, req: AiRequest, signal: AbortSignal): Promise<void> {
  if (!req.model) {
    send(target, id, { type: 'error', message: 'Choose a local model in Settings → AI assistant.' });
    return;
  }
  const tools = req.tools.map((t: AiToolDef) => ({ type: 'function', function: { name: t.name, description: t.description, parameters: t.input_schema } }));
  let res: Response;
  try {
    res = await net.fetch(`${baseUrlOf(req)}/chat/completions`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ model: req.model, messages: toOpenAi(req.system, req.messages), tools, stream: true, max_tokens: req.maxTokens ?? 8192 }),
      signal,
    });
  } catch (e) {
    if (signal.aborted) throw e;
    send(target, id, { type: 'error', message: `Couldn't reach the local AI server at ${baseUrlOf(req)}. Is Ollama or LM Studio running?` });
    return;
  }
  if (!res.ok) {
    send(target, id, { type: 'error', message: await errorText(res, 'local') });
    return;
  }
  const calls = new Map<number, { id: string; name: string; args: string }>();
  let stopReason = 'end_turn';
  // Some local models write tool calls as text (e.g. Qwen3-Coder's <function=…> XML): hold that
  // back from the chat and turn it into real calls when the reply ends
  const textCalls = new ToolCallStream(req.tools);
  await readSse(res, (_event, data) => {
    if (data.trim() === '[DONE]') return;
    let j: Record<string, any>;
    try {
      j = JSON.parse(data);
    } catch {
      return;
    }
    const choice = j.choices?.[0];
    const delta = choice?.delta ?? {};
    if (typeof delta.content === 'string' && delta.content) {
      const shown = textCalls.push(delta.content);
      if (shown) send(target, id, { type: 'text', text: shown });
    }
    for (const tc of delta.tool_calls ?? []) {
      // Servers that leave out `index` send each call whole, with its own id
      const known = tc.id ? [...calls.entries()].find(([, c]) => c.id === tc.id)?.[0] : undefined;
      const i = typeof tc.index === 'number' ? tc.index : (known ?? (tc.id ? calls.size : Math.max(0, calls.size - 1)));
      const cur = calls.get(i) ?? { id: tc.id ?? `call_${i}_${Date.now()}`, name: '', args: '' };
      if (tc.id) cur.id = tc.id;
      const name = tc.function?.name;
      if (name && cur.name !== name) cur.name += name;
      const args = tc.function?.arguments;
      if (typeof args === 'string') cur.args += args;
      else if (args && typeof args === 'object') cur.args = JSON.stringify(args);
      calls.set(i, cur);
    }
    if (choice?.finish_reason) stopReason = choice.finish_reason === 'tool_calls' ? 'tool_use' : choice.finish_reason;
  });
  const rest = textCalls.finish();
  if (rest.text) send(target, id, { type: 'text', text: rest.text });
  for (const c of calls.values()) {
    let input: unknown = {};
    try {
      input = c.args ? JSON.parse(c.args) : {};
    } catch {
      input = {};
    }
    send(target, id, { type: 'tool_use', id: c.id, name: c.name, input });
  }
  // Calls written as text count only when the server sent no real ones
  const fromText = calls.size ? [] : rest.calls;
  fromText.forEach((c, i) => send(target, id, { type: 'tool_use', id: `call_text_${Date.now().toString(36)}_${i}`, name: c.name, input: c.input }));
  send(target, id, { type: 'done', stopReason: calls.size || fromText.length ? 'tool_use' : stopReason });
}

// ---------- IPC ----------

export function setupAi(devServerUrl?: string): void {
  devUrl = devServerUrl;
  ipcMain.handle('ai:chat', async (e, id: string, req: AiRequest) => {
    if (!fromApp(e) || typeof id !== 'string') return;
    const ctl = new AbortController();
    running.set(id, ctl);
    try {
      if (req.provider === 'local') await chatLocal(e.sender, id, req, ctl.signal);
      else await chatClaude(e.sender, id, req, ctl.signal);
    } catch (err) {
      if (ctl.signal.aborted) send(e.sender, id, { type: 'done', stopReason: 'cancelled' });
      else send(e.sender, id, { type: 'error', message: `Couldn't reach the AI service: ${(err as Error).message}` });
    } finally {
      running.delete(id);
    }
  });
  ipcMain.on('ai:cancel', (e, id: string) => {
    if (fromApp(e)) running.get(id)?.abort();
  });
  ipcMain.handle('ai:set-key', async (e, key: string | null) => {
    if (!fromApp(e)) return { ok: false, error: 'Not allowed.' };
    try {
      if (!key) {
        await fs.rm(keyPath(), { force: true });
        return { ok: true };
      }
      const k = String(key).trim();
      // Messages here never repeat the key
      if (!validKeyFormat(k)) return { ok: false, error: "That doesn't look like an API key. Copy the whole key from console.anthropic.com and paste it again." };
      if (!safeStorage.isEncryptionAvailable() || (process.platform === 'linux' && safeStorage.getSelectedStorageBackend() === 'basic_text')) {
        return { ok: false, error: "This computer can't store the key securely, so it wasn't saved." };
      }
      await fs.writeFile(keyPath(), safeStorage.encryptString(k), { mode: 0o600 });
      return { ok: true };
    } catch {
      return { ok: false, error: "The key couldn't be saved." };
    }
  });
  ipcMain.handle('ai:has-key', async (e) => fromApp(e) && !!(await readKey()));
  ipcMain.handle('ai:local-models', async (e, baseUrl: string) => {
    if (!fromApp(e)) return { ok: false, models: [], error: 'Not allowed.' };
    try {
      const res = await net.fetch(`${baseUrlOf({ baseUrl })}/models`);
      if (!res.ok) return { ok: false, models: [], error: `The server answered ${res.status}.` };
      const j = (await res.json()) as { data?: { id: string }[]; models?: { name: string }[] };
      const models = (j.data?.map((m) => m.id) ?? j.models?.map((m) => m.name) ?? []).sort();
      return { ok: true, models };
    } catch {
      return { ok: false, models: [], error: `Couldn't reach ${baseUrlOf({ baseUrl })}. Is Ollama or LM Studio running?` };
    }
  });
}
