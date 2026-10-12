// The AI assistant in the app: one conversation per workbook tab, kept while the panel is closed.
// Reads are logged in the chat; changes go in with an Undo button (or wait for Apply / Reject).
import { AgentError, Conversation, type AgentOptions, type AiTransport, type Proposal, type ProposalOutcome } from '../../core/ai/agent';
import { applyChanges, previewChanges, previewLoading, revertChanges } from '../../core/ai/changes';
import { formatRange, parseRange, type Range } from '../../core/address';
import { external } from '../../core/data/external';
import { displayText } from '../../core/ops/data';
import { workbookOverview } from '../../core/ai/workbook';
import type { Doc, DocEvent } from '../../core/doc';
import type { Transaction } from '../../core/history';
import type { Workbook } from '../../core/workbook';
import type { AiEvent } from '../../shared/api';
import { containsSecret, HIDDEN, redactSecrets } from '../../shared/secrets';
import type { App } from '../app';
import { isDesktop, platform } from '../platform';

export type ChatItem =
  | { kind: 'user'; text: string }
  | { kind: 'assistant'; text: string }
  | { kind: 'reads'; lines: { text: string; error: boolean }[] }
  | ProposalItem
  | { kind: 'error'; text: string; action?: 'key' | 'settings' }
  | { kind: 'notice'; text: string; action?: 'settings' };

export interface ProposalItem {
  kind: 'proposal';
  p: Proposal;
  state: 'pending' | 'applied' | 'rejected' | 'undone' | 'stopped' | 'failed';
  tx?: Transaction;
  wb: Workbook;
  note?: string;
  resolve?: (o: ProposalOutcome) => void;
  /** Undone since the model last heard from the user (told with the next message). */
  undoneUnreported?: boolean;
}

export const CLAUDE_MODELS: [string, string][] = [
  ['claude-sonnet-5-5', 'Claude Sonnet 5.5'],
  ['claude-opus-5-5', 'Claude Opus 5.5'],
  ['claude-haiku-5-5', 'Claude Haiku 5.5'],
];

// Routes streamed events from the main process to the request they belong to
const routes = new Map<string, (e: AiEvent) => void>();
let routed = false;
export const platformTransport: AiTransport = {
  async chat(req, onEvent, signal) {
    if (!routed) {
      platform.ai.onEvent((id, e) => routes.get(id)?.(e));
      routed = true;
    }
    const id = `c${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
    let finish!: () => void;
    const ended = new Promise<void>((r) => (finish = r));
    routes.set(id, (e) => {
      onEvent(e);
      if (e.type === 'done' || e.type === 'error') finish();
    });
    const cancel = () => platform.ai.cancel(id);
    signal.addEventListener('abort', cancel);
    try {
      await platform.ai.chat(id, req);
      // Streamed events travel apart from the reply and can arrive after it: wait for the last one
      await Promise.race([ended, new Promise((r) => setTimeout(r, 5000))]);
    } finally {
      routes.delete(id);
      signal.removeEventListener('abort', cancel);
    }
  },
};

/** One chat about one workbook. Each workbook tab has its own; it keeps going in the background. */
interface ChatSession {
  doc: Doc;
  items: ChatItem[];
  conv: Conversation;
  running: boolean;
  ctl: AbortController | null;
  current: { kind: 'assistant'; text: string } | null;
  /** The workbook the chat is about (opening another file into the same tab starts over). */
  wb: Workbook;
  /** Bumped by New chat, so a request still winding down can't touch the new chat. */
  gen: number;
}

export class AiController {
  hasKey: boolean | null = null;
  /** Repaints the panel; `item` alone changed when given. */
  onChange: ((item?: ChatItem) => void) | null = null;
  private sessions = new Map<Doc, ChatSession>();

  constructor(
    private app: App,
    private transport: AiTransport = platformTransport,
  ) {
    // Previews run on a trial copy, so their downloads don't count as sheet changes: listen directly
    external.onChange(() => this.dataArrived());
  }

  /** The chat of the workbook tab being shown. */
  private get chat(): ChatSession {
    return this.session(this.app.doc);
  }

  get items(): ChatItem[] {
    return this.chat.items;
  }

  get running(): boolean {
    return this.chat.running;
  }

  get usage(): { input: number; output: number } {
    return this.chat.conv.usage;
  }

  private session(doc: Doc): ChatSession {
    // Chats of closed tabs go away
    for (const [d, s] of this.sessions) {
      if (!this.app.books.some((b) => b.doc === d)) {
        s.ctl?.abort();
        this.sessions.delete(d);
      }
    }
    let s = this.sessions.get(doc);
    if (!s) {
      const fresh: ChatSession = { doc, items: [], running: false, ctl: null, current: null, wb: doc.wb, gen: 0, conv: undefined as unknown as Conversation };
      fresh.conv = this.makeConversation(fresh);
      this.sessions.set(doc, fresh);
      doc.on((e) => this.docChanged(fresh, e));
      s = fresh;
    }
    return s;
  }

  private shown(s: ChatSession): boolean {
    return s.doc === this.app.doc;
  }

  /** Repaints when the chat is the one on screen. */
  private notify(s: ChatSession, item?: ChatItem): void {
    if (this.shown(s)) this.onChange?.(item);
  }

  private push(s: ChatSession, item: ChatItem): void {
    s.items.push(item);
    this.notify(s);
  }

  /** A conversation whose callbacks only reach its chat while it's that chat's current one. */
  private makeConversation(s: ChatSession): Conversation {
    const gen = s.gen;
    const live = () => gen === s.gen && this.sessions.get(s.doc) === s;
    return new Conversation(s.doc, this.transport, {
      onStep: () => {
        if (live()) s.current = null;
      },
      onText: (t) => {
        if (!live()) return;
        if (!s.current) {
          s.current = { kind: 'assistant', text: '' };
          this.push(s, s.current);
        }
        s.current.text += t;
        this.notify(s, s.current);
      },
      onRead: (text, error) => {
        if (!live()) return;
        s.current = null;
        const last = s.items[s.items.length - 1];
        if (last?.kind === 'reads') {
          last.lines.push({ text, error });
          this.notify(s, last);
        } else this.push(s, { kind: 'reads', lines: [{ text, error }] });
      },
      propose: (p, signal) =>
        new Promise<ProposalOutcome>((resolve) => {
          if (!live()) return resolve({ status: 'stopped' });
          s.current = null;
          const item: ProposalItem = { kind: 'proposal', p, state: 'pending', wb: s.doc.wb, resolve };
          signal.addEventListener('abort', () => this.settle(item, 'stopped', { status: 'stopped' }));
          this.push(s, item);
          if (this.autoApply) this.apply(item, true);
        }),
      onNotice: (text) => {
        if (live()) this.push(s, { kind: 'notice', text, action: 'settings' });
      },
      overview: () => {
        const app = this.app;
        return workbookOverview(s.doc, this.shown(s) ? { sheet: app.sheet, ranges: app.sel.ranges, active: app.sel.active } : undefined);
      },
    });
  }

  /** Changes go straight into the sheet (with Undo) unless the user chose to approve them first. */
  get autoApply(): boolean {
    return !this.app.settings.bool('ai.confirm');
  }

  /** Still in development: on the Developer branch and in dev builds, desktop app only. */
  get available(): boolean {
    return isDesktop && (this.app.settings.bool('updates.devBranch') || __APP_VERSION__.includes('-'));
  }

  get provider(): 'claude' | 'local' {
    return this.app.settings.str('ai.provider') === 'local' ? 'local' : 'claude';
  }

  get modelLabel(): string {
    if (this.provider === 'local') return this.app.settings.str('ai.localModel') || 'Local model';
    const id = this.app.settings.str('ai.model');
    return CLAUDE_MODELS.find((m) => m[0] === id)?.[1] ?? id;
  }

  /** Whether the assistant can be used now; else what's missing. */
  get setupNeeded(): 'key' | 'model' | null {
    if (this.provider === 'claude') return this.hasKey === false ? 'key' : null;
    return this.app.settings.str('ai.localModel') ? null : 'model';
  }

  async refreshKey(): Promise<void> {
    this.hasKey = await platform.ai.hasKey();
    this.onChange?.();
  }

  private options(): AgentOptions {
    const s = this.app.settings;
    const autoApply = this.autoApply;
    return this.provider === 'local'
      ? { provider: 'local', model: s.str('ai.localModel'), baseUrl: s.str('ai.localUrl'), contextSize: Number(s.str('ai.contextSize')) || 16384, maxCells: s.num('ai.maxCells'), autoApply }
      : { provider: 'claude', model: s.str('ai.model'), maxCells: s.num('ai.maxCells'), autoApply };
  }

  async send(text: string): Promise<void> {
    text = text.trim();
    let s = this.chat;
    if (!text || s.running) return;
    if (containsSecret(text)) {
      this.push(s, { kind: 'error', text: "That message contains an API key, so it wasn't sent. Keys never go in the chat: add yours in Settings → AI assistant, where it's stored encrypted.", action: 'key' });
      return;
    }
    if (s.doc.wb !== s.wb) {
      this.newChat();
      s = this.chat;
    }
    this.smallModelTip(s);
    const gen = s.gen;
    s.running = true;
    s.current = null;
    this.push(s, { kind: 'user', text });
    const ctl = new AbortController();
    s.ctl = ctl;
    // The model only learns about Undo clicks from us
    const undone = s.items.filter((i): i is ProposalItem => i.kind === 'proposal' && !!i.undoneUnreported && i.state === 'undone');
    for (const i of undone) i.undoneUnreported = false;
    const prefix = undone.length ? `(Since your last reply I undid: ${undone.map((i) => `"${i.p.title}"`).join(', ')}. Those changes are no longer in the workbook.)\n\n` : '';
    try {
      await s.conv.send(prefix + text, this.options(), ctl.signal);
    } catch (e) {
      const msg = e instanceof AgentError ? e.message : `Something went wrong: ${(e as Error).message || String(e)}`;
      if (gen === s.gen) this.push(s, { kind: 'error', text: msg, action: /API key/i.test(msg.split(HIDDEN).join('')) ? 'key' : /Settings/.test(msg) ? 'settings' : undefined });
    } finally {
      // Unless New chat was clicked meanwhile (that chat has its own state)
      if (gen === s.gen) {
        s.ctl = null;
        s.running = false;
        s.current = null;
        this.notify(s);
      }
    }
  }

  private tipped = new Set<string>();

  /** Models under about 7B parameters rarely manage spreadsheet tools; say so once per model. */
  private smallModelTip(s: ChatSession): void {
    if (this.provider !== 'local') return;
    const model = this.app.settings.str('ai.localModel');
    const m = /(?:^|[:\-_ ])(\d+(?:\.\d+)?)b\b/i.exec(model);
    if (!m || Number(m[1]) >= 7 || this.tipped.has(model)) return;
    this.tipped.add(model);
    this.push(s, { kind: 'notice', text: `${model} is a small model. Models under about 7B often can't work with a spreadsheet reliably — they may answer with code or invented numbers instead of changing the sheet. For better results, try a larger model such as qwen3:8b or qwen3-coder:30b.`, action: 'settings' });
  }

  stop(): void {
    this.chat.ctl?.abort();
  }

  /** Starts over in the shown tab's chat. */
  newChat(): void {
    const s = this.chat;
    s.ctl?.abort();
    for (const it of s.items) if (it.kind === 'proposal' && it.state === 'pending') this.settle(it, 'stopped', { status: 'stopped' });
    s.gen++;
    s.ctl = null;
    s.running = false;
    s.items = [];
    s.current = null;
    s.wb = s.doc.wb;
    s.conv = this.makeConversation(s);
    this.notify(s);
  }

  /** The chat a proposal card belongs to. */
  private owner(item: ChatItem): ChatSession | undefined {
    for (const s of this.sessions.values()) if (s.items.includes(item)) return s;
    return undefined;
  }

  private settle(item: ProposalItem, state: ProposalItem['state'], outcome: ProposalOutcome, note?: string): void {
    if (item.state !== 'pending') return;
    item.state = state;
    item.note = note;
    const r = item.resolve;
    item.resolve = undefined;
    r?.(outcome);
    const s = this.owner(item);
    if (s) this.notify(s, item);
  }

  /** Puts a proposal into its own workbook (which may be a tab in the background). */
  apply(item: ProposalItem, auto = false): void {
    if (item.state !== 'pending') return;
    const s = this.owner(item);
    const doc = s?.doc;
    if (!doc || doc.wb !== item.wb) {
      this.settle(item, 'failed', { status: 'failed', error: 'That workbook has been closed or replaced.' }, 'That workbook has been closed or replaced.');
      return;
    }
    if (this.shown(s) && this.app.isEditing && !this.app.commitEdit()) return;
    try {
      const tx = applyChanges(doc, item.p.title, item.p.changes);
      item.tx = tx;
      this.settle(item, 'applied', { status: 'applied', label: tx.label, auto });
      if (this.shown(s)) this.show(item);
    } catch (e) {
      const msg = (e as Error).message || String(e);
      this.settle(item, 'failed', { status: 'failed', error: msg }, msg);
    }
  }

  /** The block of cells a change wrote on its main sheet (including tables its formulas fill). */
  changedRange(item: ProposalItem): { sheet: string; range: Range } | null {
    const edits = item.p.preview.edits;
    if (!edits.length) return null;
    const sheet = edits[0].sheet;
    let box: Range | null = null;
    for (const d of edits) {
      if (d.sheet !== sheet) continue;
      for (const ref of [d.cell, d.spill]) {
        const r = ref ? parseRange(ref) : null;
        if (!r) continue;
        box = box ? { r1: Math.min(box.r1, r.r1), c1: Math.min(box.c1, r.c1), r2: Math.max(box.r2, r.r2), c2: Math.max(box.c2, r.c2) } : r;
      }
    }
    return box ? { sheet, range: box } : null;
  }

  /** Selects what a change wrote, so it's easy to see on the sheet. */
  show(item: ProposalItem): void {
    const c = this.changedRange(item);
    if (c) this.app.goTo(`'${c.sheet.replace(/'/g, "''")}'!${formatRange(c.range)}`);
  }

  /** Advanced format on the cells a change wrote. */
  format(item: ProposalItem): void {
    if (item.state !== 'applied' || this.app.doc.wb !== item.wb) return;
    const c = this.changedRange(item);
    if (!c || !this.app.goTo(`'${c.sheet.replace(/'/g, "''")}'!${formatRange(c.range)}`)) return;
    this.app.advancedFormat();
  }

  reject(item: ProposalItem, feedback?: string): void {
    const f = feedback?.trim() ? redactSecrets(feedback.trim()) : undefined;
    this.settle(item, 'rejected', { status: 'rejected', feedback: f }, f);
  }

  undo(item: ProposalItem): void {
    if (item.state !== 'applied' || !item.tx) return;
    const s = this.owner(item);
    if (!s || s.doc.wb !== item.wb) return;
    if (this.shown(s) && this.app.isEditing && !this.app.commitEdit()) return;
    const latest = s.doc.history.peekUndo() === item.tx;
    const r = revertChanges(s.doc, item.tx);
    item.state = 'undone';
    item.undoneUnreported = true;
    item.note = !latest && r.skipped ? `Cells you changed afterwards were kept (${r.skipped}).` : undefined;
    this.notify(s, item);
  }

  /** Live data arrived: fill in previews that were waiting for it. */
  private dataArrived(): void {
    for (const s of this.sessions.values()) {
      const doc = s.doc;
      for (const it of s.items) {
        if (it.kind !== 'proposal' || it.wb !== doc.wb || !previewLoading(it.p.preview)) continue;
        if (it.state === 'pending') {
          const next = previewChanges(doc, it.p.changes);
          if (!next.error) it.p.preview = next;
        } else if (it.state === 'applied') {
          // Already in the sheet: show what the cells hold now
          for (const d of it.p.preview.edits) {
            const sh = doc.wb.sheetByName(d.sheet);
            const at = parseRange(d.cell);
            if (!sh || !at || !d.after.startsWith('#BUSY!')) continue;
            d.after = displayText(doc, sh, at.r1, at.c1);
            const sp = doc.engine.spillAt(sh, at.r1, at.c1);
            if (sp && (sp.r2 > sp.r1 || sp.c2 > sp.c1)) d.spill = formatRange(sp);
          }
        } else continue;
        this.notify(s, it);
      }
    }
  }

  /** Keeps cards in step with Undo/Redo, and stops when another file is opened into the tab. */
  private docChanged(s: ChatSession, e: DocEvent): void {
    if (e.type === 'reset' && s.doc.wb !== s.wb) {
      for (const it of s.items) if (it.kind === 'proposal' && it.state === 'pending') this.settle(it, 'stopped', { status: 'stopped' }, 'Another workbook was opened.');
      s.ctl?.abort();
      return;
    }
    if (e.type !== 'change' || e.direction === 'do') return;
    for (const it of s.items) {
      if (it.kind !== 'proposal' || it.tx !== e.tx) continue;
      if (e.direction === 'undo' && it.state === 'applied') {
        it.state = 'undone';
        it.undoneUnreported = true;
      } else if (e.direction === 'redo' && it.state === 'undone') {
        it.state = 'applied';
        it.undoneUnreported = false;
      } else continue;
      it.note = undefined;
      this.notify(s, it);
    }
  }

  /** Where an edited cell is, as a clickable address. */
  static address(sheet: string, cell: string): string {
    return `${/^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`}!${cell}`;
  }
}
