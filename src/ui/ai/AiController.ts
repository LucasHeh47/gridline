// The AI assistant in the app: one conversation per window, kept while the panel is closed. Reads
// are logged in the chat; proposals wait for Apply or Reject; applied changes keep an Undo button.
import { AgentError, Conversation, type AgentOptions, type AiTransport, type Proposal, type ProposalOutcome } from '../../core/ai/agent';
import { applyChanges, previewChanges, previewLoading, revertChanges } from '../../core/ai/changes';
import { formatRange, parseRange, type Range } from '../../core/address';
import { external } from '../../core/data/external';
import { displayText } from '../../core/ops/data';
import { workbookOverview } from '../../core/ai/workbook';
import type { DocEvent } from '../../core/doc';
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

export class AiController {
  items: ChatItem[] = [];
  running = false;
  hasKey: boolean | null = null;
  /** Repaints the panel; `item` alone changed when given. */
  onChange: ((item?: ChatItem) => void) | null = null;
  private conv: Conversation;
  private ctl: AbortController | null = null;
  private wb: Workbook;
  private current: { kind: 'assistant'; text: string } | null = null;
  /** Bumped by New chat, so a request still winding down can't touch the new chat. */
  private gen = 0;

  constructor(
    private app: App,
    private transport: AiTransport = platformTransport,
  ) {
    this.wb = app.doc.wb;
    this.conv = this.makeConversation();
    app.doc.on((e) => this.docChanged(e));
    // Previews run on a trial copy, so their downloads don't count as sheet changes: listen directly
    external.onChange(() => this.dataArrived());
  }

  /** A conversation whose callbacks only reach the chat while it's the current one. */
  private makeConversation(): Conversation {
    const gen = this.gen;
    const live = () => gen === this.gen;
    return new Conversation(this.app.doc, this.transport, {
      onStep: () => {
        if (live()) this.current = null;
      },
      onText: (t) => {
        if (!live()) return;
        if (!this.current) {
          this.current = { kind: 'assistant', text: '' };
          this.push(this.current);
        }
        this.current.text += t;
        this.onChange?.(this.current);
      },
      onRead: (text, error) => {
        if (!live()) return;
        this.current = null;
        const last = this.items[this.items.length - 1];
        if (last?.kind === 'reads') {
          last.lines.push({ text, error });
          this.onChange?.(last);
        } else this.push({ kind: 'reads', lines: [{ text, error }] });
      },
      propose: (p, signal) =>
        new Promise<ProposalOutcome>((resolve) => {
          if (!live()) return resolve({ status: 'stopped' });
          this.current = null;
          const item: ProposalItem = { kind: 'proposal', p, state: 'pending', wb: this.app.doc.wb, resolve };
          signal.addEventListener('abort', () => this.settle(item, 'stopped', { status: 'stopped' }));
          this.push(item);
          if (this.autoApply) this.apply(item, true);
        }),
      onNotice: (text) => {
        if (live()) this.push({ kind: 'notice', text, action: 'settings' });
      },
      overview: () => {
        const app = this.app;
        return workbookOverview(app.doc, { sheet: app.sheet, ranges: app.sel.ranges, active: app.sel.active });
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

  private push(item: ChatItem): void {
    this.items.push(item);
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
    if (!text || this.running) return;
    if (containsSecret(text)) {
      this.push({ kind: 'error', text: "That message contains an API key, so it wasn't sent. Keys never go in the chat: add yours in Settings → AI assistant, where it's stored encrypted.", action: 'key' });
      return;
    }
    if (this.app.doc.wb !== this.wb) this.newChat();
    this.smallModelTip();
    const gen = this.gen;
    this.running = true;
    this.current = null;
    this.push({ kind: 'user', text });
    const ctl = new AbortController();
    this.ctl = ctl;
    // The model only learns about Undo clicks from us
    const undone = this.items.filter((i): i is ProposalItem => i.kind === 'proposal' && !!i.undoneUnreported && i.state === 'undone');
    for (const i of undone) i.undoneUnreported = false;
    const prefix = undone.length ? `(Since your last reply I undid: ${undone.map((i) => `"${i.p.title}"`).join(', ')}. Those changes are no longer in the workbook.)\n\n` : '';
    try {
      await this.conv.send(prefix + text, this.options(), ctl.signal);
    } catch (e) {
      const msg = e instanceof AgentError ? e.message : `Something went wrong: ${(e as Error).message || String(e)}`;
      if (gen === this.gen) this.push({ kind: 'error', text: msg, action: /API key/i.test(msg.split(HIDDEN).join('')) ? 'key' : /Settings/.test(msg) ? 'settings' : undefined });
    } finally {
      // Unless New chat was clicked meanwhile (that chat has its own state)
      if (gen === this.gen) {
        this.ctl = null;
        this.running = false;
        this.current = null;
        this.onChange?.();
      }
    }
  }

  private tipped = new Set<string>();

  /** Models under about 7B parameters rarely manage spreadsheet tools; say so once per model. */
  private smallModelTip(): void {
    if (this.provider !== 'local') return;
    const model = this.app.settings.str('ai.localModel');
    const m = /(?:^|[:\-_ ])(\d+(?:\.\d+)?)b\b/i.exec(model);
    if (!m || Number(m[1]) >= 7 || this.tipped.has(model)) return;
    this.tipped.add(model);
    this.push({ kind: 'notice', text: `${model} is a small model. Models under about 7B often can't work with a spreadsheet reliably — they may answer with code or invented numbers instead of changing the sheet. For better results, try a larger model such as qwen3:8b or qwen3-coder:30b.`, action: 'settings' });
  }

  stop(): void {
    this.ctl?.abort();
  }

  newChat(): void {
    this.stop();
    for (const it of this.items) if (it.kind === 'proposal' && it.state === 'pending') this.settle(it, 'stopped', { status: 'stopped' });
    this.gen++;
    this.ctl = null;
    this.running = false;
    this.items = [];
    this.conv = this.makeConversation();
    this.wb = this.app.doc.wb;
    this.current = null;
    this.onChange?.();
  }

  get usage(): { input: number; output: number } {
    return this.conv.usage;
  }

  private settle(item: ProposalItem, state: ProposalItem['state'], outcome: ProposalOutcome, note?: string): void {
    if (item.state !== 'pending') return;
    item.state = state;
    item.note = note;
    const r = item.resolve;
    item.resolve = undefined;
    r?.(outcome);
    this.onChange?.(item);
  }

  apply(item: ProposalItem, auto = false): void {
    if (item.state !== 'pending') return;
    const doc = this.app.doc;
    if (doc.wb !== item.wb) {
      this.settle(item, 'failed', { status: 'failed', error: 'A different workbook is open now.' }, 'A different workbook is open now.');
      return;
    }
    if (this.app.isEditing && !this.app.commitEdit()) return;
    try {
      const tx = applyChanges(doc, item.p.title, item.p.changes);
      item.tx = tx;
      this.settle(item, 'applied', { status: 'applied', label: tx.label, auto });
      this.show(item);
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
    if (this.app.doc.wb !== item.wb) return;
    if (this.app.isEditing && !this.app.commitEdit()) return;
    const latest = this.app.doc.history.peekUndo() === item.tx;
    const r = revertChanges(this.app.doc, item.tx);
    item.state = 'undone';
    item.undoneUnreported = true;
    item.note = !latest && r.skipped ? `Cells you changed afterwards were kept (${r.skipped}).` : undefined;
    this.onChange?.(item);
  }

  /** Live data arrived: fill in previews that were waiting for it. */
  private dataArrived(): void {
    const doc = this.app.doc;
    for (const it of this.items) {
      if (it.kind !== 'proposal' || it.wb !== doc.wb || !previewLoading(it.p.preview)) continue;
      if (it.state === 'pending') {
        const next = previewChanges(doc, it.p.changes);
        if (!next.error) it.p.preview = next;
      } else if (it.state === 'applied') {
        // Already in the sheet: show what the cells hold now
        for (const d of it.p.preview.edits) {
          const s = doc.wb.sheetByName(d.sheet);
          const at = parseRange(d.cell);
          if (!s || !at || !d.after.startsWith('#BUSY!')) continue;
          d.after = displayText(doc, s, at.r1, at.c1);
          const sp = doc.engine.spillAt(s, at.r1, at.c1);
          if (sp && (sp.r2 > sp.r1 || sp.c2 > sp.c1)) d.spill = formatRange(sp);
        }
      } else continue;
      this.onChange?.(it);
    }
  }

  /** Keeps cards in step with Undo/Redo, and starts over when another file is opened. */
  private docChanged(e: DocEvent): void {
    if (e.type === 'reset' && this.app.doc.wb !== this.wb) {
      for (const it of this.items) if (it.kind === 'proposal' && it.state === 'pending') this.settle(it, 'stopped', { status: 'stopped' }, 'Another workbook was opened.');
      this.stop();
      return;
    }
    if (e.type !== 'change' || e.direction === 'do') return;
    for (const it of this.items) {
      if (it.kind !== 'proposal' || it.tx !== e.tx) continue;
      if (e.direction === 'undo' && it.state === 'applied') {
        it.state = 'undone';
        it.undoneUnreported = true;
      } else if (e.direction === 'redo' && it.state === 'undone') {
        it.state = 'applied';
        it.undoneUnreported = false;
      }
      else continue;
      it.note = undefined;
      this.onChange?.(it);
    }
  }

  /** Where an edited cell is, as a clickable address. */
  static address(sheet: string, cell: string): string {
    return `${/^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`}!${cell}`;
  }
}
