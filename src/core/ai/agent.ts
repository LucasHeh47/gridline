// The AI assistant's conversation loop: send the request with the workbook overview, run the tools
// the model asks for (reads run at once and are logged; proposals wait for the user), and repeat
// until the model answers. The model talks to us through a transport, so tests can script it.
import type { AiBlock, AiEvent, AiMessage, AiProvider, AiRequest } from '../../shared/api';
import type { Doc } from '../doc';
import { previewChanges, type Change, type Preview } from './changes';
import { runReadTool, TOOLS, type ProposalInput } from './tools';

export interface AiTransport {
  /** Streams one model response. Resolves when the response is finished (or cancelled). */
  chat(req: AiRequest, onEvent: (e: AiEvent) => void, signal: AbortSignal): Promise<void>;
}

export interface AgentOptions {
  provider: AiProvider;
  model: string;
  baseUrl?: string;
  /** Most cells one read returns. */
  maxCells: number;
  /** Most model calls for one message. */
  maxSteps?: number;
  /** For the system prompt ("today is …"). */
  today?: Date;
}

export interface Proposal extends ProposalInput {
  id: string;
  preview: Preview;
}

export type ProposalOutcome = { status: 'applied'; label: string } | { status: 'rejected'; feedback?: string } | { status: 'failed'; error: string } | { status: 'stopped' };

export interface AgentHooks {
  /** A new model response begins (the UI starts a new message bubble). */
  onStep?(): void;
  /** Streamed text from the model. */
  onText(delta: string): void;
  /** A read finished (shown as a line in the chat's read log). */
  onRead(log: string, isError: boolean): void;
  /** Shows a proposal and waits for the user to apply or reject it. */
  propose(p: Proposal, signal: AbortSignal): Promise<ProposalOutcome>;
  /** The workbook overview sent with each message. */
  overview(): string;
}

export class AgentError extends Error {}

const OVERVIEW_RE = /^<workbook>[\s\S]*?<\/workbook>\n\n/;

export function systemPrompt(today = new Date()): string {
  const d = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  return `You are the AI assistant built into Gridline, a spreadsheet app that opens and saves Excel files. You help the user understand, check and change their workbook. Today is ${d}.

How you work:
- Each user message starts with an overview of the workbook (sheets, headings, formulas, charts, rules). It is a summary: read the actual cells with your tools before relying on specific values. Every read is shown to the user.
- Use evaluate to compute totals, counts and comparisons (SUM, COUNTIFS, SUMPRODUCT, FILTER, UNIQUE, XLOOKUP…) instead of reading thousands of cells and doing arithmetic yourself. Never add up many numbers in your head.
- To change anything, call propose_changes. Nothing changes until the user approves the preview. Never say a change was made unless propose_changes reported that it was applied. If the user rejects a proposal, use their feedback or ask what they would like instead.
- If the user says not to change or delete something, respect it: answer, or propose only additive changes (highlighting, a new column, a new sheet).
- Prefer formulas over pasted results so the workbook stays live (e.g. a category column with IFS or XLOOKUP, a summary with SUMIFS). Use functions Excel 365 supports. Test a formula with evaluate before filling it down.
- Put new results beside the data (the next empty column, or a new sheet) rather than over existing cells, unless the user asks otherwise. Before writing to cells, make sure they are empty or that overwriting them is intended.
- Match the workbook's look: give new headings the same style and new numbers the same number format as their neighbors.
- For questions like "why doesn't this total match", trace with describe_cell and compare with evaluate. Common causes: rows left out of a SUM range, numbers stored as text, duplicates, hidden or filtered rows, rounding, references to the wrong sheet, mismatched criteria.
- For duplicates, decide what makes two rows the same (e.g. same date, amount and description), count them with COUNTIFS, and flag or highlight them; never delete rows unless asked.
- For charts, build a small summary table with formulas when the data isn't already summarized, then chart that table.
- Dates are serial numbers in formulas (DATE, EOMONTH, TEXT work as in Excel); reads show them as displayed.
- Changes and reads without a sheet name use the sheet the user is looking at; name the sheet whenever the workbook has more than one.
- When explaining, refer to cells as Sheet!A1 (the user can click them). Be concise: short paragraphs or lists, no large tables in the chat. Don't describe your tool calls step by step — the user sees them.`;
}

let seq = 0;
const newId = () => `p${Date.now().toString(36)}${(seq++).toString(36)}`;

/** One chat about a workbook. Keeps the message history between requests. */
export class Conversation {
  messages: AiMessage[] = [];
  usage = { input: 0, output: 0 };

  constructor(
    private doc: Doc,
    private transport: AiTransport,
    private hooks: AgentHooks,
  ) {}

  clear(): void {
    this.messages = [];
    this.usage = { input: 0, output: 0 };
  }

  /** Sends a user message and works until the model gives its answer. */
  async send(text: string, opts: AgentOptions, signal: AbortSignal): Promise<void> {
    // Only the newest message carries the (current) overview
    for (const m of this.messages) {
      if (m.role !== 'user') continue;
      for (const b of m.content) if (b.type === 'text') b.text = b.text.replace(OVERVIEW_RE, '');
    }
    this.messages.push({ role: 'user', content: [{ type: 'text', text: `<workbook>\n${this.hooks.overview()}\n</workbook>\n\n${text}` }] });
    const maxSteps = opts.maxSteps ?? 30;
    try {
      for (let step = 0; step < maxSteps; step++) {
        if (signal.aborted) return;
        this.shortenOldResults();
        const blocks = await this.callModel(opts, signal);
        if (!blocks.length) return;
        this.messages.push({ role: 'assistant', content: blocks });
        const calls = blocks.filter((b): b is Extract<AiBlock, { type: 'tool_use' }> => b.type === 'tool_use');
        if (!calls.length || signal.aborted) return;
        const results: AiBlock[] = [];
        let stopped = false;
        for (const call of calls) {
          if (stopped || signal.aborted) {
            results.push({ type: 'tool_result', tool_use_id: call.id, content: 'Stopped by the user.', is_error: true });
            continue;
          }
          const input = (call.input && typeof call.input === 'object' ? call.input : {}) as Record<string, unknown>;
          if (call.name === 'propose_changes') {
            const r = await this.propose(input, signal);
            if (r.stopped) stopped = true;
            results.push({ type: 'tool_result', tool_use_id: call.id, content: r.content, ...(r.isError ? { is_error: true } : {}) });
          } else {
            const r = runReadTool(this.doc, call.name, input, opts.maxCells);
            this.hooks.onRead(r.log, !!r.isError);
            results.push({ type: 'tool_result', tool_use_id: call.id, content: r.content, ...(r.isError ? { is_error: true } : {}) });
          }
        }
        this.messages.push({ role: 'user', content: results });
        if (stopped) return;
      }
      this.hooks.onStep?.();
      this.hooks.onText('I stopped here because this took more steps than allowed. Ask me to continue if you want me to keep going.');
    } finally {
      this.repair();
    }
  }

  /** Every tool request must be answered, or the next message is refused. */
  private repair(): void {
    const last = this.messages[this.messages.length - 1];
    if (last?.role !== 'assistant') return;
    const calls = last.content.filter((b) => b.type === 'tool_use') as Extract<AiBlock, { type: 'tool_use' }>[];
    if (calls.length) this.messages.push({ role: 'user', content: calls.map((c) => ({ type: 'tool_result' as const, tool_use_id: c.id, content: 'Stopped by the user.', is_error: true })) });
  }

  /** Older read results take up space on every request; keep the start of each. */
  private shortenOldResults(): void {
    const keep = 6;
    for (let i = 0; i < this.messages.length - keep; i++) {
      for (const b of this.messages[i].content) {
        if (b.type === 'tool_result' && b.content.length > 1500) b.content = b.content.slice(0, 400) + '\n… [older result shortened — read it again if needed]';
      }
    }
  }

  private async callModel(opts: AgentOptions, signal: AbortSignal): Promise<AiBlock[]> {
    const blocks: AiBlock[] = [];
    let error: string | null = null;
    this.hooks.onStep?.();
    await this.transport.chat(
      { provider: opts.provider, model: opts.model, baseUrl: opts.baseUrl, system: systemPrompt(opts.today), messages: this.messages, tools: TOOLS },
      (e) => {
        switch (e.type) {
          case 'text': {
            const last = blocks[blocks.length - 1];
            if (last?.type === 'text') last.text += e.text;
            else blocks.push({ type: 'text', text: e.text });
            this.hooks.onText(e.text);
            break;
          }
          case 'tool_use':
            blocks.push({ type: 'tool_use', id: e.id, name: e.name, input: e.input });
            break;
          case 'done':
            if (e.usage) {
              this.usage.input += e.usage.input;
              this.usage.output += e.usage.output;
            }
            break;
          case 'error':
            error = e.message;
            break;
        }
      },
      signal,
    );
    if (error) throw new AgentError(error);
    // Text-only whitespace replies aren't accepted back by the API
    return blocks.filter((b) => b.type !== 'text' || b.text.trim());
  }

  private async propose(input: Record<string, unknown>, signal: AbortSignal): Promise<{ content: string; isError?: boolean; stopped?: boolean }> {
    const changes = Array.isArray(input.changes) ? (input.changes.filter((c) => c && typeof c === 'object' && typeof (c as { op?: unknown }).op === 'string') as Change[]) : [];
    if (!changes.length) return { content: 'propose_changes needs a non-empty "changes" list.', isError: true };
    const title = String(input.title ?? '').trim() || 'Changes';
    const summary = String(input.summary ?? '').trim();
    const preview = previewChanges(this.doc, changes);
    if (preview.error) {
      this.hooks.onRead(`Proposal "${title}" couldn't be applied: ${preview.error}`, true);
      return { content: `This proposal couldn't be applied, so it wasn't shown to the user: ${preview.error} Nothing changed. Fix it and propose again.`, isError: true };
    }
    const outcome = await this.hooks.propose({ id: newId(), title, summary, changes, preview }, signal);
    switch (outcome.status) {
      case 'applied':
        return { content: `The user approved. The changes are now in the workbook as one undo step ("${outcome.label}"): ${preview.notes.join(' ')}` };
      case 'rejected':
        return { content: `The user rejected this proposal; nothing changed.${outcome.feedback ? ` Their feedback: ${outcome.feedback}` : ' Ask what they would like instead.'}` };
      case 'failed':
        return { content: `Applying failed and nothing changed: ${outcome.error}`, isError: true };
      case 'stopped':
        return { content: 'The user stopped the conversation before deciding; nothing changed.', isError: true, stopped: true };
    }
  }
}
