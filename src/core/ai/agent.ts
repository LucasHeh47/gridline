// The AI assistant's conversation loop: send the request with the workbook overview, run the tools
// the model asks for (reads run at once and are logged; changes are applied, or wait for the user
// when they've asked to approve first), and repeat until the model answers. The model talks to us through a transport, so tests can script it.
import type { AiBlock, AiEvent, AiMessage, AiProvider, AiRequest } from '../../shared/api';
import { redactSecrets } from '../../shared/secrets';
import type { Doc } from '../doc';
import { previewChanges, type Change, type Preview } from './changes';
import { FUNCTIONS } from '../formula/functions/registry';
import { runReadToolAsync, TOOLS, type ProposalInput } from './tools';
import { readRange } from './workbook';
import { external } from '../data/external';
import { previewLoading } from './changes';
import { formatRange, parseRange } from '../address';

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
  /** Changes go straight into the workbook (the user can undo them) instead of waiting for approval. */
  autoApply?: boolean;
  /** Local models: how many tokens they may read at once. */
  contextSize?: number;
}

export interface Proposal extends ProposalInput {
  id: string;
  preview: Preview;
}

export type ProposalOutcome = { status: 'applied'; label: string; auto?: boolean } | { status: 'rejected'; feedback?: string } | { status: 'failed'; error: string } | { status: 'stopped' };

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
  /** Something the user should know about the conversation itself (e.g. it no longer fits the model). */
  onNotice?(text: string): void;
}

export class AgentError extends Error {}

const OVERVIEW_RE = /^<workbook>[\s\S]*?<\/workbook>\n\n/;

/** Gridline's own functions (beyond Excel's), listed from the function registry. */
export function gridlineFunctions(): string {
  const groups: [string, string][] = [
    ['NFL', 'NFL (nflverse data, 1999 to today; season defaults to the current one)'],
    ['Personal finance', 'Live prices and personal finance'],
    ['World', 'Weather and world'],
  ];
  const out: string[] = [];
  for (const [cat, title] of groups) {
    const fns = [...FUNCTIONS.values()].filter((f) => f.category === cat).sort((a, b) => a.name.localeCompare(b.name));
    if (!fns.length) continue;
    out.push(`${title}:`);
    for (const f of fns) out.push(`- ${f.name}(${f.args}): ${f.desc}`);
  }
  return out.join('\n');
}

export function systemPrompt(today = new Date(), autoApply = false): string {
  const d = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  const changes = autoApply
    ? 'Changes are applied as soon as you call propose_changes; the user sees every changed cell and can undo it in one click.'
    : 'Nothing changes until the user approves your proposal. If they reject it, use their feedback or ask what they would like instead.';
  return `You are the AI assistant inside Gridline, a spreadsheet app (like Excel). You are connected to the user's open workbook: you can read it and change it, but only through your tools. Today is ${d}.

The most important rules:
1. To change the workbook, call the propose_changes tool. Never write changes, JSON, code or tables in your reply instead — they do nothing. ${changes}
2. Never make up data: no invented sports stats, prices, rates or weather. Gridline has live-data functions for these (listed below), e.g. =NFL.LEADERS("Passing Yards", 2025, 15, "QB") fills a ranked table of real quarterback stats. Write such formulas into cells; type numbers only when the user gave them to you.
3. Never say something was done unless the tool result says so, and describe only what it lists.

How you work:
- Each user message starts with <workbook>…</workbook>: an overview of the workbook (sheets, headings, formulas, charts, rules). Don't repeat it back. Read actual cells with read_range before relying on specific values; every read is shown to the user.
- Use evaluate to compute totals, counts and comparisons (SUM, COUNTIFS, FILTER, XLOOKUP…) and to try a formula before using it. Never add up many numbers in your head.
- A function that returns a table (NFL.LEADERS, NFL.PLAYERSTATS, NFL.SCHEDULE, FILTER, SORT, UNIQUE…) goes in ONE cell: it fills the cells to the right and below by itself, including its own header row. Never fill it down or across, don't type headers over it, and keep that area empty.
- Prefer formulas over pasted results so the workbook stays live (a category column with IFS or XLOOKUP, a summary with SUMIFS). Use functions Excel 365 supports.
- If the user says not to change or delete something, respect it: answer, or make only additive changes (highlighting, a new column, a new sheet).
- Put new results beside the data (the next empty column, or a new sheet), not over existing cells, unless asked. Check that target cells are empty first.
- Match the workbook's look: new headings like the existing ones, numbers in the same number format.
- For "why doesn't this total match", trace with describe_cell and compare with evaluate. Common causes: rows left out of a SUM range, numbers stored as text, duplicates, hidden or filtered rows, rounding, wrong sheet references.
- For duplicates, decide what makes two rows the same, count them with COUNTIFS, and highlight or flag them; never delete rows unless asked.
- For charts, build a small summary table with formulas when the data isn't summarized yet, then chart it.
- Never ask for API keys, passwords or other secrets. Values that look like keys appear as [hidden API key]; leave those cells alone.
- #BUSY! in a cell means live data is still downloading; it fills in by itself, so it isn't an error.
- Dates are serial numbers in formulas (DATE, EOMONTH, TEXT work as in Excel); reads show them as displayed.
- Changes and reads without a sheet name use the sheet the user is looking at; name the sheet when the workbook has more than one.
- When explaining, refer to cells as Sheet!A1 (the user can click them). Be concise: short paragraphs or lists, no big tables in the chat, and don't narrate your tool calls — the user sees them.

Gridline's own functions, on top of Excel's. They download live data, keep it offline and refresh on their own. Functions that return a table spill into the cells below and to the right (like FILTER): write the formula in one cell and leave that area empty. NFL stat names come from =NFL.STATS(), e.g. "Passing Yards", "Passing TDs", "Interceptions", "Completion %", "Passer Rating", "Rushing Yards", "Receptions", "Fantasy Points PPR". To fill a grid, pass ranges: =NFL.PLAYER("Josh Allen", B1:E1, A2:A6) gives seasons down and stats across; =NFL.PLAYER(A2:A16, B1:E1, 2025) gives players down and stats across.
${gridlineFunctions()}`;
}

/** Waits (up to 25 s) for live-data downloads until `done()` says enough has arrived. */
async function waitForData(done: () => boolean, signal: AbortSignal, ms = 25_000): Promise<void> {
  const deadline = Date.now() + ms;
  while (!done() && Date.now() < deadline && !signal.aborted) {
    await new Promise<void>((resolve) => {
      const timer = setTimeout(finish, Math.max(0, deadline - Date.now()));
      const off = external.onChange(finish);
      signal.addEventListener('abort', finish);
      function finish() {
        clearTimeout(timer);
        off();
        signal.removeEventListener('abort', finish);
        resolve();
      }
    });
  }
}

/** Topics Gridline has live data for (numbers about them should come from its functions). */
const LIVE_TOPIC = /\b(nfl|football|quarterbacks?|qbs?|running ?backs?|receivers?|touchdowns?|passing|rushing|fantasy|stocks?|shares?|tickers?|etfs?|crypto|bitcoin|btc|ethereum|eth|exchange rates?|currenc(?:y|ies)|weather|forecast|temperatures?|holidays?|population|gdp)\b/i;

/** Numbers a proposal types into cells (not formulas). */
function typedNumbers(changes: Change[]): number[] {
  const out: number[] = [];
  for (const c of changes) {
    if (c.op !== 'set_values' || !Array.isArray(c.values)) continue;
    for (const row of c.values) {
      for (const v of Array.isArray(row) ? row : [row]) {
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'string' && /^\s*-?[\d,]*\.?\d+%?\s*$/.test(v) && !v.trim().startsWith('=')) out.push(Number(v.replace(/[,%\s]/g, '')));
      }
    }
  }
  return out;
}

function numbersIn(text: string): Set<number> {
  return new Set((text.match(/-?\d[\d,]*\.?\d*/g) ?? []).map((n) => Number(n.replace(/,/g, ''))));
}

let seq = 0;
const newId = () => `p${Date.now().toString(36)}${(seq++).toString(36)}`;

/** One chat about a workbook. Keeps the message history between requests. */
export class Conversation {
  messages: AiMessage[] = [];
  usage = { input: 0, output: 0 };
  private userTexts: string[] = [];
  private nudged = false;
  private errorRetries = 0;
  private warned = false;

  constructor(
    private doc: Doc,
    private transport: AiTransport,
    private hooks: AgentHooks,
  ) {}

  clear(): void {
    this.messages = [];
    this.usage = { input: 0, output: 0 };
    this.userTexts = [];
    this.warned = false;
  }

  /** Sends a user message and works until the model gives its answer. */
  async send(text: string, opts: AgentOptions, signal: AbortSignal): Promise<void> {
    // Only the newest message carries the (current) overview
    for (const m of this.messages) {
      if (m.role !== 'user') continue;
      for (const b of m.content) if (b.type === 'text') b.text = b.text.replace(OVERVIEW_RE, '');
    }
    // Anything that looks like an API key (in the workbook or the message) is hidden from the model
    this.messages.push({ role: 'user', content: [{ type: 'text', text: `<workbook>\n${redactSecrets(this.hooks.overview())}\n</workbook>\n\n${redactSecrets(text)}` }] });
    this.userTexts.push(text);
    this.nudged = false;
    this.errorRetries = 0;
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
            const r = await this.propose(input, signal, !!opts.autoApply);
            if (r.stopped) stopped = true;
            results.push({ type: 'tool_result', tool_use_id: call.id, content: redactSecrets(r.content), ...(r.isError ? { is_error: true } : {}) });
          } else {
            const r = await runReadToolAsync(this.doc, call.name, input, opts.maxCells);
            this.hooks.onRead(r.log, !!r.isError);
            results.push({ type: 'tool_result', tool_use_id: call.id, content: redactSecrets(r.content), ...(r.isError ? { is_error: true } : {}) });
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
      { provider: opts.provider, model: opts.model, baseUrl: opts.baseUrl, contextSize: opts.contextSize, system: systemPrompt(opts.today, opts.autoApply), messages: this.messages, tools: TOOLS },
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
            if (e.warning && !this.warned) {
              this.warned = true;
              this.hooks.onNotice?.(e.warning);
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

  /** What the changed cells show now (first rows), so the model can check its work. */
  private snapshot(preview: Preview): string {
    const first = preview.edits[0];
    if (!first) return '';
    let box: { r1: number; c1: number; r2: number; c2: number } | null = null;
    for (const d of preview.edits) {
      if (d.sheet !== first.sheet) continue;
      for (const ref of [d.cell, d.spill]) {
        const r = ref ? parseRange(ref) : null;
        if (r) box = box ? { r1: Math.min(box.r1, r.r1), c1: Math.min(box.c1, r.c1), r2: Math.max(box.r2, r.r2), c2: Math.max(box.c2, r.c2) } : r;
      }
    }
    if (!box) return '';
    const shown = { ...box, r2: Math.min(box.r2, box.r1 + 5), c2: Math.min(box.c2, box.c1 + 9) };
    try {
      const t = readRange(this.doc, first.sheet, formatRange(shown), 'values', 80).text;
      return `\nThe changed cells now show (first rows of ${first.sheet}!${formatRange(box)}):\n${t}`;
    } catch {
      return '';
    }
  }

  private async propose(input: Record<string, unknown>, signal: AbortSignal, autoApply: boolean): Promise<{ content: string; isError?: boolean; stopped?: boolean }> {
    let raw = input.changes;
    // Some models send the list as a JSON string
    if (typeof raw === 'string') {
      try {
        raw = JSON.parse(raw);
      } catch {
        raw = [];
      }
    }
    const changes = Array.isArray(raw) ? (raw.filter((c) => c && typeof c === 'object' && typeof (c as { op?: unknown }).op === 'string') as Change[]) : [];
    if (!changes.length) return { content: 'propose_changes needs a non-empty "changes" list.', isError: true };
    const title = String(input.title ?? '').trim() || 'Changes';
    const summary = String(input.summary ?? '').trim();
    // Typing in stats or prices from memory: point the model at live data (once per message)
    const said = numbersIn(this.userTexts.join(' '));
    const typed = typedNumbers(changes).filter((n) => !said.has(n));
    if (!this.nudged && typed.length >= 8 && LIVE_TOPIC.test(this.userTexts[this.userTexts.length - 1] ?? '')) {
      this.nudged = true;
      this.hooks.onRead('Asked the AI to use live data instead of numbers typed from memory', true);
      return {
        content:
          `Not applied: this types ${typed.length} numbers into cells from memory, and they may be wrong or out of date. Gridline has live data for this — write its functions into the cells instead. ` +
          'For example =NFL.LEADERS("Passing Yards", 2025, 15, "QB") fills a ranked table of real quarterback stats by itself (leave the cells it spills into empty), and =NFL.PLAYER(A2:A16, B1:E1, 2025) fills a grid of players × stats. ' +
          'Check the formula with evaluate, then call propose_changes again. Only type numbers the user gave you.',
        isError: true,
      };
    }
    let preview = previewChanges(this.doc, changes);
    // Live data (NFL stats, prices…) still downloading: wait, so the preview and the error check
    // see real results instead of #BUSY!
    if (!preview.error && previewLoading(preview)) {
      await waitForData(() => !previewLoading(previewChanges(this.doc, changes)), signal);
      preview = previewChanges(this.doc, changes);
    }
    if (preview.error) {
      this.hooks.onRead(`Proposal "${title}" couldn't be applied: ${preview.error}`, true);
      return { content: `This proposal couldn't be applied, so it wasn't shown to the user: ${preview.error} Nothing changed. Fix it and propose again.`, isError: true };
    }
    // Formulas that would show errors are never applied: the model fixes them first
    if (preview.errorCount) {
      this.errorRetries++;
      const giveUp = this.errorRetries >= 3;
      this.hooks.onRead(`"${title}" would show ${preview.errorCount} formula error${preview.errorCount === 1 ? '' : 's'}; asked the AI to fix ${preview.errorCount === 1 ? 'it' : 'them'} first`, true);
      const list = preview.errors.map((e) => `${e.sheet}!${e.cell} ${e.code}${e.message ? ` (${e.message})` : ''}`).join('; ');
      return {
        content:
          `Not applied: ${preview.errorCount} of the formulas you wrote would show errors: ${list}${preview.errorCount > preview.errors.length ? '; …' : ''}. Nothing changed. ` +
          (giveUp
            ? 'This has failed several times, so stop trying: tell the user what you wanted to do and what went wrong.'
            : 'Fix them and call propose_changes again. Tips: a function that returns a table (NFL.LEADERS, NFL.PLAYERSTATS, FILTER, SORT, UNIQUE…) goes in ONE cell and fills the cells to the right and below by itself, including its own header row — don\'t fill it down or across, don\'t type headers over it, and leave that area empty. #SPILL! means cells in that area aren\'t empty: put the formula somewhere empty, or clear them in the same call. Try formulas with evaluate first.'),
        isError: true,
      };
    }
    const outcome = await this.hooks.propose({ id: newId(), title, summary, changes, preview }, signal);
    // Let downloads the change started finish, so the model sees results, not #BUSY!
    if (outcome.status === 'applied') await waitForData(() => !this.snapshot(preview).includes('#BUSY!'), signal);
    switch (outcome.status) {
      case 'applied': {
        const errs = preview.errorCount ? ` Warning: ${preview.errorCount} cell${preview.errorCount === 1 ? ' shows an error' : 's show errors'} (${preview.errors.map((e) => `${e.sheet}!${e.cell} ${e.code}`).join(', ')}); tell the user and offer to fix ${preview.errorCount === 1 ? 'it' : 'them'}.` : '';
        return {
          content: `${outcome.auto ? 'Applied' : 'The user approved'}. These changes are now in the workbook as one undo step ("${outcome.label}")${outcome.auto ? ', and the user can see and undo them' : ''}: ${preview.notes.join(' ')} Nothing else was changed, so when you tell the user what you did, describe only these.${errs}${this.snapshot(preview)}`,
        };
      }
      case 'rejected':
        return { content: `The user rejected this proposal; nothing changed.${outcome.feedback ? ` Their feedback: ${outcome.feedback}` : ' Ask what they would like instead.'}` };
      case 'failed':
        return { content: `Applying failed and nothing changed: ${outcome.error}`, isError: true };
      case 'stopped':
        return { content: 'The user stopped the conversation before deciding; nothing changed.', isError: true, stopped: true };
    }
  }
}
