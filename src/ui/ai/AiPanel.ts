// The AI assistant panel: chat, a log of everything it read, and proposal cards that show exactly
// what would change (cells before → after, and formula results elsewhere) with Apply / Reject.
import * as L from 'lucide';
import type { CellDiff } from '../../core/ai/changes';
import type { App } from '../app';
import { h, icon } from '../dom';
import type { PanelContent } from '../panels/SidePanel';
import { platform } from '../platform';
import { dialog, toast } from '../widgets';
import { AiController, type ChatItem, type ProposalItem } from './AiController';
import { renderMarkdown, withRefs, type RefLinks } from './markdown';
import { validKeyFormat } from '../../shared/secrets';

const EXAMPLES = [
  "Find duplicate transactions, but don't delete anything.",
  'Build a monthly budget using these columns.',
  "Explain why my revenue total doesn't match the report.",
  'Create a chart showing which products are losing money.',
  'Add a formula that categorizes these transactions.',
];

/** Asks for the Claude API key and stores it encrypted on this computer. */
export async function askForKey(app: App): Promise<boolean> {
  const input = h('input', { class: 'input', type: 'password', placeholder: 'sk-ant-…', spellcheck: false, autocomplete: 'off', style: { width: '100%' } }) as HTMLInputElement;
  const body = h(
    'div',
    null,
    h('p', null, 'The assistant uses your own Claude account. Create a key at ', h('a', { href: '#', onclick: (e: Event) => {
      e.preventDefault();
      platform.openExternal('https://console.anthropic.com/settings/keys');
    } }, 'console.anthropic.com'), ' and paste it here. Usage is billed to that account.'),
    h('div', { class: 'field' }, h('label', null, 'Claude API key'), input),
    h('p', { class: 'hint' }, "The key is encrypted with your computer's keychain and only sent to Anthropic. Gridline never shows it again."),
  );
  setTimeout(() => input.focus(), 0);
  const ok = await dialog('Claude API key', body, [
    { label: 'Save key', value: true, primary: true },
    { label: 'Cancel', value: false },
  ]);
  // Take the key out of the page as soon as it's read; it's stored by the main process only
  const key = input.value.trim();
  input.value = '';
  if (!ok || !key) return false;
  if (!validKeyFormat(key)) {
    toast("That doesn't look like an API key. Copy the whole key from console.anthropic.com and paste it again.", { kind: 'error' });
    return false;
  }
  const r = await platform.ai.setKey(key);
  if (!r.ok) {
    toast(r.error ?? "Couldn't save the key.", { kind: 'error' });
    return false;
  }
  toast('Claude API key saved.', { kind: 'success' });
  await app.ai.refreshKey();
  return true;
}

export async function removeKey(app: App): Promise<void> {
  const r = await platform.ai.setKey(null);
  if (r.ok) toast('Claude API key removed from this computer.');
  await app.ai.refreshKey();
}

export class AiPanel implements PanelContent {
  title = 'AI assistant';
  readonly el: HTMLElement;
  private log: HTMLElement;
  private input: HTMLTextAreaElement;
  private sendBtn: HTMLButtonElement;
  private modelBtn: HTMLButtonElement;
  private usageEl: HTMLElement;
  private rendered = new Map<ChatItem, HTMLElement>();
  private expanded = new Set<string>();
  private links: RefLinks;

  constructor(private app: App) {
    const ai = app.ai;
    this.links = {
      valid: (ref) => {
        const bang = ref.lastIndexOf('!');
        if (bang < 0) return true;
        const name = ref.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
        return !!app.doc.wb.sheetByName(name);
      },
      open: (ref) => {
        if (!app.goTo(ref)) toast(`Couldn't find ${ref}.`);
      },
    };
    this.modelBtn = h('button', { class: 'btn ai-model', title: 'Change the model (Settings → AI assistant)', onclick: () => app.openSettings('ai') }, icon(L.Sparkles, 14), h('span', null, ai.modelLabel)) as HTMLButtonElement;
    const newChat = h('button', { class: 'btn', title: 'New chat', 'aria-label': 'New chat', onclick: () => {
      ai.newChat();
      this.input.focus();
    } }, icon(L.SquarePen, 16));
    this.usageEl = h('span', { class: 'ai-usage' });
    this.log = h('div', { class: 'ai-log', role: 'log', 'aria-live': 'polite' });
    this.input = h('textarea', { class: 'input ai-input', rows: 2, placeholder: 'Ask about this workbook…', spellcheck: true }) as HTMLTextAreaElement;
    this.input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
        e.preventDefault();
        this.submit();
      }
    });
    this.input.addEventListener('input', () => this.grow());
    this.sendBtn = h('button', { class: 'btn primary ai-send', onclick: () => (ai.running ? ai.stop() : this.submit()) }) as HTMLButtonElement;
    const compose = h('div', { class: 'ai-compose' }, h('div', { class: 'ai-input-row' }, this.input, this.sendBtn), h('div', { class: 'ai-foot' }, h('span', null, 'Nothing changes until you apply it.'), this.usageEl));
    this.el = h('div', { class: 'ai-panel' }, h('div', { class: 'ai-bar' }, this.modelBtn, h('span', { style: { flex: '1' } }), newChat), this.log, compose);
    ai.onChange = (item) => this.update(item);
    if (ai.hasKey === null) void ai.refreshKey();
    this.update();
    setTimeout(() => this.input.focus(), 0);
  }

  dispose(): void {
    if (this.app.ai.onChange) this.app.ai.onChange = null;
  }

  refresh(): void {
    (this.modelBtn.lastChild as HTMLElement).textContent = this.app.ai.modelLabel;
  }

  focusInput(): void {
    this.input.focus();
  }

  private grow(): void {
    this.input.style.height = 'auto';
    this.input.style.height = Math.min(160, this.input.scrollHeight + 2) + 'px';
  }

  private submit(text = this.input.value): void {
    const ai = this.app.ai;
    if (ai.running || !text.trim()) return;
    if (ai.setupNeeded) {
      this.update();
      return;
    }
    this.input.value = '';
    this.grow();
    void ai.send(text);
  }

  // ---------- rendering ----------

  private atBottom(): boolean {
    return this.log.scrollHeight - this.log.scrollTop - this.log.clientHeight < 60;
  }

  private update(item?: ChatItem): void {
    const ai = this.app.ai;
    const stick = this.atBottom();
    if (item && this.rendered.has(item)) {
      const old = this.rendered.get(item)!;
      const next = this.renderItem(item);
      old.replaceWith(next);
      this.rendered.set(item, next);
    } else {
      // Rebuild the list, reusing what's already drawn
      const keep = new Map<ChatItem, HTMLElement>();
      const nodes: HTMLElement[] = [];
      for (const it of ai.items) {
        const el = this.rendered.get(it) ?? this.renderItem(it);
        keep.set(it, el);
        nodes.push(el);
      }
      this.rendered = keep;
      this.log.replaceChildren(...(ai.items.length ? nodes : [this.empty()]));
      const setup = this.setupCard();
      if (setup) this.log.append(setup);
      if (ai.running) this.log.append(h('div', { class: 'ai-thinking' }, h('span'), h('span'), h('span')));
    }
    this.sendBtn.replaceChildren(ai.running ? icon(L.Square, 14) : icon(L.ArrowUp, 16));
    this.sendBtn.title = ai.running ? 'Stop' : 'Send (Enter)';
    this.sendBtn.setAttribute('aria-label', this.sendBtn.title);
    this.sendBtn.classList.toggle('stop', ai.running);
    const u = ai.usage;
    this.usageEl.textContent = u.input || u.output ? `${tokens(u.input)} in · ${tokens(u.output)} out` : '';
    this.usageEl.title = 'Tokens used in this chat';
    this.refresh();
    if (stick) this.log.scrollTop = this.log.scrollHeight;
  }

  private empty(): HTMLElement {
    return h(
      'div',
      { class: 'ai-empty' },
      icon(L.Sparkles, 28),
      h('h3', null, 'Ask about this workbook'),
      h('p', null, 'The assistant reads the sheets, formulas, charts and rules it needs — every read is listed here — and proposes changes you review before anything is modified. Applied changes undo in one click.'),
      h('div', { class: 'ai-examples' }, ...EXAMPLES.map((t) => h('button', { class: 'ai-example', onclick: () => this.submit(t) }, t))),
    );
  }

  private setupCard(): HTMLElement | null {
    const ai = this.app.ai;
    const need = ai.setupNeeded;
    if (!need) return null;
    if (need === 'key') {
      return h(
        'div',
        { class: 'ai-setup' },
        h('strong', null, 'Connect Claude'),
        h('p', null, 'Add your Claude API key to start. Or run a model on this computer with Ollama or LM Studio.'),
        h('div', { class: 'ai-row' }, h('button', { class: 'btn primary', onclick: () => void askForKey(this.app) }, icon(L.KeyRound, 14), 'Add API key'), h('button', { class: 'btn outline', onclick: () => this.app.openSettings('ai') }, 'Use a local model')),
      );
    }
    return h(
      'div',
      { class: 'ai-setup' },
      h('strong', null, 'Choose a local model'),
      h('p', null, 'Pick which model to use from your local AI server (Ollama or LM Studio) in Settings.'),
      h('div', { class: 'ai-row' }, h('button', { class: 'btn primary', onclick: () => this.app.openSettings('ai') }, 'Open AI settings')),
    );
  }

  private renderItem(it: ChatItem): HTMLElement {
    switch (it.kind) {
      case 'user':
        return h('div', { class: 'ai-msg user' }, it.text);
      case 'assistant':
        return h('div', { class: 'ai-msg bot' }, renderMarkdown(it.text, this.links));
      case 'reads':
        return h(
          'div',
          { class: 'ai-reads' },
          ...it.lines.map((l) => h('div', { class: `ai-read${l.error ? ' error' : ''}` }, icon(l.error ? L.CircleAlert : readIcon(l.text), 13), h('span', null, l.text))),
        );
      case 'error':
        return h(
          'div',
          { class: 'ai-error' },
          icon(L.CircleAlert, 16),
          h('div', null, h('div', null, it.text), it.action === 'key' ? h('button', { class: 'btn outline', onclick: () => void askForKey(this.app) }, 'Update API key') : it.action === 'settings' ? h('button', { class: 'btn outline', onclick: () => this.app.openSettings('ai') }, 'Open AI settings') : null),
        );
      case 'proposal':
        return this.card(it);
    }
  }

  private card(it: ProposalItem): HTMLElement {
    const ai = this.app.ai;
    const p = it.p;
    const pv = p.preview;
    const badge: Record<ProposalItem['state'], [string, string]> = {
      pending: ['Waiting for you', 'pending'],
      applied: ['Applied', 'ok'],
      rejected: ['Rejected', 'muted'],
      undone: ['Undone', 'muted'],
      stopped: ['Stopped', 'muted'],
      failed: ["Couldn't apply", 'bad'],
    };
    const [label, cls] = badge[it.state];
    const card = h(
      'div',
      { class: `ai-card ${it.state}` },
      h('div', { class: 'ai-card-head' }, icon(L.WandSparkles, 16), h('strong', null, p.title), h('span', { class: `ai-badge ${cls}` }, label)),
      p.summary ? h('p', { class: 'ai-card-sum' }, p.summary) : null,
      h('ul', { class: 'ai-notes' }, ...pv.notes.map((n) => h('li', null, ...withRefs(n, this.links)))),
    );
    if (pv.editCount) card.append(this.diff(`${p.id}:e`, `${pv.editCount.toLocaleString()} cell${pv.editCount === 1 ? '' : 's'} change`, pv.edits, pv.editCount, 6, true));
    if (pv.recalcCount) card.append(this.diff(`${p.id}:r`, `${pv.recalcCount.toLocaleString()} other result${pv.recalcCount === 1 ? '' : 's'} change`, pv.recalcs, pv.recalcCount, 4, false));
    const extra: string[] = [];
    if (pv.formatCount) extra.push(`Formatting changes in ${pv.formatCount.toLocaleString()} cell${pv.formatCount === 1 ? '' : 's'}.`);
    if (pv.structural) extra.push('Rows or columns move, so other formulas recalculate after applying.');
    if (!pv.editCount && !pv.formatCount && !pv.structural) extra.push('No cell values change.');
    for (const x of extra) card.append(h('p', { class: 'ai-card-extra' }, x));
    if (it.state === 'pending') {
      const feedback = h('input', { class: 'input ai-feedback', placeholder: 'Or say what to change instead…' }) as HTMLInputElement;
      feedback.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && feedback.value.trim()) {
          e.preventDefault();
          ai.reject(it, feedback.value);
        }
      });
      card.append(
        h(
          'div',
          { class: 'ai-row' },
          h('button', { class: 'btn primary', onclick: () => ai.apply(it) }, icon(L.Check, 14), 'Apply'),
          h('button', { class: 'btn outline', onclick: () => ai.reject(it, feedback.value) }, 'Reject'),
        ),
        feedback,
      );
    } else if (it.state === 'applied') {
      card.append(h('div', { class: 'ai-row' }, h('button', { class: 'btn outline', title: 'Undo everything this change did', onclick: () => ai.undo(it) }, icon(L.Undo2, 14), 'Undo'), h('span', { class: 'ai-card-note' }, 'One undo step: ', h('em', null, it.tx?.label ?? p.title))));
    }
    if (it.note) card.append(h('p', { class: 'ai-card-note' }, it.state === 'rejected' ? `Your feedback: “${it.note}”` : it.note));
    return card;
  }

  private diff(key: string, title: string, rows: CellDiff[], total: number, first: number, formulas: boolean): HTMLElement {
    const open = this.expanded.has(key);
    const shown = open ? rows : rows.slice(0, first);
    const body = h('tbody', null);
    for (const d of shown) {
      const ref = AiController.address(d.sheet, d.cell);
      body.append(
        h(
          'tr',
          null,
          h('td', null, h('a', { class: 'ai-ref', href: '#', title: `Go to ${ref}`, onclick: (e: Event) => {
            e.preventDefault();
            this.app.goTo(ref);
          } }, this.app.doc.wb.sheets.length > 1 || d.sheet !== this.app.sheet.name ? ref : d.cell)),
          h('td', { class: 'before' }, cellText(d.before, formulas ? d.beforeFormula : undefined)),
          h('td', { class: 'arrow' }, '→'),
          h('td', { class: 'after' }, cellText(d.after, formulas ? d.afterFormula : undefined)),
        ),
      );
    }
    const more = total - shown.length;
    const toggle =
      rows.length > first
        ? h('button', { class: 'btn link ai-more', onclick: () => {
            if (open) this.expanded.delete(key);
            else this.expanded.add(key);
            const item = this.app.ai.items.find((x) => x.kind === 'proposal' && key.startsWith(x.p.id + ':'));
            if (item) this.update(item);
          } }, open ? 'Show fewer' : `Show all ${total.toLocaleString()}`)
        : null;
    return h(
      'div',
      { class: 'ai-diff' },
      h('div', { class: 'ai-diff-title' }, title),
      h('div', { class: `ai-diff-scroll${open ? ' open' : ''}` }, h('table', null, body)),
      toggle,
      open && more > 0 ? h('div', { class: 'ai-card-note' }, `…and ${more.toLocaleString()} more not listed.`) : null,
    );
  }
}

function cellText(text: string, formula?: string): (HTMLElement | string)[] {
  const out: (HTMLElement | string)[] = [text === '' ? h('span', { class: 'faint' }, 'empty') : h('span', { class: 'val' }, text)];
  if (formula) out.push(h('div', { class: 'ai-formula', title: formula }, formula));
  return out;
}

function readIcon(text: string) {
  if (text.startsWith('Searched')) return L.Search;
  if (text.startsWith('Calculated')) return L.Calculator;
  if (text.startsWith('Traced')) return L.Waypoints;
  if (text.startsWith('Looked over')) return L.Table2;
  return L.Eye;
}

function tokens(n: number): string {
  return n >= 1000 ? `${(n / 1000).toFixed(n >= 100_000 ? 0 : 1)}k` : String(n);
}
