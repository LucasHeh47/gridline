import { describe, expect, it } from 'vitest';
import { Conversation, type AiTransport, type AgentHooks, type Proposal, type ProposalOutcome } from '../src/core/ai/agent';
import { applyChanges, previewChanges, revertChanges, type Change } from '../src/core/ai/changes';
import { runReadTool, TOOLS } from '../src/core/ai/tools';
import { workbookOverview } from '../src/core/ai/workbook';
import type { AiEvent, AiRequest } from '../src/shared/api';
import { makeDoc, set, val } from './helpers';

/** A small ledger: Date, Item, Qty, Price, Total (=C*D), and a grand total. */
function ledger() {
  const doc = makeDoc({
    A1: 'Date', B1: 'Item', C1: 'Qty', D1: 'Price', E1: 'Total',
    A2: '1/5/2026', B2: 'Coffee', C2: '2', D2: '3.5', E2: '=C2*D2',
    A3: '1/6/2026', B3: 'Bagel', C3: '1', D3: '2.25', E3: '=C3*D3',
    A4: '1/6/2026', B4: 'Coffee', C4: '2', D4: '3.5', E4: '=C4*D4',
    A5: '1/9/2026', B5: 'Juice', C5: '3', D5: '4', E5: '=C5*D5',
    G1: 'Grand total', G2: '=SUM(E2:E5)',
  });
  return doc;
}

const read = (doc: ReturnType<typeof makeDoc>, name: string, input: Record<string, unknown>, max = 2000) => runReadTool(doc, name, input, max);

describe('AI assistant: reading the workbook', () => {
  it('overview lists headings, formula runs and the selection', () => {
    const doc = ledger();
    const o = workbookOverview(doc);
    expect(o).toContain('Headers in row 1: A=Date, B=Item, C=Qty, D=Price, E=Total');
    expect(o).toContain('E2:E5 =C2*D2 (same formula in 4 rows)');
    expect(o).toContain('G2 =SUM(E2:E5)');
  });

  it('finds separate blocks of data on a sheet', () => {
    const doc = ledger();
    set(doc, { I1: 'Item', J1: 'Spent', I2: 'Coffee', J2: '=SUMIFS(E:E,B:B,I2)', I3: 'Juice', J3: '=SUMIFS(E:E,B:B,I3)' });
    expect(workbookOverview(doc)).toContain('Separate blocks of data: A1:E5 (5×5, first row: Date | Item | Qty | Price | Total); I1:J3 (3×2, first row: Item | Spent).');
  });

  it('reads ranges as tables, with formulas on request, and stops at the cell limit', () => {
    const doc = ledger();
    const r = read(doc, 'read_range', { range: 'A1:E3' });
    expect(r.content.split('\n')).toEqual(['Sheet1!A1:E3', 'row\tA\tB\tC\tD\tE', '1\tDate\tItem\tQty\tPrice\tTotal', '2\t1/5/2026\tCoffee\t2\t3.5\t7', '3\t1/6/2026\tBagel\t1\t2.25\t2.25']);
    expect(r.log).toBe('Read Sheet1!A1:E3 — 15 cells');
    expect(read(doc, 'read_range', { range: 'E2', show: 'both' }).content).toContain('7 [=C2*D2]');
    const cut = read(doc, 'read_range', { range: 'A:E' }, 10);
    expect(cut.content).toContain('… 3 more rows not shown');
    expect(read(doc, 'read_range', { range: 'Nope!A1' }).isError).toBe(true);
  });

  it('evaluates formulas, finds cells and traces dependencies', () => {
    const doc = ledger();
    expect(read(doc, 'evaluate', { formula: '=SUM(E2:E5)' }).content).toBe('28.25');
    expect(read(doc, 'evaluate', { formula: 'COUNTIFS(B2:B5,"Coffee",C2:C5,2)' }).content).toBe('2');
    expect(read(doc, 'evaluate', { formula: '=UNIQUE(B2:B5)' }).content).toContain('Array 3×1');
    expect(read(doc, 'find_cells', { query: 'coffee' }).content).toBe('Sheet1!B2: Coffee\nSheet1!B4: Coffee');
    expect(read(doc, 'find_cells', { query: 'SUM(', in_formulas: true }).content).toBe('Sheet1!G2: =SUM(E2:E5)');
    const cell = read(doc, 'describe_cell', { cell: 'E3' }).content;
    expect(cell).toContain('Formula: =C3*D3');
    expect(cell).toContain('Sheet1!C3 = 1');
    expect(cell).toContain('Used by: Sheet1!G2');
    const sheet = read(doc, 'describe_sheet', {}).content;
    expect(sheet).toContain('E "Total": 4 numbers (min 2.25, max 12, sum 28.25) 4 formulas');
  });

  it('every tool has a schema', () => {
    for (const t of TOOLS) expect(t.input_schema.type).toBe('object');
    expect(TOOLS.map((t) => t.name)).toEqual(['read_range', 'find_cells', 'describe_sheet', 'describe_cell', 'evaluate', 'propose_changes']);
  });
});

describe('AI assistant: previews, applying and undo', () => {
  const addCategory: Change[] = [
    { op: 'set_values', start: 'F1', values: [['Category']] },
    { op: 'fill_formula', range: 'F2:F5', formula: '=IF(B2="Coffee","Drinks","Food")' },
  ];

  it('previews exactly without changing anything', () => {
    const doc = ledger();
    const version = doc.history.version;
    const p = previewChanges(doc, [...addCategory, { op: 'set_values', start: 'C3', values: [[4]] }]);
    expect(p.error).toBeUndefined();
    expect(p.editCount).toBe(6);
    expect(p.edits.find((e) => e.cell === 'F2')).toMatchObject({ before: '', after: 'Drinks', afterFormula: '=IF(B2="Coffee","Drinks","Food")' });
    expect(p.edits.find((e) => e.cell === 'C3')).toMatchObject({ before: '1', after: '4' });
    // Side effects: E3 and the grand total recalculate
    expect(p.recalcs).toEqual([
      { sheet: 'Sheet1', cell: 'E3', before: '2.25', after: '9' },
      { sheet: 'Sheet1', cell: 'G2', before: '28.25', after: '35' },
    ]);
    expect(p.notes[1]).toBe('Fills Sheet1!F2:F5 with =IF(B2="Coffee","Drinks","Food") (adjusted for each cell).');
    // Nothing changed
    expect(doc.history.version).toBe(version);
    expect(val(doc, 'F2')).toBe(null);
    expect(val(doc, 'G2')).toBe(28.25);
  });

  it('ignores rewrites that change nothing', () => {
    const doc = ledger();
    const p = previewChanges(doc, [{ op: 'set_values', start: 'B2', values: [['Coffee', 2]] }, { op: 'fill_formula', range: 'E2:E5', formula: '=C2*D2' }]);
    expect([p.editCount, p.formatCount, p.recalcCount]).toEqual([0, 0, 0]);
  });

  it('reports a bad proposal instead of half-applying it', () => {
    const doc = ledger();
    const p = previewChanges(doc, [{ op: 'set_values', start: 'F1', values: [['x']] }, { op: 'rename_sheet', sheet: 'Missing', name: 'Y' }]);
    expect(p.error).toContain('Missing');
    expect(val(doc, 'F1')).toBe(null);
  });

  it('applies as one undo step and undoes with one click, even after later edits', () => {
    const doc = ledger();
    const tx = applyChanges(doc, 'Add a Category column', addCategory);
    expect(tx.label).toBe('AI: Add a Category column');
    expect(val(doc, 'F4')).toBe('Drinks');
    expect(val(doc, 'F3')).toBe('Food');
    // Latest change: a normal undo
    expect(revertChanges(doc, tx).skipped).toBe(0);
    expect(val(doc, 'F1')).toBe(null);
    // Not the latest: the user's own later edit to F3 is kept
    const tx2 = applyChanges(doc, 'Add a Category column', addCategory);
    set(doc, { F3: 'Bakery', H1: 'note' });
    expect(revertChanges(doc, tx2)).toMatchObject({ skipped: 1 });
    expect(val(doc, 'F2')).toBe(null);
    expect(val(doc, 'F3')).toBe('Bakery');
    expect(val(doc, 'H1')).toBe('note');
  });

  it('shows new rows as empty before', () => {
    const doc = ledger();
    const p = previewChanges(doc, [
      { op: 'insert_rows', before_row: 3 },
      { op: 'set_values', start: 'A3', values: [['1/5/2026', 'Tea', 1, 2]] },
    ]);
    expect(p.structural).toBe(true);
    expect(p.edits.filter((e) => e.cell.endsWith('3')).map((e) => e.before)).toEqual(['', '', '', '']);
    expect(val(doc, 'B3')).toBe('Bagel');
  });

  it('adds highlights, charts, sheets and dropdowns', () => {
    const doc = ledger();
    const tx = applyChanges(doc, 'Extras', [
      { op: 'highlight', range: 'B2:B5', rule: 'duplicates', color: 'yellow' },
      { op: 'add_sheet', name: 'Summary' },
      { op: 'set_values', sheet: 'Summary', start: 'A1', values: [['Item', 'Total'], ['Coffee', '=SUMIFS(Sheet1!E:E,Sheet1!B:B,A2)']] },
      { op: 'add_chart', sheet: 'Summary', type: 'column', data: 'A1:B2', title: 'By item' },
      { op: 'dropdown', range: 'B2:B5', choices: ['Coffee', 'Bagel', 'Juice'] },
    ]);
    const summary = doc.wb.sheetByName('Summary')!;
    expect(val(doc, 'B2', summary)).toBe(14);
    expect(summary.charts[0].title).toBe('By item');
    expect(doc.wb.sheets[0].cf).toHaveLength(1);
    expect(doc.wb.sheets[0].validations).toHaveLength(1);
    revertChanges(doc, tx);
    expect(doc.wb.sheetByName('Summary')).toBeUndefined();
    expect(doc.wb.sheets[0].cf).toHaveLength(0);
  });
});

/** Replays scripted model responses and records each request. */
function scripted(turns: AiEvent[][]): AiTransport & { requests: AiRequest[] } {
  const requests: AiRequest[] = [];
  return {
    requests,
    async chat(req, onEvent) {
      requests.push(JSON.parse(JSON.stringify(req)));
      for (const e of turns.shift() ?? [{ type: 'text', text: '(no more turns)' }, { type: 'done', stopReason: 'end_turn' }]) onEvent(e);
    },
  };
}

function hooks(doc: ReturnType<typeof makeDoc>, decide: (p: Proposal) => ProposalOutcome | Promise<ProposalOutcome>) {
  const log = { text: '', reads: [] as string[], proposals: [] as Proposal[] };
  const h: AgentHooks = {
    onText: (t) => (log.text += t),
    onRead: (l) => log.reads.push(l),
    propose: async (p) => {
      log.proposals.push(p);
      return decide(p);
    },
    overview: () => workbookOverview(doc),
  };
  return { h, log };
}

const opts = { provider: 'claude' as const, model: 'test', maxCells: 2000, today: new Date(2026, 9, 11) };

describe('AI assistant: the conversation loop', () => {
  it('reads, proposes, waits for approval, then answers', async () => {
    const doc = ledger();
    const t = scripted([
      [{ type: 'text', text: 'Let me look.' }, { type: 'tool_use', id: 't1', name: 'read_range', input: { range: 'A1:E5' } }, { type: 'done', stopReason: 'tool_use' }],
      [{ type: 'tool_use', id: 't2', name: 'propose_changes', input: { title: 'Add a Category column', summary: 'Labels each row.', changes: [{ op: 'set_values', start: 'F1', values: [['Category']] }, { op: 'fill_formula', range: 'F2:F5', formula: '=IF(B2="Coffee","Drinks","Food")' }] } }, { type: 'done', stopReason: 'tool_use' }],
      [{ type: 'text', text: 'Done — see Sheet1!F2.' }, { type: 'done', stopReason: 'end_turn' }],
    ]);
    const { h, log } = hooks(doc, (p) => {
      applyChanges(doc, p.title, p.changes);
      return { status: 'applied', label: 'AI: ' + p.title };
    });
    const conv = new Conversation(doc, t, h);
    await conv.send('Add a formula that categorizes these transactions.', opts, new AbortController().signal);
    expect(log.text).toBe('Let me look.Done — see Sheet1!F2.');
    expect(log.reads).toEqual(['Read Sheet1!A1:E5 — 25 cells']);
    expect(log.proposals[0].preview.editCount).toBe(5);
    expect(val(doc, 'F2')).toBe('Drinks');
    // The model saw the overview, the read and the decision
    expect(t.requests[0].messages[0].content[0]).toMatchObject({ type: 'text' });
    expect((t.requests[0].messages[0].content[0] as { text: string }).text).toMatch(/^<workbook>\n[\s\S]*<\/workbook>\n\nAdd a formula/);
    expect(t.requests[0].system).toContain('Today is 2026-10-11');
    const results = t.requests[2].messages[4].content[0] as { type: string; content: string };
    expect(results.content).toContain('The user approved');
    expect(conv.messages.map((m) => m.role)).toEqual(['user', 'assistant', 'user', 'assistant', 'user', 'assistant']);
  });

  it('passes rejection feedback back and drops old overviews', async () => {
    const doc = ledger();
    const t = scripted([
      [{ type: 'tool_use', id: 't1', name: 'propose_changes', input: { title: 'Delete duplicates', summary: '', changes: [{ op: 'delete_rows', row: 4 }] } }, { type: 'done', stopReason: 'tool_use' }],
      [{ type: 'text', text: 'OK, I will highlight them instead.' }, { type: 'done', stopReason: 'end_turn' }],
      [{ type: 'text', text: 'Hi.' }, { type: 'done', stopReason: 'end_turn' }],
    ]);
    const { h } = hooks(doc, () => ({ status: 'rejected', feedback: "Don't delete anything." }));
    const conv = new Conversation(doc, t, h);
    await conv.send('Find duplicate transactions.', opts, new AbortController().signal);
    expect(val(doc, 'B4')).toBe('Coffee');
    const res = t.requests[1].messages[2].content[0] as { content: string };
    expect(res.content).toBe("The user rejected this proposal; nothing changed. Their feedback: Don't delete anything.");
    await conv.send('Thanks', opts, new AbortController().signal);
    const msgs = t.requests[2].messages;
    expect((msgs[0].content[0] as { text: string }).text).toBe('Find duplicate transactions.');
    expect((msgs[msgs.length - 1].content[0] as { text: string }).text).toContain('<workbook>');
  });

  it('a broken proposal goes back to the model, not to the user', async () => {
    const doc = ledger();
    const t = scripted([
      [{ type: 'tool_use', id: 't1', name: 'propose_changes', input: { title: 'x', summary: '', changes: [{ op: 'set_values', start: 'ZZZZ1', values: [[1]] }] } }, { type: 'done', stopReason: 'tool_use' }],
      [{ type: 'text', text: 'Sorry.' }, { type: 'done', stopReason: 'end_turn' }],
    ]);
    const { h, log } = hooks(doc, () => ({ status: 'applied', label: '' }));
    await new Conversation(doc, t, h).send('Do it', opts, new AbortController().signal);
    expect(log.proposals).toHaveLength(0);
    expect((t.requests[1].messages[2].content[0] as { is_error?: boolean }).is_error).toBe(true);
  });

  it('stopping while a proposal waits leaves a valid history', async () => {
    const doc = ledger();
    const ctl = new AbortController();
    const t = scripted([[{ type: 'tool_use', id: 't1', name: 'propose_changes', input: { title: 'x', summary: '', changes: [{ op: 'set_values', start: 'F1', values: [[1]] }] } }, { type: 'tool_use', id: 't2', name: 'read_range', input: { range: 'A1' } }, { type: 'done', stopReason: 'tool_use' }]]);
    const { h } = hooks(doc, () => {
      ctl.abort();
      return { status: 'stopped' };
    });
    const conv = new Conversation(doc, t, h);
    await conv.send('Do it', opts, ctl.signal);
    const last = conv.messages[conv.messages.length - 1];
    expect(last.role).toBe('user');
    expect(last.content.map((b) => (b as { tool_use_id: string }).tool_use_id)).toEqual(['t1', 't2']);
    expect(val(doc, 'F1')).toBe(null);
  });

  it('model errors surface as exceptions', async () => {
    const doc = ledger();
    const t = scripted([[{ type: 'error', message: 'Your Claude API key was rejected.' }]]);
    const { h } = hooks(doc, () => ({ status: 'stopped' }));
    await expect(new Conversation(doc, t, h).send('Hi', opts, new AbortController().signal)).rejects.toThrow('API key was rejected');
  });
});
