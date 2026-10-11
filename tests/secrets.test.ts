import { describe, expect, it } from 'vitest';
import { Conversation, type AiTransport } from '../src/core/ai/agent';
import { workbookOverview } from '../src/core/ai/workbook';
import type { AiRequest } from '../src/shared/api';
import { containsSecret, HIDDEN, redactSecrets, validKeyFormat } from '../src/shared/secrets';
import { makeDoc } from './helpers';

// Built at run time so these obviously fake keys don't look like real ones in the source
const fakeAnthropic = 'sk-ant-api03-' + 'Ab3_'.repeat(12);
const fakeGithub = 'ghp_' + 'a1B2'.repeat(9);

describe('API keys never reach the model', () => {
  it('recognizes and hides common key formats', () => {
    expect(containsSecret(`key: ${fakeAnthropic}`)).toBe(true);
    expect(containsSecret(fakeGithub)).toBe(true);
    expect(containsSecret('AKIA' + 'ABCDEFGHIJKLMNOP')).toBe(true);
    expect(containsSecret('Find duplicate transactions in sk-ant column')).toBe(false);
    expect(containsSecret('ask-ant-colony 2026')).toBe(false);
    expect(redactSecrets(`a ${fakeAnthropic} b ${fakeGithub}`)).toBe(`a ${HIDDEN} b ${HIDDEN}`);
    // An exact key is hidden whatever it looks like
    expect(redactSecrets('error: bad header "my-own-custom-key-123"', 'my-own-custom-key-123')).toBe(`error: bad header "${HIDDEN}"`);
  });

  it('only accepts keys that can safely go in a header', () => {
    expect(validKeyFormat(fakeAnthropic)).toBe(true);
    expect(validKeyFormat(fakeAnthropic + '\n')).toBe(false);
    expect(validKeyFormat('sk-ant with spaces in it 1234')).toBe(false);
    expect(validKeyFormat('short')).toBe(false);
  });

  it('keys in cells, messages and feedback are hidden from every request', async () => {
    const doc = makeDoc({ A1: 'Service', B1: 'Key', A2: 'Claude', B2: fakeAnthropic, A3: 'GitHub', B3: fakeGithub });
    const requests: AiRequest[] = [];
    const turns = [
      [{ type: 'tool_use' as const, id: 't1', name: 'read_range', input: { range: 'A1:B3' } }, { type: 'done' as const, stopReason: 'tool_use' }],
      [{ type: 'tool_use' as const, id: 't2', name: 'find_cells', input: { query: 'sk-ant' } }, { type: 'done' as const, stopReason: 'tool_use' }],
      [{ type: 'tool_use' as const, id: 't3', name: 'describe_sheet', input: {} }, { type: 'done' as const, stopReason: 'tool_use' }],
      [{ type: 'text' as const, text: 'Done.' }, { type: 'done' as const, stopReason: 'end_turn' }],
    ];
    const transport: AiTransport = {
      async chat(req, onEvent) {
        requests.push(JSON.parse(JSON.stringify(req)));
        for (const e of turns.shift() ?? []) onEvent(e);
      },
    };
    const conv = new Conversation(doc, transport, { onText: () => undefined, onRead: () => undefined, propose: async () => ({ status: 'stopped' }), overview: () => workbookOverview(doc) });
    await conv.send(`Here is my key ${fakeAnthropic}, what's in B2?`, { provider: 'claude', model: 'test', maxCells: 2000 }, new AbortController().signal);
    const sent = JSON.stringify(requests);
    expect(requests).toHaveLength(4);
    expect(sent).not.toContain(fakeAnthropic);
    expect(sent).not.toContain(fakeGithub);
    expect(sent).not.toContain('Ab3_Ab3_');
    expect(sent).toContain(HIDDEN);
    // The workbook itself is unchanged
    expect(doc.engine.getValue(doc.wb.activeSheet, 1, 1)).toBe(fakeAnthropic);
  });
});
