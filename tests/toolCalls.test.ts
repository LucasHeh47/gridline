import { describe, expect, it } from 'vitest';
import { TOOLS } from '../src/core/ai/tools';
import { parseToolCalls, ToolCallStream } from '../src/shared/toolCalls';

// What qwen3-coder:30b sent through Ollama: the call written as text (no opening <tool_call>)
const QWEN = `I'll create a table with NFL quarterback stats. Let me first check what data we might need and set up the table structure.

<function=propose_changes>
<parameter=title>
Add NFL Quarterback Stats Table
</parameter>
<parameter=summary>
Creating a table with sample NFL quarterback statistics.
</parameter>
<parameter=changes>
[{"op":"set_values","start":"A1","values":[["Quarterback","Team","Games","Passing Yards"],["Joshua Dobbs","Miami Dolphins",16,2894]]}]
</parameter>
</function>
</tool_call>`;

/** Feeds text in random-sized chunks, like a stream. */
function stream(text: string, seed = 7): { shown: string; calls: ReturnType<ToolCallStream['finish']>['calls']; tail: string } {
  const s = new ToolCallStream(TOOLS);
  let shown = '';
  let x = seed;
  for (let i = 0; i < text.length; ) {
    x = (x * 1103515245 + 12345) % 2147483648;
    const n = 1 + (x % 9);
    shown += s.push(text.slice(i, i + n));
    i += n;
  }
  const end = s.finish();
  return { shown: shown + end.text, calls: end.calls, tail: end.text };
}

describe('tool calls written as text by local models', () => {
  it("turns Qwen3-Coder's XML into a real call, whatever the chunk sizes", () => {
    for (const seed of [1, 7, 42, 99, 1234]) {
      const r = stream(QWEN, seed);
      expect(r.shown).toBe("I'll create a table with NFL quarterback stats. Let me first check what data we might need and set up the table structure.\n\n");
      expect(r.calls).toEqual([
        {
          name: 'propose_changes',
          input: {
            title: 'Add NFL Quarterback Stats Table',
            summary: 'Creating a table with sample NFL quarterback statistics.',
            changes: [{ op: 'set_values', start: 'A1', values: [['Quarterback', 'Team', 'Games', 'Passing Yards'], ['Joshua Dobbs', 'Miami Dolphins', 16, 2894]] }],
          },
        },
      ]);
    }
  });

  it('reads the other common formats', () => {
    const hermes = parseToolCalls('Checking.\n<tool_call>\n{"name": "read_range", "arguments": {"range": "A1:D20"}}\n</tool_call>', TOOLS);
    expect(hermes).toEqual({ text: 'Checking.', calls: [{ name: 'read_range', input: { range: 'A1:D20' } }] });
    const llama = parseToolCalls('{"name": "evaluate", "parameters": {"formula": "=SUM(B:B)"}}', TOOLS);
    expect(llama.calls).toEqual([{ name: 'evaluate', input: { formula: '=SUM(B:B)' } }]);
    const fenced = parseToolCalls('```json\n{"name": "describe_sheet", "arguments": "{\\"sheet\\": \\"Data\\"}"}\n```', TOOLS);
    expect(fenced.calls).toEqual([{ name: 'describe_sheet', input: { sheet: 'Data' } }]);
    const two = parseToolCalls('<function=read_range><parameter=range>A1:B2</parameter></function><function=evaluate><parameter=formula>=1+1</parameter></function>', TOOLS);
    expect(two.calls.map((c) => c.name)).toEqual(['read_range', 'evaluate']);
    // Typed parameters follow the tool's schema
    const typed = parseToolCalls('<function=find_cells>\n<parameter=query>\n2024\n</parameter>\n<parameter=in_formulas>\ntrue\n</parameter>\n</function>', TOOLS);
    expect(typed.calls[0].input).toEqual({ query: '2024', in_formulas: true });
  });

  it('leaves ordinary text and code alone', () => {
    const plain = 'Use `=SUMIFS(C:C,B:B,"Food")` here.\n\n```\n=XLOOKUP(A2,Data!A:A,Data!C:C)\n```\nDone.';
    expect(stream(plain).shown).toBe(plain);
    expect(stream(plain).calls).toEqual([]);
    // A JSON example that isn't a call to one of our tools stays as text
    const example = 'Like this: {"name": "Bob", "age": 3}';
    expect(parseToolCalls(example, TOOLS)).toEqual({ text: example, calls: [] });
    // An unknown function name is not a call
    expect(parseToolCalls('<function=delete_everything></function>', TOOLS).calls).toEqual([]);
  });

  it('hides <think> reasoning', () => {
    const r = stream('<think>\nThe user wants totals. I should evaluate.\n</think>\n\nThe total is 42.');
    expect(r.shown).toBe('The total is 42.');
  });
});
