import { describe, expect, it } from 'vitest';
import type { CellAddr, Range } from '../src/core/address';
import type { Doc } from '../src/core/doc';
import type { MacroUI } from '../src/core/macros/host';
import { MacroProject } from '../src/core/macros/project';
import { MacroRunner } from '../src/core/macros/runner';
import { compileStepMacro, describeStep, newStep, STEP_EXAMPLES, variablesSetIn, type Step } from '../src/core/macros/steps';
import { addSheet } from '../src/core/ops/sheets';
import type { Sheet } from '../src/core/workbook';
import { makeDoc, val } from './helpers';

function ui(doc: Doc, out: string[], answers: string[]): MacroUI {
  let sel: { sheet: Sheet; ranges: Range[]; active: CellAddr } | null = null;
  return {
    msgBox: async (p) => {
      out.push('msg:' + p);
      return 1;
    },
    inputBox: async () => answers.shift() ?? '',
    print: (t) => out.push(t),
    status: () => undefined,
    selection: () => sel ?? { sheet: doc.wb.activeSheet, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: { row: 0, col: 0 } },
    select: (sheet, ranges, active) => {
      sel = { sheet, ranges, active };
    },
    activate: (sheet) => {
      doc.wb.activeSheetId = sheet.id;
      sel = null;
    },
    save: async () => true,
  };
}

/** A step with some fields filled in, and steps inside it. */
function s(type: string, args: Record<string, string> = {}, body?: Step[], elseBody?: Step[]): Step {
  const step = newStep(type);
  step.args = { ...step.args, ...args };
  if (body) step.body = body;
  if (elseBody) step.elseBody = elseBody;
  return step;
}

async function run(doc: Doc, steps: Step[], answers: string[] = []) {
  const project = (doc.wb.macros ??= new MacroProject());
  project.stepMacros = [{ id: 'm', name: 'Test', steps }];
  project.changed();
  const out: string[] = [];
  const runner = new MacroRunner(
    () => doc,
    () => ui(doc, out, answers),
  );
  runner.onPrint = (t) => out.push(t);
  const result = await runner.run('Test');
  if (result.error) throw new Error(`${result.error.message}\n${compileStepMacro('Test', steps).code}`);
  return { out, problems: runner.errors().map((e) => e.message) };
}

describe('step macro variables', () => {
  it('sets, calculates and changes numbers', async () => {
    const doc = makeDoc({ B2: '40' });
    await run(doc, [
      s('setVar', { name: 'Total', value: '10' }),
      s('mathVar', { name: 'Total', op: 'add', amount: '[B2]' }),
      s('mathVar', { name: 'Total', op: 'mul', amount: '1.5' }),
      s('setValue', { range: 'A1', value: '{Total}' }),
      s('mathVar', { name: 'Total', op: 'div', amount: '7' }),
      s('mathVar', { name: 'Total', op: 'round', amount: '2' }),
      s('setValue', { range: 'A2', value: '{Total}' }),
      s('setVar', { name: 'Half', value: '=ROUND({Total}/2, 0)' }),
      s('setValue', { range: 'A3', value: '{Half}' }),
      s('mathVar', { name: 'Never set', op: 'add', amount: '1' }),
      s('mathVar', { name: 'Never set', op: 'sub', amount: '5' }),
      s('mathVar', { name: 'Never set', op: 'abs' }),
      s('setValue', { range: 'A4', value: '{Never set}' }),
      s('setVar', { name: 'R', value: '17' }),
      s('mathVar', { name: 'R', op: 'mod', amount: '5' }),
      s('setValue', { range: 'A5', value: '{R}' }),
      s('mathVar', { name: 'R', op: 'atleast', amount: '10' }),
      s('setValue', { range: 'A6', value: '{R}' }),
    ]);
    expect([val(doc, 'A1'), val(doc, 'A2'), val(doc, 'A3'), val(doc, 'A4'), val(doc, 'A5'), val(doc, 'A6')]).toEqual([75, 10.71, 5, 4, 2, 10]);
  });

  it('treats answers like "1,200" as numbers and rounds like the spreadsheet', async () => {
    const doc = makeDoc();
    await run(
      doc,
      [
        s('ask', { question: 'How much?', name: 'Amount' }),
        s('mathVar', { name: 'Amount', op: 'add', amount: '0.5' }),
        s('mathVar', { name: 'Amount', op: 'round', amount: '0' }),
        s('setValue', { range: 'A1', value: '{Amount}' }),
        s('setVar', { name: 'Doubled', value: '={Amount}*2' }),
        s('setValue', { range: 'A2', value: '{Doubled}' }),
      ],
      ['1,200'],
    );
    expect(val(doc, 'A1')).toBe(1201);
    expect(val(doc, 'A2')).toBe(2402);
  });

  it('changes text', async () => {
    const doc = makeDoc();
    await run(doc, [
      s('setVar', { name: 'Name', value: '  ada   lovelace ' }),
      s('textVar', { name: 'Name', op: 'trim' }),
      s('textVar', { name: 'Name', op: 'proper' }),
      s('setValue', { range: 'A1', value: '{Name}' }),
      s('textVar', { name: 'Name', op: 'replace', text: 'LOVELACE', with: 'King' }),
      s('textVar', { name: 'Name', op: 'prepend', text: 'Countess ' }),
      s('textVar', { name: 'Name', op: 'append', text: '!' }),
      s('setValue', { range: 'A2', value: '{Name}' }),
      s('textVar', { name: 'Name', op: 'left', count: '8' }),
      s('textVar', { name: 'Name', op: 'upper' }),
      s('setValue', { range: 'A3', value: '{Name}' }),
      s('setVar', { name: 'Code', value: 'ab-12-cd' }),
      s('textVar', { name: 'Code', op: 'remove', text: '-' }),
      s('textVar', { name: 'Code', op: 'right', count: '4' }),
      s('setValue', { range: 'A4', value: '{Code}' }),
      // A text variable in a formula goes in quoted; inside quotes it goes in as it is
      s('setValue', { range: 'A5', value: '=LEN({Code})' }),
      s('setValue', { range: 'A6', value: '="<{Code}>"' }),
    ]);
    expect([val(doc, 'A1'), val(doc, 'A2'), val(doc, 'A3'), val(doc, 'A4'), val(doc, 'A5'), val(doc, 'A6')]).toEqual(['Ada Lovelace', 'Countess Ada King!', 'COUNTESS', '12cd', 4, '<12cd>']);
  });

  it('changes dates', async () => {
    const doc = makeDoc({ B1: '1/30/2026' });
    await run(doc, [
      s('setVar', { name: 'Due', value: '[B1]' }),
      s('dateVar', { name: 'Due', op: 'add', amount: '1', unit: 'm' }),
      s('setValue', { range: 'A1', value: '{Due}' }),
      s('dateVar', { name: 'Due', op: 'sub', amount: '2', unit: 'ww' }),
      s('setValue', { range: 'A2', value: '{Due}' }),
      s('setVar', { name: 'Friday', value: '1/2/2026' }),
      s('dateVar', { name: 'Friday', op: 'add', amount: '1', unit: 'wd' }),
      s('setValue', { range: 'A3', value: '{Friday}' }),
    ]);
    const serial = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
    expect(val(doc, 'A1')).toBe(serial(2026, 2, 28));
    expect(val(doc, 'A2')).toBe(serial(2026, 2, 14));
    expect(val(doc, 'A3')).toBe(serial(2026, 1, 5)); // Friday + 1 working day is Monday
  });

  it('gets totals and counts, and finds values', async () => {
    const doc = makeDoc({ A1: 'Item', A2: 'Rent', A3: 'Food', A4: 'Fun', B2: '1000', B3: '300', B4: 'n/a' });
    await run(doc, [
      s('measureVar', { what: 'sum', range: 'B2:B10', name: 'Sum' }),
      s('measureVar', { what: 'count', range: 'B2:B10', name: 'Numbers' }),
      s('measureVar', { what: 'counta', range: 'A1:A10', name: 'Filled' }),
      s('measureVar', { what: 'lastrow', range: 'A:A', name: 'Last' }),
      s('measureVar', { what: 'nextrow', range: 'D:D', name: 'Next' }),
      s('findVar', { find: 'food', range: 'A:A', what: 'right', name: 'Food' }),
      s('findVar', { find: 'Nope', range: 'A:A', what: 'found', name: 'Has nope' }),
      s('findVar', { find: 'Fun', range: 'A:A', what: 'row', name: 'Fun row' }),
      s('setValue', { range: 'D1', value: '{Sum} | {Numbers} | {Filled} | {Last} | {Next} | {Food} | {Has nope} | {Fun row}' }),
    ]);
    expect(val(doc, 'D1')).toBe('1300 | 2 | 4 | 4 | 1 | 300 | False | 4');
  });

  it('picks random numbers in range', async () => {
    const doc = makeDoc();
    const loop = s('repeat', { times: '20' }, [s('randomVar', { from: '1', to: '3', name: 'Roll' }), s('setValue', { range: 'A{Count}', value: '{Roll}' })]);
    await run(doc, [loop]);
    for (let r = 1; r <= 20; r++) expect([1, 2, 3]).toContain(val(doc, `A${r}`));
  });

  it('uses variables in rules, with a second rule', async () => {
    const doc = makeDoc({ A1: '5', B1: 'yes' });
    const rule = (args: Record<string, string>, cell: string) => s('if', args, [s('setValue', { range: cell, value: 'T' })], [s('setValue', { range: cell, value: 'F' })]);
    await run(doc, [
      s('setVar', { name: 'Limit', value: '10' }),
      s('ask', { question: '?', name: 'Answer' }),
      rule({ range: '{Limit}', op: '>', value: '[A1]' }, 'C1'),
      rule({ range: '{Answer}', op: '>', value: '9' }, 'C2'), // "12" as text is still more than 9
      rule({ range: 'A1', op: '>', value: '{Limit}' }, 'C3'),
      rule({ range: 'A1', op: '<', value: '{Limit}', join: 'and', range2: 'B1', op2: '=', value2: 'YES' }, 'C4'),
      rule({ range: 'A1', op: '>', value: '{Limit}', join: 'or', range2: 'B1', op2: 'contains', value2: 'e' }, 'C5'),
      rule({ range: 'A1', op: '>', value: '{Limit}', join: 'or', range2: 'B9', op2: 'notempty' }, 'C6'),
      rule({ range: 'B9', op: '<', value: '100' }, 'C7'), // a blank cell isn't compared with numbers
      rule({ range: '{Today}', op: '>', value: '1/1/2020' }, 'C8'),
    ], ['12']);
    expect(['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8'].map((a) => val(doc, a))).toEqual(['T', 'T', 'F', 'T', 'T', 'F', 'F', 'T']);
  });

  it('warns about variables that are never set, with a guess', async () => {
    const doc = makeDoc();
    const { problems } = await run(doc, [s('setVar', { name: 'Total', value: '1' }), s('setValue', { range: 'A1', value: '{Totl}' }), s('message', { text: '{Other}' }), s('if', { range: 'Total', op: '>', value: '0' })]);
    expect(problems).toContain('{Totl} is used but never set. Did you mean {Total}?');
    expect(problems).toContain('{Other} is used but never set. Add a "Set a variable" step before it.');
    expect(problems).toContain('"Total" isn\'t a cell. To check a variable, write {Total}.');
  });
});

describe('step macro loops', () => {
  it('repeats while a rule holds, and leaves loops early', async () => {
    const doc = makeDoc();
    await run(doc, [
      s('setVar', { name: 'N', value: '1' }),
      s('repeatWhile', { mode: 'while', range: '{N}', op: '<', value: '100' }, [s('mathVar', { name: 'N', op: 'mul', amount: '3' })]),
      s('setValue', { range: 'A1', value: '{N}' }),
      s('setVar', { name: 'Tries', value: '0' }),
      s('repeatWhile', { mode: 'until', range: '{Tries}', op: '>=', value: '1000' }, [s('setVar', { name: 'Tries', value: '{Count}' }), s('if', { range: '{Count}', op: '=', value: '4' }, [s('exitLoop')])]),
      s('setValue', { range: 'A2', value: '{Tries}' }),
      s('repeat', { times: '10' }, [s('if', { range: '{Count}', op: '>', value: '3' }, [s('exitLoop')]), s('setValue', { range: 'B{Count}', value: 'x' })]),
    ]);
    expect(val(doc, 'A1')).toBe(243);
    expect(val(doc, 'A2')).toBe(4);
    expect([val(doc, 'B3'), val(doc, 'B4')]).toEqual(['x', null]);
  });

  it('goes through each cell, each item of a list, and each sheet', async () => {
    const doc = makeDoc({ B1: '5', B2: 'hello', B3: '-2' });
    doc.transact('add', (tx) => addSheet(tx, doc.wb, 'Second', 1));
    doc.wb.activeSheetId = doc.wb.sheets[0].id;
    await run(doc, [
      s('eachCell', { range: 'B1:B3' }, [
        s('if', { range: '{Value}', op: '<', value: '0' }, [s('setValue', { range: '{Cell}', value: '0' })]),
        s('textVar', { name: 'Seen', op: 'append', text: '{Cell}={Value};' }),
      ]),
      s('setValue', { range: 'C1', value: '{Seen}' }),
      s('setVar', { name: 'Count', value: '1' }),
      s('eachItem', { list: 'Jan, Feb,, Mar,', sep: ',' }, [s('setValue', { range: 'D{Count}', value: '{Item}' }), s('mathVar', { name: 'Count', op: 'add', amount: '1' })]),
      s('eachSheet', {}, [s('setValue', { range: 'Z1', value: 'On {Sheet}' })]),
    ]);
    expect(val(doc, 'B3')).toBe(0);
    expect(val(doc, 'C1')).toBe('B1=5;B2=hello;B3=0;');
    // {Count} isn't a loop name here, so it's the variable the steps set
    expect([val(doc, 'D1'), val(doc, 'D2'), val(doc, 'D3'), val(doc, 'D4')]).toEqual(['Jan', 'Feb', 'Mar', null]);
    expect(val(doc, 'Z1')).toBe(`On ${doc.wb.sheets[0].name}`);
    expect(doc.engine.getValue(doc.wb.sheetByName('Second')!, 0, 25)).toBe('On Second');
    expect(doc.wb.activeSheetId).toBe(doc.wb.sheets[0].id);
  });

  it('can skip failing steps', async () => {
    const doc = makeDoc();
    await run(doc, [s('onError', { mode: 'skip' }), s('goSheet', { sheet: 'No such sheet' }), s('setValue', { range: 'A1', value: 'kept going' })]);
    expect(val(doc, 'A1')).toBe('kept going');
  });

  it('notes and output', async () => {
    const doc = makeDoc();
    const { out } = await run(doc, [s('note', { text: 'Just a note' }), s('setVar', { name: 'X', value: '42' }), s('log', { text: 'X is {X}' })]);
    expect(out).toContain('X is 42');
    expect(compileStepMacro('T', [s('note', { text: 'Just a note' })]).code).toContain("' Note: Just a note");
  });
});

describe('step macro builder helpers', () => {
  it('describes only the fields that apply', () => {
    expect(describeStep(s('if', { range: 'A1', op: 'empty', value: 'x' }))).toBe('If A1 is empty');
    expect(describeStep(s('if', { range: 'A1', op: '=', value: '5', join: 'or', range2: '{X}', op2: '>', value2: '2' }))).toBe('If A1 is 5 or {X} is more than 2');
    expect(describeStep(s('textVar', { name: 'Name', op: 'upper' }))).toBe('Change text in Name: make it UPPERCASE');
    expect(describeStep(s('textVar', { name: 'Name', op: 'left', count: '2' }))).toBe('Change text in Name: keep the first characters: 2');
  });

  it('lists the variables a macro sets', () => {
    const steps = [s('setVar', { name: 'Total' }), s('repeat', {}, [s('ask', { name: 'Answer' }), s('mathVar', { name: 'total' })])];
    expect(variablesSetIn(steps)).toEqual(['Total', 'Answer']);
  });

  it('runs the new examples', async () => {
    const doc = makeDoc({ C2: '250', C3: '50', C4: '20', D2: 'Paid', D3: 'Due', D4: 'paid', A2: 'x', A3: 'y', A4: 'z' });
    const { out } = await run(doc, STEP_EXAMPLES.find((e) => e.name === 'Add up paid amounts')!.steps());
    expect(out).toContain('msg:2 paid rows, 270 in total');
    const doc2 = makeDoc();
    await run(doc2, STEP_EXAMPLES.find((e) => e.name === 'A sheet per quarter')!.steps());
    expect(doc2.wb.sheets.map((x) => x.name).slice(1)).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(doc2.engine.getValue(doc2.wb.sheetByName('Q3')!, 0, 0)).toBe('Q3 report');
  });
});

