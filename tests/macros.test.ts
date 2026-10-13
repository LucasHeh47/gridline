import { describe, expect, it } from 'vitest';
import type { CellAddr, Range } from '../src/core/address';
import type { Doc } from '../src/core/doc';
import type { MacroUI } from '../src/core/macros/host';
import { MacroProject } from '../src/core/macros/project';
import { r1c1ToA1 } from '../src/core/macros/range';
import { MacroRunner } from '../src/core/macros/runner';
import { compileStepMacro, newStep, stepHelperCode, STEP_EXAMPLES, STEPS } from '../src/core/macros/steps';
import type { Sheet } from '../src/core/workbook';
import { makeDoc, val } from './helpers';

function fakeUI(doc: Doc, out: string[] = [], answers: (string | number | null)[] = []): MacroUI {
  let sel: { sheet: Sheet; ranges: Range[]; active: CellAddr } | null = null;
  return {
    msgBox: async (p) => {
      out.push('msg:' + p);
      return (answers.shift() as number) ?? 1;
    },
    inputBox: async (p) => {
      out.push('ask:' + p);
      return (answers.shift() as string | null) ?? '';
    },
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

/** Puts `code` in Module1 and runs Main. */
async function runCode(doc: Doc, code: string, opts: { out?: string[]; answers?: (string | number | null)[]; macro?: string } = {}) {
  const project = (doc.wb.macros ??= new MacroProject());
  if (!project.module('Module1')) project.addModule('standard', 'Module1', code);
  else project.setCode('Module1', code);
  const out = opts.out ?? [];
  const runner = new MacroRunner(
    () => doc,
    () => fakeUI(doc, out, opts.answers ?? []),
  );
  runner.onPrint = (t) => out.push(t);
  const result = await runner.run(opts.macro ?? 'Main');
  return { result, out, runner };
}

describe('Excel object model', () => {
  it('reads and writes cells with Range and Cells', async () => {
    const doc = makeDoc({ A1: '10', A2: '20' });
    const { result } = await runCode(
      doc,
      `Sub Main()
  Range("B1").Value = Range("A1").Value * 2
  Cells(2, 2) = Cells(2, 1) + 5
  Range("C1:C3").Value = 7
  Range("D1").Formula = "=SUM(A1:A2)"
  Range("E1") = "hello"
  [F1] = 3
End Sub`,
    );
    expect(result.ok).toBe(true);
    expect(val(doc, 'B1')).toBe(20);
    expect(val(doc, 'B2')).toBe(25);
    expect(val(doc, 'C3')).toBe(7);
    expect(val(doc, 'D1')).toBe(30);
    expect(val(doc, 'E1')).toBe('hello');
    expect(val(doc, 'F1')).toBe(3);
  });

  it('undoes a whole macro in one step', async () => {
    const doc = makeDoc({ A1: '1' });
    await runCode(doc, 'Sub Main()\n  Dim i\n  For i = 1 To 50\n    Cells(i, 2) = i\n  Next\nEnd Sub');
    expect(val(doc, 'B50')).toBe(50);
    doc.undo();
    expect(val(doc, 'B50')).toBe(null);
    expect(val(doc, 'A1')).toBe(1);
  });

  it('sees formula results right after writing', async () => {
    const doc = makeDoc({ B1: '=A1*10' });
    const { out } = await runCode(doc, 'Sub Main()\n  Range("A1") = 4\n  Debug.Print Range("B1").Value\nEnd Sub');
    expect(out).toEqual(['40']);
  });

  it('finds the last row with End(xlUp) and loops over rows', async () => {
    const doc = makeDoc({ A1: 'Name', A2: 'a', A3: 'b', A4: 'c', B2: '1', B3: '2', B4: '3' });
    const { out } = await runCode(
      doc,
      `Sub Main()
  Dim last As Long, r As Long, total
  last = Cells(Rows.Count, 1).End(xlUp).Row
  For r = 2 To last
    total = total + Cells(r, 2).Value
  Next r
  Debug.Print last, total
  Debug.Print Range("A1").End(xlDown).Address, Range("A1").CurrentRegion.Address
  Debug.Print Range("A2").Offset(1, 1).Address(False, False), Range("A1:B2").Resize(3).Address
  Debug.Print Range("A1:B4").Rows.Count, Range("A1:B4").Columns.Count, Range("A1:B4").Count
End Sub`,
    );
    expect(out).toEqual(['4 6', '$A$4 $A$1:$B$4', 'B3 $A$1:$B$3', '4 2 8']);
  });

  it('formats cells', async () => {
    const doc = makeDoc({ A1: '5' });
    await runCode(
      doc,
      `Sub Main()
  With Range("A1:B2")
    .Font.Bold = True
    .Interior.Color = RGB(255, 0, 0)
    .NumberFormat = "0.00"
    .HorizontalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
  End With
  Range("A1").Font.Color = vbBlue
  Columns("C").ColumnWidth = 20
End Sub`,
    );
    const sheet = doc.wb.activeSheet;
    const st = doc.wb.styles.get(sheet.styleIdAt(1, 1));
    expect(st.bold).toBe(true);
    expect(st.fill).toBe('#ff0000');
    expect(st.numFmt).toBe('0.00');
    expect(st.hAlign).toBe('center');
    expect(st.borderTop?.style).toBe('thin');
    expect(doc.wb.styles.get(sheet.styleIdAt(0, 0)).color).toBe('#0000ff');
    expect(sheet.colWidth(2)).toBe(145);
  });

  it('works with sheets', async () => {
    const doc = makeDoc();
    const { out, result } = await runCode(
      doc,
      `Sub Main()
  Dim ws As Worksheet
  Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
  ws.Name = "Data"
  ws.Range("A1") = "x"
  Worksheets("Sheet1").Activate
  Debug.Print Worksheets.Count, ActiveSheet.Name, Sheets(2).Name, ws.Index
  For Each ws In Worksheets
    Debug.Print ws.Name
  Next
  Worksheets("Data").Range("B1").Value = Worksheets("Data").Range("A1").Value & "y"
End Sub`,
    );
    expect(result.ok).toBe(true);
    expect(out).toEqual(['2 Sheet1 Data 2', 'Sheet1', 'Data']);
    expect(val(doc, 'B1', doc.wb.sheetByName('Data'))).toBe('xy');
  });

  it('reads and writes arrays', async () => {
    const doc = makeDoc({ A1: '1', A2: '2', B1: '3', B2: '4' });
    const { out } = await runCode(
      doc,
      `Sub Main()
  Dim v
  v = Range("A1:B2").Value
  Debug.Print UBound(v, 1), UBound(v, 2), v(2, 1)
  v(1, 1) = 100
  Range("D1:E2").Value = v
  Range("D4:F4").Value = Array("a", "b", "c")
End Sub`,
    );
    expect(out).toEqual(['2 2 2']);
    expect(val(doc, 'D1')).toBe(100);
    expect(val(doc, 'E2')).toBe(4);
    expect(val(doc, 'F4')).toBe('c');
  });

  it('calls worksheet functions', async () => {
    const doc = makeDoc({ A1: 'apple', B1: '3', A2: 'pear', B2: '5' });
    const { out } = await runCode(
      doc,
      `Sub Main()
  Debug.Print WorksheetFunction.Sum(Range("B1:B2")), Application.WorksheetFunction.VLookup("pear", Range("A1:B2"), 2, False)
  Debug.Print WorksheetFunction.Max(1, 9, 4), Application.Evaluate("2*21"), WorksheetFunction.CountIf(Range("B1:B2"), ">3")
  On Error Resume Next
  Dim x
  x = WorksheetFunction.VLookup("kiwi", Range("A1:B2"), 2, False)
  Debug.Print Err.Number
End Sub`,
    );
    expect(out).toEqual(['8 5', '9 42 1', '1004']);
  });

  it('copies, sorts, finds and deletes rows', async () => {
    const doc = makeDoc({ A1: 'Name', B1: 'Score', A2: 'Cy', B2: '7', A3: 'Al', B3: '9', A4: 'Bo', B4: '2' });
    const { out } = await runCode(
      doc,
      `Sub Main()
  Range("A1:B4").Sort Key1:=Range("B1"), Order1:=xlDescending, Header:=xlYes
  Debug.Print Range("A2").Value, Range("A4").Value
  Dim f As Range
  Set f = Range("A1:A4").Find("bo")
  Debug.Print f.Address
  Range("A1:B4").Copy Destination:=Range("D1")
  Rows(2).Delete
  Debug.Print Range("A2").Value, Range("E3").Value
End Sub`,
    );
    expect(out).toEqual(['Al Bo', '$A$4', 'Cy 2']);
  });

  it('converts R1C1 formulas', () => {
    expect(r1c1ToA1('=SUM(R[-3]C:R[-1]C)', 4, 1)).toBe('=SUM(B2:B4)');
    expect(r1c1ToA1('=RC[-1]*R1C1', 2, 2)).toBe('=B3*$A$1');
    expect(r1c1ToA1('="RC"&ROUND(RC[1],0)', 0, 0)).toBe('="RC"&ROUND(B1,0)');
  });

  it('writes FormulaR1C1 like the macro recorder', async () => {
    const doc = makeDoc({ A1: '2', A2: '3' });
    await runCode(doc, 'Sub Main()\n  Range("B1:B2").FormulaR1C1 = "=RC[-1]*10"\nEnd Sub');
    expect(val(doc, 'B2')).toBe(30);
  });

  it('reports errors with module and line', async () => {
    const doc = makeDoc();
    const { result } = await runCode(doc, 'Sub Main()\n  Dim ws\n  Set ws = Worksheets("Nope")\nEnd Sub');
    expect(result.ok).toBe(false);
    expect(result.error).toMatchObject({ number: 9, module: 'Module1', line: 3 });
  });

  it('runs Worksheet_Change for edits but not for macro changes', async () => {
    const doc = makeDoc();
    const project = (doc.wb.macros = new MacroProject());
    const runner = new MacroRunner(
      () => doc,
      () => fakeUI(doc),
    );
    // Creates the Sheet1 document module
    await runCode(doc, 'Sub Main()\nEnd Sub');
    project.setCode('Sheet1', 'Private Sub Worksheet_Change(ByVal Target As Range)\n  Range("Z1").Value = Range("Z1").Value + 1\n  Target.Offset(0, 1).Value = "seen"\nEnd Sub');
    await runner.fireChange(doc.wb.activeSheet, [{ r1: 0, c1: 0, r2: 0, c2: 0 }]);
    expect(val(doc, 'Z1')).toBe(1);
    expect(val(doc, 'B1')).toBe('seen');
  });

  it('can cancel saving from Workbook_BeforeSave', async () => {
    const doc = makeDoc();
    await runCode(doc, 'Sub Main()\nEnd Sub');
    doc.wb.macros!.setCode('ThisWorkbook', 'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\n  Cancel = True\nEnd Sub');
    const runner = new MacroRunner(
      () => doc,
      () => fakeUI(doc),
    );
    expect((await runner.fireBeforeSave(false)).cancel).toBe(true);
  });
});

describe('step macros', () => {
  it('compiles every step type to valid VBA', () => {
    for (const def of STEPS) {
      const step = newStep(def.type);
      const { code, helpers } = compileStepMacro('Test', [step]);
      const project = new MacroProject();
      project.addModule('standard', 'M', code + '\r\n\r\n' + stepHelperCode(helpers));
      const runner = new MacroRunner(
        () => {
          const d = makeDoc();
          d.wb.macros = project;
          return d;
        },
        () => ({}) as MacroUI,
      );
      expect(runner.errors(), `${def.type}:\n${code}`).toEqual([]);
    }
  });

  it('runs the example step macros', async () => {
    const doc = makeDoc({ A1: 'Date', A2: '1/1/2026', B1: 'Item', C2: '250', C3: '50', D2: 'Paid', D3: 'Due', A3: '1/2/2026' });
    const project = (doc.wb.macros = new MacroProject());
    const byName = (n: string) => STEP_EXAMPLES.find((e) => e.name === n)!;
    for (const n of ['Log today', 'Highlight big expenses', 'Mark paid rows']) project.stepMacros.push({ id: n, name: n, steps: byName(n).steps() });
    project.changed();
    const runner = new MacroRunner(
      () => doc,
      () => fakeUI(doc),
    );
    expect(runner.list().map((m) => m.label)).toEqual(['Log today', 'Highlight big expenses', 'Mark paid rows']);
    const sheet = doc.wb.activeSheet;
    const fill = (r: number, c: number) => doc.wb.styles.get(sheet.styleIdAt(r, c)).fill;
    expect((await runner.run('Log today')).error).toBeUndefined();
    expect(val(doc, 'B4')).toBe('New entry');
    expect(typeof val(doc, 'A4')).toBe('number');
    expect((await runner.run('Highlight big expenses')).error).toBeUndefined();
    expect(fill(1, 2)).toBe('#fef08a');
    expect(fill(2, 2)).toBeUndefined();
    expect((await runner.run('Mark paid rows')).error).toBeUndefined();
    expect(fill(1, 0)).toBe('#bbf7d0');
    expect(fill(2, 0)).toBeUndefined();
  });

  it('asks questions and uses the answers', async () => {
    const doc = makeDoc();
    const project = (doc.wb.macros = new MacroProject());
    const ask = newStep('ask');
    ask.args = { question: 'Name?', name: 'Who' };
    const set = newStep('setValue');
    set.args = { range: 'A1', value: 'Hi {Who}!' };
    project.stepMacros.push({ id: 'x', name: 'Greet', steps: [ask, set] });
    project.changed();
    const out: string[] = [];
    const runner = new MacroRunner(
      () => doc,
      () => fakeUI(doc, out, ['Sam']),
    );
    await runner.run('Greet');
    expect(val(doc, 'A1')).toBe('Hi Sam!');
  });
});
