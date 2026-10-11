import { describe, expect, it } from 'vitest';
import { VbaProgram, type VbaHost } from '../src/core/macros/vba/interpreter';
import { tokenize } from '../src/core/macros/vba/lexer';
import { parseModule } from '../src/core/macros/vba/parser';
import { toStr, VbArray, VbError, type Value } from '../src/core/macros/vba/values';

/** A host with no workbook: Debug.Print and MsgBox go to `out`. */
function stubHost(out: string[], answers: (number | string | null)[] = []): VbaHost {
  return {
    application: { vbType: 'Application', get: () => undefined, set: () => undefined },
    global: () => undefined,
    isGlobal: () => false,
    documentObject: () => null,
    msgBox: async (p) => {
      out.push('msg:' + p);
      return (answers.shift() as number) ?? 1;
    },
    inputBox: async (p) => {
      out.push('input:' + p);
      return (answers.shift() as string | null) ?? '';
    },
    print: (t) => out.push(t),
    tick: async () => undefined,
    evaluate: () => 0,
  };
}

async function run(code: string, macro = 'Main', answers?: (number | string | null)[]): Promise<string[]> {
  const out: string[] = [];
  const program = new VbaProgram([{ name: 'Module1', code, kind: 'standard' }]);
  if (program.errors.length) throw new Error(program.errors.map((e) => `${e.line}: ${e.message}`).join('\n'));
  await program.run(stubHost(out, answers), macro);
  return out;
}

async function evalExpr(expr: string): Promise<string> {
  const out = await run(`Sub Main()\n  Debug.Print ${expr}\nEnd Sub`);
  return out[0];
}

describe('VBA lexer', () => {
  it('joins line continuations and drops comments', () => {
    const toks = tokenize('x = 1 + _\n  2 \' comment\nRem another\ny = "a""b"');
    const text = toks.filter((t) => t.type !== 'nl' && t.type !== 'eof').map((t) => t.text);
    expect(text).toEqual(['x', '=', '1', '+', '2', 'y', '=', 'a"b']);
  });

  it('reads hex, dates and type suffixes', () => {
    const toks = tokenize('a = &HFF + 10& + #1/2/2026#');
    expect(toks.find((t) => t.text === '&HFF')?.value).toBe(255);
    expect(toks.find((t) => t.type === 'date')?.value).toBe(46024);
    expect(toks.find((t) => t.text === '10&')?.value).toBe(10);
  });
});

describe('VBA parser', () => {
  it('reads procedures, module variables and options', () => {
    const ast = parseModule(`Option Explicit
Private total As Long
Public Const RATE = 0.05
Sub One()
End Sub
Private Function Two(ByVal a As Long, Optional b = 2) As Long
  Two = a + b
End Function`);
    expect(ast.optionExplicit).toBe(true);
    expect(ast.vars[0].dim.name).toBe('total');
    expect(ast.consts[0].name).toBe('RATE');
    expect(ast.procs.map((p) => p.name)).toEqual(['One', 'Two']);
    expect(ast.procs[1].params[0].byVal).toBe(true);
    expect(ast.procs[1].params[1].optional).toBe(true);
  });

  it('reports syntax errors with a line number', () => {
    expect(() => parseModule('Sub A()\n  If x Then\nEnd Sub')).toThrow();
    try {
      parseModule('Sub A()\n  x = (1 + \nEnd Sub');
    } catch (e) {
      expect((e as { line: number }).line).toBe(2);
    }
  });
});

describe('VBA interpreter', () => {
  it('does arithmetic and string operations', async () => {
    expect(await evalExpr('1 + 2 * 3')).toBe('7');
    expect(await evalExpr('2 ^ 3 ^ 2')).toBe('64');
    expect(await evalExpr('-2 ^ 2')).toBe('-4');
    expect(await evalExpr('7 \\ 2 & "|" & 7 Mod 3')).toBe('3|1');
    expect(await evalExpr('"a" & 1 + 2')).toBe('a3');
    expect(await evalExpr('"5" + 3')).toBe('8');
    expect(await evalExpr('UCase(Left("hello", 3)) & Mid("abcdef", 2, 3)')).toBe('HELbcd');
    expect(await evalExpr('InStr("banana", "an")')).toBe('2');
    expect(await evalExpr('Replace("a-b-c", "-", "+")')).toBe('a+b+c');
    expect(await evalExpr('Format(1234.5, "#,##0.00")')).toBe('1,234.50');
    expect(await evalExpr('"abc" Like "a*"')).toBe('True');
    expect(await evalExpr('Round(2.5) & Round(3.5)')).toBe('24');
    expect(await evalExpr('CInt("12") + CDbl("0.5")')).toBe('12.5');
    expect(await evalExpr('Not True Or 5 > 3 And "b" > "a"')).toBe('True');
    expect(await evalExpr('IIf(1 > 2, "yes", "no")')).toBe('no');
  });

  it('runs loops, ifs and select case', async () => {
    const out = await run(`Sub Main()
  Dim i As Long, total As Long
  For i = 1 To 10
    If i Mod 2 = 0 Then total = total + i
  Next i
  Debug.Print total
  Dim n: n = 0
  Do While n < 5
    n = n + 2
  Loop
  Debug.Print n
  Do
    n = n - 1
  Loop Until n <= 0
  Debug.Print n
  For i = 10 To 1 Step -3: Debug.Print i;: Next
  Select Case 15
    Case Is < 10: Debug.Print "small"
    Case 10 To 20: Debug.Print "medium"
    Case Else: Debug.Print "large"
  End Select
  Dim s As String
  For Each x In Array("a", "b", "c")
    s = s & x
    If x = "b" Then Exit For
  Next
  Debug.Print s
  While n < 3
    n = n + 1
  Wend
  Debug.Print n
End Sub`);
    expect(out).toEqual(['30', '6', '0', '10', '7', '4', '1', 'medium', 'ab', '3']);
  });

  it('calls subs and functions with ByRef and ByVal', async () => {
    const out = await run(`Function Square(x)
  Square = x * x
End Function
Sub AddOne(n)
  n = n + 1
End Sub
Sub AddOneVal(ByVal n)
  n = n + 1
End Sub
Function Fact(n As Long) As Long
  If n <= 1 Then
    Fact = 1
  Else
    Fact = n * Fact(n - 1)
  End If
End Function
Function Greet(Optional name As String = "you") As String
  Greet = "Hi " & name
End Function
Sub Main()
  Dim a: a = 4
  AddOne a
  Debug.Print a
  AddOneVal a
  Debug.Print a
  Call AddOne(a)
  Debug.Print a
  Debug.Print Square(3), Fact(5)
  Debug.Print Greet(), Greet("Sam"), Greet(name:="Ana")
End Sub`);
    expect(out).toEqual(['5', '5', '6', '9 120', 'Hi you Hi Sam Hi Ana']);
  });

  it('handles arrays, ReDim Preserve and Split/Join', async () => {
    const out = await run(`Sub Main()
  Dim a(1 To 3) As Long
  a(1) = 10: a(2) = 20: a(3) = 30
  Debug.Print a(2), LBound(a), UBound(a)
  Dim b() As String
  ReDim b(0 To 1)
  b(0) = "x": b(1) = "y"
  ReDim Preserve b(0 To 2)
  b(2) = "z"
  Debug.Print Join(b, ",")
  Dim parts
  parts = Split("1;2;3", ";")
  Debug.Print UBound(parts), parts(2)
  Dim grid(2, 2)
  grid(1, 2) = "mid"
  Debug.Print grid(1, 2)
End Sub`);
    expect(out).toEqual(['20 1 3', 'x,y,z', '2 3', 'mid']);
  });

  it('supports Collection and Dictionary', async () => {
    const out = await run(`Sub Main()
  Dim c As New Collection
  c.Add "apple", "a"
  c.Add "banana"
  Debug.Print c.Count, c(1), c("a"), c.Item(2)
  Dim d As Object
  Set d = CreateObject("Scripting.Dictionary")
  d("x") = 1
  d.Add "y", 2
  d("x") = d("x") + 10
  Debug.Print d.Count, d("x"), d.Exists("y"), d.Exists("z")
  Dim k
  For Each k In d.Keys
    Debug.Print k
  Next
End Sub`);
    expect(out).toEqual(['2 apple apple banana', '2 11 True False', 'x', 'y']);
  });

  it('handles errors with On Error', async () => {
    const out = await run(`Sub Main()
  On Error Resume Next
  Dim x
  x = 1 / 0
  Debug.Print Err.Number
  Err.Clear
  On Error GoTo Handler
  x = 1 / 0
  Debug.Print "skipped"
Done:
  Debug.Print "done"
  Exit Sub
Handler:
  Debug.Print "caught " & Err.Number
  Resume Done
End Sub`);
    expect(out).toEqual(['11', 'caught 11', 'done']);
  });

  it('reports runtime errors with the line', async () => {
    const program = new VbaProgram([{ name: 'Module1', code: 'Sub Main()\n  Dim a(2)\n  a(5) = 1\nEnd Sub', kind: 'standard' }]);
    await expect(program.run(stubHost([]), 'Main')).rejects.toMatchObject({ number: 9, line: 3 });
  });

  it('enforces Option Explicit', async () => {
    const program = new VbaProgram([{ name: 'Module1', code: 'Option Explicit\nSub Main()\n  y = 1\nEnd Sub', kind: 'standard' }]);
    await expect(program.run(stubHost([]), 'Main')).rejects.toBeInstanceOf(VbError);
  });

  it('uses module-level variables and constants across modules', async () => {
    const out: string[] = [];
    const program = new VbaProgram([
      { name: 'Module1', code: 'Public counter As Long\nPublic Const NAME_ = "Gridline"\nSub Bump()\n  counter = counter + 1\nEnd Sub', kind: 'standard' },
      { name: 'Module2', code: 'Sub Main()\n  Bump\n  Bump\n  Module1.Bump\n  Debug.Print counter, NAME_\nEnd Sub', kind: 'standard' },
    ]);
    await program.run(stubHost(out), 'Main');
    expect(out).toEqual(['3 Gridline']);
  });

  it('asks the user with MsgBox and InputBox', async () => {
    const out = await run(
      `Sub Main()
  Dim r
  r = MsgBox("Continue?", vbYesNo)
  If r = vbYes Then
    Dim name As String
    name = InputBox("Your name?")
    MsgBox "Hello " & name
  End If
End Sub`,
      'Main',
      [6, 'Pat'],
    );
    expect(out).toEqual(['msg:Continue?', 'input:Your name?', 'msg:Hello Pat']);
  });

  it('runs With blocks on objects and class modules', async () => {
    const out: string[] = [];
    const program = new VbaProgram([
      {
        name: 'Counter',
        kind: 'class',
        code: `Private count_ As Long
Public Label As String
Private Sub Class_Initialize()
  count_ = 100
End Sub
Public Sub Add(n As Long)
  count_ = count_ + n
End Sub
Public Property Get Total() As Long
  Total = count_
End Property`,
      },
      {
        name: 'Module1',
        kind: 'standard',
        code: `Sub Main()
  Dim c As New Counter
  With c
    .Add 5
    .Add 7
    .Label = "tally"
  End With
  Debug.Print c.Total, c.Label
End Sub`,
      },
    ]);
    await program.run(stubHost(out), 'Main');
    expect(out).toEqual(['112 tally']);
  });

  it('stops when asked', async () => {
    const program = new VbaProgram([{ name: 'Module1', code: 'Sub Main()\n  Do\n  Loop\nEnd Sub', kind: 'standard' }]);
    const host = stubHost([]);
    let ticks = 0;
    host.tick = async () => {
      if (++ticks === 3) program.stop();
    };
    await expect(program.run(host, 'Main')).rejects.toMatchObject({ reason: 'cancel' });
  });

  it('lists runnable macros', () => {
    const program = new VbaProgram([{ name: 'Module1', code: 'Sub A()\nEnd Sub\nPrivate Sub B()\nEnd Sub\nFunction C()\nEnd Function\nSub D(x)\nEnd Sub', kind: 'standard' }]);
    expect(program.macros().filter((m) => m.runnable).map((m) => m.name)).toEqual(['A']);
  });

  it('formats dates', async () => {
    expect(await evalExpr('Year(DateSerial(2026, 3, 14)) & "-" & Month(#3/14/2026#)')).toBe('2026-3');
    expect(await evalExpr('Format(DateSerial(2026, 3, 4), "yyyy-mm-dd")')).toBe('2026-03-04');
    expect(await evalExpr('DateDiff("d", #1/1/2026#, #3/1/2026#)')).toBe('59');
    expect(await evalExpr('Format(DateAdd("m", 1, #1/31/2026#), "m/d/yyyy")')).toBe('2/28/2026');
  });

  it('converts values to text like VBA', () => {
    expect(toStr(0.1 + 0.2)).toBe('0.3');
    expect(toStr(true)).toBe('True');
    const arr = new VbArray([{ lb: 0, ub: 1 }]);
    expect(arr.data.length).toBe(2);
    const v: Value = 1e21;
    expect(toStr(v)).toBe('1E+21');
  });
});
