import { describe, expect, it } from 'vitest';
import type { Doc } from '../src/core/doc';
import { userFunctions } from '../src/core/macros/cellFunctions';
import { MacroProject } from '../src/core/macros/project';
import { addSheet } from '../src/core/ops/sheets';
import { applyStyle } from '../src/core/ops/cells';
import { parseRange } from '../src/core/address';
import { makeDoc, set, val } from './helpers';

// VBA Functions of the kinds people write in Excel: looking values up on a settings sheet, reading a
// cell's number format, optional parameters with defaults, error values, loops over used ranges.
const CODE = `
Option Explicit

Function SETTING(settingName As String) As Variant
    Application.Volatile True
    Dim ws As Worksheet
    Dim rng As Range
    Set ws = ThisWorkbook.Sheets("Settings")
    On Error Resume Next
    Set rng = ws.Range("A:A").Find(What:=settingName, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0
    If Not rng Is Nothing Then
        SETTING = rng.Offset(0, 1).Value
    Else
        SETTING = CVErr(xlErrNA)
    End If
End Function

Function KIND_OF(cellRef As Variant) As String
    Dim fmt As String
    If TypeName(cellRef) = "Range" Then
        fmt = LCase(cellRef.NumberFormat) & LCase(CStr(cellRef.Text))
    Else
        fmt = LCase(CStr(cellRef))
    End If
    If InStr(fmt, "$") > 0 Then
        KIND_OF = "money"
    ElseIf InStr(fmt, "%") > 0 Then
        KIND_OF = "percent"
    Else
        KIND_OF = ""
    End If
End Function

Function FUZZY(key As String, _
               Optional sourceSheet As String = "Settings") As Variant
    Application.Volatile True
    Dim ws As Worksheet, c As Range, found As Range
    Dim k As String, label As String
    On Error GoTo SafeExit
    Set ws = ThisWorkbook.Sheets(sourceSheet)
    k = UCase$(Trim$(key))
    If Right$(k, 1) = "S" Then k = Left$(k, Len(k) - 1)
    For Each c In Intersect(ws.UsedRange, ws.Columns(1))
        If Len(c.Value) > 0 Then
            label = Trim$(Replace(UCase$(CStr(c.Value)), " BUDGET", ""))
            If InStr(1, label, k, vbTextCompare) > 0 Then
                Set found = c.Offset(0, 1)
                Exit For
            End If
        End If
    Next c
    If found Is Nothing Then
        FUZZY = ""
        Exit Function
    End If
    FUZZY = found.Value
    Exit Function
SafeExit:
    FUZZY = "error"
End Function

Public Function LINEAR(ByVal x As Double, _
    Optional ByVal m As Double = 0.5, _
    Optional ByVal b As Double = -10) As Double
    LINEAR = m * x + b
    If LINEAR < 0 Then LINEAR = 0
End Function

Function TWICE(values As Variant) As Variant
    Dim out() As Variant, r As Long
    ReDim out(1 To values.Rows.Count, 1 To 1)
    For r = 1 To values.Rows.Count
        out(r, 1) = values.Cells(r, 1).Value * 2
    Next r
    TWICE = out
End Function

Function WHERE_AM_I() As String
    WHERE_AM_I = Application.Caller.Address(False, False)
End Function

Function TRIES_TO_WRITE() As Variant
    Range("Z1").Value = 1
    TRIES_TO_WRITE = "done"
End Function

Function TRIES_TO_ASK() As Variant
    MsgBox "hello"
    TRIES_TO_ASK = "done"
End Function

Private Function HIDDEN_ONE() As Long
    HIDDEN_ONE = 1
End Function
`;

const R = (a: string) => parseRange(a)!;

function workbook(): Doc {
  const doc = makeDoc();
  const settings = doc.transact('add', (tx) => addSheet(tx, doc.wb, 'Settings', 1)).wb.sheetByName('Settings')!;
  set(doc, { A1: 'Rent Budget', B1: '1500', A2: 'Savings rate', B2: '15%', A3: 'Groceries', B3: '$420', A4: 'Note', B4: 'hello' }, settings);
  doc.wb.activeSheetId = doc.wb.sheets[0].id;
  const project = (doc.wb.macros ??= new MacroProject());
  project.addModule('standard', 'Module1', CODE);
  doc.setUserFunctions(userFunctions(doc));
  return doc;
}

describe('VBA Functions in formulas', () => {
  it('looks values up with Find, Offset and Application.Volatile', () => {
    const doc = workbook();
    set(doc, { A1: '=SETTING("Groceries")', A2: '=SETTING("Nope")', A3: '=SETTING("Rent Budget")*12' });
    expect(val(doc, 'A1')).toBe(420);
    expect(val(doc, 'A2')).toBe('#N/A');
    expect(val(doc, 'A3')).toBe(18000);
    // Volatile: changing the settings sheet updates the result, though no formula refers to it
    const settings = doc.wb.sheetByName('Settings')!;
    set(doc, { B3: '$500' }, settings);
    expect(val(doc, 'A1')).toBe(500);
  });

  it('receives a Range for Variant parameters (number formats, text)', () => {
    const doc = workbook();
    set(doc, { B1: '$12.50', B2: '7%', B3: '42' });
    set(doc, { C1: '=KIND_OF(B1)', C2: '=KIND_OF(B2)', C3: '=KIND_OF(B3)', C4: '=KIND_OF("5%")' });
    expect([val(doc, 'C1'), val(doc, 'C2'), val(doc, 'C3'), val(doc, 'C4')]).toEqual(['money', 'percent', '', 'percent']);
    // Not volatile, but recalculated when its argument changes, including its format
    doc.transact('fmt', (tx) => applyStyle(tx, doc.wb, doc.wb.activeSheet, [R('B3')], { numFmt: '"$"#,##0' }));
    set(doc, { B3: '43' });
    expect(val(doc, 'C3')).toBe('money');
  });

  it('runs loops, string functions, optional parameters and On Error GoTo', () => {
    const doc = workbook();
    set(doc, { A1: '=FUZZY("rents")', A2: '=FUZZY("savings")', A3: '=FUZZY("nothing")', A4: '=FUZZY("rent", "Missing sheet")' });
    expect(val(doc, 'A1')).toBe(1500);
    expect(val(doc, 'A2')).toBe(0.15);
    expect(val(doc, 'A3')).toBe('');
    expect(val(doc, 'A4')).toBe('error');
  });

  it('converts typed parameters and uses defaults for missing ones', () => {
    const doc = workbook();
    set(doc, { B1: '100', A1: '=LINEAR(B1)', A2: '=LINEAR(B1, 2)', A3: '=LINEAR(10)', A4: '=LINEAR("abc")', A5: '=LINEAR(B1:B2)' });
    expect([val(doc, 'A1'), val(doc, 'A2'), val(doc, 'A3')]).toEqual([40, 190, 0]);
    expect(val(doc, 'A4')).toBe('#VALUE!');
    expect(val(doc, 'A5')).toBe('#VALUE!');
  });

  it('returns arrays that spill, and knows its own cell', () => {
    const doc = workbook();
    set(doc, { A1: '1', A2: '2', A3: '3', C1: '=TWICE(A1:A3)', E5: '=WHERE_AM_I()' });
    expect([val(doc, 'C1'), val(doc, 'C2'), val(doc, 'C3')]).toEqual([2, 4, 6]);
    expect(val(doc, 'E5')).toBe('E5');
  });

  it("can't change the workbook or show messages (like Excel), and private functions stay private", () => {
    const doc = workbook();
    set(doc, { A1: '=TRIES_TO_WRITE()', A2: '=TRIES_TO_ASK()', A3: '=HIDDEN_ONE()' });
    expect(val(doc, 'A1')).toBe('#VALUE!');
    expect(val(doc, 'Z1')).toBe(null);
    expect(val(doc, 'A2')).toBe('#VALUE!');
    expect(val(doc, 'A3')).toBe('#NAME?');
  });

  it('only runs while macros are allowed', () => {
    const doc = workbook();
    set(doc, { A1: '=LINEAR(100)' });
    expect(val(doc, 'A1')).toBe(40);
    doc.setUserFunctions(null);
    expect(val(doc, 'A1')).toBe('#NAME?');
    const e = doc.engine.getValue(doc.wb.activeSheet, 0, 0) as { message?: string };
    expect(e.message).toContain("LINEAR is a function in this workbook's macros");
    // Edited code takes effect after a pause in typing; switching on again recalculates now
    doc.wb.macros!.setCode('Module1', CODE.replace('LINEAR = m * x + b', 'LINEAR = m * x + b + 1'));
    doc.setUserFunctions(userFunctions(doc));
    expect(val(doc, 'A1')).toBe(41);
  });
});
