# Macros in Gridline

A **macro** is a list of instructions that Gridline carries out for you: clean up a table, color the
rows that need attention, add today's date to a log, build a monthly report sheet… anything you would
otherwise do by hand, step by step.

Gridline runs **VBA macros**, the same kind Excel uses, and opens and saves macro-enabled **.xlsm**
files. Macros you write in Gridline work in Excel, and most Excel macros work in Gridline.

- [Running a macro](#running-a-macro)
- [Writing a macro (VBA)](#writing-a-macro-vba)
- [Step macros (no code, Beta)](#step-macros-no-code-beta)
- [Buttons, shortcuts and automatic macros](#buttons-shortcuts-and-automatic-macros)
- [Files: .xlsm and .xlsx](#files-xlsm-and-xlsx)
- [Safety](#safety)
- [VBA reference: what's supported](#vba-reference-whats-supported)

---

## Running a macro

- **Automate → Run macro** lists every macro in the workbook. Pick one to run it.
- **Alt+F8** opens the same list.
- A button on the sheet or a keyboard shortcut can run a macro too (see below).

While a macro runs, a small **Running… Stop** bar appears at the bottom of the window. Press **Esc** or
click **Stop** to stop it.

**Undo:** everything a macro changes is a single undo step. Press **Ctrl+Z** once to undo the whole run.

## Writing a macro (VBA)

1. Click **Automate → New macro** (or **Macro editor**, **Alt+F11**).
2. The macro editor opens under the grid, so you can see your sheet while you work. A new module has
   a small example to start from:

   ```vb
   Sub HelloWorld()
       Range("A1").Value = "Hello from Gridline!"
       MsgBox "Done!"
   End Sub
   ```

3. Press **F5** (or click **Run**) to run the macro your cursor is in.

**Tips:**
- Typing `Sub Name()` and pressing Enter adds the matching `End Sub`. Tab and Shift+Tab indent and outdent.
- Mistakes are underlined in red, and the message appears in the output area under the code. Click a
  message to jump to the line.
- `Debug.Print` writes to the output area, which is handy for checking values.
- If a macro stops with an error, the message says which module and line. **Show code** jumps there.

**Modules:**
- **Code modules** (Module1, Module2…) hold macros you run yourself.
- **This workbook** and one module **per sheet** hold code that runs on its own when something happens
  (see [Automatic macros](#automatic-macros)).
- **Class modules** let you define your own objects (`Dim c As New MyClass`).
- Right-click a module to rename or delete it.

### A few useful examples

```vb
' Add up column B, from row 2 down to the last row with data
Sub TotalColumnB()
    Dim lastRow As Long, total As Double, r As Long
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    For r = 2 To lastRow
        total = total + Cells(r, "B").Value
    Next r
    MsgBox "The total is " & Format(total, "$#,##0.00")
End Sub

' Color every overdue row red
Sub MarkOverdue()
    Dim r As Long
    For r = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(r, 3).Value < Date And Cells(r, 4).Value <> "Done" Then
            Range("A" & r & ":D" & r).Interior.Color = RGB(255, 205, 205)
        End If
    Next r
End Sub

' Ask for a name and add a dated row to a log
Sub LogEntry()
    Dim who As String, nextRow As Long
    who = InputBox("Who is this entry for?")
    If who = "" Then Exit Sub
    nextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    Cells(nextRow, 1).Value = Date
    Cells(nextRow, 2).Value = who
End Sub
```

## Your own functions in cells

A `Function` in a module works in formulas too, just like in Excel:

```vb
Function TAX(gross As Double, Optional rate As Double = 0.25) As Double
    TAX = gross * rate
End Function
```

Then `=TAX(B2)` or `=TAX(B2, 0.3)` in any cell. Gridline uses the same rules as Excel:

- **Which functions:** `Public` (or plain) Functions in standard modules (Module1…). Private functions, and functions in sheet or ThisWorkbook modules, can't be used in cells.
- **Arguments:** a cell reference passed to a `Variant` or `Range` parameter arrives as a Range, so the code can read `.Value`, `.NumberFormat`, `.Text` or loop over the cells. A typed parameter (`As String`, `As Double`…) gets the cell's value.
- **Results:** a function can return a number, text, True/False, a date, an error (`CVErr(xlErrNA)` shows `#N/A`) or an array, which spills into the cells next to it.
- **Recalculating:** a function recalculates when its arguments change. If it reads other cells itself (like a lookup on a settings sheet), add `Application.Volatile`, and it recalculates after every change.
- **Read-only:** like in Excel, a function used in a cell can read the whole workbook but can't change cells, sheets or formatting, and it can't show message boxes. `Application.Caller` is the cell it's calculating.
- **Security:** these functions only run while the workbook's macros are on. Until you click **Enable macros**, their cells show the result last saved in the file.

## Step macros (no code, Beta)

Most people don't write code, so Gridline can also build macros from **steps**, like the Shortcuts app
on a phone. Each step reads like a sentence, and you fill in the blanks:

> **Color** `B2:B50` **that** `is more than` `100` **in** `Yellow`

Step macros are **in development**. They're available when **Settings → Updates → Developer branch**
is on, and in any workbook that already has step macros.

1. Click **Automate → Step macro**, or pick one of the **Examples**.
2. Click **Add step** and choose what should happen. Steps are grouped into Cells & data, Formatting,
   Sheets, Logic & loops, Messages & questions, and Workbook.
3. Fill in the blanks. For cells, click the **⬚** button next to a box to use the cells you've
   selected on the sheet.
4. Click **Run**.

**Blocks:** *If*, *Repeat* and *For each row* hold other steps. Inside *For each row*, write `{Row}`
for the current row number. For example, `A{Row}:D{Row}` is "this row, columns A to D".

**Remembering values:** *Ask a question* and *Remember a value* store an answer under a name. Use it
later in any box as `{Name}`. `{Today}` and `{Now}` are always available.

**Show code** displays the VBA the steps turn into, next to the steps. It's a good way to learn VBA,
and the code works in Excel too.

### Examples included

| Example | What it does |
| --- | --- |
| Log today | Adds a row with today's date below the data in column A. |
| Highlight big expenses | Colors every amount over 100 in column C yellow. |
| Clear the input cells | Asks first, then empties B2:B20 for next month. |
| Mark paid rows | Goes through each row and colors it green when column D says "Paid". |
| Monthly report sheet | Asks for a month, copies the Template sheet and names it after the month. |

## Buttons, shortcuts and automatic macros

### Buttons on the sheet
**Automate → Button** puts a button next to the selected cell. Click it to run the macro.
- **Drag** a button to move it, or drag its bottom-right corner to resize it.
- **Right-click** it to change the text, the macro, the color, or to delete it.

### Keyboard shortcuts
**Automate → Shortcut** lets you press a key combination for a macro, such as **Ctrl+Shift+M**.
In this workbook, the shortcut runs the macro instead of any built-in command. Ctrl+letter shortcuts
are also saved in a way Excel understands.

### Automatic macros
**Automate → Automatically** runs a macro by itself:

- when the workbook opens,
- when cells change (on any sheet, or only on one sheet or in a range like `B2:B50`),
- when a sheet is shown,
- before the workbook is saved.

VBA code can do the same with event procedures. Put them in **This workbook** or a **sheet module**.
The macro editor's *add an event…* link writes the first line for you:

| Event | Where | Runs when |
| --- | --- | --- |
| `Workbook_Open` | This workbook | the file is opened (and macros are enabled) |
| `Workbook_BeforeSave(SaveAsUI, Cancel)` | This workbook | before saving; set `Cancel = True` to stop the save |
| `Workbook_SheetChange(Sh, Target)` | This workbook | cells change on any sheet |
| `Workbook_SheetActivate(Sh)` | This workbook | another sheet is shown |
| `Worksheet_Change(Target)` | a sheet | cells on that sheet change (`Target` is the changed range) |
| `Worksheet_Activate()` | a sheet | that sheet is shown |
| `Worksheet_SelectionChange(Target)` | a sheet | the selection moves on that sheet |

Changes made **by a macro** don't start change events, so a `Worksheet_Change` macro can safely write
to the sheet. You can turn off all automatic macros in **Settings → Macros**.

## Files: .xlsm and .xlsx

- **.xlsm** files hold VBA code. Gridline opens them and saves the code back. Code that you didn't change
  is saved exactly as it was, and code you changed is recompiled by Excel when it opens the file.
- When a workbook has VBA code and you save it as `.xlsx`, Gridline offers to save an `.xlsm` instead.
  A regular `.xlsx` can't hold VBA code.
- **Step macros, buttons, shortcuts and automatic macros** are stored inside the workbook (in a
  part Excel keeps but ignores), in both `.xlsx` and `.xlsm` files.
- UserForms in a file are kept when you save as `.xlsm`, but Gridline can't show them yet.

## Safety

Macros are programs, and a macro from a stranger can make a mess of a workbook. Gridline protects you
in several ways:

- When you open a file with macros, they're **off** until you click **Enable macros** in the bar at
  the top. **Always trust this file** skips the question next time for that file.
- **Settings → Macros** can ask every time (the default), always run macros, or never run them.
- Macros in Gridline **can only work with the workbook**. They can't open other files or programs,
  delete files, run commands, send keystrokes or reach the internet. Things like `Shell`, `Kill`,
  `CreateObject("WScript.Shell")` and file I/O stop with a clear message.

## VBA reference: what's supported

### The language
- `Sub`, `Function`, `Property Get/Let/Set`, `Optional` and `ParamArray` arguments, `ByRef`/`ByVal`
- `Dim`, `Static`, `Const`, `ReDim [Preserve]`, `Option Explicit`, `Option Base`, `Option Compare Text`, `Enum`
- `If … Then … ElseIf … Else … End If` (and single-line If), `Select Case` (with `Is`, `To`, lists)
- `For … Next` (with `Step`), `For Each … Next`, `Do … Loop` (`While`/`Until`, before or after), `While … Wend`
- `With … End With`, `Exit Sub/Function/For/Do`, `GoTo`, `End`
- `On Error Resume Next`, `On Error GoTo label`, `On Error GoTo 0`, `Resume`, `Resume Next`, `Err.Number`, `Err.Description`, `Err.Raise`
- Arrays (any number of dimensions, `LBound`/`UBound`), `Collection`, `Scripting.Dictionary`, class modules
- Line continuation (` _`), `:` between statements, `[A1]` shorthand, date literals `#3/14/2026#`

### Built-in functions
- **Text:** Len, Left, Right, Mid, InStr, InStrRev, Replace, Split, Join, Trim, LCase, UCase, StrComp, StrConv, Format, FormatNumber, FormatCurrency, FormatPercent, FormatDateTime, Space, String, Chr, Asc, Like
- **Numbers:** Abs, Int, Fix, Round, Sqr, Exp, Log, Sin, Cos, Tan, Atn, Sgn, Rnd, Val, Hex, Oct, Mod
- **Conversion:** CStr, CInt, CLng, CDbl, CBool, CDate, CVErr
- **Dates:** Now, Date, Time, Timer, Year, Month, Day, Hour, Minute, Second, Weekday, WeekdayName, MonthName, DateSerial, TimeSerial, DateValue, TimeValue, DateAdd, DateDiff, DatePart
- **Checks:** IsNumeric, IsDate, IsEmpty, IsNull, IsArray, IsObject, IsMissing, IsError, TypeName, VarType
- **Other:** Array, Filter, IIf, Choose, Switch, RGB, QBColor, MsgBox, InputBox, DoEvents
- All `vb…` and the common `xl…` constants (`vbYesNo`, `vbCrLf`, `xlUp`, `xlCenter`, `xlContinuous`…)

### The Excel objects
| Object | Supported |
| --- | --- |
| **Application** | ActiveSheet, ActiveCell, Selection, ActiveWorkbook, ThisWorkbook, Worksheets, Sheets, Range, Cells, Rows, Columns, WorksheetFunction, Union, Intersect, Evaluate, Run, InputBox (with `Type`), Wait, Calculate, ScreenUpdating, StatusBar, DisplayAlerts, EnableEvents, CutCopyMode, Goto |
| **Workbook** | Name, FullName, Path, Sheets, Worksheets, ActiveSheet, Names, Save, Saved |
| **Worksheets / Sheets** | Count, Item, Add (Before/After/Count), For Each |
| **Worksheet** | Name, CodeName, Index, Range, Cells, Rows, Columns, UsedRange, Activate, Select, Delete, Copy and Move (Before/After), Visible, Tab.Color, Next, Previous, Paste, Sort (SortFields, SetRange, Header, Apply), AutoFilterMode, ShowAllData, Hyperlinks.Add, Evaluate |
| **Range** | Value, Value2, Formula, FormulaR1C1, Text, NumberFormat, Address, Row, Column, Rows, Columns, Count, Cells, Item, Offset, Resize, End, CurrentRegion, EntireRow, EntireColumn, Areas, MergeArea, Interior, Font, Borders, BorderAround, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, ColumnWidth, RowHeight, Hidden, MergeCells, Merge, UnMerge, Clear, ClearContents, ClearFormats, ClearComments, Delete, Insert, Copy, Cut, PasteSpecial, Select, Activate, AutoFit, Sort, Find, FindNext, FindPrevious, Replace, FillDown, FillUp, FillRight, FillLeft, AutoFill, SpecialCells, RemoveDuplicates, AutoFilter, AddComment, Comment, NoteText, Hyperlinks, Name |
| **Interior** | Color, ColorIndex, Pattern |
| **Font** | Bold, Italic, Underline, Strikethrough, Size, Name, Color, ColorIndex |
| **Borders / Border** | LineStyle, Weight, Color, ColorIndex; `Borders(xlEdgeBottom)` and the other edges |
| **WorksheetFunction** | Every function Gridline has (Sum, VLookup, Match, Index, CountIf, SumIfs, …) |
| **Names** | Add, Item, Count, RefersTo, RefersToRange, Delete |
| **ActiveWindow** | FreezePanes, SplitRow, SplitColumn, DisplayGridlines, Zoom |

### Not supported (yet)
- UserForms and ActiveX controls (they're kept in .xlsm files, but not shown)
- Charts, shapes and pivot tables from code
- Opening, creating or closing other workbooks; reading and writing files; other programs
- `Application.OnTime` (use **Automate → Automatically** instead)
- `Declare` (Windows API calls) and `Type` blocks are skipped

If a macro uses something that isn't supported, it stops with a message that names it, for example
"Gridline doesn't support Range.PivotTable yet."
