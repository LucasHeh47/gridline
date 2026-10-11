// No-code macros: a list of plain-English steps (like the Shortcuts app) that compiles to VBA.
// Text fields can use {Variables}: {Row} inside "For each row", {Count} inside "Repeat", answers
// saved by "Ask a question", plus {Today} and {Now}.
import { toIdentifier } from './project';

export interface Step {
  id: string;
  type: string;
  args: Record<string, string>;
  /** Steps inside a block (If, Repeat, For each row). */
  body?: Step[];
  /** The "Otherwise" part of an If. */
  elseBody?: Step[];
  /** Skipped when running (but kept). */
  off?: boolean;
}

export type FieldKind = 'range' | 'cell' | 'text' | 'number' | 'sheet' | 'color' | 'choice' | 'macro' | 'variable' | 'column';

export interface StepField {
  key: string;
  label: string;
  kind: FieldKind;
  /** For 'choice'. */
  options?: [value: string, label: string][];
  dflt?: string;
  placeholder?: string;
  /** Can be left empty. */
  optional?: boolean;
}

export type StepCategory = 'Cells' | 'Format' | 'Sheets' | 'Logic' | 'Talk' | 'Workbook';

export interface StepDef {
  type: string;
  category: StepCategory;
  title: string;
  /** Sentence shown in the list; {key} is replaced by the field value. */
  sentence: string;
  icon: string;
  fields: StepField[];
  block?: boolean;
  hasElse?: boolean;
  /** Variable made available to the steps inside (shown as a chip). */
  provides?: string;
  compile(a: Record<string, string>, c: Compiler, step: Step): string[];
}

// ---------- shared field definitions ----------

const COMPARE: [string, string][] = [
  ['=', 'is'],
  ['<>', 'is not'],
  ['>', 'is more than'],
  ['<', 'is less than'],
  ['>=', 'is at least'],
  ['<=', 'is at most'],
  ['contains', 'contains'],
  ['empty', 'is empty'],
  ['notempty', 'is not empty'],
];

const COLORS: [string, string][] = [
  ['#fef08a', 'Yellow'],
  ['#bbf7d0', 'Green'],
  ['#fecaca', 'Red'],
  ['#bfdbfe', 'Blue'],
  ['#fed7aa', 'Orange'],
  ['#e9d5ff', 'Purple'],
  ['#e5e7eb', 'Gray'],
  ['none', 'No color'],
];

const range = (key = 'range', label = 'Cells', dflt = 'A1'): StepField => ({ key, label, kind: 'range', dflt, placeholder: 'A1 or A1:C10' });

// ---------- the step catalog ----------

export const STEPS: StepDef[] = [
  // Cells
  {
    type: 'setValue',
    category: 'Cells',
    title: 'Set cells',
    sentence: 'Set {range} to {value}',
    icon: 'pencil',
    fields: [range(), { key: 'value', label: 'Value', kind: 'text', placeholder: 'Text, a number, or a formula like =SUM(A1:A5)' }],
    compile: (a, c) => {
      const v = c.value(a.value);
      return [`${c.range(a.range)}.${v.formula ? 'Formula' : 'Value'} = ${v.expr}`];
    },
  },
  {
    type: 'stamp',
    category: 'Cells',
    title: 'Put the date or time',
    sentence: 'Put the current {what} in {range}',
    icon: 'calendar-clock',
    fields: [
      { key: 'what', label: 'What', kind: 'choice', options: [['date', 'date'], ['time', 'time'], ['both', 'date and time']], dflt: 'date' },
      range('range', 'Cell'),
    ],
    compile: (a, c) => [`${c.range(a.range)}.Value = ${a.what === 'time' ? 'Time' : a.what === 'both' ? 'Now' : 'Date'}`],
  },
  {
    type: 'addTo',
    category: 'Cells',
    title: 'Add to a number',
    sentence: 'Add {amount} to {range}',
    icon: 'plus',
    fields: [{ key: 'amount', label: 'Amount', kind: 'number', dflt: '1' }, range('range', 'Cell')],
    compile: (a, c) => {
      const r = c.range(a.range);
      return [`${r}.Value = Val(${r}.Value) + ${c.num(a.amount)}`];
    },
  },
  {
    type: 'appendRow',
    category: 'Cells',
    title: 'Add a row at the bottom',
    sentence: 'Add a row below the data in column {column} with {values}',
    icon: 'list-plus',
    fields: [
      { key: 'column', label: 'First column of the table', kind: 'column', dflt: 'A' },
      { key: 'values', label: 'Values (separate with |)', kind: 'text', placeholder: '{Today} | Coffee | 4.50' },
    ],
    compile: (a, c) => {
      const col = c.column(a.column);
      const row = c.temp('nextRow');
      const out = [`${row} = ${c.sheetPrefix(a.column)}Cells(Rows.Count, "${col}").End(xlUp).Row + 1`, `If IsEmpty(${c.sheetPrefix(a.column)}Cells(1, "${col}").Value) Then ${row} = 1`];
      (a.values ?? '').split('|').forEach((part, i) => {
        const v = c.value(part.trim());
        out.push(`${c.sheetPrefix(a.column)}Cells(${row}, "${col}").Offset(0, ${i}).${v.formula ? 'Formula' : 'Value'} = ${v.expr}`);
      });
      return out;
    },
  },
  {
    type: 'copy',
    category: 'Cells',
    title: 'Copy cells',
    sentence: 'Copy {range} to {dest} ({how})',
    icon: 'copy',
    fields: [range('range', 'From', 'A1:A10'), range('dest', 'To', 'C1'), { key: 'how', label: 'Copy', kind: 'choice', options: [['all', 'everything'], ['values', 'values only'], ['formats', 'formatting only']], dflt: 'all' }],
    compile: (a, c) => {
      if (a.how === 'values') return [`${c.range(a.range)}.Copy`, `${c.range(a.dest)}.PasteSpecial xlPasteValues`];
      if (a.how === 'formats') return [`${c.range(a.range)}.Copy`, `${c.range(a.dest)}.PasteSpecial xlPasteFormats`];
      return [`${c.range(a.range)}.Copy Destination:=${c.range(a.dest)}`];
    },
  },
  {
    type: 'clear',
    category: 'Cells',
    title: 'Clear cells',
    sentence: 'Clear {what} in {range}',
    icon: 'eraser',
    fields: [{ key: 'what', label: 'Clear', kind: 'choice', options: [['contents', 'the contents'], ['formats', 'the formatting'], ['all', 'everything']], dflt: 'contents' }, range('range', 'Cells', 'A1:C10')],
    compile: (a, c) => [`${c.range(a.range)}.${a.what === 'formats' ? 'ClearFormats' : a.what === 'all' ? 'Clear' : 'ClearContents'}`],
  },
  {
    type: 'replace',
    category: 'Cells',
    title: 'Find and replace',
    sentence: 'Replace {find} with {with} in {range}',
    icon: 'replace',
    fields: [{ key: 'find', label: 'Find', kind: 'text' }, { key: 'with', label: 'Replace with', kind: 'text', optional: true }, range('range', 'In', 'A:Z')],
    compile: (a, c) => [`${c.range(a.range)}.Replace What:=${c.text(a.find)}, Replacement:=${c.text(a.with ?? '')}, LookAt:=xlPart, MatchCase:=False`],
  },
  {
    type: 'sort',
    category: 'Cells',
    title: 'Sort',
    sentence: 'Sort {range} by column {column}, {order}',
    icon: 'arrow-down-up',
    fields: [
      range('range', 'Table', 'A1:D100'),
      { key: 'column', label: 'By column', kind: 'column', dflt: 'A' },
      { key: 'order', label: 'Order', kind: 'choice', options: [['asc', 'smallest first (A→Z)'], ['desc', 'largest first (Z→A)']], dflt: 'asc' },
      { key: 'header', label: 'First row is a header', kind: 'choice', options: [['yes', 'Yes'], ['no', 'No']], dflt: 'yes' },
    ],
    compile: (a, c) => [
      `${c.range(a.range)}.Sort Key1:=${c.range(a.range)}.Columns(${c.columnIndexIn(a.column, a.range)}), Order1:=${a.order === 'desc' ? 'xlDescending' : 'xlAscending'}, Header:=${a.header === 'no' ? 'xlNo' : 'xlYes'}`,
    ],
  },
  {
    type: 'dedupe',
    category: 'Cells',
    title: 'Remove duplicate rows',
    sentence: 'Remove duplicate rows in {range}',
    icon: 'copy-minus',
    fields: [range('range', 'Table', 'A1:D100'), { key: 'header', label: 'First row is a header', kind: 'choice', options: [['yes', 'Yes'], ['no', 'No']], dflt: 'yes' }],
    compile: (a, c) => {
      const r = c.range(a.range);
      const n = c.temp('cols');
      const i = c.temp('i');
      return [
        `ReDim ${n}(0 To ${r}.Columns.Count - 1)`,
        `For ${i} = 0 To ${r}.Columns.Count - 1: ${n}(${i}) = ${i} + 1: Next ${i}`,
        `${r}.RemoveDuplicates Columns:=(${n}), Header:=${a.header === 'no' ? 'xlNo' : 'xlYes'}`,
      ];
    },
  },
  {
    type: 'deleteRowsIf',
    category: 'Cells',
    title: 'Delete rows that match',
    sentence: 'Delete rows in {range} where column {column} {op} {value}',
    icon: 'trash-2',
    fields: [
      range('range', 'Table', 'A2:D100'),
      { key: 'column', label: 'Column', kind: 'column', dflt: 'A' },
      { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '=' },
      { key: 'value', label: 'Value', kind: 'text', optional: true },
    ],
    compile: (a, c) => {
      const r = c.range(a.range);
      const i = c.temp('i');
      const cell = `${r}.Cells(${i}, ${c.columnIndexIn(a.column, a.range)})`;
      // Bottom to top, so deleting a row doesn't skip the next one
      return [`For ${i} = ${r}.Rows.Count To 1 Step -1`, `    If ${c.condition(cell, a.op, a.value)} Then ${r}.Rows(${i}).EntireRow.Delete`, `Next ${i}`];
    },
  },
  {
    type: 'insertRows',
    category: 'Cells',
    title: 'Insert rows',
    sentence: 'Insert {count} row(s) above row {row}',
    icon: 'between-horizontal-start',
    fields: [{ key: 'count', label: 'How many', kind: 'number', dflt: '1' }, { key: 'row', label: 'Above row', kind: 'number', dflt: '2' }],
    compile: (a, c) => [`Rows(${c.num(a.row)} & ":" & ${c.num(a.row)} + ${c.num(a.count)} - 1).Insert`],
  },
  {
    type: 'fillDown',
    category: 'Cells',
    title: 'Fill down',
    sentence: 'Fill the top row of {range} down',
    icon: 'arrow-down-to-line',
    fields: [range('range', 'Cells', 'C2:C100')],
    compile: (a, c) => [`${c.range(a.range)}.FillDown`],
  },
  {
    type: 'select',
    category: 'Cells',
    title: 'Select cells',
    sentence: 'Select {range}',
    icon: 'mouse-pointer-2',
    fields: [range()],
    compile: (a, c) => {
      const sheet = c.sheetOf(a.range);
      return [...(sheet ? [`Worksheets(${c.str(sheet)}).Activate`] : []), `${c.range(a.range)}.Select`];
    },
  },

  // Format
  {
    type: 'fill',
    category: 'Format',
    title: 'Fill color',
    sentence: 'Color {range} {color}',
    icon: 'paint-bucket',
    fields: [range(), { key: 'color', label: 'Color', kind: 'color', options: COLORS, dflt: '#fef08a' }],
    compile: (a, c) => [a.color === 'none' ? `${c.range(a.range)}.Interior.ColorIndex = xlNone` : `${c.range(a.range)}.Interior.Color = ${c.color(a.color)}`],
  },
  {
    type: 'fontColor',
    category: 'Format',
    title: 'Text color',
    sentence: 'Make the text in {range} {color}',
    icon: 'baseline',
    fields: [range(), { key: 'color', label: 'Color', kind: 'color', options: [['#dc2626', 'Red'], ['#16a34a', 'Green'], ['#2563eb', 'Blue'], ['#ea580c', 'Orange'], ['#6b7280', 'Gray'], ['none', 'Automatic']], dflt: '#dc2626' }],
    compile: (a, c) => [a.color === 'none' ? `${c.range(a.range)}.Font.ColorIndex = xlAutomatic` : `${c.range(a.range)}.Font.Color = ${c.color(a.color)}`],
  },
  {
    type: 'fontStyle',
    category: 'Format',
    title: 'Bold, italic, size',
    sentence: 'Make {range} {style}',
    icon: 'bold',
    fields: [
      range(),
      { key: 'style', label: 'Style', kind: 'choice', options: [['bold', 'bold'], ['notbold', 'not bold'], ['italic', 'italic'], ['underline', 'underlined'], ['big', 'bigger (14 pt)'], ['normal', 'normal size (11 pt)']], dflt: 'bold' },
    ],
    compile: (a, c) => {
      const f = `${c.range(a.range)}.Font`;
      switch (a.style) {
        case 'notbold':
          return [`${f}.Bold = False`];
        case 'italic':
          return [`${f}.Italic = True`];
        case 'underline':
          return [`${f}.Underline = xlUnderlineStyleSingle`];
        case 'big':
          return [`${f}.Size = 14`];
        case 'normal':
          return [`${f}.Size = 11`];
        default:
          return [`${f}.Bold = True`];
      }
    },
  },
  {
    type: 'numberFormat',
    category: 'Format',
    title: 'Number format',
    sentence: 'Show {range} as {format}',
    icon: 'dollar-sign',
    fields: [
      range(),
      {
        key: 'format',
        label: 'Format',
        kind: 'choice',
        options: [
          ['"$"#,##0.00', 'money'],
          ['0.0%', 'percent'],
          ['#,##0', 'whole number'],
          ['#,##0.00', 'number with 2 decimals'],
          ['m/d/yyyy', 'date'],
          ['h:mm AM/PM', 'time'],
          ['@', 'text'],
          ['General', 'general'],
        ],
        dflt: '"$"#,##0.00',
      },
    ],
    compile: (a, c) => [`${c.range(a.range)}.NumberFormat = ${c.str(a.format)}`],
  },
  {
    type: 'borders',
    category: 'Format',
    title: 'Borders',
    sentence: 'Give {range} {which}',
    icon: 'square-dashed',
    fields: [range('range', 'Cells', 'A1:D10'), { key: 'which', label: 'Borders', kind: 'choice', options: [['all', 'all borders'], ['outline', 'an outline'], ['bottom', 'a bottom border'], ['none', 'no borders']], dflt: 'all' }],
    compile: (a, c) => {
      const r = c.range(a.range);
      if (a.which === 'none') return [`${r}.Borders.LineStyle = xlNone`];
      if (a.which === 'outline') return [`${r}.BorderAround LineStyle:=xlContinuous, Weight:=xlThin`];
      if (a.which === 'bottom') return [`${r}.Borders(xlEdgeBottom).LineStyle = xlContinuous`];
      return [`${r}.Borders.LineStyle = xlContinuous`];
    },
  },
  {
    type: 'highlightIf',
    category: 'Format',
    title: 'Highlight matching cells',
    sentence: 'Color cells in {range} that {op} {value} {color}',
    icon: 'highlighter',
    fields: [range('range', 'Cells', 'B2:B100'), { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '>' }, { key: 'value', label: 'Value', kind: 'text', optional: true, dflt: '100' }, { key: 'color', label: 'Color', kind: 'color', options: COLORS, dflt: '#fef08a' }],
    compile: (a, c) => {
      const cell = c.temp('cell', 'Range');
      return [
        `For Each ${cell} In ${c.range(a.range)}.Cells`,
        `    If ${c.condition(cell, a.op, a.value)} Then ${a.color === 'none' ? `${cell}.Interior.ColorIndex = xlNone` : `${cell}.Interior.Color = ${c.color(a.color)}`}`,
        `Next ${cell}`,
      ];
    },
  },
  {
    type: 'autofit',
    category: 'Format',
    title: 'Fit column widths',
    sentence: 'Make columns {range} fit their contents',
    icon: 'move-horizontal',
    fields: [range('range', 'Columns', 'A:F')],
    compile: (a, c) => [`${c.range(a.range)}.EntireColumn.AutoFit`],
  },

  // Sheets
  {
    type: 'goSheet',
    category: 'Sheets',
    title: 'Go to a sheet',
    sentence: 'Go to sheet {sheet}',
    icon: 'arrow-right-to-line',
    fields: [{ key: 'sheet', label: 'Sheet', kind: 'sheet' }],
    compile: (a, c) => [`Worksheets(${c.text(a.sheet)}).Activate`],
  },
  {
    type: 'addSheet',
    category: 'Sheets',
    title: 'Add a sheet',
    sentence: 'Add a sheet named {name}',
    icon: 'file-plus',
    fields: [{ key: 'name', label: 'Name', kind: 'text', placeholder: 'Report {Today}' }],
    compile: (a, c) => [`Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = ${c.text(a.name)}`],
  },
  {
    type: 'copySheet',
    category: 'Sheets',
    title: 'Duplicate a sheet',
    sentence: 'Duplicate {sheet} as {name}',
    icon: 'files',
    fields: [{ key: 'sheet', label: 'Sheet', kind: 'sheet' }, { key: 'name', label: 'New name', kind: 'text' }],
    compile: (a, c) => [`Worksheets(${c.text(a.sheet)}).Copy After:=Worksheets(Worksheets.Count)`, `ActiveSheet.Name = ${c.text(a.name)}`],
  },
  {
    type: 'renameSheet',
    category: 'Sheets',
    title: 'Rename a sheet',
    sentence: 'Rename {sheet} to {name}',
    icon: 'text-cursor-input',
    fields: [{ key: 'sheet', label: 'Sheet', kind: 'sheet' }, { key: 'name', label: 'New name', kind: 'text' }],
    compile: (a, c) => [`Worksheets(${c.text(a.sheet)}).Name = ${c.text(a.name)}`],
  },
  {
    type: 'deleteSheet',
    category: 'Sheets',
    title: 'Delete a sheet',
    sentence: 'Delete sheet {sheet}',
    icon: 'trash',
    fields: [{ key: 'sheet', label: 'Sheet', kind: 'sheet' }],
    compile: (a, c) => [`Worksheets(${c.text(a.sheet)}).Delete`],
  },

  // Logic
  {
    type: 'if',
    category: 'Logic',
    title: 'If',
    sentence: 'If {range} {op} {value}',
    icon: 'split',
    block: true,
    hasElse: true,
    fields: [range('range', 'Cell'), { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '=' }, { key: 'value', label: 'Value', kind: 'text', optional: true }],
    compile: (a, c, step) => {
      const out = [`If ${c.condition(c.range(a.range), a.op, a.value)} Then`, ...c.block(step.body)];
      if (step.elseBody?.length) out.push('Else', ...c.block(step.elseBody));
      out.push('End If');
      return out;
    },
  },
  {
    type: 'repeat',
    category: 'Logic',
    title: 'Repeat',
    sentence: 'Repeat {times} times',
    icon: 'repeat',
    block: true,
    provides: 'Count',
    fields: [{ key: 'times', label: 'Times', kind: 'number', dflt: '5' }],
    compile: (a, c, step) => {
      const i = c.loopVar('Count');
      return [`For ${i} = 1 To ${c.num(a.times)}`, ...c.block(step.body), `Next ${i}`, ...c.endLoop('Count')];
    },
  },
  {
    type: 'eachRow',
    category: 'Logic',
    title: 'For each row',
    sentence: 'For each row from {from} to {to}',
    icon: 'rows-3',
    block: true,
    provides: 'Row',
    fields: [
      { key: 'from', label: 'First row', kind: 'number', dflt: '2' },
      { key: 'to', label: 'Last row', kind: 'text', dflt: 'last', placeholder: 'A number, or "last" for the last row with data' },
      { key: 'column', label: 'Find the last row using column', kind: 'column', dflt: 'A', optional: true },
    ],
    compile: (a, c, step) => {
      const r = c.loopVar('Row');
      const last = !a.to || /^last$/i.test(a.to.trim()) ? `Cells(Rows.Count, "${c.column(a.column || 'A')}").End(xlUp).Row` : c.num(a.to);
      return [`For ${r} = ${c.num(a.from)} To ${last}`, ...c.block(step.body), `Next ${r}`, ...c.endLoop('Row')];
    },
  },
  {
    type: 'runMacro',
    category: 'Logic',
    title: 'Run another macro',
    sentence: 'Run {macro}',
    icon: 'play',
    fields: [{ key: 'macro', label: 'Macro', kind: 'macro' }],
    compile: (a, c) => [`Application.Run ${c.str(macroProcName(a.macro ?? ''))}`],
  },
  {
    type: 'stop',
    category: 'Logic',
    title: 'Stop',
    sentence: 'Stop the macro',
    icon: 'octagon-x',
    fields: [],
    compile: () => ['End'],
  },
  {
    type: 'wait',
    category: 'Logic',
    title: 'Wait',
    sentence: 'Wait {seconds} seconds',
    icon: 'hourglass',
    fields: [{ key: 'seconds', label: 'Seconds', kind: 'number', dflt: '1' }],
    compile: (a, c) => [`Application.Wait Now + ${c.num(a.seconds)} / 86400`],
  },
  {
    type: 'setVar',
    category: 'Logic',
    title: 'Remember a value',
    sentence: 'Remember {value} as {name}',
    icon: 'variable',
    fields: [{ key: 'value', label: 'Value', kind: 'text', placeholder: 'Text, a number, or a cell like [B2]' }, { key: 'name', label: 'Name', kind: 'variable', dflt: 'Total' }],
    compile: (a, c) => [`${c.variable(a.name)} = ${c.value(a.value, true).expr}`],
  },

  // Talk
  {
    type: 'message',
    category: 'Talk',
    title: 'Show a message',
    sentence: 'Show message {text}',
    icon: 'message-square',
    fields: [{ key: 'text', label: 'Message', kind: 'text', placeholder: 'Done! Total is {Total}' }],
    compile: (a, c) => [`MsgBox ${c.text(a.text)}`],
  },
  {
    type: 'ask',
    category: 'Talk',
    title: 'Ask a question',
    sentence: 'Ask {question} and remember the answer as {name}',
    icon: 'message-circle-question',
    fields: [{ key: 'question', label: 'Question', kind: 'text', placeholder: 'How many rows?' }, { key: 'name', label: 'Remember as', kind: 'variable', dflt: 'Answer' }],
    compile: (a, c) => {
      const v = c.variable(a.name);
      // An empty answer (or Cancel) stops the macro
      return [`${v} = InputBox(${c.text(a.question)})`, `If ${v} = "" Then End`];
    },
  },
  {
    type: 'confirm',
    category: 'Talk',
    title: 'Ask to continue',
    sentence: 'Ask {question} and stop if the answer is No',
    icon: 'circle-help',
    fields: [{ key: 'question', label: 'Question', kind: 'text', placeholder: 'Clear last month’s data?' }],
    compile: (a, c) => [`If MsgBox(${c.text(a.question)}, vbYesNo + vbQuestion) = vbNo Then End`],
  },
  {
    type: 'status',
    category: 'Talk',
    title: 'Status bar text',
    sentence: 'Show {text} in the status bar',
    icon: 'panel-bottom',
    fields: [{ key: 'text', label: 'Text', kind: 'text', placeholder: 'Working on row {Row}…' }],
    compile: (a, c) => [`Application.StatusBar = ${c.text(a.text)}`],
  },

  // Workbook
  {
    type: 'save',
    category: 'Workbook',
    title: 'Save the workbook',
    sentence: 'Save the workbook',
    icon: 'save',
    fields: [],
    compile: () => ['ThisWorkbook.Save'],
  },
  {
    type: 'recalc',
    category: 'Workbook',
    title: 'Recalculate',
    sentence: 'Recalculate all formulas',
    icon: 'refresh-cw',
    fields: [],
    compile: () => ['Application.Calculate'],
  },
];

export const STEP_BY_TYPE = new Map(STEPS.map((s) => [s.type, s]));

export const STEP_CATEGORIES: StepCategory[] = ['Cells', 'Format', 'Sheets', 'Logic', 'Talk', 'Workbook'];

/** Fills in a step's sentence for the list ("Set A1 to 5"). */
export function describeStep(step: Step): string {
  const def = STEP_BY_TYPE.get(step.type);
  if (!def) return `Unknown step "${step.type}"`;
  return def.sentence.replace(/\{(\w+)\}/g, (_m, key: string) => {
    const field = def.fields.find((f) => f.key === key);
    const v = step.args[key] ?? field?.dflt ?? '';
    if (field?.options) return field.options.find(([o]) => o === v)?.[1] ?? v;
    return v === '' ? '…' : v;
  });
}

export function newStep(type: string): Step {
  const def = STEP_BY_TYPE.get(type);
  if (!def) throw new Error(`Unknown step type ${type}`);
  const args: Record<string, string> = {};
  for (const f of def.fields) if (f.dflt !== undefined) args[f.key] = f.dflt;
  return { id: 's' + Math.random().toString(36).slice(2, 9), type, args, ...(def.block ? { body: [] } : {}) };
}

/** The VBA procedure name for a macro name ("Monthly report" → Monthly_report; "Module1.Main" stays). */
export function macroProcName(name: string): string {
  return name
    .split('.')
    .map((p) => toIdentifier(p))
    .join('.');
}

// ---------- compiler ----------

const IDENT = /[^A-Za-z0-9_]/g;

/** Turns steps into VBA. One instance compiles one macro. */
export class Compiler {
  private temps = new Map<string, string>();
  private loops: { name: string; v: string }[] = [];
  private loopCount = 0;
  private vars = new Set<string>();
  private depth = 1;
  readonly problems: string[] = [];

  /** A VBA string literal. */
  str(s: string): string {
    return `"${(s ?? '').replace(/"/g, '""')}"`;
  }

  /** Text with {Variables} as a VBA string expression: "Total: " & v_Total */
  text(s: string | undefined): string {
    const src = s ?? '';
    const parts: string[] = [];
    let last = 0;
    for (const m of src.matchAll(/\{([A-Za-z][A-Za-z0-9 _]*)\}/g)) {
      if (m.index! > last) parts.push(this.str(src.slice(last, m.index)));
      parts.push(this.ref(m[1]));
      last = m.index! + m[0].length;
    }
    if (last < src.length || !parts.length) parts.push(this.str(src.slice(last)));
    return parts.join(' & ');
  }

  /** A value typed by the user: a number, TRUE/FALSE, a formula, [A1] for a cell's value, or text. */
  value(s: string | undefined, forVariable = false): { expr: string; formula: boolean } {
    const t = (s ?? '').trim();
    if (t === '') return { expr: forVariable ? '""' : 'Empty', formula: false };
    if (/^-?\d+(\.\d+)?$/.test(t)) return { expr: t, formula: false };
    if (/^(true|false)$/i.test(t)) return { expr: t[0].toUpperCase() + t.slice(1).toLowerCase(), formula: false };
    if (t.startsWith('=') && !forVariable) return { expr: this.text(t), formula: true };
    const cell = /^\[([^\]]+)\]$/.exec(t);
    if (cell) return { expr: `${this.range(cell[1])}.Value`, formula: false };
    const single = /^\{([A-Za-z][A-Za-z0-9 _]*)\}$/.exec(t);
    if (single) return { expr: this.ref(single[1]), formula: false };
    return { expr: this.text(s), formula: false };
  }

  /** A number field, which may be a {Variable}. */
  num(s: string | undefined): string {
    const t = (s ?? '').trim();
    if (/^-?\d+(\.\d+)?$/.test(t)) return t;
    if (t === '') return '0';
    return `Val(${this.text(t)})`;
  }

  /** The sheet named in "Sheet2!A1", if any. */
  sheetOf(addr: string | undefined): string | null {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!/.exec((addr ?? '').trim());
    return m ? (m[1] ?? m[2]).replace(/''/g, "'") : null;
  }

  sheetPrefix(addr: string | undefined): string {
    const sheet = this.sheetOf(addr);
    return sheet ? `Worksheets(${this.str(sheet)}).` : '';
  }

  /** Range("A1"), Worksheets("Data").Range("B2:B9") or Range("C" & r1) with {Row}. */
  range(addr: string | undefined): string {
    let a = (addr ?? '').trim();
    if (!a) {
      this.problems.push('A step is missing its cells.');
      a = 'A1';
    }
    const sheet = this.sheetOf(a);
    const local = sheet ? a.slice(a.indexOf('!') + 1) : a;
    return `${sheet ? `Worksheets(${this.str(sheet)}).` : ''}Range(${this.text(local.replace(/\$/g, ''))})`;
  }

  /** Column letters from "B", "Sheet2!B" or "B:B". */
  column(s: string | undefined): string {
    const t = (s ?? 'A').trim().replace(/^.*!/, '').replace(/\$/g, '').split(':')[0].replace(/\d+/g, '');
    if (!/^[A-Za-z]{1,3}$/.test(t)) {
      this.problems.push(`"${s}" isn't a column letter.`);
      return 'A';
    }
    return t.toUpperCase();
  }

  /** Position of a column within a table range (column C in B2:E9 is 2). */
  columnIndexIn(col: string | undefined, table: string | undefined): string {
    const letters = this.column(col);
    const first = /([A-Za-z]{1,3})\$?\d*/.exec((table ?? 'A1').replace(/^.*!/, '').replace(/\$/g, ''))?.[1] ?? 'A';
    const n = (x: string) => [...x.toUpperCase()].reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);
    const idx = n(letters) - n(first) + 1;
    if (idx < 1) this.problems.push(`Column ${letters} isn't inside ${table}.`);
    return String(Math.max(1, idx));
  }

  /** VBA color number for '#rrggbb'. */
  color(hex: string | undefined): string {
    const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '');
    if (!m) return 'RGB(255, 240, 138)';
    const n = parseInt(m[1], 16);
    return `RGB(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255})`;
  }

  /** Condition comparing a cell (expression) with a typed value. */
  condition(cellExpr: string, op: string | undefined, raw: string | undefined): string {
    const v = `${cellExpr}.Value`;
    if (op === 'empty') return `Trim(CStr(${v})) = ""`;
    if (op === 'notempty') return `Trim(CStr(${v})) <> ""`;
    if (op === 'contains') return `InStr(1, CStr(${v}), ${this.text(raw)}, vbTextCompare) > 0`;
    const t = (raw ?? '').trim();
    const numeric = /^-?\d+(\.\d+)?$/.test(t);
    const rhs = numeric ? t : this.value(raw, true).expr;
    const sym = ['=', '<>', '>', '<', '>=', '<='].includes(op ?? '') ? op : '=';
    if (numeric) return `IsNumeric(${v}) And Not IsEmpty(${v}) And ${v} ${sym} ${rhs}`;
    if (sym === '=' || sym === '<>') return `LCase(CStr(${v})) ${sym} LCase(${rhs})`;
    return `${v} ${sym} ${rhs}`;
  }

  /** A user variable ({Total} → v_Total). */
  variable(name: string | undefined): string {
    const id = 'v_' + (name ?? 'Value').trim().replace(IDENT, '_');
    this.vars.add(id);
    return id;
  }

  /** {Name} inside text: loop counters, {Today}/{Now}, or a remembered value. */
  private ref(name: string): string {
    const n = name.trim();
    for (let i = this.loops.length - 1; i >= 0; i--) if (this.loops[i].name.toLowerCase() === n.toLowerCase()) return this.loops[i].v;
    if (/^today$/i.test(n)) return 'Format(Date, "m/d/yyyy")';
    if (/^now$/i.test(n)) return 'Format(Now, "m/d/yyyy h:mm AM/PM")';
    if (/^(row|count)$/i.test(n)) {
      this.problems.push(`{${n}} only works inside a "${/^row$/i.test(n) ? 'For each row' : 'Repeat'}" block.`);
      return '0';
    }
    return this.variable(n);
  }

  /** A helper variable declared once per macro. */
  temp(base: string, type?: string): string {
    const key = base + (type ?? '');
    let v = this.temps.get(key);
    if (!v) {
      v = `${base}${this.temps.size ? this.temps.size + 1 : ''}`;
      this.temps.set(key, v);
      this.typed.set(v, base === 'cols' ? '()' : type ?? 'Long');
    }
    return v;
  }

  private typed = new Map<string, string>();

  loopVar(name: string): string {
    const v = `${name === 'Row' ? 'r' : 'n'}${++this.loopCount}`;
    this.typed.set(v, 'Long');
    this.loops.push({ name, v });
    return v;
  }

  endLoop(name: string): string[] {
    const i = this.loops.map((l) => l.name).lastIndexOf(name);
    if (i >= 0) this.loops.splice(i, 1);
    return [];
  }

  block(steps: Step[] | undefined): string[] {
    this.depth++;
    try {
      const out = this.steps(steps ?? []);
      return out.length ? out.map((l) => '    ' + l) : ["    ' (nothing yet)"];
    } finally {
      this.depth--;
    }
  }

  steps(steps: Step[]): string[] {
    const out: string[] = [];
    for (const step of steps) {
      if (step.off) continue;
      const def = STEP_BY_TYPE.get(step.type);
      if (!def) {
        this.problems.push(`Unknown step "${step.type}".`);
        continue;
      }
      out.push(`' ${describeStep(step).replace(/\r?\n/g, ' ')}`);
      out.push(...def.compile(step.args, this, step));
    }
    return out;
  }

  /** Declarations for the variables the compiled steps use. */
  declarations(): string[] {
    const out: string[] = [];
    for (const [v, t] of this.typed) out.push(t === '()' ? `Dim ${v}()` : `Dim ${v} As ${t}`);
    for (const v of [...this.vars].sort()) out.push(`Dim ${v} As Variant`);
    return out;
  }
}

/** The VBA for one step macro, as a Sub named `procName`. */
export function compileStepMacro(procName: string, steps: Step[], description?: string): { code: string; problems: string[] } {
  const c = new Compiler();
  const body = c.steps(steps);
  const lines = [`Sub ${procName}()`];
  if (description) lines.push(`    ' ${description.replace(/\r?\n/g, ' ')}`);
  for (const d of c.declarations()) lines.push('    ' + d);
  if (c.declarations().length) lines.push('');
  for (const l of body) lines.push('    ' + l);
  lines.push('    Application.StatusBar = False');
  lines.push('End Sub');
  return { code: lines.join('\r\n'), problems: c.problems };
}

/** Ready-made step macros shown as examples in the builder. */
export const STEP_EXAMPLES: { name: string; description: string; steps: () => Step[] }[] = [
  {
    name: 'Log today',
    description: 'Adds a row with today’s date to the bottom of a list.',
    steps: () => [withArgs(newStep('appendRow'), { column: 'A', values: '{Today} | New entry' })],
  },
  {
    name: 'Highlight big expenses',
    description: 'Colors every amount over 100 in column C yellow.',
    steps: () => [withArgs(newStep('highlightIf'), { range: 'C2:C200', op: '>', value: '100', color: '#fef08a' })],
  },
  {
    name: 'Clear the input cells',
    description: 'Asks first, then empties B2:B20 so the sheet is ready for next month.',
    steps: () => [withArgs(newStep('confirm'), { question: 'Clear all the inputs?' }), withArgs(newStep('clear'), { what: 'contents', range: 'B2:B20' })],
  },
  {
    name: 'Mark paid rows',
    description: 'Goes through each row and colors it green when column D says "Paid".',
    steps: () => {
      const loop = withArgs(newStep('eachRow'), { from: '2', to: 'last', column: 'A' });
      const cond = withArgs(newStep('if'), { range: 'D{Row}', op: '=', value: 'Paid' });
      cond.body = [withArgs(newStep('fill'), { range: 'A{Row}:D{Row}', color: '#bbf7d0' })];
      loop.body = [cond];
      return [loop];
    },
  },
  {
    name: 'Monthly report sheet',
    description: 'Asks for a month, copies the Template sheet and names it after the month.',
    steps: () => [withArgs(newStep('ask'), { question: 'Which month is this report for?', name: 'Month' }), withArgs(newStep('copySheet'), { sheet: 'Template', name: '{Month}' })],
  },
];

function withArgs(step: Step, args: Record<string, string>): Step {
  step.args = { ...step.args, ...args };
  return step;
}
