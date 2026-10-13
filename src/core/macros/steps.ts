// No-code macros: a list of plain-English steps (like the Shortcuts app) that compiles to VBA.
// Text fields can use {Variables}: names set by "Set a variable", "Ask a question" and the other
// Variables steps, what a loop provides ({Row}, {Count}, {Cell}, {Item}, {Sheet}), plus {Today} and {Now}.
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
  /** Only shown (and used) when another field has one of these values, or none of `not`. */
  when?: { key: string; is?: string[]; not?: string[] };
}

export type StepCategory = 'Cells' | 'Format' | 'Sheets' | 'Variables' | 'Logic' | 'Talk' | 'Workbook';

/** A name a block makes available to the steps inside it, like {Row}. */
export interface Provided {
  name: string;
  hint: string;
}

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
  /** Names the steps inside can use (shown as chips). */
  provides?: Provided[];
  /** The field naming the variable this step stores something in. */
  sets?: string;
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
const variable = (dflt = 'Total', key = 'name', label = 'Variable'): StepField => ({ key, label, kind: 'variable', dflt });
/** A comparison's value, hidden for "is empty" and "is not empty". */
const compareValue = (key = 'value', opKey = 'op', dflt?: string): StepField => ({ key, label: 'Value', kind: 'text', optional: true, dflt, placeholder: 'Text, a number, {Variable} or [B2]', when: { key: opKey, not: ['empty', 'notempty'] } });
/** The left side of a rule: a cell, or a {Variable}. */
const subject = (key = 'range', dflt = 'A1'): StepField => ({ key, label: 'Cell or {Variable}', kind: 'range', dflt, placeholder: 'A1 or {Total}' });

const NUMBER_HINT = 'A number, {Variable} or [B2]';

const MATH: [string, string][] = [
  ['add', 'add'],
  ['sub', 'subtract'],
  ['mul', 'multiply by'],
  ['div', 'divide by'],
  ['pow', 'raise to the power of'],
  ['mod', 'keep the remainder after dividing by'],
  ['round', 'round to decimal places:'],
  ['int', 'drop the decimals'],
  ['abs', 'make it positive'],
  ['atmost', 'limit it to at most'],
  ['atleast', 'limit it to at least'],
];

const TEXT_OPS: [string, string][] = [
  ['append', 'add to the end:'],
  ['prepend', 'add to the start:'],
  ['replace', 'replace'],
  ['remove', 'remove'],
  ['upper', 'make it UPPERCASE'],
  ['lower', 'make it lowercase'],
  ['proper', 'Capitalize Each Word'],
  ['trim', 'remove extra spaces'],
  ['left', 'keep the first characters:'],
  ['right', 'keep the last characters:'],
];

// ---------- the step catalog ----------

export const STEPS: StepDef[] = [
  // Cells
  {
    type: 'setValue',
    category: 'Cells',
    title: 'Set cells',
    sentence: 'Set {range} to {value}',
    icon: 'pencil',
    fields: [range(), { key: 'value', label: 'Value', kind: 'text', placeholder: 'Text, a number, {Variable} or a formula like =SUM(A1:A5)' }],
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
    fields: [{ key: 'amount', label: 'Amount', kind: 'number', dflt: '1', placeholder: NUMBER_HINT }, range('range', 'Cell')],
    compile: (a, c) => {
      const r = c.range(a.range);
      return [`${r}.Value = ${c.number(`${r}.Value`)} + ${c.num(a.amount)}`];
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
    fields: [range('range', 'Table', 'A2:D100'), { key: 'column', label: 'Column', kind: 'column', dflt: 'A' }, { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '=' }, compareValue()],
    compile: (a, c) => {
      const r = c.range(a.range);
      const i = c.temp('i');
      const cell = `${r}.Cells(${i}, ${c.columnIndexIn(a.column, a.range)}).Value`;
      // Bottom to top, so deleting a row doesn't skip the next one
      return [`For ${i} = ${r}.Rows.Count To 1 Step -1`, `    If ${c.test(cell, a.op, a.value)} Then ${r}.Rows(${i}).EntireRow.Delete`, `Next ${i}`];
    },
  },
  {
    type: 'insertRows',
    category: 'Cells',
    title: 'Insert rows',
    sentence: 'Insert {count} row(s) above row {row}',
    icon: 'between-horizontal-start',
    fields: [
      { key: 'count', label: 'How many', kind: 'number', dflt: '1', placeholder: NUMBER_HINT },
      { key: 'row', label: 'Above row', kind: 'number', dflt: '2', placeholder: NUMBER_HINT },
    ],
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
      return [...(sheet ? [`Worksheets(${c.text(sheet)}).Activate`] : []), `${c.range(a.range)}.Select`];
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
    fields: [range('range', 'Cells', 'B2:B100'), { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '>' }, compareValue('value', 'op', '100'), { key: 'color', label: 'Color', kind: 'color', options: COLORS, dflt: '#fef08a' }],
    compile: (a, c) => {
      const cell = c.temp('cell', 'Range');
      return [
        `For Each ${cell} In ${c.range(a.range)}.Cells`,
        `    If ${c.test(`${cell}.Value`, a.op, a.value)} Then ${a.color === 'none' ? `${cell}.Interior.ColorIndex = xlNone` : `${cell}.Interior.Color = ${c.color(a.color)}`}`,
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

  // Variables
  {
    type: 'setVar',
    category: 'Variables',
    title: 'Set a variable',
    sentence: 'Set {name} to {value}',
    icon: 'variable',
    sets: 'name',
    fields: [variable(), { key: 'value', label: 'Value', kind: 'text', optional: true, placeholder: 'Text, a number, [B2] for a cell, or a formula like ={Price}*2' }],
    compile: (a, c) => {
      const value = c.value(a.value, true).expr;
      return [`${c.setVariable(a.name)} = ${value}`];
    },
  },
  {
    type: 'mathVar',
    category: 'Variables',
    title: 'Change a number',
    sentence: 'Change {name}: {op} {amount}',
    icon: 'calculator',
    sets: 'name',
    fields: [variable(), { key: 'op', label: 'How', kind: 'choice', options: MATH, dflt: 'add' }, { key: 'amount', label: 'Amount', kind: 'number', dflt: '1', placeholder: NUMBER_HINT, when: { key: 'op', not: ['int', 'abs'] } }],
    compile: (a, c) => {
      const n = c.num(a.amount);
      const v = c.setVariable(a.name);
      const x = c.number(v);
      switch (a.op) {
        case 'sub':
          return [`${v} = ${x} - ${n}`];
        case 'mul':
          return [`${v} = ${x} * ${n}`];
        case 'div':
          return [`${v} = ${x} / ${n}`];
        case 'pow':
          return [`${v} = ${x} ^ ${n}`];
        case 'mod':
          // Like the MOD function (VBA's Mod only works with whole numbers)
          return [`${v} = ${x} - ${n} * Int(${x} / ${n})`];
        case 'round':
          // The spreadsheet ROUND (VBA's Round rounds 2.5 down to 2)
          return [`${v} = WorksheetFunction.Round(${x}, ${n})`];
        case 'int':
          return [`${v} = Fix(${x})`];
        case 'abs':
          return [`${v} = Abs(${x})`];
        case 'atmost':
          return [`${v} = WorksheetFunction.Min(${x}, ${n})`];
        case 'atleast':
          return [`${v} = WorksheetFunction.Max(${x}, ${n})`];
        default:
          return [`${v} = ${x} + ${n}`];
      }
    },
  },
  {
    type: 'textVar',
    category: 'Variables',
    title: 'Change text',
    sentence: 'Change text in {name}: {op} {text}{count}',
    icon: 'type',
    sets: 'name',
    fields: [
      variable('Name'),
      { key: 'op', label: 'How', kind: 'choice', options: TEXT_OPS, dflt: 'append' },
      { key: 'text', label: 'Text', kind: 'text', optional: true, placeholder: 'Text or {Variable}', when: { key: 'op', is: ['append', 'prepend', 'replace', 'remove'] } },
      { key: 'count', label: 'Characters', kind: 'number', dflt: '3', placeholder: NUMBER_HINT, when: { key: 'op', is: ['left', 'right'] } },
      { key: 'with', label: 'With', kind: 'text', optional: true, placeholder: 'New text', when: { key: 'op', is: ['replace'] } },
    ],
    compile: (a, c) => {
      const t = c.text(a.text);
      const n = c.num(a.count);
      const w = c.text(a.with);
      const v = c.setVariable(a.name);
      switch (a.op) {
        case 'prepend':
          return [`${v} = ${t} & ${v}`];
        case 'replace':
          return [`${v} = Replace(${v}, ${t}, ${w}, 1, -1, vbTextCompare)`];
        case 'remove':
          return [`${v} = Replace(${v}, ${t}, "", 1, -1, vbTextCompare)`];
        case 'upper':
          return [`${v} = UCase(${v})`];
        case 'lower':
          return [`${v} = LCase(${v})`];
        case 'proper':
          return [`${v} = StrConv(${v}, vbProperCase)`];
        case 'trim':
          // The spreadsheet TRIM also turns double spaces inside into one
          return [`${v} = WorksheetFunction.Trim(${v})`];
        case 'left':
          return [`${v} = Left(${v}, ${n})`];
        case 'right':
          return [`${v} = Right(${v}, ${n})`];
        default:
          return [`${v} = ${v} & ${t}`];
      }
    },
  },
  {
    type: 'dateVar',
    category: 'Variables',
    title: 'Change a date',
    sentence: 'Change date {name}: {op} {amount} {unit}',
    icon: 'calendar-plus',
    sets: 'name',
    fields: [
      variable('Due'),
      { key: 'op', label: 'How', kind: 'choice', options: [['add', 'add'], ['sub', 'subtract']], dflt: 'add' },
      { key: 'amount', label: 'Amount', kind: 'number', dflt: '1', placeholder: NUMBER_HINT },
      { key: 'unit', label: 'Unit', kind: 'choice', options: [['d', 'days'], ['wd', 'working days'], ['ww', 'weeks'], ['m', 'months'], ['yyyy', 'years']], dflt: 'd' },
    ],
    compile: (a, c) => {
      const n = a.op === 'sub' ? `-(${c.num(a.amount)})` : c.num(a.amount);
      const v = c.setVariable(a.name);
      // A variable that was never set starts from today
      const start = c.temp('start', 'Date');
      const out = [`If IsEmpty(${v}) Then ${start} = Date Else ${start} = CDate(${v})`];
      if (a.unit === 'wd') return [...out, `${v} = CDate(WorksheetFunction.WorkDay(${start}, ${n}))`];
      return [...out, `${v} = DateAdd(${c.str(['d', 'ww', 'm', 'yyyy'].includes(a.unit) ? a.unit : 'd')}, ${n}, ${start})`];
    },
  },
  {
    type: 'measureVar',
    category: 'Variables',
    title: 'Get a total or count',
    sentence: 'Remember the {what} of {range} as {name}',
    icon: 'sigma',
    sets: 'name',
    fields: [
      {
        key: 'what',
        label: 'What',
        kind: 'choice',
        options: [
          ['sum', 'sum'],
          ['average', 'average'],
          ['max', 'largest value'],
          ['min', 'smallest value'],
          ['count', 'count of numbers'],
          ['counta', 'count of filled cells'],
          ['countblank', 'count of empty cells'],
          ['lastrow', 'last row with data'],
          ['nextrow', 'first empty row below the data'],
        ],
        dflt: 'sum',
      },
      range('range', 'Cells', 'B2:B100'),
      variable(),
    ],
    compile: (a, c) => {
      const r = c.range(a.range);
      const v = c.setVariable(a.name);
      if (a.what === 'lastrow' || a.what === 'nextrow') {
        const last = `${c.sheetPrefix(a.range)}Cells(Rows.Count, ${r}.Column).End(xlUp).Row`;
        if (a.what === 'lastrow') return [`${v} = ${last}`];
        return [`${v} = ${last} + 1`, `If IsEmpty(${c.sheetPrefix(a.range)}Cells(1, ${r}.Column).Value) Then ${v} = 1`];
      }
      const fn = { sum: 'Sum', average: 'Average', max: 'Max', min: 'Min', count: 'Count', counta: 'CountA', countblank: 'CountBlank' }[a.what ?? 'sum'] ?? 'Sum';
      // Application.Average gives #DIV/0! for no numbers instead of stopping the macro
      return [`${v} = Application.${fn}(${r})`];
    },
  },
  {
    type: 'findVar',
    category: 'Variables',
    title: 'Find a value',
    sentence: 'Find {find} in {range} and remember its {what} as {name}',
    icon: 'search',
    sets: 'name',
    fields: [
      { key: 'find', label: 'Find', kind: 'text', placeholder: 'Text, a number or {Variable}' },
      range('range', 'In', 'A:A'),
      { key: 'what', label: 'Remember', kind: 'choice', options: [['row', 'row number'], ['address', 'address'], ['right', 'value one cell to the right'], ['found', 'TRUE or FALSE (was it found?)']], dflt: 'row' },
      variable('Found'),
      { key: 'match', label: 'Match', kind: 'choice', options: [['whole', 'the whole cell'], ['part', 'part of the cell']], dflt: 'whole' },
    ],
    compile: (a, c) => {
      const f = c.temp('found', 'Range');
      const what = c.value(a.find, true).expr;
      const v = c.setVariable(a.name);
      const [missing, hit] =
        a.what === 'address' ? ['""', `${f}.Address(False, False)`] : a.what === 'right' ? ['""', `${f}.Offset(0, 1).Value`] : a.what === 'found' ? ['False', 'True'] : ['0', `${f}.Row`];
      return [
        `Set ${f} = ${c.range(a.range)}.Find(What:=${what}, LookIn:=xlValues, LookAt:=${a.match === 'part' ? 'xlPart' : 'xlWhole'}, MatchCase:=False)`,
        `If ${f} Is Nothing Then ${v} = ${missing} Else ${v} = ${hit}`,
      ];
    },
  },
  {
    type: 'randomVar',
    category: 'Variables',
    title: 'Pick a random number',
    sentence: 'Pick a random whole number from {from} to {to} as {name}',
    icon: 'dices',
    sets: 'name',
    fields: [{ key: 'from', label: 'From', kind: 'number', dflt: '1', placeholder: NUMBER_HINT }, { key: 'to', label: 'To', kind: 'number', dflt: '6', placeholder: NUMBER_HINT }, variable('Roll')],
    compile: (a, c) => {
      const lo = c.num(a.from);
      const hi = c.num(a.to);
      return [`${c.setVariable(a.name)} = WorksheetFunction.RandBetween(${lo}, ${hi})`];
    },
  },

  // Logic
  {
    type: 'if',
    category: 'Logic',
    title: 'If',
    sentence: 'If {range} {op} {value} {join} {range2} {op2} {value2}',
    icon: 'split',
    block: true,
    hasElse: true,
    fields: [
      subject(),
      { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '=' },
      compareValue(),
      { key: 'join', label: 'Another rule', kind: 'choice', options: [['', '+ and/or'], ['and', 'and'], ['or', 'or']], dflt: '', optional: true },
      { ...subject('range2', 'B1'), when: { key: 'join', is: ['and', 'or'] } },
      { key: 'op2', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '=', when: { key: 'join', is: ['and', 'or'] } },
      { ...compareValue('value2', 'op2'), when: { key: 'join', is: ['and', 'or'] } },
    ],
    compile: (a, c, step) => {
      const out = [`If ${c.rules(a)} Then`, ...c.block(step.body)];
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
    provides: [{ name: 'Count', hint: 'the repeat number (1, 2, 3…)' }],
    fields: [{ key: 'times', label: 'Times', kind: 'number', dflt: '5', placeholder: NUMBER_HINT }],
    compile: (a, c, step) => {
      const times = c.num(a.times);
      const i = c.counter('n');
      const body = c.loop('For', [{ name: 'Count', text: i, numeric: true }], step.body);
      return [`For ${i} = 1 To ${times}`, ...body, `Next ${i}`];
    },
  },
  {
    type: 'repeatWhile',
    category: 'Logic',
    title: 'Repeat while…',
    sentence: 'Repeat {mode} {range} {op} {value}',
    icon: 'refresh-ccw',
    block: true,
    provides: [{ name: 'Count', hint: 'how many times it has repeated (1, 2, 3…)' }],
    fields: [{ key: 'mode', label: 'Repeat', kind: 'choice', options: [['while', 'while'], ['until', 'until']], dflt: 'while' }, subject('range', '{Total}'), { key: 'op', label: 'Rule', kind: 'choice', options: COMPARE, dflt: '<' }, compareValue('value', 'op', '100')],
    compile: (a, c, step) => {
      const rule = c.rules(a);
      const n = c.counter('n');
      const body = c.loop('Do', [{ name: 'Count', text: n, numeric: true }], step.body);
      return [`${n} = 0`, `Do ${a.mode === 'until' ? 'Until' : 'While'} ${rule}`, `    ${n} = ${n} + 1`, ...body, 'Loop'];
    },
  },
  {
    type: 'eachRow',
    category: 'Logic',
    title: 'For each row',
    sentence: 'For each row from {from} to {to}',
    icon: 'rows-3',
    block: true,
    provides: [{ name: 'Row', hint: 'the row number, e.g. D{Row}' }],
    fields: [
      { key: 'from', label: 'First row', kind: 'number', dflt: '2', placeholder: NUMBER_HINT },
      { key: 'to', label: 'Last row', kind: 'text', dflt: 'last', placeholder: 'A number, or "last" for the last row with data' },
      { key: 'column', label: 'Find the last row using column', kind: 'column', dflt: 'A', optional: true },
    ],
    compile: (a, c, step) => {
      const from = c.num(a.from);
      const last = !a.to || /^last$/i.test(a.to.trim()) ? `Cells(Rows.Count, "${c.column(a.column || 'A')}").End(xlUp).Row` : c.num(a.to);
      const r = c.counter('r');
      const body = c.loop('For', [{ name: 'Row', text: r, numeric: true }], step.body);
      return [`For ${r} = ${from} To ${last}`, ...body, `Next ${r}`];
    },
  },
  {
    type: 'eachCell',
    category: 'Logic',
    title: 'For each cell',
    sentence: 'For each cell in {range}',
    icon: 'grid-3x3',
    block: true,
    provides: [
      { name: 'Cell', hint: 'the cell, e.g. Set {Cell} to …' },
      { name: 'Value', hint: 'what’s in it' },
    ],
    fields: [range('range', 'Cells', 'B2:B20')],
    compile: (a, c, step) => {
      const r = c.range(a.range);
      const cell = c.counter('cell', 'Range');
      const body = c.loop(
        'For',
        [
          { name: 'Cell', text: `${cell}.Address(False, False)`, range: cell },
          { name: 'Value', text: `${cell}.Value` },
        ],
        step.body,
      );
      return [`For Each ${cell} In ${r}.Cells`, ...body, `Next ${cell}`];
    },
  },
  {
    type: 'eachItem',
    category: 'Logic',
    title: 'For each item in a list',
    sentence: 'For each item in {list}',
    icon: 'list',
    block: true,
    provides: [{ name: 'Item', hint: 'the item' }],
    fields: [
      { key: 'list', label: 'List', kind: 'text', dflt: 'Jan, Feb, Mar', placeholder: 'Jan, Feb, Mar or {List}' },
      { key: 'sep', label: 'Separated by', kind: 'choice', options: [[',', 'commas'], [';', 'semicolons'], ['|', 'bars ( | )'], [' ', 'spaces']], dflt: ',' },
    ],
    compile: (a, c, step) => {
      const list = c.text(a.list);
      const item = c.counter('item', 'Variant');
      const body = c.loop('For', [{ name: 'Item', text: `Trim(${item})` }], step.body);
      // Blank items ("a,,b" or a comma at the end) are skipped
      return [`For Each ${item} In Split(${list}, ${c.str(a.sep || ',')})`, `    If Trim(${item}) <> "" Then`, ...body.map((l) => '    ' + l), '    End If', `Next ${item}`];
    },
  },
  {
    type: 'eachSheet',
    category: 'Logic',
    title: 'For each sheet',
    sentence: 'For each sheet in the workbook',
    icon: 'layers',
    block: true,
    provides: [{ name: 'Sheet', hint: 'the sheet’s name (steps inside work on that sheet)' }],
    fields: [],
    compile: (_a, c, step) => {
      const ws = c.counter('ws', 'Worksheet');
      const back = c.temp('startSheet', 'Worksheet');
      const body = c.loop('For', [{ name: 'Sheet', text: `${ws}.Name` }], step.body);
      // Hidden sheets are skipped; afterwards the sheet you started on is shown again
      return [`Set ${back} = ActiveSheet`, `For Each ${ws} In Worksheets`, `    If ${ws}.Visible = xlSheetVisible Then`, `        ${ws}.Activate`, ...body.map((l) => '    ' + l), '    End If', `Next ${ws}`, `${back}.Activate`];
    },
  },
  {
    type: 'exitLoop',
    category: 'Logic',
    title: 'Stop repeating',
    sentence: 'Stop repeating (leave the loop)',
    icon: 'log-out',
    fields: [],
    compile: (_a, c) => [c.exitLoop()],
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
    fields: [{ key: 'seconds', label: 'Seconds', kind: 'number', dflt: '1', placeholder: NUMBER_HINT }],
    compile: (a, c) => [`Application.Wait Now + ${c.num(a.seconds)} / 86400`],
  },
  {
    type: 'onError',
    category: 'Logic',
    title: 'If a step fails',
    sentence: 'If a later step fails, {mode}',
    icon: 'shield-alert',
    fields: [{ key: 'mode', label: 'Then', kind: 'choice', options: [['skip', 'skip it and keep going'], ['stop', 'stop and show the error']], dflt: 'skip' }],
    compile: (a) => [a.mode === 'stop' ? 'On Error GoTo 0' : 'On Error Resume Next'],
  },
  {
    type: 'note',
    category: 'Logic',
    title: 'Note',
    sentence: 'Note: {text}',
    icon: 'sticky-note',
    fields: [{ key: 'text', label: 'Note', kind: 'text', optional: true, placeholder: 'What the next steps do (for you, not run)' }],
    // describeStep already writes it as a comment
    compile: () => [],
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
    sets: 'name',
    fields: [{ key: 'question', label: 'Question', kind: 'text', placeholder: 'How many rows?' }, { key: 'name', label: 'Remember as', kind: 'variable', dflt: 'Answer' }],
    compile: (a, c) => {
      const q = c.text(a.question);
      const v = c.setVariable(a.name);
      // An empty answer (or Cancel) stops the macro
      return [`${v} = InputBox(${q})`, `If ${v} = "" Then End`];
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
  {
    type: 'log',
    category: 'Talk',
    title: 'Write to the output',
    sentence: 'Write {text} to the output panel',
    icon: 'terminal',
    fields: [{ key: 'text', label: 'Text', kind: 'text', placeholder: 'Total so far: {Total}' }],
    compile: (a, c) => [`Debug.Print ${c.text(a.text)}`],
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

export const STEP_CATEGORIES: StepCategory[] = ['Cells', 'Format', 'Sheets', 'Variables', 'Logic', 'Talk', 'Workbook'];

/** Whether a field applies, given the other fields (the replacement text only matters for "replace"). */
export function fieldShown(f: StepField, args: Record<string, string>, def: StepDef): boolean {
  if (!f.when) return true;
  const other = def.fields.find((x) => x.key === f.when!.key);
  const v = args[f.when.key] ?? other?.dflt ?? '';
  if (other && !fieldShown(other, args, def)) return false;
  if (f.when.is && !f.when.is.includes(v)) return false;
  if (f.when.not?.includes(v)) return false;
  return true;
}

/** Fills in a step's sentence for the list ("Set A1 to 5"). */
export function describeStep(step: Step): string {
  const def = STEP_BY_TYPE.get(step.type);
  if (!def) return `Unknown step "${step.type}"`;
  return def.sentence
    .replace(/\{(\w+)\}/g, (_m, key: string) => {
      const field = def.fields.find((f) => f.key === key);
      if (field && !fieldShown(field, step.args, def)) return '';
      const v = step.args[key] ?? field?.dflt ?? '';
      if (field?.options) return v === '' ? '' : (field.options.find(([o]) => o === v)?.[1] ?? v);
      return v === '' ? '…' : v;
    })
    .replace(/ {2,}/g, ' ')
    .trim();
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

/** The variable names (as typed) that steps store values in, in order, without repeats. */
export function variablesSetIn(steps: Step[]): string[] {
  const seen = new Map<string, string>();
  const walk = (list: Step[]) => {
    for (const s of list) {
      const def = STEP_BY_TYPE.get(s.type);
      const name = def?.sets ? (s.args[def.sets] ?? def.fields.find((f) => f.key === def.sets)?.dflt ?? '').trim() : '';
      const id = varId(name).toLowerCase();
      if (name && !seen.has(id)) seen.set(id, name);
      walk(s.body ?? []);
      walk(s.elseBody ?? []);
    }
  };
  walk(steps);
  return [...seen.values()];
}

// ---------- compiler ----------

const IDENT = /[^A-Za-z0-9_]/g;
const VAR_REF = /\{([A-Za-z][A-Za-z0-9 _]*)\}/g;
const NUMBER = /^-?\d+(\.\d+)?$/;
/** A cell or range address, maybe on another sheet, maybe with {Variables} for parts ("D{Row}"). */
const CELLISH = /^(?:(?:'(?:[^']|'')+'|[^!'"\s]+)!)?\$?[A-Za-z]{1,3}\$?(?:\d+|\{[^}]+\})(?::\$?[A-Za-z]{1,3}\$?(?:\d+|\{[^}]+\}))?$/;

const varId = (name: string) => 'v_' + name.trim().replace(IDENT, '_');

/** A name a loop makes available, and what it means as text and (for {Cell}) as a range. */
interface LoopName {
  name: string;
  text: string;
  range?: string;
  /** Already a number (loop counters), so it needs no converting. */
  numeric?: boolean;
}

// Small functions added to the generated code when steps need them (shown with "Show code", so the
// code still works when pasted into Excel).
const HELPERS: Record<string, string> = {
  StepNumber: `Private Function StepNumber(ByVal v As Variant) As Double
    ' A value as a number ("1,200" and "$5" work too); anything else counts as 0
    If IsError(v) Then
        StepNumber = 0
    ElseIf IsNumeric(v) Then
        StepNumber = CDbl(v)
    ElseIf IsDate(v) Then
        StepNumber = CDbl(CDate(v))
    End If
End Function`,
  StepTest: `Private Function StepTest(ByVal a As Variant, ByVal op As String, ByVal b As Variant) As Boolean
    ' Compares the way a person would: numbers as numbers, dates as dates, text ignoring case
    Dim d As Double
    If IsError(a) Or IsError(b) Then Exit Function
    Select Case op
        Case "empty"
            StepTest = Trim(CStr(a)) = ""
            Exit Function
        Case "notempty"
            StepTest = Trim(CStr(a)) <> ""
            Exit Function
        Case "contains"
            StepTest = InStr(1, CStr(a), CStr(b), vbTextCompare) > 0
            Exit Function
    End Select
    If IsNumeric(b) And Trim(CStr(b)) <> "" And VarType(b) <> vbBoolean Then
        ' Blank cells and text are never compared with a number
        If Not IsNumeric(a) Or Trim(CStr(a)) = "" Then Exit Function
        d = CDbl(a) - CDbl(b)
    ElseIf IsDate(a) And IsDate(b) Then
        d = CDbl(CDate(a)) - CDbl(CDate(b))
    Else
        d = StrComp(Trim(CStr(a)), Trim(CStr(b)), vbTextCompare)
    End If
    Select Case op
        Case "<>"
            StepTest = d <> 0
        Case ">"
            StepTest = d > 0
        Case "<"
            StepTest = d < 0
        Case ">="
            StepTest = d >= 0
        Case "<="
            StepTest = d <= 0
        Case Else
            StepTest = d = 0
    End Select
End Function`,
  StepLiteral: `Private Function StepLiteral(ByVal v As Variant) As String
    ' A value written the way a formula needs it: numbers as they are, text in quotes
    If IsEmpty(v) Then
        StepLiteral = "0"
    ElseIf IsError(v) Then
        StepLiteral = "NA()"
    ElseIf VarType(v) = vbBoolean Then
        StepLiteral = IIf(v, "TRUE", "FALSE")
    ElseIf VarType(v) = vbDate Then
        StepLiteral = Trim(Str(CDbl(v)))
    ElseIf VarType(v) = vbString Then
        StepLiteral = """" & Replace(v, """", """""") & """"
        ' Numbers typed as text (an answer to a question) count as numbers, but "007" stays text
        If IsNumeric(v) Then
            If CStr(CDbl(v)) = Trim(v) Then StepLiteral = Trim(Str(CDbl(v)))
        End If
    Else
        StepLiteral = Trim(Str(v))
    End If
End Function`,
};

/** VBA for the helper functions with these names. */
export function stepHelperCode(names: Iterable<string>): string {
  const want = new Set(names);
  return Object.keys(HELPERS)
    .filter((n) => want.has(n))
    .map((n) => HELPERS[n].replace(/\n/g, '\r\n'))
    .join('\r\n\r\n');
}

/** Turns steps into VBA. One instance compiles one macro. */
export class Compiler {
  private temps = new Map<string, string>();
  private typed = new Map<string, string>();
  private loops: { kind: 'For' | 'Do'; names: LoopName[] }[] = [];
  private loopCount = 0;
  private vars = new Set<string>();
  /** Variables (VBA names) that steps store something in, and the ones text refers to. */
  private setVars = new Set<string>();
  private usedVars = new Map<string, string>();
  readonly helpers = new Set<string>();
  readonly problems: string[] = [];

  /** `known`: the variables the whole macro sets, so a loop can use one that's set further down. */
  constructor(known: string[] = []) {
    for (const name of known) this.setVars.add(varId(name).toLowerCase());
  }

  /** A VBA string literal. */
  str(s: string): string {
    return `"${(s ?? '').replace(/"/g, '""')}"`;
  }

  /** Text with {Variables} as a VBA string expression: "Total: " & v_Total */
  text(s: string | undefined): string {
    const src = s ?? '';
    const parts: string[] = [];
    let last = 0;
    for (const m of src.matchAll(VAR_REF)) {
      if (m.index! > last) parts.push(this.str(src.slice(last, m.index)));
      parts.push(this.ref(m[1]));
      last = m.index! + m[0].length;
    }
    if (last < src.length || !parts.length) parts.push(this.str(src.slice(last)));
    return parts.join(' & ');
  }

  /**
   * A formula with {Variables} as a VBA string expression. A variable becomes a value the formula
   * understands: "Sam" in quotes, 5 as a number. Where it's part of a reference (B{Row}, {Sheet}!A1)
   * or inside quotes in the formula, its text goes in as it is.
   */
  formula(s: string): string {
    const parts: string[] = [];
    let last = 0;
    for (const m of s.matchAll(VAR_REF)) {
      const at = m.index!;
      if (at > last) parts.push(this.str(s.slice(last, at)));
      const x = this.ref(m[1]);
      const before = s.slice(0, at);
      const prev = before.slice(-1);
      const next = s.charAt(at + m[0].length);
      const inQuotes = (before.match(/"/g)?.length ?? 0) % 2 === 1;
      if (inQuotes) parts.push(`Replace(${x}, """", """""")`);
      else if (next === '!') parts.push(`"'" & Replace(${x}, "'", "''") & "'"`);
      else if (/[A-Za-z$:]/.test(prev) || /[0-9:$]/.test(next)) parts.push(x);
      else parts.push(`${this.helper('StepLiteral')}(${x})`);
      last = at + m[0].length;
    }
    if (last < s.length || !parts.length) parts.push(this.str(s.slice(last)));
    return parts.join(' & ');
  }

  /**
   * A value typed by the user: a number, TRUE/FALSE, a formula, [A1] for a cell's value, a {Variable},
   * or text. For variables a formula is calculated (=ROUND({Total}, 2)); for cells it's put in the cell.
   */
  value(s: string | undefined, forVariable = false): { expr: string; formula: boolean } {
    const t = (s ?? '').trim();
    if (t === '') return { expr: forVariable ? '""' : 'Empty', formula: false };
    if (NUMBER.test(t)) return { expr: t, formula: false };
    if (/^(true|false)$/i.test(t)) return { expr: t[0].toUpperCase() + t.slice(1).toLowerCase(), formula: false };
    if (t.startsWith('=')) return forVariable ? { expr: `Evaluate(${this.formula(t.slice(1))})`, formula: false } : { expr: this.formula(t), formula: true };
    const cell = /^\[([^\]]+)\]$/.exec(t);
    if (cell) return { expr: `${this.range(cell[1])}.Value`, formula: false };
    const single = /^\{([A-Za-z][A-Za-z0-9 _]*)\}$/.exec(t);
    if (single) return { expr: this.ref(single[1]), formula: false };
    return { expr: this.text(s), formula: false };
  }

  /** A number field: a number, a {Variable}, [B2] or a formula. */
  num(s: string | undefined): string {
    const t = (s ?? '').trim();
    if (NUMBER.test(t)) return t;
    if (t === '') return '0';
    const single = /^\{([A-Za-z][A-Za-z0-9 _]*)\}$/.exec(t);
    const loop = single && this.loopName(single[1]);
    if (loop && loop.numeric) return loop.text;
    return this.number(this.value(t, true).expr);
  }

  /** An expression converted to a number. */
  number(expr: string): string {
    return `${this.helper('StepNumber')}(${expr})`;
  }

  /** The sheet named in "Sheet2!A1", if any. */
  sheetOf(addr: string | undefined): string | null {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!/.exec((addr ?? '').trim());
    return m ? (m[1] ?? m[2]).replace(/''/g, "'") : null;
  }

  sheetPrefix(addr: string | undefined): string {
    const sheet = this.sheetOf(addr);
    return sheet ? `Worksheets(${this.text(sheet)}).` : '';
  }

  /** Range("A1"), Worksheets("Data").Range("B2:B9"), Range("C" & r1) with {Row}, or the cell of a loop. */
  range(addr: string | undefined): string {
    let a = (addr ?? '').trim();
    if (!a) {
      this.problems.push('A step is missing its cells.');
      a = 'A1';
    }
    const single = /^\{([A-Za-z][A-Za-z0-9 _]*)\}$/.exec(a);
    const loop = single && this.loopName(single[1]);
    if (loop && loop.range) return loop.range;
    const sheet = this.sheetOf(a);
    const local = sheet ? a.slice(a.indexOf('!') + 1) : a;
    return `${sheet ? `Worksheets(${this.text(sheet)}).` : ''}Range(${this.text(local.replace(/\$/g, ''))})`;
  }

  /** What a rule checks: a cell's value (A1, D{Row}, [B2]) or a {Variable}. */
  subject(raw: string | undefined): string {
    const t = (raw ?? '').trim();
    if (!t) {
      this.problems.push('A rule is missing the cell or {Variable} to check.');
      return 'Empty';
    }
    const single = /^\{([A-Za-z][A-Za-z0-9 _]*)\}$/.exec(t);
    if (single) {
      const loop = this.loopName(single[1]);
      if (loop?.range) return `${loop.range}.Value`;
      if (/^today$/i.test(single[1].trim())) return 'Date';
      if (/^now$/i.test(single[1].trim())) return 'Now';
      return this.ref(single[1]);
    }
    if (CELLISH.test(t)) return `${this.range(t)}.Value`;
    if (/^\[[^\]]+\]$/.test(t)) return this.value(t, true).expr;
    if (/^[A-Za-z][A-Za-z0-9 _]*$/.test(t)) this.problems.push(`"${t}" isn't a cell. To check a variable, write {${t}}.`);
    return this.value(t, true).expr;
  }

  /** A rule comparing a value (VBA expression) with what the user typed. */
  test(valueExpr: string, op: string | undefined, raw: string | undefined): string {
    const o = COMPARE.some(([v]) => v === op) ? op! : '=';
    const rhs = o === 'empty' || o === 'notempty' ? '""' : o === 'contains' ? this.text(raw) : this.value(raw, true).expr;
    return `${this.helper('StepTest')}(${valueExpr}, ${this.str(o)}, ${rhs})`;
  }

  /** The rule(s) of an If or Repeat while: one, or two joined by And/Or. */
  rules(a: Record<string, string>): string {
    const first = this.test(this.subject(a.range), a.op, a.value);
    if (a.join !== 'and' && a.join !== 'or') return first;
    return `${first} ${a.join === 'and' ? 'And' : 'Or'} ${this.test(this.subject(a.range2), a.op2, a.value2)}`;
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

  /** A user variable ({Total} → v_Total). */
  variable(name: string | undefined): string {
    const id = varId(name || 'Value');
    this.vars.add(id);
    return id;
  }

  /** A variable a step stores something in. */
  setVariable(name: string | undefined): string {
    if (!(name ?? '').trim()) this.problems.push('A step is missing the name of its variable.');
    const id = this.variable(name);
    this.setVars.add(id.toLowerCase());
    return id;
  }

  /** A helper function the code needs (see HELPERS). */
  helper(name: string): string {
    this.helpers.add(name);
    return name;
  }

  private loopName(name: string): LoopName | undefined {
    const n = name.trim().toLowerCase();
    for (let i = this.loops.length - 1; i >= 0; i--) {
      const hit = this.loops[i].names.find((l) => l.name.toLowerCase() === n);
      if (hit) return hit;
    }
    return undefined;
  }

  /** {Name} inside text: what a loop provides, {Today}/{Now}, or a variable. */
  private ref(name: string): string {
    const n = name.trim();
    const loop = this.loopName(n);
    if (loop) return loop.text;
    if (/^today$/i.test(n)) return 'Format(Date, "m/d/yyyy")';
    if (/^now$/i.test(n)) return 'Format(Now, "m/d/yyyy h:mm AM/PM")';
    const id = this.variable(n);
    const provider = STEPS.find((d) => d.provides?.some((p) => p.name.toLowerCase() === n.toLowerCase()));
    if (provider && !this.setVars.has(id.toLowerCase())) {
      this.problems.push(`{${n}} only works inside a "${provider.title}" block.`);
      return '0';
    }
    if (!this.usedVars.has(id.toLowerCase())) this.usedVars.set(id.toLowerCase(), n);
    return id;
  }

  /** A helper variable declared once per macro. */
  temp(base: string, type?: string): string {
    const key = base + (type ?? '');
    let v = this.temps.get(key);
    if (!v) {
      let n = this.temps.size;
      do v = `${base}${n ? n + 1 : ''}`;
      while (this.typed.has(v) && ++n);
      this.temps.set(key, v);
      this.typed.set(v, base === 'cols' ? '()' : (type ?? 'Long'));
    }
    return v;
  }

  /** A new loop variable (r1, n2, cell3…). */
  counter(base: string, type = 'Long'): string {
    let v: string;
    do v = `${base}${++this.loopCount}`;
    while (this.typed.has(v));
    this.typed.set(v, type);
    return v;
  }

  /** Compiles a loop's steps with the names it provides. */
  loop(kind: 'For' | 'Do', names: LoopName[], steps: Step[] | undefined): string[] {
    this.loops.push({ kind, names });
    try {
      return this.block(steps);
    } finally {
      this.loops.pop();
    }
  }

  /** Leaves the innermost loop. */
  exitLoop(): string {
    const inner = this.loops[this.loops.length - 1];
    if (!inner) {
      this.problems.push('"Stop repeating" only works inside Repeat or For each.');
      return "' (not inside a loop)";
    }
    return `Exit ${inner.kind}`;
  }

  block(steps: Step[] | undefined): string[] {
    const out = this.steps(steps ?? []);
    return out.length ? out.map((l) => '    ' + l) : ["    ' (nothing yet)"];
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

  /** Variables used in text but never given a value, with a guess at what was meant. */
  checkVariables(): void {
    const known = [...this.setVars];
    for (const [id, name] of this.usedVars) {
      if (this.setVars.has(id)) continue;
      const guess = known.find((k) => close(k, id));
      this.problems.push(`{${name}} is used but never set.${guess ? ` Did you mean {${this.displayName(guess)}}?` : ' Add a "Set a variable" step before it.'}`);
    }
  }

  private displayName(id: string): string {
    for (const v of this.vars) if (v.toLowerCase() === id) return v.slice(2).replace(/_/g, ' ');
    return id.slice(2);
  }

  /** Declarations for the variables the compiled steps use. */
  declarations(): string[] {
    const out: string[] = [];
    for (const [v, t] of this.typed) out.push(t === '()' ? `Dim ${v}()` : `Dim ${v} As ${t}`);
    for (const v of [...this.vars].sort()) out.push(`Dim ${v} As Variant`);
    return out;
  }
}

/** Names that differ by a typo: one letter added, dropped, changed or two swapped. */
function close(a: string, b: string): boolean {
  if (Math.abs(a.length - b.length) > 1) return false;
  const d: number[][] = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++)
    for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
    }
  return d[a.length][b.length] <= (a.length > 5 ? 2 : 1);
}

/**
 * The VBA for one step macro, as a Sub named `procName`. `helpers` names the helper functions it
 * calls; put stepHelperCode(helpers) once in the same module.
 */
export function compileStepMacro(procName: string, steps: Step[], description?: string): { code: string; problems: string[]; helpers: string[] } {
  const c = new Compiler(variablesSetIn(steps));
  const body = c.steps(steps);
  c.checkVariables();
  const lines = [`Sub ${procName}()`];
  if (description) lines.push(`    ' ${description.replace(/\r?\n/g, ' ')}`);
  for (const d of c.declarations()) lines.push('    ' + d);
  if (c.declarations().length) lines.push('');
  for (const l of body) lines.push('    ' + l);
  lines.push('    Application.StatusBar = False');
  lines.push('End Sub');
  return { code: lines.join('\r\n'), problems: [...new Set(c.problems)], helpers: [...c.helpers] };
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
    name: 'Add up paid amounts',
    description: 'Adds up column C for the rows marked "Paid" in column D, and shows the total.',
    steps: () => {
      const loop = withArgs(newStep('eachRow'), { from: '2', to: 'last', column: 'A' });
      const cond = withArgs(newStep('if'), { range: 'D{Row}', op: '=', value: 'Paid' });
      cond.body = [withArgs(newStep('mathVar'), { name: 'Paid', op: 'add', amount: '[C{Row}]' }), withArgs(newStep('mathVar'), { name: 'Rows', op: 'add', amount: '1' })];
      loop.body = [cond];
      return [
        withArgs(newStep('setVar'), { name: 'Paid', value: '0' }),
        withArgs(newStep('setVar'), { name: 'Rows', value: '0' }),
        loop,
        withArgs(newStep('message'), { text: '{Rows} paid rows, {Paid} in total' }),
      ];
    },
  },
  {
    name: 'Monthly report sheet',
    description: 'Asks for a month, copies the Template sheet and names it after the month.',
    steps: () => [withArgs(newStep('ask'), { question: 'Which month is this report for?', name: 'Month' }), withArgs(newStep('copySheet'), { sheet: 'Template', name: '{Month}' })],
  },
  {
    name: 'A sheet per quarter',
    description: 'Adds a sheet for each name in a list (Q1 to Q4) with a heading in A1.',
    steps: () => {
      const loop = withArgs(newStep('eachItem'), { list: 'Q1, Q2, Q3, Q4', sep: ',' });
      loop.body = [withArgs(newStep('addSheet'), { name: '{Item}' }), withArgs(newStep('setValue'), { range: 'A1', value: '{Item} report' }), withArgs(newStep('fontStyle'), { range: 'A1', style: 'bold' })];
      return [loop];
    },
  },
];

function withArgs(step: Step, args: Record<string, string>): Step {
  step.args = { ...step.args, ...args };
  return step;
}
