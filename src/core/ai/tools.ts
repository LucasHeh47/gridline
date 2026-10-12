// The tools the AI assistant can call. Read tools run immediately (each read is logged in the chat);
// propose_changes only previews — nothing changes until the user approves.
import type { AiToolDef } from '../../shared/api';
import { external } from '../data/external';
import type { Doc } from '../doc';
import type { Change } from './changes';
import { describeCell, describeSheet, evaluateFormula, findCells, readRange, ReadError, type ReadMode } from './workbook';

const sheetProp = { type: 'string', description: 'Sheet name. Optional: defaults to the sheet the user is looking at. An address like "Data!A1:C9" also names the sheet.' };

const CHANGE_SCHEMA = {
  type: 'object',
  properties: {
    op: {
      type: 'string',
      enum: ['set_values', 'fill_formula', 'clear', 'insert_rows', 'delete_rows', 'insert_columns', 'delete_columns', 'add_sheet', 'rename_sheet', 'format', 'column_width', 'add_chart', 'highlight', 'sort', 'filter', 'freeze', 'add_note', 'dropdown'],
      description: [
        'set_values: write a 2-D array `values` starting at cell `start` (text starting with "=" is a formula; null clears a cell).',
        'fill_formula: put `formula` in every cell of `range`, written for the FIRST cell; relative references adjust per cell like Excel fill-down.',
        'clear: clear `range` (`what`: contents | formats | all).',
        'insert_rows: insert `count` rows above row number `before_row`. delete_rows: delete `count` rows starting at row number `row`.',
        'insert_columns: insert `count` columns before column letter `before_column`. delete_columns: delete `count` columns starting at `column`.',
        'add_sheet: new sheet `name` (optionally `after` another sheet). rename_sheet: rename `sheet` to `name`.',
        'format: style `range` with bold, italic, fill, color (hex like "#1f4e78" or null), number_format (Excel codes like "$#,##0.00", "0.0%", "yyyy-mm-dd"), align, wrap, font_size, border.',
        'column_width: set `width` (characters, or "auto" to fit) for `columns` like "B" or "B:E".',
        'add_chart: chart of `data` (a range whose first row/column are labels) of `type`, with optional `title`, placed at cell `at`.',
        'highlight: conditional formatting on `range` with `rule` (duplicates, unique, greater_than, less_than, between, equal, text_contains, formula, blanks, errors, top, bottom), `value`/`value2` and `color` (red, yellow, green, blue). For rule "formula", `value` is a formula written for the top-left cell of the range that returns TRUE for cells to highlight.',
        'sort: sort `range` by `by` columns (letters); `header` true when the first row is headings.',
        'filter: add filter buttons to `range`. freeze: freeze the top `rows` rows and the first `columns` columns.',
        'add_note: attach a note `text` to `cell`. dropdown: a list of `choices` for cells in `range`.',
      ].join('\n'),
    },
    sheet: sheetProp,
    start: { type: 'string', description: 'set_values: top-left cell.' },
    values: { type: 'array', items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } }, description: 'set_values: rows of cell values.' },
    range: { type: 'string', description: 'A1-style range like "B2:B200" or "D:D".' },
    formula: { type: 'string', description: 'fill_formula: formula for the first cell of the range, starting with "=".' },
    what: { type: 'string', enum: ['contents', 'formats', 'all'] },
    before_row: { type: 'integer', minimum: 1 },
    row: { type: 'integer', minimum: 1 },
    before_column: { type: 'string' },
    column: { type: 'string' },
    count: { type: 'integer', minimum: 1 },
    name: { type: 'string' },
    after: { type: 'string' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    fill: { type: ['string', 'null'] },
    color: { type: ['string', 'null'], description: 'format: text color as hex. highlight: red, yellow, green or blue.' },
    number_format: { type: 'string' },
    align: { type: 'string', enum: ['left', 'center', 'right'] },
    wrap: { type: 'boolean' },
    font_size: { type: 'number' },
    border: { type: 'string', enum: ['all', 'outline', 'bottom', 'none'] },
    columns: { type: 'string' },
    width: { type: ['number', 'string'], description: 'Characters, or "auto".' },
    type: { type: 'string', enum: ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'] },
    data: { type: 'string', description: 'add_chart: the data range including labels.' },
    title: { type: 'string' },
    at: { type: 'string', description: 'add_chart: top-left cell for the chart.' },
    rule: { type: 'string', enum: ['duplicates', 'unique', 'greater_than', 'less_than', 'between', 'equal', 'text_contains', 'formula', 'blanks', 'errors', 'top', 'bottom'] },
    value: { type: ['string', 'number'] },
    value2: { type: ['string', 'number'] },
    by: { type: 'array', items: { type: 'object', properties: { column: { type: 'string' }, descending: { type: 'boolean' } }, required: ['column'] } },
    header: { type: 'boolean' },
    rows: { type: 'integer', minimum: 0 },
    cell: { type: 'string' },
    text: { type: 'string' },
    choices: { type: 'array', items: { type: 'string' } },
  },
  required: ['op'],
};

export const TOOLS: AiToolDef[] = [
  {
    name: 'read_range',
    description: 'Read cells from the workbook as a table (row numbers down the side, column letters across). Use "formulas" or "both" to see how values are calculated. Large ranges are cut off at the reading limit; read the next rows in another call, or use evaluate for totals and counts.',
    input_schema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Like "A1:F40", "Data!A1:L500" or "C:C" (whole columns stop at the last used row).' },
        sheet: sheetProp,
        show: { type: 'string', enum: ['values', 'formulas', 'both'], description: 'Default "values" (as shown on screen).' },
      },
      required: ['range'],
    },
  },
  {
    name: 'find_cells',
    description: 'Search every sheet (or one) for cells whose shown text contains the query, or whose formula contains it. Returns addresses and text.',
    input_schema: {
      type: 'object',
      properties: {
        query: { type: 'string' },
        sheet: sheetProp,
        in_formulas: { type: 'boolean', description: 'Search formulas instead of shown values (e.g. "VLOOKUP" or "Data!").' },
        exact: { type: 'boolean', description: 'Whole-cell match instead of "contains".' },
      },
      required: ['query'],
    },
  },
  {
    name: 'describe_sheet',
    description: "Summary of one sheet: each column's heading, type of data, min/max/sum, blanks and errors, the formulas used (grouped into runs), merged cells, charts, highlight rules and dropdowns.",
    input_schema: { type: 'object', properties: { sheet: sheetProp }, required: [] },
  },
  {
    name: 'describe_cell',
    description: 'Everything about one cell: value, how it is shown, formula, number format, note, the cells its formula uses (with their values) and the formulas that use it. Use it to trace why a number is what it is.',
    input_schema: { type: 'object', properties: { cell: { type: 'string', description: 'Like "D14" or "Report!B7".' }, sheet: sheetProp }, required: ['cell'] },
  },
  {
    name: 'evaluate',
    description: "Calculate any formula against the workbook without writing it anywhere, e.g. =SUM(Data!F:F), =COUNTIFS(...), =FILTER(...), =UNIQUE(...) or Gridline's live-data functions like =NFL.LEADERS(\"Passing Yards\", 2025, 10, \"QB\"). Use it to check totals, count matches, find mismatches, look up real data and test a formula before proposing it. References without a sheet name refer to the given (or current) sheet.",
    input_schema: { type: 'object', properties: { formula: { type: 'string' }, sheet: sheetProp }, required: ['formula'] },
  },
  {
    name: 'propose_changes',
    description:
      'Propose changes to the workbook. Nothing changes yet: the user sees a preview with every cell that changes (before → after) and every formula result elsewhere that changes, then approves or rejects. The result tells you what they decided. Group everything needed for one request into a single proposal. Changes apply in order, and addresses in later changes refer to the sheet after earlier ones — put row/column insertions and deletions first.',
    input_schema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Short name, e.g. "Add a Category column" (becomes the undo label).' },
        summary: { type: 'string', description: 'One or two sentences on what this does and why.' },
        changes: { type: 'array', items: CHANGE_SCHEMA, minItems: 1 },
      },
      required: ['title', 'summary', 'changes'],
    },
  },
];

export interface ToolOutcome {
  /** Text returned to the model. */
  content: string;
  isError?: boolean;
  /** Short line for the chat's read log. */
  log: string;
}

export interface ProposalInput {
  title: string;
  summary: string;
  changes: Change[];
}

function str(v: unknown): string | undefined {
  return typeof v === 'string' && v.trim() ? v.trim() : undefined;
}

/**
 * Like runReadTool, but a formula that needs live data (NFL stats, prices, weather) waits for the
 * download instead of answering #BUSY!.
 */
export async function runReadToolAsync(doc: Doc, name: string, input: Record<string, unknown>, maxCells: number, waitMs = 25_000): Promise<ToolOutcome> {
  let r = runReadTool(doc, name, input, maxCells);
  if (name !== 'evaluate') return r;
  const deadline = Date.now() + waitMs;
  while (r.content.includes('#BUSY!') && Date.now() < deadline) {
    await new Promise<void>((resolve) => {
      const timer = setTimeout(done, Math.max(0, deadline - Date.now()));
      const off = external.onChange(done);
      function done() {
        clearTimeout(timer);
        off();
        resolve();
      }
    });
    r = runReadTool(doc, name, input, maxCells);
  }
  return r;
}

/** Runs a read tool. propose_changes is handled by the agent (it waits for the user). */
export function runReadTool(doc: Doc, name: string, input: Record<string, unknown>, maxCells: number): ToolOutcome {
  try {
    switch (name) {
      case 'read_range': {
        const mode = (['values', 'formulas', 'both'].includes(String(input.show)) ? input.show : 'values') as ReadMode;
        const r = readRange(doc, str(input.sheet), String(input.range ?? ''), mode, maxCells);
        return { content: r.text, log: `Read ${r.label}${mode !== 'values' ? ' (formulas)' : ''} — ${r.cells.toLocaleString()} cell${r.cells === 1 ? '' : 's'}` };
      }
      case 'find_cells': {
        const q = String(input.query ?? '');
        if (!q) throw new ReadError('Give some text to search for.');
        const r = findCells(doc, q, { sheet: str(input.sheet), inFormulas: !!input.in_formulas, exact: !!input.exact });
        return { content: r.text, log: `Searched ${input.in_formulas ? 'formulas' : 'cells'}${input.sheet ? ` in ${input.sheet}` : ''} for "${q}" — ${r.hits} match${r.hits === 1 ? '' : 'es'}` };
      }
      case 'describe_sheet': {
        const text = describeSheet(doc, str(input.sheet));
        return { content: text, log: `Looked over sheet ${str(input.sheet) ?? doc.wb.activeSheet.name}` };
      }
      case 'describe_cell': {
        const text = describeCell(doc, str(input.sheet), String(input.cell ?? ''));
        return { content: text, log: `Traced ${text.split('\n')[0]}` };
      }
      case 'evaluate': {
        const f = String(input.formula ?? '').trim();
        if (!f) throw new ReadError('Give a formula to calculate.');
        const text = evaluateFormula(doc, f, str(input.sheet));
        return { content: text, log: `Calculated ${f.startsWith('=') ? f : '=' + f}` };
      }
    }
    return { content: `There's no tool called "${name}".`, isError: true, log: `Unknown tool ${name}` };
  } catch (e) {
    const msg = e instanceof ReadError ? e.message : `Error: ${(e as Error).message || String(e)}`;
    return { content: msg, isError: true, log: `${name.replace(/_/g, ' ')} failed: ${msg}` };
  }
}
