// Starter workbooks: budgeting, loans, debt, savings, investing and NFL tracking.
// Each template is built with the same operations the UI uses, so formulas and styles are real.
import { colToName, formatCell, parseCell, parseRange, quoteSheetName } from './address';
import { currentSeason } from './data/nfl';
import { dateToSerial } from './dates';
import { Doc } from './doc';
import type { Transaction } from './history';
import { applyStyle, setCellInput } from './ops/cells';
import type { StylePatch } from './styles';
import { Sheet, Workbook, type CFRuleSpec, type ChartDef } from './workbook';

export interface TemplateInfo {
  id: string;
  name: string;
  description: string;
  /** lucide icon name, resolved by the UI. */
  icon: string;
  group: 'Money' | 'NFL';
}

export const TEMPLATES: TemplateInfo[] = [
  { id: 'budget', name: 'Monthly budget', description: 'Plan income and spending by category. Log purchases on the Transactions sheet and the budget totals them for you.', icon: 'Wallet', group: 'Money' },
  { id: 'loan', name: 'Loan & mortgage planner', description: 'Monthly payment, total interest, payoff date and a full payment schedule. See how much an extra payment saves.', icon: 'House', group: 'Money' },
  { id: 'debt', name: 'Debt payoff planner', description: 'Pay off credit cards and loans faster. Compares the avalanche and snowball methods.', icon: 'TrendingDown', group: 'Money' },
  { id: 'networth', name: 'Net worth & savings', description: 'Track what you own and owe, plan a savings goal and project long-term growth.', icon: 'PiggyBank', group: 'Money' },
  { id: 'portfolio', name: 'Investment portfolio', description: 'Live stock, ETF and crypto prices with gains, losses and allocation.', icon: 'ChartCandlestick', group: 'Money' },
  { id: 'nfl', name: 'NFL season tracker', description: "Your team's schedule and results, standings, a QB comparison and weekly fantasy points — all live.", icon: 'Trophy', group: 'NFL' },
];

// ---------- building blocks ----------

const MONEY = '"$"#,##0.00';
const MONEY0 = '"$"#,##0';
const PCT = '0.0%';
const DATE = 'mmm d, yyyy';

const C = {
  title: { bold: true, fontSize: 18 } as StylePatch,
  subtitle: { italic: true, color: '#6b7280' } as StylePatch,
  header: { bold: true, fill: '#dbeafe', borderBottom: { style: 'thin', color: '#93c5fd' } } as StylePatch,
  section: { bold: true, fontSize: 13 } as StylePatch,
  input: { fill: '#fef3c7' } as StylePatch,
  total: { bold: true, borderTop: { style: 'thin' } } as StylePatch,
  big: { bold: true, fontSize: 14 } as StylePatch,
};
const GOOD = { color: '#15803d', fill: '#dcfce7' };
const BAD = { color: '#b91c1c', fill: '#fee2e2' };

class SheetBuilder {
  constructor(
    private doc: Doc,
    private tx: Transaction,
    readonly sheet: Sheet,
  ) {}

  /** Writes a value or formula, exactly as if typed. */
  set(addr: string, raw: string | number | null): this {
    const a = parseCell(addr)!;
    if (raw !== null) setCellInput(this.tx, this.doc, this.sheet, a.row, a.col, String(raw));
    return this;
  }

  /** Writes values left to right starting at addr. */
  row(addr: string, values: (string | number | null)[]): this {
    const a = parseCell(addr)!;
    values.forEach((v, i) => this.set(formatCell(a.row, a.col + i), v));
    return this;
  }

  /** Writes rows of values starting at addr. */
  table(addr: string, rows: (string | number | null)[][]): this {
    const a = parseCell(addr)!;
    rows.forEach((r, i) => this.row(formatCell(a.row + i, a.col), r));
    return this;
  }

  style(ref: string, patch: StylePatch): this {
    applyStyle(this.tx, this.doc.wb, this.sheet, ref.split(',').map((r) => parseRange(r.trim())!), patch);
    return this;
  }

  widths(w: Record<string, number>): this {
    for (const [col, px] of Object.entries(w)) this.sheet.colWidths.set(parseCell(col + '1')!.col, px);
    return this;
  }

  cf(ref: string, spec: CFRuleSpec): this {
    this.sheet.cf.push({ ...spec, id: 'cf' + Math.random().toString(36).slice(2, 9), ranges: ref.split(',').map((r) => parseRange(r.trim())!) });
    return this;
  }

  /** An absolute reference to a range on this sheet, for chart series. */
  ref(range: string): string {
    const r = parseRange(range)!;
    const a = formatCell(r.r1, r.c1, true, true);
    const b = formatCell(r.r2, r.c2, true, true);
    return `${quoteSheetName(this.sheet.name)}!${a === b ? a : `${a}:${b}`}`;
  }

  chart(def: Omit<ChartDef, 'id' | 'legend' | 'anchor'> & { legend?: ChartDef['legend']; at: string; w?: number; h?: number }): this {
    const a = parseCell(def.at)!;
    const { at: _at, w, h, ...rest } = def;
    void _at;
    this.sheet.charts.push({ legend: 'bottom', ...rest, id: 'chart-' + Math.random().toString(36).slice(2, 10), anchor: { row: a.row, col: a.col, dx: 8, dy: 8, w: w ?? 480, h: h ?? 300 } });
    return this;
  }

  note(addr: string, text: string): this {
    const a = parseCell(addr)!;
    const cell = this.sheet.getCell(a.row, a.col) ?? {};
    this.tx.setCell(this.sheet, a.row, a.col, { ...cell, note: text });
    return this;
  }

  select(addr: string): this {
    const a = parseCell(addr)!;
    this.sheet.view.activeRow = a.row;
    this.sheet.view.activeCol = a.col;
    return this;
  }
}

function build(sheetNames: string[], fn: (...sheets: SheetBuilder[]) => void): Workbook {
  const wb = new Workbook(true);
  for (const n of sheetNames) wb.sheets.push(new Sheet(n));
  wb.activeSheetId = wb.sheets[0].id;
  const doc = new Doc(wb);
  try {
    doc.transact('Template', (tx) => fn(...wb.sheets.map((s) => new SheetBuilder(doc, tx, s))));
  } finally {
    doc.dispose();
  }
  return wb;
}

function today(): { y: number; m: number; d: number } {
  const t = new Date();
  return { y: t.getFullYear(), m: t.getMonth() + 1, d: t.getDate() };
}

// ---------- templates ----------

function budget(): Workbook {
  const { y, m } = today();
  return build(['Budget', 'Transactions'], (b, t) => {
    const income: [string, number][] = [
      ['Paycheck', 4200],
      ['Side income', 400],
      ['Other income', 0],
    ];
    const expenses: [string, number][] = [
      ['Rent / mortgage', 1500],
      ['Utilities', 220],
      ['Groceries', 550],
      ['Transportation', 260],
      ['Insurance', 180],
      ['Phone & internet', 120],
      ['Dining out', 220],
      ['Entertainment', 120],
      ['Subscriptions', 60],
      ['Health', 100],
      ['Debt payments', 350],
      ['Savings & investing', 600],
      ['Personal & shopping', 150],
      ['Miscellaneous', 100],
    ];
    // Transactions log
    const tr = 'Transactions';
    const sumifs = (catCell: string) => `=SUMIFS(${tr}!$D$2:$D$5000, ${tr}!$C$2:$C$5000, ${catCell}, ${tr}!$A$2:$A$5000, ">="&$B$2, ${tr}!$A$2:$A$5000, "<="&EOMONTH($B$2, 0))`;
    t.row('A1', ['Date', 'Description', 'Category', 'Amount']).style('A1:D1', C.header);
    const sample: [number, string, string, number][] = [
      [1, 'Paycheck', 'Paycheck', 2100],
      [1, 'Rent', 'Rent / mortgage', 1500],
      [15, 'Paycheck', 'Paycheck', 2100],
      [10, 'Freelance project', 'Side income', 400],
      [2, 'Grocery store', 'Groceries', 132.48],
      [3, 'Gas', 'Transportation', 48.2],
      [4, 'Streaming service', 'Subscriptions', 15.99],
      [5, 'Electric bill', 'Utilities', 96.4],
      [6, 'Lunch with friends', 'Dining out', 38.75],
      [7, 'Transfer to savings', 'Savings & investing', 300],
      [8, 'Grocery store', 'Groceries', 87.1],
    ];
    sample.forEach(([d, desc, cat, amt], i) => {
      const r = i + 2;
      t.set(`A${r}`, `${y}-${String(m).padStart(2, '0')}-${String(Math.min(d, 28)).padStart(2, '0')}`).set(`B${r}`, desc).set(`C${r}`, cat).set(`D${r}`, amt);
    });
    t.style('A2:A5000', { numFmt: DATE }).style('D2:D5000', { numFmt: MONEY });
    t.widths({ A: 110, B: 220, C: 170, D: 110, F: 360 });
    t.set('F1', 'How this works').style('F1', C.section);
    t.set('F2', 'Add every paycheck and purchase here. Use the exact category names from the Budget sheet.');
    t.set('F3', 'The Budget sheet adds them up for the month shown at the top of the Budget sheet.');
    t.style('F2:F3', { wrap: true, color: '#6b7280' });
    t.sheet.freeze = { rows: 1, cols: 0 };
    t.sheet.tabColor = '#3fcf8e';

    b.set('A1', 'Monthly Budget').style('A1', C.title);
    b.set('A2', 'Month').set('B2', `=DATE(YEAR(TODAY()), MONTH(TODAY()), 1)`).style('B2', { ...C.input, numFmt: 'mmmm yyyy', bold: true, hAlign: 'left' });
    b.note('B2', 'Type any date in the month you want to see, like 3/1/2026.');
    b.row('A4', ['Income', 'Planned', 'Actual', 'Difference']).style('A4:D4', C.header);
    const i0 = 5;
    income.forEach(([name, amt], i) => {
      const r = i0 + i;
      b.row(`A${r}`, [name, amt, sumifs(`A${r}`), `=C${r}-B${r}`]);
    });
    const iT = i0 + income.length;
    b.row(`A${iT}`, ['Total income', `=SUM(B${i0}:B${iT - 1})`, `=SUM(C${i0}:C${iT - 1})`, `=C${iT}-B${iT}`]).style(`A${iT}:D${iT}`, C.total);
    b.style(`B${i0}:B${iT - 1}`, C.input);

    const e0 = iT + 3;
    b.row(`A${e0 - 1}`, ['Expenses', 'Budget', 'Actual', 'Left to spend', '% of income']).style(`A${e0 - 1}:E${e0 - 1}`, C.header);
    expenses.forEach(([name, amt], i) => {
      const r = e0 + i;
      b.row(`A${r}`, [name, amt, sumifs(`A${r}`), `=B${r}-C${r}`, `=IF($B$${iT}=0, 0, B${r}/$B$${iT})`]);
    });
    const eT = e0 + expenses.length;
    b.row(`A${eT}`, ['Total expenses', `=SUM(B${e0}:B${eT - 1})`, `=SUM(C${e0}:C${eT - 1})`, `=B${eT}-C${eT}`, `=IF($B$${iT}=0, 0, B${eT}/$B$${iT})`]).style(`A${eT}:E${eT}`, C.total);
    b.style(`B${e0}:B${eT - 1}`, C.input);
    b.style(`B${i0}:D${eT}`, { numFmt: MONEY }).style(`E${e0}:E${eT}`, { numFmt: PCT });

    const sR = eT + 2;
    const savingsRow = e0 + expenses.findIndex(([n]) => n.startsWith('Savings'));
    b.set(`A${sR}`, 'Summary').style(`A${sR}`, C.section);
    b.row(`A${sR + 1}`, ['', 'Planned', 'Actual']).style(`B${sR + 1}:C${sR + 1}`, { bold: true });
    b.row(`A${sR + 2}`, ['Left over', `=B${iT}-B${eT}`, `=C${iT}-C${eT}`]);
    b.row(`A${sR + 3}`, ['Savings rate', `=IF(B${iT}=0, 0, (B${savingsRow}+B${sR + 2})/B${iT})`, `=IF(C${iT}=0, 0, (C${savingsRow}+C${sR + 2})/C${iT})`]);
    b.style(`B${sR + 2}:C${sR + 2}`, { ...C.big, numFmt: MONEY }).style(`B${sR + 3}:C${sR + 3}`, { bold: true, numFmt: PCT });
    b.note(`A${sR + 3}`, 'Money saved or invested plus whatever is left over, as a share of income.');

    b.cf(`D${e0}:D${eT}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: BAD });
    b.cf(`D${i0}:D${iT}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: BAD });
    b.cf(`B${sR + 2}:C${sR + 2}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: BAD });
    b.cf(`C${e0}:C${eT - 1}`, { type: 'dataBar', color: '#4ea1ff', min: { type: 'num', value: 0 }, max: { type: 'max' }, gradient: false, showValue: true });

    b.widths({ A: 190, B: 110, C: 110, D: 110, E: 100 });
    b.chart({
      type: 'bar',
      title: 'Budget vs. actual',
      series: [
        { name: '=' + b.ref(`B${e0 - 1}`), categories: b.ref(`A${e0}:A${eT - 1}`), values: b.ref(`B${e0}:B${eT - 1}`) },
        { name: '=' + b.ref(`C${e0 - 1}`), categories: b.ref(`A${e0}:A${eT - 1}`), values: b.ref(`C${e0}:C${eT - 1}`) },
      ],
      at: 'G4',
      w: 520,
      h: 420,
    });
    b.sheet.tabColor = '#4ea1ff';
    b.select(`B${i0}`);
  });
}

function loan(): Workbook {
  const { y, m } = today();
  const next = m === 12 ? [y + 1, 1] : [y, m + 1];
  return build(['Loan'], (b) => {
    b.set('A1', 'Loan & Mortgage Planner').style('A1', C.title);
    b.set('A2', 'Change the yellow cells — everything else updates.').style('A2', C.subtitle);
    b.table('A4', [
      ['Loan amount', 350000],
      ['Interest rate (yearly)', '6.5%'],
      ['Length (years)', 30],
      ['First payment', null],
      ['Extra payment each month', 0],
    ]);
    setDate(b, 'B7', dateToSerial(next[0], next[1], 1));
    b.style('B4:B8', C.input).style('B4', { numFmt: MONEY0 }).style('B5', { numFmt: '0.00%' }).style('B8', { numFmt: MONEY0 });

    b.table('D4', [
      ['Monthly payment', '=PMT(B5/12, B6*12, -B4)'],
      ['Paid off in (months)', '=PAYOFFMONTHS(B5, B4, E4+B8)'],
      ['Payoff date', '=EDATE(B7, E5-1)'],
      ['Total interest', '=PAYOFFINTEREST(B5, B4, E4+B8)'],
      ['Saved by extra payment', '=PAYOFFINTEREST(B5, B4, E4)-E7'],
      ['Total paid', '=B4+E7'],
    ]);
    b.style('E4', { ...C.big, numFmt: MONEY }).style('E5', { numFmt: '0' }).style('E6', { numFmt: 'mmmm yyyy' }).style('E7:E9', { numFmt: MONEY });
    b.cf('E8', { type: 'cellIs', operator: 'greaterThan', values: ['0'], style: GOOD });

    b.set('A11', 'Payment schedule').style('A11', C.section);
    b.set('A12', '=AMORTIZATION(B5, B6, B4, B7, B8)');
    b.style('A12:G12', C.header);
    b.widths({ A: 190, B: 120, C: 110, D: 170, E: 120, F: 100, G: 130 });
    b.chart({
      type: 'line',
      title: 'Balance over time',
      series: [{ name: 'Balance', categories: b.ref('B13:B1213'), values: b.ref('G13:G1213') }],
      legend: 'none',
      at: 'I3',
      w: 520,
      h: 300,
    });
    b.sheet.tabColor = '#ff9f43';
    b.select('B4');
  });
}

function setDate(b: SheetBuilder, addr: string, serial: number): void {
  b.set(addr, serial).style(addr, { numFmt: DATE, hAlign: 'left' });
}

function debt(): Workbook {
  return build(['Debt payoff'], (b) => {
    b.set('A1', 'Debt Payoff Planner').style('A1', C.title);
    b.set('A2', 'List your debts, set a monthly budget and see when you\'ll be debt-free.').style('A2', C.subtitle);
    b.row('A4', ['Debt', 'Balance', 'Interest rate', 'Minimum payment']).style('A4:D4', C.header);
    b.table('A5', [
      ['Credit card', 5200, '24.99%', 150],
      ['Store card', 1200, '19.99%', 45],
      ['Personal loan', 3000, '11.5%', 95],
      ['Car loan', 14500, '6.9%', 330],
      ['Student loan', 22000, '5.5%', 250],
      [null, null, null, null],
    ]);
    b.style('A5:D10', C.input).style('B5:B11,D5:D11', { numFmt: MONEY }).style('C5:C10', { numFmt: '0.00%' });
    b.row('A11', ['Total', '=SUM(B5:B10)', '', '=SUM(D5:D10)']).style('A11:D11', C.total);

    b.row('A13', ['Monthly budget for debt', 1100]).style('B13', { ...C.input, numFmt: MONEY });
    b.row('A14', ['Method', 'avalanche']).style('B14', C.input);
    b.note('B14', 'avalanche: pay the highest interest rate first (saves the most money).\nsnowball: pay the smallest balance first (quick wins).');

    // DEBTPAYOFF skips blank rows by reading only filled ones
    const args = 'FILTER(A5:A10, B5:B10<>""), FILTER(B5:B10, B5:B10<>""), FILTER(C5:C10, B5:B10<>""), FILTER(D5:D10, B5:B10<>"")';
    b.set('A16', 'Your plan').style('A16', C.section);
    b.set('A17', `=DEBTPAYOFF(${args}, B13, B14)`);
    b.style('A17:C17', C.header);

    b.set('F4', 'Compare methods').style('F4', C.section);
    b.row('F5', ['', 'Months', 'Interest']).style('F5:H5', C.header);
    b.row('F6', ['Avalanche', `=INDEX(TAKE(DEBTPAYOFF(${args}, B13, "avalanche"), -1), 1, 2)`, `=INDEX(TAKE(DEBTPAYOFF(${args}, B13, "avalanche"), -1), 1, 3)`]);
    b.row('F7', ['Snowball', `=INDEX(TAKE(DEBTPAYOFF(${args}, B13, "snowball"), -1), 1, 2)`, `=INDEX(TAKE(DEBTPAYOFF(${args}, B13, "snowball"), -1), 1, 3)`]);
    b.row('F8', ['Minimums only', '=MAX(IFERROR(PAYOFFMONTHS(FILTER(C5:C10, B5:B10<>""), FILTER(B5:B10, B5:B10<>""), FILTER(D5:D10, B5:B10<>"")), 1200))', '=SUM(IFERROR(PAYOFFINTEREST(FILTER(C5:C10, B5:B10<>""), FILTER(B5:B10, B5:B10<>""), FILTER(D5:D10, B5:B10<>"")), 0))']);
    b.style('G6:G8', { numFmt: '0' }).style('H6:H8', { numFmt: MONEY });
    b.note('F8', 'Paying only the minimum on every debt, for comparison.');
    b.cf('H6:H8', { type: 'top', rank: 1, percent: false, bottom: true, style: GOOD });

    b.widths({ A: 190, B: 170, C: 110, D: 140, F: 130, G: 90, H: 120 });
    b.sheet.tabColor = '#e8618c';
    b.select('A5');
  });
}

function networth(): Workbook {
  return build(['Net worth', 'Savings'], (n, s) => {
    n.set('A1', 'Net Worth').style('A1', C.title);
    n.row('A2', ['As of', '=TODAY()']).style('B2', { numFmt: DATE, hAlign: 'left' });
    n.row('A4', ['Assets', 'Value']).style('A4:B4', C.header);
    const assets: [string, number][] = [
      ['Checking', 3500],
      ['Savings', 12000],
      ['Retirement (401k)', 45000],
      ['IRA', 18000],
      ['Brokerage', 9500],
      ['Home', 0],
      ['Car', 14000],
      ['Other', 0],
    ];
    const debts: [string, number][] = [
      ['Mortgage', 0],
      ['Car loan', 9800],
      ['Student loans', 21000],
      ['Credit cards', 1800],
      ['Other', 0],
    ];
    n.table('A5', assets);
    const aT = 5 + assets.length;
    n.row(`A${aT}`, ['Total assets', `=SUM(B5:B${aT - 1})`]).style(`A${aT}:B${aT}`, C.total);
    n.row('D4', ['Debts', 'Owed']).style('D4:E4', C.header);
    n.table('D5', debts);
    n.row(`D${aT}`, ['Total debts', `=SUM(E5:E${aT - 1})`]).style(`D${aT}:E${aT}`, C.total);
    n.style(`B5:B${aT - 1},E5:E${aT - 1}`, C.input).style(`B5:B${aT},E5:E${aT}`, { numFmt: MONEY0 });
    n.row(`A${aT + 2}`, ['Net worth', `=B${aT}-E${aT}`]).style(`A${aT + 2}:B${aT + 2}`, { ...C.big, numFmt: MONEY0 });
    n.row(`A${aT + 3}`, ['Debt-to-asset ratio', `=IF(B${aT}=0, 0, E${aT}/B${aT})`]).style(`B${aT + 3}`, { numFmt: PCT });
    n.note(`A${aT + 3}`, 'Below 50% is generally considered healthy; lower is better.');
    n.cf(`B${aT + 2}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: BAD });
    n.cf(`B${aT + 2}`, { type: 'cellIs', operator: 'greaterThanOrEqual', values: ['0'], style: GOOD });

    n.set(`A${aT + 5}`, 'Emergency fund').style(`A${aT + 5}`, C.section);
    n.row(`A${aT + 6}`, ['Monthly living expenses', 3200]).style(`B${aT + 6}`, { ...C.input, numFmt: MONEY0 });
    n.row(`A${aT + 7}`, ['Months covered by savings', `=IF(B${aT + 6}=0, 0, (B5+B6)/B${aT + 6})`]).style(`B${aT + 7}`, { numFmt: '0.0' });
    n.note(`A${aT + 7}`, 'Checking plus savings divided by monthly expenses. 3–6 months is a common target.');
    n.cf(`B${aT + 7}`, { type: 'cellIs', operator: 'lessThan', values: ['3'], style: BAD });
    n.cf(`B${aT + 7}`, { type: 'cellIs', operator: 'greaterThanOrEqual', values: ['6'], style: GOOD });

    n.widths({ A: 190, B: 120, C: 24, D: 150, E: 120 });
    n.chart({ type: 'doughnut', title: 'Where your money is', series: [{ name: 'Assets', categories: n.ref(`A5:A${aT - 1}`), values: n.ref(`B5:B${aT - 1}`) }], legend: 'right', at: 'G3', w: 440, h: 300 });
    n.sheet.tabColor = '#3fcf8e';
    n.select('B5');

    s.set('A1', 'Savings Planner').style('A1', C.title);
    s.set('A3', 'Savings goal').style('A3', C.section);
    s.table('A4', [
      ['Goal amount', 25000],
      ['Years to reach it', 3],
      ['Interest rate (yearly)', '4.5%'],
      ['Already saved', 2000],
    ]);
    s.row('A8', ['Save each month', '=SAVINGSGOAL(B4, B5, B6, B7)']).style('A8:B8', C.big);
    s.style('B4:B7', C.input).style('B4,B7,B8', { numFmt: MONEY }).style('B6', { numFmt: '0.00%' });

    s.set('D3', 'Long-term growth').style('D3', C.section);
    s.table('D4', [
      ['Starting balance', 10000],
      ['Added each month', 500],
      ['Expected return (yearly)', '7%'],
      ['Years', 30],
    ]);
    s.row('D8', ['Ending balance', '=INDEX(TAKE(SAVINGS.PROJECTION(E4, E5, E6, E7), -1), 1, 4)']).style('D8:E8', C.big);
    s.style('E4:E7', C.input).style('E4:E5,E8', { numFmt: MONEY0 }).style('E6', { numFmt: '0.0%' });

    s.set('A10', '=SAVINGS.PROJECTION(E4, E5, E6, E7)');
    s.style('A10:D10', C.header);
    s.widths({ A: 170, B: 120, C: 24, D: 190, E: 120 });
    s.chart({
      type: 'column',
      stacked: true,
      title: 'Growth over time',
      series: [
        { name: 'Contributed', categories: s.ref('A11:A110'), values: s.ref('B11:B110') },
        { name: 'Growth', categories: s.ref('A11:A110'), values: s.ref('C11:C110') },
      ],
      at: 'G3',
      w: 520,
      h: 320,
    });
    s.sheet.tabColor = '#a78bfa';
  });
}

function portfolio(): Workbook {
  return build(['Portfolio'], (b) => {
    b.set('A1', 'Investment Portfolio').style('A1', C.title);
    b.set('A2', 'Prices are live (about 15 minutes delayed). Press Ctrl+Alt+F5 to refresh.').style('A2', C.subtitle);
    b.row('A4', ['Ticker', 'Name', 'Shares', 'Cost per share', 'Price', 'Today', 'Market value', 'Cost basis', 'Gain / loss', 'Gain %', 'Weight']).style('A4:K4', C.header);
    const holdings: [string, number, number][] = [
      ['VTI', 25, 210],
      ['VXUS', 40, 55],
      ['BND', 30, 72],
      ['AAPL', 10, 150],
      ['MSFT', 5, 300],
    ];
    const rows = 8;
    holdings.forEach(([t, sh, cost], i) => b.row(`A${5 + i}`, [t, null, sh, cost]));
    const last = 5 + rows - 1;
    for (let r = 5; r <= last; r++) {
      b.row(`B${r}`, [`=IF(A${r}="", "", STOCK(A${r}, "Name"))`]);
      b.row(`E${r}`, [
        `=IF(A${r}="", "", STOCK(A${r}))`,
        `=IF(A${r}="", "", STOCK(A${r}, "Change %"))`,
        `=IF(A${r}="", "", C${r}*E${r})`,
        `=IF(A${r}="", "", C${r}*D${r})`,
        `=IF(A${r}="", "", G${r}-H${r})`,
        `=IF(OR(A${r}="", H${r}=0), "", I${r}/H${r})`,
        `=IF(A${r}="", "", G${r}/$G$${last + 1})`,
      ]);
    }
    const tot = last + 1;
    b.row(`A${tot}`, ['Total', '', '', '', '', '', `=SUM(G5:G${last})`, `=SUM(H5:H${last})`, `=G${tot}-H${tot}`, `=IF(H${tot}=0, 0, I${tot}/H${tot})`, '']).style(`A${tot}:K${tot}`, C.total);
    b.style(`A5:A${last},C5:D${last}`, C.input);
    b.style(`D5:E${last},G5:I${tot}`, { numFmt: MONEY }).style(`F5:F${last},J5:K${tot}`, { numFmt: '0.00%' });
    b.cf(`F5:F${last},I5:J${tot}`, { type: 'cellIs', operator: 'greaterThan', values: ['0'], style: { color: '#15803d' } });
    b.cf(`F5:F${last},I5:J${tot}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: { color: '#b91c1c' } });

    const c0 = tot + 3;
    b.set(`A${c0 - 1}`, 'Crypto').style(`A${c0 - 1}`, C.section);
    b.row(`A${c0}`, ['Coin', 'Amount', 'Cost per coin', 'Price', 'Value', 'Gain / loss']).style(`A${c0}:F${c0}`, C.header);
    const coins: [string, number, number][] = [
      ['BTC', 0.05, 42000],
      ['ETH', 0.8, 2400],
    ];
    coins.forEach(([sym, amt, cost], i) => {
      const r = c0 + 1 + i;
      b.row(`A${r}`, [sym, amt, cost, `=IF(A${r}="", "", CRYPTO(A${r}))`, `=IF(A${r}="", "", B${r}*D${r})`, `=IF(A${r}="", "", E${r}-B${r}*C${r})`]);
    });
    const cl = c0 + coins.length + 1;
    b.row(`A${cl}`, ['', '', '', '', `=IF(A${cl}="", "", B${cl}*D${cl})`, `=IF(A${cl}="", "", E${cl}-B${cl}*C${cl})`]);
    b.row(`D${cl}`, [`=IF(A${cl}="", "", CRYPTO(A${cl}))`]);
    b.style(`A${c0 + 1}:C${cl}`, C.input).style(`C${c0 + 1}:F${cl}`, { numFmt: MONEY }).style(`B${c0 + 1}:B${cl}`, { numFmt: '0.0000' });
    b.cf(`F${c0 + 1}:F${cl}`, { type: 'cellIs', operator: 'greaterThan', values: ['0'], style: { color: '#15803d' } });
    b.cf(`F${c0 + 1}:F${cl}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: { color: '#b91c1c' } });

    b.row(`A${cl + 2}`, ['Everything', '', '', '', `=G${tot}+SUM(E${c0 + 1}:E${cl})`]).style(`A${cl + 2}:E${cl + 2}`, { ...C.big, numFmt: MONEY });

    b.widths({ A: 90, B: 220, C: 100, D: 120, E: 110, F: 90, G: 130, H: 120, I: 120, J: 90, K: 90 });
    b.chart({ type: 'pie', title: 'Allocation', series: [{ name: 'Market value', categories: b.ref(`A5:A${last}`), values: b.ref(`G5:G${last}`) }], legend: 'right', at: `H${c0 - 1}`, w: 420, h: 280 });
    b.sheet.freeze = { rows: 4, cols: 1 };
    b.sheet.tabColor = '#ffd43b';
    b.select(`A${5 + holdings.length}`);
  });
}

function nfl(prefs: TemplatePrefs): Workbook {
  const season = currentSeason();
  const team = prefs.team || 'Chiefs';
  const fantasyStat = prefs.fantasyStat || 'Fantasy Points PPR';
  return build(['My team', 'Standings', 'QB compare', 'Fantasy'], (t, st, qb, fan) => {
    t.set('A1', 'NFL Season Tracker').style('A1', C.title);
    t.row('A2', ['Team', team]).row('A3', ['Season', season]).style('B2:B3', { ...C.input, hAlign: 'left', bold: true });
    t.note('B2', 'Type any team: "Lions", "DAL", "Green Bay"…');
    t.set('A5', '=NFL.SCHEDULE(B2, B3)');
    t.style('A5:G5', C.header);
    t.cf('D6:D30', { type: 'text', operator: 'beginsWith', text: 'W', style: GOOD });
    t.cf('D6:D30', { type: 'text', operator: 'beginsWith', text: 'L', style: BAD });
    t.set('I1', 'Season so far').style('I1', C.section);
    const recs = ['Record', 'Win %', 'Points For', 'Points Against', 'Point Differential', 'Turnover Differential'];
    recs.forEach((r, i) => t.row(`I${2 + i}`, [r, `=NFL.TEAM($B$2, I${2 + i}, $B$3)`]));
    t.style('J2', { bold: true, hAlign: 'right' }).style('J3', { numFmt: '0.000' });
    t.widths({ A: 70, B: 110, C: 220, D: 70, E: 80, F: 70, G: 90, I: 170, J: 90 });
    t.sheet.tabColor = '#e8618c';
    t.select('B2');

    st.set('A1', 'Standings').style('A1', C.title);
    st.row('A2', ['Season', "='My team'!B3"]).style('B2', { bold: true, hAlign: 'left' });
    st.set('A4', 'AFC').style('A4', C.section);
    st.set('A5', '=NFL.STANDINGS($B$2, "AFC")');
    st.set('A23', 'NFC').style('A23', C.section);
    st.set('A24', '=NFL.STANDINGS($B$2, "NFC")');
    st.style('A5:J5,A24:J24', C.header);
    st.cf('I6:I21,I25:I40', { type: 'colorScale', stops: [{ at: { type: 'min' }, color: '#f8696b' }, { at: { type: 'percentile', value: 50 }, color: '#ffffff' }, { at: { type: 'max' }, color: '#63be7b' }] });
    st.widths({ A: 60, B: 200, F: 70, J: 110 });
    st.sheet.tabColor = '#4ea1ff';

    qb.set('A1', 'Quarterback Comparison').style('A1', C.title);
    qb.row('A2', ['Season', season]).style('B2', { ...C.input, hAlign: 'left', bold: true });
    qb.set('D2', 'Change the names or stats in the blue row — formulas pick them up.').style('D2', C.subtitle);
    const qbStats = ['Passing Yards', 'Passing TDs', 'Interceptions Thrown', 'Completion %', 'Passer Rating', 'Rushing Yards', 'Rushing TDs', 'Fantasy Points'];
    qb.row('A4', ['Player', ...qbStats]).style(`A4:${colToName(qbStats.length)}4`, C.header);
    const qbs = ['Patrick Mahomes', 'Josh Allen', 'Lamar Jackson', 'Joe Burrow', 'Jalen Hurts', 'Jayden Daniels', 'C.J. Stroud', 'Jared Goff'];
    qbs.forEach((name, i) => {
      const r = 5 + i;
      qb.row(`A${r}`, [name, ...qbStats.map((_, j) => `=NFL.PLAYER($A${r}, ${colToName(j + 1)}$4, $B$2)`)]);
    });
    const qL = 4 + qbs.length;
    qb.style(`A5:A${qL}`, C.input);
    qb.cf(`F5:F${qL}`, { type: 'colorScale', stops: [{ at: { type: 'min' }, color: '#f8696b' }, { at: { type: 'percentile', value: 50 }, color: '#ffeb84' }, { at: { type: 'max' }, color: '#63be7b' }] });
    qb.cf(`B5:B${qL}`, { type: 'dataBar', color: '#4ea1ff', min: { type: 'num', value: 0 }, max: { type: 'max' }, gradient: false, showValue: true });
    qb.widths({ A: 170, B: 120, C: 100, D: 150, E: 120, F: 110, G: 120, H: 110, I: 120 });
    qb.chart({ type: 'bar', title: 'Passing TDs', series: [{ name: '=' + qb.ref('C4'), categories: qb.ref(`A5:A${qL}`), values: qb.ref(`C5:C${qL}`) }], legend: 'none', at: `B${qL + 2}`, w: 520, h: 280 });
    qb.sheet.freeze = { rows: 4, cols: 1 };
    qb.sheet.tabColor = '#ff9f43';

    fan.set('A1', 'Fantasy Points by Week').style('A1', C.title);
    fan.row('A2', ['Season', season]).row('A3', ['Scoring', fantasyStat]).style('B2:B3', { ...C.input, hAlign: 'left', bold: true });
    fan.note('B3', 'Use "Fantasy Points" (standard), "Fantasy Points Half PPR" or "Fantasy Points PPR".');
    const weeks = Array.from({ length: 18 }, (_, i) => i + 1);
    const lastWeekCol = colToName(weeks.length);
    const totCol = colToName(weeks.length + 1);
    const avgCol = colToName(weeks.length + 2);
    fan.row('A5', ['Player', ...weeks, 'Total', 'Avg']).style(`A5:${avgCol}5`, C.header).style(`B5:${lastWeekCol}5`, { hAlign: 'center', numFmt: '"Wk "0' });
    const players = ['Josh Allen', 'Lamar Jackson', 'Bijan Robinson', 'Jahmyr Gibbs', "Ja'Marr Chase", 'Justin Jefferson', 'CeeDee Lamb', 'Puka Nacua', 'Brock Bowers'];
    players.forEach((p, i) => {
      const r = 6 + i;
      fan.row(`A${r}`, [p, ...weeks.map((_, j) => `=IFNA(NFL.PLAYER($A${r}, $B$3, $B$2, ${colToName(j + 1)}$5), "")`), `=SUM(B${r}:${lastWeekCol}${r})`, `=IFERROR(AVERAGE(B${r}:${lastWeekCol}${r}), "")`]);
    });
    const fL = 5 + players.length;
    fan.style(`A6:A${fL}`, C.input).style(`B6:${avgCol}${fL}`, { numFmt: '0.0' }).style(`${totCol}6:${avgCol}${fL}`, { bold: true });
    fan.cf(`B6:${lastWeekCol}${fL}`, { type: 'colorScale', stops: [{ at: { type: 'min' }, color: '#f8696b' }, { at: { type: 'percentile', value: 50 }, color: '#ffeb84' }, { at: { type: 'max' }, color: '#63be7b' }] });
    fan.set(`A${fL + 2}`, 'Blank weeks are byes, injuries or games not played yet. Drag the last row down to add players.').style(`A${fL + 2}`, C.subtitle);
    fan.widths({ A: 160, ...Object.fromEntries(weeks.map((_, j) => [colToName(j + 1), 54])), [totCol]: 70, [avgCol]: 60 });
    fan.sheet.freeze = { rows: 5, cols: 1 };
    fan.sheet.tabColor = '#3fcf8e';
  });
}

/** Personal defaults used by some templates. */
export interface TemplatePrefs {
  /** NFL team nickname, e.g. "Lions". */
  team?: string;
  /** Stat used for weekly fantasy points. */
  fantasyStat?: string;
}

const BUILDERS: Record<string, (prefs: TemplatePrefs) => Workbook> = { budget, loan, debt, networth, portfolio, nfl };

export function buildTemplate(id: string, prefs: TemplatePrefs = {}): Workbook | null {
  return BUILDERS[id]?.(prefs) ?? null;
}
