// Starter workbooks (presets): personal finance, investing, sports and planning.
// Each template is built with the same operations the UI uses, so formulas and styles are real.
import { colToName, formatCell, parseCell, parseRange, quoteSheetName } from './address';
import { currentSeason } from './data/nfl';
import { resolveTeam, TEAMS, teamInfo } from './data/nflTeams';
import { dateToSerial } from './dates';
import { Doc } from './doc';
import type { Transaction } from './history';
import { applyStyle, setCellInput } from './ops/cells';
import type { StylePatch } from './styles';
import { Sheet, Workbook, type CFRuleSpec, type ChartDef } from './workbook';

export type TemplateCategory = 'Personal finance' | 'Investing' | 'Sports' | 'Planning';

export interface TemplateInfo {
  id: string;
  name: string;
  description: string;
  /** lucide icon name, resolved by the UI. */
  icon: string;
  category: TemplateCategory;
  /** Accent color for the template's card. */
  color: string;
  /** Uses live data (needs an internet connection the first time). */
  live?: boolean;
}

export const TEMPLATE_CATEGORIES: TemplateCategory[] = ['Personal finance', 'Investing', 'Sports', 'Planning'];

export const TEMPLATES: TemplateInfo[] = [
  { id: 'budget', name: 'Monthly budget', description: 'Plan income and spending by category. Log purchases on the Transactions sheet and the budget totals them for you.', icon: 'Wallet', category: 'Personal finance', color: '#3fcf8e' },
  { id: 'loan', name: 'Loan & mortgage calculator', description: 'Monthly payment, total interest, payoff date and a full payment schedule. See how much an extra payment saves.', icon: 'House', category: 'Personal finance', color: '#ff9f43' },
  { id: 'debt', name: 'Debt payoff planner', description: 'Pay off credit cards and loans faster. Compares the avalanche and snowball methods.', icon: 'TrendingDown', category: 'Personal finance', color: '#e8618c' },
  { id: 'networth', name: 'Net worth & savings', description: 'Track what you own and owe, plan a savings goal and project long-term growth.', icon: 'PiggyBank', category: 'Personal finance', color: '#a78bfa' },
  { id: 'stocks', name: 'Stock portfolio', description: 'Your holdings with live prices, gains and allocation, a watchlist, and a year of price history.', icon: 'ChartCandlestick', category: 'Investing', color: '#4ea1ff', live: true },
  { id: 'crypto', name: 'Crypto portfolio', description: 'Live coin prices, the value of your holdings, gains and losses, and the last 7 days of prices.', icon: 'Bitcoin', category: 'Investing', color: '#ffd43b', live: true },
  { id: 'nfl', name: 'NFL team tracker', description: "Pick a team to see its schedule and results, season stats, its division race and league standings.", icon: 'Trophy', category: 'Sports', color: '#e8618c', live: true },
  { id: 'fantasy', name: 'Fantasy football', description: 'Weekly fantasy points for your roster, side-by-side player comparisons and the top players at each position.', icon: 'Medal', category: 'Sports', color: '#22b8cf', live: true },
  { id: 'travel', name: 'Trip planner', description: 'A trip budget in any currency converted at live rates, plus the weather forecast and public holidays at your destination.', icon: 'Plane', category: 'Planning', color: '#22b8cf', live: true },
  { id: 'project', name: 'Project tracker', description: 'Tasks with owners, status, priority and due dates. Overdue work is highlighted and progress is totaled for you.', icon: 'ListChecks', category: 'Planning', color: '#94d82d' },
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

function stocks(): Workbook {
  return build(['Portfolio', 'Watchlist', 'History'], (b, w, hist) => {
    b.set('A1', 'Stock Portfolio').style('A1', C.title);
    b.set('A2', 'Prices are live (about 15 minutes delayed). Change the yellow cells to your own holdings.').style('A2', C.subtitle);
    b.row('A4', ['Ticker', 'Name', 'Shares', 'Cost per share', 'Price', 'Today', 'Market value', 'Cost basis', 'Gain / loss', 'Gain %', 'Weight']).style('A4:K4', C.header);
    const holdings: [string, number, number][] = [
      ['VTI', 25, 210],
      ['VXUS', 40, 55],
      ['BND', 30, 72],
      ['AAPL', 10, 150],
      ['MSFT', 5, 300],
    ];
    const rows = 10;
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
    b.widths({ A: 80, B: 230, C: 80, D: 120, E: 100, F: 80, G: 130, H: 120, I: 120, J: 90, K: 80 });
    b.chart({ type: 'doughnut', title: 'Allocation', series: [{ name: 'Market value', categories: b.ref(`A5:A${last}`), values: b.ref(`G5:G${last}`) }], legend: 'right', at: `B${tot + 2}`, w: 460, h: 280 });
    b.sheet.freeze = { rows: 4, cols: 1 };
    b.sheet.tabColor = '#4ea1ff';
    b.select(`A${5 + holdings.length}`);

    w.set('A1', 'Watchlist').style('A1', C.title);
    w.set('A2', 'Add any ticker in column A: stocks, ETFs or indexes like ^GSPC.').style('A2', C.subtitle);
    w.row('A4', ['Ticker', 'Name', 'Price', 'Change', 'Change %', 'Day low', 'Day high', '52-week low', '52-week high', 'Below 52-week high']).style('A4:J4', C.header);
    const watch = ['SPY', 'QQQ', 'DIA', 'NVDA', 'AMZN', 'GOOGL', 'TSLA', 'META'];
    const wl = 4 + watch.length + 4;
    for (let r = 5; r <= wl; r++) {
      const t = watch[r - 5] ?? null;
      w.row(`A${r}`, [
        t,
        `=IF(A${r}="", "", STOCK(A${r}, "Name"))`,
        `=IF(A${r}="", "", STOCK(A${r}))`,
        `=IF(A${r}="", "", STOCK(A${r}, "Change"))`,
        `=IF(A${r}="", "", STOCK(A${r}, "Change %"))`,
        `=IF(A${r}="", "", STOCK(A${r}, "Day Low"))`,
        `=IF(A${r}="", "", STOCK(A${r}, "Day High"))`,
        `=IF(A${r}="", "", STOCK(A${r}, "52 Week Low"))`,
        `=IF(A${r}="", "", STOCK(A${r}, "52 Week High"))`,
        `=IF(A${r}="", "", C${r}/I${r}-1)`,
      ]);
    }
    w.style(`A5:A${wl}`, C.input).style(`C5:D${wl},F5:I${wl}`, { numFmt: '#,##0.00' }).style(`E5:E${wl},J5:J${wl}`, { numFmt: '0.00%' });
    w.cf(`D5:E${wl}`, { type: 'cellIs', operator: 'greaterThan', values: ['0'], style: { color: '#15803d' } });
    w.cf(`D5:E${wl}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: { color: '#b91c1c' } });
    w.cf(`J5:J${wl}`, { type: 'dataBar', color: '#ff6b6b', min: { type: 'min' }, max: { type: 'num', value: 0 }, gradient: false, showValue: true });
    w.widths({ A: 80, B: 230, C: 100, D: 90, E: 90, F: 100, G: 100, H: 110, I: 110, J: 140 });
    w.sheet.freeze = { rows: 4, cols: 1 };
    w.sheet.tabColor = '#22b8cf';

    hist.set('A1', 'Price History').style('A1', C.title);
    hist.row('A2', ['Ticker', 'SPY']).style('B2', { ...C.input, bold: true, hAlign: 'left' });
    hist.note('B2', 'Type any ticker to see its last year of weekly prices.');
    hist.set('A4', '=STOCKHISTORY(B2, EDATE(TODAY(), -12), TODAY(), "weekly")');
    hist.style('A4:F4', C.header);
    hist.widths({ A: 110, B: 90, C: 90, D: 90, E: 90, F: 120 });
    hist.chart({ type: 'line', title: 'Weekly close', series: [{ name: 'Close', categories: hist.ref('A5:A60'), values: hist.ref('E5:E60') }], legend: 'none', at: 'H2', w: 560, h: 320 });
    hist.sheet.tabColor = '#a78bfa';
  });
}

function crypto(prefs: TemplatePrefs): Workbook {
  const cur = prefs.currency || 'USD';
  return build(['Portfolio', 'Prices'], (b, p) => {
    b.set('A1', 'Crypto Portfolio').style('A1', C.title);
    b.row('A2', ['Currency', cur]).style('B2', { ...C.input, bold: true, hAlign: 'left' });
    b.note('B2', 'Prices are shown in this currency: USD, EUR, GBP, CAD…');
    b.set('D2', 'Prices update live. Change the yellow cells to your own coins.').style('D2', C.subtitle);
    b.row('A4', ['Coin', 'Amount', 'Cost per coin', 'Price', 'Value', 'Cost', 'Gain / loss', 'Gain %', 'Weight']).style('A4:I4', C.header);
    const coins: [string, number, number][] = [
      ['BTC', 0.05, 42000],
      ['ETH', 0.8, 2400],
      ['SOL', 6, 95],
      ['ADA', 500, 0.45],
      ['DOGE', 1500, 0.08],
    ];
    const last = 4 + coins.length + 3;
    for (let r = 5; r <= last; r++) {
      const c = coins[r - 5];
      b.row(`A${r}`, [
        c ? c[0] : null,
        c ? c[1] : null,
        c ? c[2] : null,
        `=IF(A${r}="", "", CRYPTO(A${r}, $B$2))`,
        `=IF(A${r}="", "", B${r}*D${r})`,
        `=IF(A${r}="", "", B${r}*C${r})`,
        `=IF(A${r}="", "", E${r}-F${r})`,
        `=IF(OR(A${r}="", F${r}=0), "", G${r}/F${r})`,
        `=IF(A${r}="", "", E${r}/$E$${last + 1})`,
      ]);
    }
    const tot = last + 1;
    b.row(`A${tot}`, ['Total', '', '', '', `=SUM(E5:E${last})`, `=SUM(F5:F${last})`, `=E${tot}-F${tot}`, `=IF(F${tot}=0, 0, G${tot}/F${tot})`, '']).style(`A${tot}:I${tot}`, C.total);
    b.style(`A5:C${last}`, C.input).style(`B5:B${last}`, { numFmt: '#,##0.####' }).style(`C5:G${tot}`, { numFmt: '#,##0.00' }).style(`H5:I${tot}`, { numFmt: '0.0%' });
    b.cf(`G5:H${tot}`, { type: 'cellIs', operator: 'greaterThan', values: ['0'], style: { color: '#15803d' } });
    b.cf(`G5:H${tot}`, { type: 'cellIs', operator: 'lessThan', values: ['0'], style: { color: '#b91c1c' } });
    b.widths({ A: 80, B: 110, C: 120, D: 120, E: 120, F: 120, G: 120, H: 90, I: 80 });
    b.chart({ type: 'pie', title: 'Allocation', series: [{ name: 'Value', categories: b.ref(`A5:A${last}`), values: b.ref(`E5:E${last}`) }], legend: 'right', at: `B${tot + 2}`, w: 440, h: 280 });
    b.sheet.freeze = { rows: 4, cols: 1 };
    b.sheet.tabColor = '#ffd43b';
    b.select(`A${5 + coins.length}`);

    p.set('A1', 'Last 7 Days').style('A1', C.title);
    p.set('A2', 'Daily prices for the coins in the blue row. Change them to any symbol.').style('A2', C.subtitle);
    const syms = ['BTC', 'ETH', 'SOL'];
    p.row('A4', ['Date', ...syms]).style('A4:D4', C.header);
    for (let i = 0; i < 7; i++) {
      const r = 5 + i;
      p.row(`A${r}`, [`=TODAY()-${6 - i}`, ...syms.map((_, j) => `=CRYPTO(${colToName(j + 1)}$4, Portfolio!$B$2, $A${r})`)]);
    }
    p.row('A12', ['7-day change', ...syms.map((_, j) => `=${colToName(j + 1)}11/${colToName(j + 1)}5-1`)]).style('A12:D12', C.total);
    p.style('A5:A11', { numFmt: 'ddd, mmm d', hAlign: 'left' }).style('B5:D11', { numFmt: '#,##0.00' }).style('B12:D12', { numFmt: '0.0%' });
    p.cf('B12:D12', { type: 'cellIs', operator: 'greaterThan', values: ['0'], style: GOOD });
    p.cf('B12:D12', { type: 'cellIs', operator: 'lessThan', values: ['0'], style: BAD });
    p.widths({ A: 120, B: 120, C: 120, D: 120 });
    p.chart({ type: 'line', title: '=' + p.ref('B4'), series: [{ name: '=' + p.ref('B4'), categories: p.ref('A5:A11'), values: p.ref('B5:B11') }], legend: 'none', at: 'F2', w: 480, h: 280 });
    p.sheet.tabColor = '#ff9f43';
  });
}

/** The nicknames of the teams in the same division as `team` (Chiefs → Chiefs, Broncos, Raiders, Chargers). */
function divisionOf(team: string): string[] {
  const r = resolveTeam(team);
  const info = 'abbr' in r ? teamInfo(r.abbr) : undefined;
  if (!info) return ['Chiefs', 'Broncos', 'Raiders', 'Chargers'];
  return [info.nickname, ...TEAMS.filter((t) => t.conf === info.conf && t.div === info.div && t.abbr !== info.abbr).map((t) => t.nickname)];
}

function nfl(prefs: TemplatePrefs): Workbook {
  const season = currentSeason();
  const team = prefs.team || 'Chiefs';
  return build(['My team', 'Division', 'Standings'], (t, dv, st) => {
    t.set('A1', 'NFL Team Tracker').style('A1', C.title);
    t.row('A2', ['Team', team]).row('A3', ['Season', season]).style('B2:B3', { ...C.input, hAlign: 'left', bold: true });
    t.note('B2', 'Type any team: "Lions", "DAL", "Green Bay"…');
    t.set('A5', '=NFL.SCHEDULE(B2, B3)');
    t.style('A5:G5', C.header);
    t.cf('D6:D30', { type: 'text', operator: 'beginsWith', text: 'W', style: GOOD });
    t.cf('D6:D30', { type: 'text', operator: 'beginsWith', text: 'L', style: BAD });
    t.set('I1', 'Season so far').style('I1', C.section);
    const recs = ['Record', 'Win %', 'Points For', 'Points Against', 'Point Differential', 'Offensive Yards', 'Yards Allowed', 'Turnover Differential'];
    recs.forEach((r, i) => t.row(`I${2 + i}`, [r, `=NFL.TEAM($B$2, I${2 + i}, $B$3)`]));
    t.style('J2', { bold: true, hAlign: 'right' }).style('J3', { numFmt: '0.000' }).style('J7:J8', { numFmt: '#,##0' });
    t.widths({ A: 70, B: 110, C: 220, D: 70, E: 80, F: 70, G: 90, I: 170, J: 90 });
    t.sheet.tabColor = '#e8618c';
    t.select('B2');

    const div = divisionOf(team);
    dv.set('A1', 'Division Race').style('A1', C.title);
    dv.row('A2', ['Season', "='My team'!B3"]).style('B2', { bold: true, hAlign: 'left' });
    dv.set('D2', 'Teams are in the blue column; stats are in the blue row. Change either.').style('D2', C.subtitle);
    const stats = ['Record', 'Win %', 'Points For', 'Points Against', 'Point Differential', 'Passing Yards', 'Rushing Yards', 'Def Sacks', 'Takeaways'];
    dv.row('A4', ['Team', ...stats]).style(`A4:${colToName(stats.length)}4`, C.header);
    div.forEach((name, i) => {
      const r = 5 + i;
      dv.row(`A${r}`, [name, ...stats.map((_, j) => `=NFL.TEAM($A${r}, ${colToName(j + 1)}$4, $B$2)`)]);
    });
    const dl = 4 + div.length;
    dv.style(`A5:A${dl}`, C.input).style(`C5:C${dl}`, { numFmt: '0.000' }).style(`D5:${colToName(stats.length)}${dl}`, { numFmt: '#,##0' });
    dv.cf(`F5:F${dl}`, { type: 'colorScale', stops: [{ at: { type: 'min' }, color: '#f8696b' }, { at: { type: 'percentile', value: 50 }, color: '#ffeb84' }, { at: { type: 'max' }, color: '#63be7b' }] });
    dv.widths({ A: 130, B: 80, C: 70, D: 100, E: 120, F: 140, G: 120, H: 120, I: 90, J: 100 });
    dv.chart({ type: 'column', title: 'Point differential', series: [{ name: '=' + dv.ref('F4'), categories: dv.ref(`A5:A${dl}`), values: dv.ref(`F5:F${dl}`) }], legend: 'none', at: `B${dl + 2}`, w: 480, h: 260 });
    dv.sheet.tabColor = '#ff9f43';

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
  });
}

function fantasy(prefs: TemplatePrefs): Workbook {
  const season = currentSeason();
  const fantasyStat = prefs.fantasyStat || 'Fantasy Points PPR';
  return build(['Weekly points', 'Compare players', 'Top players'], (fan, cmp, top) => {
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
    fan.set(`A${fL + 2}`, 'Put your roster in column A. Blank weeks are byes, injuries or games not played yet. Drag the last row down to add players.').style(`A${fL + 2}`, C.subtitle);
    fan.widths({ A: 160, ...Object.fromEntries(weeks.map((_, j) => [colToName(j + 1), 54])), [totCol]: 70, [avgCol]: 60 });
    fan.sheet.freeze = { rows: 5, cols: 1 };
    fan.sheet.tabColor = '#3fcf8e';
    fan.select('A6');

    cmp.set('A1', 'Compare Players').style('A1', C.title);
    cmp.row('A2', ['Season', season]).style('B2', { ...C.input, hAlign: 'left', bold: true });
    cmp.set('D2', 'Change the players or the stats in the blue row — formulas pick them up.').style('D2', C.subtitle);
    const stats = [fantasyStat, 'Games', 'Passing Yards', 'Passing TDs', 'Rushing Yards', 'Rushing TDs', 'Receptions', 'Receiving Yards', 'Receiving TDs'];
    cmp.row('A4', ['Player', ...stats]).style(`A4:${colToName(stats.length)}4`, C.header);
    const compare = ['Josh Allen', 'Lamar Jackson', 'Saquon Barkley', 'Derrick Henry', "Ja'Marr Chase", 'Justin Jefferson', 'Travis Kelce', 'George Kittle'];
    compare.forEach((name, i) => {
      const r = 5 + i;
      cmp.row(`A${r}`, [name, ...stats.map((_, j) => `=NFL.PLAYER($A${r}, ${colToName(j + 1)}$4, $B$2)`)]);
    });
    const cl = 4 + compare.length;
    cmp.style(`A5:A${cl}`, C.input).style(`B5:B${cl}`, { numFmt: '0.0', bold: true }).style(`D5:${colToName(stats.length)}${cl}`, { numFmt: '#,##0' });
    cmp.cf(`B5:B${cl}`, { type: 'dataBar', color: '#4ea1ff', min: { type: 'num', value: 0 }, max: { type: 'max' }, gradient: false, showValue: true });
    cmp.widths({ A: 170, B: 150, C: 70, D: 120, E: 100, F: 120, G: 100, H: 100, I: 130, J: 120 });
    cmp.chart({ type: 'bar', title: '=' + cmp.ref('B4'), series: [{ name: '=' + cmp.ref('B4'), categories: cmp.ref(`A5:A${cl}`), values: cmp.ref(`B5:B${cl}`) }], legend: 'none', at: `B${cl + 2}`, w: 520, h: 280 });
    cmp.sheet.freeze = { rows: 4, cols: 1 };
    cmp.sheet.tabColor = '#4ea1ff';

    top.set('A1', 'Top Players by Position').style('A1', C.title);
    top.row('A2', ['Season', "='Weekly points'!B2"]).style('B2', { bold: true, hAlign: 'left' });
    const blocks: [string, string, string][] = [
      ['A4', 'QB', 'A5'],
      ['G4', 'RB', 'G5'],
      ['A17', 'WR', 'A18'],
      ['G17', 'TE', 'G18'],
    ];
    const titles: Record<string, string> = { QB: 'Quarterbacks', RB: 'Running backs', WR: 'Wide receivers', TE: 'Tight ends' };
    for (const [label, pos, at] of blocks) {
      top.set(label, titles[pos]).style(label, C.section);
      top.set(at, `=NFL.LEADERS('Weekly points'!$B$3, $B$2, 10, "${pos}")`);
      const a = parseCell(at)!;
      top.style(`${colToName(a.col)}${a.row + 1}:${colToName(a.col + 4)}${a.row + 1}`, C.header);
    }
    top.style('E6:E15,K6:K15,E19:E28,K19:K28', { numFmt: '0.0' });
    top.widths({ A: 50, B: 170, C: 55, D: 75, E: 150, F: 24, G: 50, H: 170, I: 55, J: 75, K: 150 });
    top.sheet.tabColor = '#a78bfa';
  });
}

function travel(prefs: TemplatePrefs): Workbook {
  const home = prefs.currency || 'USD';
  const { y, m, d } = today();
  const start = dateToSerial(y, m, d) + 30;
  return build(['Trip budget', 'Weather & holidays'], (b, wx) => {
    b.set('A1', 'Trip Planner').style('A1', C.title);
    b.set('A2', 'Change the yellow cells. Costs in other currencies are converted at live exchange rates.').style('A2', C.subtitle);
    b.table('A4', [
      ['Destination', 'Lisbon, Portugal'],
      ['Country', 'Portugal'],
      ['First day', null],
      ['Nights', 7],
      ['Your currency', home],
      ['Local currency', 'EUR'],
    ]);
    setDate(b, 'B6', start);
    b.style('B4:B9', { ...C.input, hAlign: 'left' });
    b.note('B4', 'A city, or "City, Country". Used for the weather forecast.');
    b.row('D4', ['Exchange rate', '=FXRATE(B8, B9)']).style('E4', { numFmt: '0.0000' });
    b.set('F4', '=IFERROR("1 "&B8&" = "&TEXT(E4, "0.00")&" "&B9, "")').style('F4', C.subtitle);
    b.row('D5', ['Last day', '=B6+B7']).style('E5', { numFmt: DATE, hAlign: 'left' });

    b.row('A11', ['Item', 'Category', 'Cost', 'Currency', `Cost in ${home}`, 'Paid']).style('A11:F11', C.header);
    const items: [string, string, number | string, string, string][] = [
      ['Flights', 'Travel', 850, home, 'Yes'],
      ['Hotel', 'Lodging', '=140*B7', 'EUR', 'No'],
      ['Food & drinks', 'Food', '=55*B7', 'EUR', 'No'],
      ['Local transport', 'Travel', 60, 'EUR', 'No'],
      ['Museums & tours', 'Activities', 120, 'EUR', 'No'],
      ['Travel insurance', 'Other', 45, home, 'Yes'],
      ['Souvenirs', 'Shopping', 80, 'EUR', 'No'],
    ];
    const i0 = 12;
    const iL = i0 + items.length + 2;
    for (let r = i0; r <= iL; r++) {
      const it = items[r - i0];
      b.row(`A${r}`, [it ? it[0] : null, it ? it[1] : null, it ? it[2] : null, it ? it[3] : null, `=IF(C${r}="", "", IF(D${r}=$B$8, C${r}, FX(C${r}, D${r}, $B$8)))`, it ? it[4] : null]);
    }
    const tot = iL + 1;
    b.row(`A${tot}`, ['Total', '', '', '', `=SUM(E${i0}:E${iL})`, '']).style(`A${tot}:F${tot}`, C.total);
    b.row(`A${tot + 1}`, ['Per day', '', '', '', `=IF($B$7=0, "", E${tot}/($B$7+1))`]);
    b.row(`A${tot + 2}`, ['Still to pay', '', '', '', `=SUMIFS(E${i0}:E${iL}, F${i0}:F${iL}, "No")`]);
    b.style(`A${i0}:D${iL},F${i0}:F${iL}`, C.input).style(`C${i0}:C${iL}`, { numFmt: '#,##0.00' }).style(`E${i0}:E${tot + 2}`, { numFmt: home === 'USD' ? MONEY : `#,##0.00 "${home}"` });
    b.style(`E${tot}:E${tot + 2}`, { bold: true });
    b.cf(`F${i0}:F${iL}`, { type: 'text', operator: 'beginsWith', text: 'Yes', style: GOOD });
    b.widths({ A: 150, B: 140, C: 100, D: 120, E: 140, F: 70 });
    b.chart({ type: 'doughnut', title: 'Where the money goes', series: [{ name: 'Cost', categories: b.ref(`A${i0}:A${iL}`), values: b.ref(`E${i0}:E${iL}`) }], legend: 'right', at: 'H4', w: 420, h: 280 });
    b.sheet.tabColor = '#22b8cf';
    b.select('B4');

    wx.set('A1', 'Weather & Holidays').style('A1', C.title);
    wx.row('A2', ['Destination', "='Trip budget'!B4"]).style('B2', { bold: true, hAlign: 'left' });
    wx.row('A3', ['Right now', '=IFERROR(TEXT(WEATHER(B2), "0")&"°, "&WEATHER(B2, "Condition"), "")']).style('B3', { hAlign: 'left' });
    wx.set('A5', '16-day forecast').style('A5', C.section);
    wx.set('A6', '=WEATHER.FORECAST(B2, 16)');
    wx.style('A6:F6', C.header);
    wx.cf('E7:E22', { type: 'dataBar', color: '#4ea1ff', min: { type: 'num', value: 0 }, max: { type: 'num', value: 1 }, gradient: false, showValue: true });
    wx.set('H5', 'Public holidays').style('H5', C.section);
    wx.set('H6', "=HOLIDAYS('Trip budget'!B5, YEAR('Trip budget'!B6))");
    wx.style('H6:K6', C.header);
    wx.set('H4', 'Holidays during your trip are highlighted. Shops and offices may be closed.').style('H4', C.subtitle);
    wx.row('M2', ['First day', "='Trip budget'!B6", "='Trip budget'!E5"]).style('N2:O2', { numFmt: DATE });
    wx.cf('H7:K40', { type: 'expression', formula: 'AND($H7>=$N$2, $H7<=$O$2)', style: { fill: '#fef3c7', bold: true } });
    wx.widths({ A: 130, B: 120, C: 70, D: 70, E: 110, F: 140, G: 24, H: 160, I: 220, J: 200, K: 100, M: 80, N: 110, O: 110 });
    wx.sheet.tabColor = '#3fcf8e';
  });
}

function project(): Workbook {
  const { y, m, d } = today();
  const t0 = dateToSerial(y, m, d);
  return build(['Tasks'], (b) => {
    b.set('A1', 'Project Tracker').style('A1', C.title);
    b.set('A2', 'One row per task. Set Status to Not started, In progress, Blocked or Done.').style('A2', C.subtitle);
    b.row('A4', ['Task', 'Owner', 'Status', 'Priority', 'Start', 'Due', '% done', 'Days left', 'Notes']).style('A4:I4', C.header);
    const tasks: [string, string, string, string, number, number, number, string][] = [
      ['Agree on goals and scope', 'Alex', 'Done', 'High', -14, -7, 1, ''],
      ['Research options', 'Sam', 'Done', 'Medium', -10, -3, 1, ''],
      ['Draft the plan', 'Alex', 'In progress', 'High', -3, 4, 0.6, 'Waiting on numbers from Sam'],
      ['Get the budget approved', 'Jordan', 'Blocked', 'High', -4, -1, 0.2, 'Needs sign-off'],
      ['Build the first version', 'Sam', 'Not started', 'Medium', 5, 19, 0, ''],
      ['Test and review', 'Jordan', 'Not started', 'Medium', 19, 26, 0, ''],
      ['Launch', 'Alex', 'Not started', 'High', 27, 28, 0, ''],
    ];
    const r0 = 5;
    const rL = r0 + tasks.length + 5;
    for (let r = r0; r <= rL; r++) {
      const t = tasks[r - r0];
      b.row(`A${r}`, [t ? t[0] : null, t ? t[1] : null, t ? t[2] : null, t ? t[3] : null]);
      if (t) {
        setDate(b, `E${r}`, t0 + t[4]);
        setDate(b, `F${r}`, t0 + t[5]);
        b.set(`G${r}`, t[6]).set(`I${r}`, t[7] || null);
      }
      b.set(`H${r}`, `=IF(OR(F${r}="", C${r}="Done"), "", IF(F${r}=TODAY(), "Due today", IF(F${r}>TODAY(), F${r}-TODAY()&IF(F${r}-TODAY()=1, " day left", " days left"), TODAY()-F${r}&IF(TODAY()-F${r}=1, " day late", " days late"))))`);
    }
    b.style(`A${r0}:I${rL}`, { borderBottom: { style: 'thin', color: '#e5e7eb' } }).style(`E${r0}:F${rL}`, { numFmt: 'mmm d', hAlign: 'right' }).style(`G${r0}:G${rL}`, { numFmt: '0%' }).style(`H${r0}:H${rL}`, { hAlign: 'right' });
    b.cf(`H${r0}:H${rL}`, { type: 'text', operator: 'contains', text: 'late', style: { color: '#b91c1c', bold: true } });
    b.style(`C${r0}:D${rL}`, { hAlign: 'center' });
    b.cf(`C${r0}:C${rL}`, { type: 'text', operator: 'beginsWith', text: 'Done', style: GOOD });
    b.cf(`C${r0}:C${rL}`, { type: 'text', operator: 'beginsWith', text: 'Blocked', style: BAD });
    b.cf(`C${r0}:C${rL}`, { type: 'text', operator: 'beginsWith', text: 'In progress', style: { color: '#1d4ed8', fill: '#dbeafe' } });
    b.cf(`D${r0}:D${rL}`, { type: 'text', operator: 'beginsWith', text: 'High', style: { color: '#b91c1c', bold: true } });
    b.cf(`A${r0}:A${rL},F${r0}:F${rL}`, { type: 'expression', formula: `AND($F${r0}<>"", $F${r0}<TODAY(), $C${r0}<>"Done")`, style: { color: '#b91c1c', bold: true } });
    b.cf(`G${r0}:G${rL}`, { type: 'dataBar', color: '#3fcf8e', min: { type: 'num', value: 0 }, max: { type: 'num', value: 1 }, gradient: false, showValue: true });

    b.set('K4', 'Summary').style('K4', C.section);
    const rng = (c: string) => `${c}${r0}:${c}${rL}`;
    b.table('K5', [
      ['Tasks', `=COUNTA(${rng('A')})`],
      ['Done', `=COUNTIF(${rng('C')}, "Done")`],
      ['In progress', `=COUNTIF(${rng('C')}, "In progress")`],
      ['Blocked', `=COUNTIF(${rng('C')}, "Blocked")`],
      ['Not started', `=COUNTIF(${rng('C')}, "Not started")`],
      ['Overdue', `=COUNTIFS(${rng('F')}, "<"&TODAY(), ${rng('C')}, "<>Done", ${rng('A')}, "<>")`],
      ['Overall progress', `=IFERROR(AVERAGEIF(${rng('A')}, "<>", ${rng('G')}), 0)`],
      ['Next deadline', `=IFERROR(MINIFS(${rng('F')}, ${rng('C')}, "<>Done", ${rng('F')}, ">="&TODAY()), "")`],
    ]);
    b.style('L11', { numFmt: '0%', bold: true }).style('L12', { numFmt: DATE }).style('L10', { color: '#b91c1c', bold: true });
    b.widths({ A: 220, B: 90, C: 110, D: 80, E: 80, F: 80, G: 110, H: 110, I: 220, J: 24, K: 140, L: 110 });
    b.chart({ type: 'doughnut', title: 'Status', series: [{ name: 'Tasks', categories: b.ref('K6:K9'), values: b.ref('L6:L9') }], legend: 'right', at: 'K14', w: 380, h: 240 });
    b.sheet.freeze = { rows: 4, cols: 1 };
    b.sheet.tabColor = '#94d82d';
    b.select(`A${r0 + tasks.length}`);
  });
}

/** Personal defaults used by some templates. */
export interface TemplatePrefs {
  /** NFL team nickname, e.g. "Lions". */
  team?: string;
  /** Stat used for weekly fantasy points. */
  fantasyStat?: string;
  /** Currency code for prices and totals, e.g. "USD". */
  currency?: string;
}

const BUILDERS: Record<string, (prefs: TemplatePrefs) => Workbook> = { budget, loan, debt, networth, stocks, crypto, nfl, fantasy, travel, project };

export function buildTemplate(id: string, prefs: TemplatePrefs = {}): Workbook | null {
  return BUILDERS[id]?.(prefs) ?? null;
}
