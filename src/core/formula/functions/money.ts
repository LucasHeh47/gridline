// Personal finance: loans, debt payoff, savings goals, plus live stock, crypto and currency prices.
import { dateToSerial, serialToParts, todaySerial } from '../../dates';
import { external, type Resource } from '../../data/external';
import { Errors, ErrorValue, isError, type Scalar } from '../../values';
import { argNum, argText, flatValues, numResult, toNumber } from '../coerce';
import { Matrix, type Arg, type Ctx } from '../types';
import { define } from './registry';

const MONEY = '"$"#,##0.00';
const DATE = 'mmm d, yyyy';

function addMonths(serial: number, months: number): number {
  const p = serialToParts(serial);
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12);
  const m = (total % 12) + 1;
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return dateToSerial(y, m, Math.min(p.d, dim));
}

const round2 = (n: number) => Math.round(n * 100) / 100;

/** Rates can be typed as 6.5 or 6.5% / 0.065. */
function annualRate(ctx: Ctx, a: Arg): number | ErrorValue {
  const r = argNum(ctx, a);
  if (isError(r)) return r;
  if (r < 0) return Errors.num('The interest rate can\'t be negative.');
  return r >= 1 ? r / 100 : r;
}

function monthlyPayment(r: number, n: number, pv: number): number {
  if (r === 0) return pv / n;
  return (pv * r) / (1 - Math.pow(1 + r, -n));
}

// ---------- loans ----------

define(
  'AMORTIZATION',
  'Personal finance',
  'annual_rate, years, loan_amount, [start_date], [extra_payment]',
  'A month-by-month loan schedule: payment, principal, interest and remaining balance. Example: =AMORTIZATION(6.5%, 30, 350000, DATE(2026,1,1)). An extra monthly payment shows how much sooner you\'d pay it off.',
  (ctx, [ra, yrs, amt, sd, ex]) => {
    const rate = annualRate(ctx, ra);
    if (isError(rate)) return rate;
    const years = argNum(ctx, yrs);
    if (isError(years)) return years;
    const principal = argNum(ctx, amt);
    if (isError(principal)) return principal;
    const start = sd === undefined ? null : argNum(ctx, sd);
    if (isError(start)) return start;
    const extra = argNum(ctx, ex, 0);
    if (isError(extra)) return extra;
    const n = Math.round(years * 12);
    if (n < 1 || n > 1200) return Errors.num('years must be between 1 month and 100 years.');
    if (principal <= 0) return Errors.num('loan_amount must be positive.');
    const r = rate / 12;
    const pay = monthlyPayment(r, n, principal);
    const rows: Scalar[][] = [['Payment #', 'Date', 'Payment', 'Principal', 'Interest', 'Extra', 'Balance']];
    let bal = principal;
    for (let i = 1; i <= n && bal > 0.005; i++) {
      const interest = bal * r;
      let princ = Math.min(bal, pay - interest);
      let extraPaid = Math.min(Math.max(0, extra), bal - princ);
      if (princ < 0) princ = 0;
      if (extraPaid < 0) extraPaid = 0;
      bal = bal - princ - extraPaid;
      if (bal < 0.005) bal = 0;
      rows.push([i, start !== null ? addMonths(start, i - 1) : '', round2(princ + interest), round2(princ), round2(interest), round2(extraPaid), round2(bal)]);
    }
    return new Matrix(rows, [undefined, DATE, MONEY, MONEY, MONEY, MONEY, MONEY]);
  },
  { custom: true, min: 3, argHelp: { annual_rate: 'Yearly rate, like 6.5% (or 6.5)', years: 'Loan length in years', extra_payment: 'Extra principal paid every month' } },
);

define(
  'LOANBALANCE',
  'Personal finance',
  'annual_rate, years, loan_amount, payments_made',
  'What you still owe on a loan after a number of monthly payments. Example: =LOANBALANCE(6.5%, 30, 350000, 60).',
  (ctx, [ra, yrs, amt, made]) => {
    const rate = annualRate(ctx, ra);
    if (isError(rate)) return rate;
    const years = argNum(ctx, yrs);
    if (isError(years)) return years;
    const pv = argNum(ctx, amt);
    if (isError(pv)) return pv;
    const k = argNum(ctx, made);
    if (isError(k)) return k;
    const n = Math.round(years * 12);
    const r = rate / 12;
    const pay = monthlyPayment(r, n, pv);
    if (k >= n) return 0;
    if (r === 0) return round2(pv - pay * k);
    return round2(pv * Math.pow(1 + r, k) - (pay * (Math.pow(1 + r, k) - 1)) / r);
  },
  { custom: true, lift: true },
);

define(
  'PAYOFFMONTHS',
  'Personal finance',
  'annual_rate, balance, monthly_payment',
  'How many months to pay off a balance (credit card, loan) with a fixed monthly payment. Example: =PAYOFFMONTHS(24.99%, 5000, 200).',
  (ctx, [ra, b, p]) => {
    const rate = annualRate(ctx, ra);
    if (isError(rate)) return rate;
    const bal = argNum(ctx, b);
    if (isError(bal)) return bal;
    const pay = argNum(ctx, p);
    if (isError(pay)) return pay;
    const r = rate / 12;
    if (bal <= 0) return 0;
    if (pay <= bal * r) return Errors.num(`A payment of ${round2(pay)} doesn't cover the monthly interest (${round2(bal * r)}), so the balance never goes down.`);
    // Trim floating-point noise so an exact PMT() payment gives a whole number of months
    const ceil = (x: number) => Math.ceil(x - 1e-9);
    if (r === 0) return ceil(bal / pay);
    return ceil(-Math.log(1 - (r * bal) / pay) / Math.log(1 + r));
  },
  { custom: true, lift: true },
);

define(
  'PAYOFFINTEREST',
  'Personal finance',
  'annual_rate, balance, monthly_payment',
  'Total interest you\'ll pay before a balance is gone, paying a fixed amount each month.',
  (ctx, [ra, b, p]) => {
    const rate = annualRate(ctx, ra);
    if (isError(rate)) return rate;
    const bal0 = argNum(ctx, b);
    if (isError(bal0)) return bal0;
    const pay = argNum(ctx, p);
    if (isError(pay)) return pay;
    const r = rate / 12;
    if (pay <= bal0 * r) return Errors.num('The payment doesn\'t cover the monthly interest, so the balance never goes down.');
    let bal = bal0;
    let interest = 0;
    for (let i = 0; i < 1200 && bal > 0.005; i++) {
      const it = bal * r;
      interest += it;
      bal = bal + it - Math.min(pay, bal + it);
    }
    return round2(interest);
  },
  { custom: true, lift: true },
);

define(
  'DEBTPAYOFF',
  'Personal finance',
  'names, balances, annual_rates, min_payments, monthly_budget, [method]',
  'Plans paying off several debts with one monthly budget. method "avalanche" (highest rate first — least interest) or "snowball" (smallest balance first). Returns when each debt is paid off and the interest paid.',
  (ctx, [nm, bs, rs, ms, budgetArg, meth]) => {
    const names = flatValues(ctx, nm);
    const bals = flatValues(ctx, bs).map((v) => toNumber(v));
    const rates = flatValues(ctx, rs).map((v) => toNumber(v));
    const mins = flatValues(ctx, ms).map((v) => toNumber(v));
    for (const list of [bals, rates, mins]) for (const v of list) if (isError(v)) return v;
    const n = bals.length;
    if (!n || rates.length !== n || mins.length !== n || names.length !== n) return Errors.value('names, balances, annual_rates and min_payments must be the same size.');
    const budget = argNum(ctx, budgetArg);
    if (isError(budget)) return budget;
    const method = argText(ctx, meth, 'avalanche');
    if (isError(method)) return method;
    const snowball = method.toLowerCase().startsWith('s');
    const debts = bals.map((b, i) => ({
      name: names[i] === null ? `Debt ${i + 1}` : String(names[i]),
      bal: b as number,
      rate: ((rates[i] as number) >= 1 ? (rates[i] as number) / 100 : (rates[i] as number)) / 12,
      min: mins[i] as number,
      interest: 0,
      month: 0,
    }));
    const totalMin = debts.reduce((s, d) => s + d.min, 0);
    if (budget < totalMin) return Errors.num(`The monthly budget (${budget}) is less than the minimum payments (${round2(totalMin)}).`);
    const order = () =>
      debts
        .filter((d) => d.bal > 0.005)
        .sort((a, b) => (snowball ? a.bal - b.bal : b.rate - a.rate || a.bal - b.bal));
    let month = 0;
    while (debts.some((d) => d.bal > 0.005) && month < 1200) {
      month++;
      for (const d of debts) {
        if (d.bal <= 0.005) continue;
        const it = d.bal * d.rate;
        d.interest += it;
        d.bal += it;
      }
      let left = budget;
      for (const d of debts) {
        if (d.bal <= 0.005) continue;
        const pay = Math.min(d.min, d.bal);
        d.bal -= pay;
        left -= pay;
      }
      for (const d of order()) {
        if (left <= 0) break;
        const pay = Math.min(left, d.bal);
        d.bal -= pay;
        left -= pay;
      }
      for (const d of debts) if (d.bal <= 0.005 && !d.month) {
        d.bal = 0;
        d.month = month;
      }
    }
    if (month >= 1200) return Errors.num('These debts would take more than 100 years to pay off with this budget.');
    const out: Scalar[][] = [['Debt', 'Paid off after (months)', 'Interest paid']];
    for (const d of debts) out.push([d.name, d.month, round2(d.interest)]);
    out.push(['Total', Math.max(...debts.map((d) => d.month)), round2(debts.reduce((s, d) => s + d.interest, 0))]);
    return new Matrix(out, [undefined, '0', MONEY]);
  },
  { custom: true, min: 5 },
);

// ---------- saving & investing ----------

define(
  'SAVINGSGOAL',
  'Personal finance',
  'goal, years, annual_rate, [current_savings]',
  'How much to save each month to reach a goal. Example: =SAVINGSGOAL(25000, 3, 4.5%, 2000).',
  (ctx, [g, y, ra, cs]) => {
    const goal = argNum(ctx, g);
    if (isError(goal)) return goal;
    const years = argNum(ctx, y);
    if (isError(years)) return years;
    const rate = annualRate(ctx, ra);
    if (isError(rate)) return rate;
    const cur = argNum(ctx, cs, 0);
    if (isError(cur)) return cur;
    const n = Math.round(years * 12);
    if (n < 1) return Errors.num('years must be at least one month.');
    const r = rate / 12;
    const grown = cur * Math.pow(1 + r, n);
    const need = goal - grown;
    if (need <= 0) return 0;
    if (r === 0) return round2(need / n);
    return round2((need * r) / (Math.pow(1 + r, n) - 1));
  },
  { custom: true, lift: true },
);

define(
  'SAVINGS.PROJECTION',
  'Personal finance',
  'starting_balance, monthly_contribution, annual_rate, years',
  'Year-by-year growth of savings or investments with monthly contributions and compound interest. Example: =SAVINGS.PROJECTION(10000, 500, 7%, 30).',
  (ctx, [sb, mc, ra, y]) => {
    const start = argNum(ctx, sb);
    if (isError(start)) return start;
    const monthly = argNum(ctx, mc);
    if (isError(monthly)) return monthly;
    const rate = annualRate(ctx, ra);
    if (isError(rate)) return rate;
    const years = argNum(ctx, y);
    if (isError(years)) return years;
    if (years < 1 || years > 100) return Errors.num('years must be between 1 and 100.');
    const r = rate / 12;
    let bal = start;
    let contributed = start;
    const out: Scalar[][] = [['Year', 'Contributed', 'Growth', 'Balance']];
    for (let yr = 1; yr <= Math.round(years); yr++) {
      for (let m = 0; m < 12; m++) {
        bal = bal * (1 + r) + monthly;
        contributed += monthly;
      }
      out.push([yr, round2(contributed), round2(bal - contributed), round2(bal)]);
    }
    return new Matrix(out, [undefined, MONEY, MONEY, MONEY]);
  },
  { custom: true },
);

define(
  'CAGR',
  'Personal finance',
  'start_value, end_value, years',
  'Compound annual growth rate between two values. Example: =CAGR(10000, 18000, 5) returns about 12.5%.',
  (ctx, [a, b, y]) => {
    const s = argNum(ctx, a);
    if (isError(s)) return s;
    const e = argNum(ctx, b);
    if (isError(e)) return e;
    const n = argNum(ctx, y);
    if (isError(n)) return n;
    if (s <= 0 || n <= 0) return Errors.num('start_value and years must be positive.');
    return numResult(Math.pow(e / s, 1 / n) - 1);
  },
  { custom: true, lift: true },
);

// ---------- live prices ----------

const YAHOO = 'https://query1.finance.yahoo.com/v8/finance/chart/';
const UA = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130 Safari/537.36' };

interface ChartResult {
  meta: Record<string, unknown>;
  timestamp?: number[];
  indicators?: { quote?: { open?: (number | null)[]; high?: (number | null)[]; low?: (number | null)[]; close?: (number | null)[]; volume?: (number | null)[] }[]; adjclose?: { adjclose?: (number | null)[] }[] };
}

function parseChart(text: string): ChartResult {
  const j = JSON.parse(text) as { chart?: { result?: ChartResult[] | null; error?: { description?: string } | null } };
  const r = j.chart?.result?.[0];
  if (!r) throw new Error(j.chart?.error?.description ?? 'No data for that ticker');
  return r;
}

function tickerArg(ctx: Ctx, a: Arg): string | ErrorValue {
  const t = argText(ctx, a);
  if (isError(t)) return t;
  const s = t.trim().toUpperCase();
  if (!/^[A-Z0-9.^=\-]{1,15}$/.test(s)) return Errors.value(`"${t}" doesn't look like a ticker symbol (e.g. "AAPL", "VOO", "^GSPC").`);
  return s;
}

const QUOTE_FIELDS = ['Price', 'Change', 'Change %', 'Previous Close', 'Day High', 'Day Low', 'Volume', '52 Week High', '52 Week Low', 'Name', 'Currency', 'Exchange', 'Updated'];

define(
  'STOCK',
  'Personal finance',
  'ticker, [field]',
  'A live (about 15-minute delayed) stock, ETF or index quote. Fields: Price, Change, Change %, Previous Close, Day High, Day Low, Volume, 52 Week High, 52 Week Low, Name, Currency, Exchange, Updated. Example: =STOCK("VOO").',
  (ctx, [t, f]) => {
    const ticker = tickerArg(ctx, t);
    if (isError(ticker)) return ticker;
    const field = argText(ctx, f, 'Price');
    if (isError(field)) return field;
    const key = field.toLowerCase().replace(/[^a-z0-9%]/g, '');
    const canonical = QUOTE_FIELDS.find((x) => x.toLowerCase().replace(/[^a-z0-9%]/g, '') === key) ?? ({ last: 'Price', close: 'Price', changepct: 'Change %', changepercent: 'Change %', prevclose: 'Previous Close', high: 'Day High', low: 'Day Low', '52whigh': '52 Week High', '52wlow': '52 Week Low', time: 'Updated' } as Record<string, string>)[key];
    if (!canonical) return Errors.value(`Unknown field "${field}". Use one of: ${QUOTE_FIELDS.join(', ')}.`);
    const res: Resource<ChartResult> = { key: `stock:${ticker}`, url: `${YAHOO}${encodeURIComponent(ticker)}?range=1d&interval=1d`, maxAgeSec: 120, label: `the ${ticker} quote`, parse: parseChart, headers: UA };
    const got = external.get(res);
    if (!got.ok) return got.error.code === '#N/A' ? Errors.na(`No quote for "${ticker}". Check the ticker symbol.`) : got.error;
    const m = got.value.meta;
    const num = (k: string) => (typeof m[k] === 'number' ? (m[k] as number) : null);
    const price = num('regularMarketPrice');
    const prev = num('chartPreviousClose') ?? num('previousClose');
    switch (canonical) {
      case 'Price':
        return price ?? Errors.na();
      case 'Change':
        return price !== null && prev !== null ? round2(price - prev) : Errors.na();
      case 'Change %':
        return price !== null && prev ? (price - prev) / prev : Errors.na();
      case 'Previous Close':
        return prev ?? Errors.na();
      case 'Day High':
        return num('regularMarketDayHigh') ?? Errors.na();
      case 'Day Low':
        return num('regularMarketDayLow') ?? Errors.na();
      case 'Volume':
        return num('regularMarketVolume') ?? Errors.na();
      case '52 Week High':
        return num('fiftyTwoWeekHigh') ?? Errors.na();
      case '52 Week Low':
        return num('fiftyTwoWeekLow') ?? Errors.na();
      case 'Name':
        return String(m.longName ?? m.shortName ?? ticker);
      case 'Currency':
        return String(m.currency ?? '');
      case 'Exchange':
        return String(m.fullExchangeName ?? m.exchangeName ?? '');
      case 'Updated': {
        const t2 = num('regularMarketTime');
        if (t2 === null) return Errors.na();
        const d = new Date(t2 * 1000);
        return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60) / 86400;
      }
    }
    return Errors.na();
  },
  { external: true, lift: [0, 1], suggest: { 1: 'stockField' } },
);

function serialToUnix(serial: number): number {
  const p = serialToParts(serial);
  return Math.floor(Date.UTC(p.y, p.m - 1, p.d) / 1000);
}

define(
  'STOCKHISTORY',
  'Personal finance',
  'ticker, start_date, [end_date], [interval]',
  'Historical prices as a table: Date, Open, High, Low, Close, Volume. interval: "daily" (default), "weekly" or "monthly". Example: =STOCKHISTORY("AAPL", DATE(2024,1,1), DATE(2024,12,31), "monthly").',
  (ctx, [t, s, e, iv]) => {
    const ticker = tickerArg(ctx, t);
    if (isError(ticker)) return ticker;
    const start = argNum(ctx, s);
    if (isError(start)) return start;
    const today = todaySerial();
    const end = e === undefined ? today : argNum(ctx, e);
    if (isError(end)) return end;
    if (end < start) return Errors.value('end_date is before start_date.');
    const ivText = argText(ctx, iv, 'daily');
    if (isError(ivText)) return ivText;
    const interval = ivText.toLowerCase().startsWith('w') ? '1wk' : ivText.toLowerCase().startsWith('m') ? '1mo' : '1d';
    const p1 = serialToUnix(Math.floor(start));
    const p2 = serialToUnix(Math.floor(end) + 1);
    const res: Resource<ChartResult> = {
      key: `stockhist:${ticker}:${p1}:${p2}:${interval}`,
      url: `${YAHOO}${encodeURIComponent(ticker)}?period1=${p1}&period2=${p2}&interval=${interval}`,
      maxAgeSec: end >= today - 1 ? 3600 : 30 * 86400,
      label: `${ticker} price history`,
      parse: parseChart,
      headers: UA,
    };
    const got = external.get(res);
    if (!got.ok) return got.error.code === '#N/A' ? Errors.na(`No price history for "${ticker}".`) : got.error;
    const r = got.value;
    const ts = r.timestamp ?? [];
    const q = r.indicators?.quote?.[0] ?? {};
    const out: Scalar[][] = [['Date', 'Open', 'High', 'Low', 'Close', 'Volume']];
    const off = typeof r.meta.gmtoffset === 'number' ? (r.meta.gmtoffset as number) : 0;
    ts.forEach((t2, i) => {
      const close = q.close?.[i];
      if (close === null || close === undefined) return;
      const d = new Date((t2 + off) * 1000);
      const serial = dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
      const v = (a?: (number | null)[]) => (a?.[i] === null || a?.[i] === undefined ? '' : round2(a[i] as number));
      out.push([serial, v(q.open), v(q.high), v(q.low), round2(close), q.volume?.[i] ?? '']);
    });
    if (out.length === 1) return Errors.na(`No ${ticker} prices between those dates.`);
    return new Matrix(out, ['yyyy-mm-dd', '#,##0.00', '#,##0.00', '#,##0.00', '#,##0.00', '#,##0']);
  },
  { external: true, min: 2, suggest: { 3: 'interval' } },
);

define(
  'CRYPTO',
  'Personal finance',
  'symbol, [currency], [date]',
  'Cryptocurrency price (from Coinbase). Example: =CRYPTO("BTC") or a past day: =CRYPTO("ETH", "USD", DATE(2024,1,1)).',
  (ctx, [s, c, d]) => {
    const sym = argText(ctx, s);
    if (isError(sym)) return sym;
    const cur = argText(ctx, c, 'USD');
    if (isError(cur)) return cur;
    const pair = `${sym.trim().toUpperCase()}-${cur.trim().toUpperCase()}`;
    if (!/^[A-Z0-9]{2,10}-[A-Z]{3,5}$/.test(pair)) return Errors.value(`"${sym}" isn't a crypto symbol like "BTC" or "ETH".`);
    let dateStr = '';
    if (d !== undefined) {
      const ds = argNum(ctx, d);
      if (isError(ds)) return ds;
      const p = serialToParts(ds);
      dateStr = `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`;
    }
    const res: Resource<number> = {
      key: `crypto:${pair}:${dateStr}`,
      url: `https://api.coinbase.com/v2/prices/${pair}/spot${dateStr ? `?date=${dateStr}` : ''}`,
      maxAgeSec: dateStr ? 365 * 86400 : 60,
      label: `the ${pair} price`,
      parse: (text) => {
        const j = JSON.parse(text) as { data?: { amount?: string } };
        const n = Number(j.data?.amount);
        if (!isFinite(n)) throw new Error('No price');
        return n;
      },
    };
    const got = external.get(res);
    if (!got.ok) return got.error.code === '#N/A' || got.error.code === '#CONNECT!' ? Errors.na(`No price for ${pair}. Check the symbol.`) : got.error;
    return got.value;
  },
  { external: true, lift: [0, 1, 2], suggest: { 1: 'currency' } },
);

function fxRate(ctx: Ctx, f: Arg, t: Arg, d: Arg): number | ErrorValue {
  const from = argText(ctx, f);
  if (isError(from)) return from;
  const to = argText(ctx, t);
  if (isError(to)) return to;
  const a = from.trim().toUpperCase();
  const b = to.trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(a) || !/^[A-Z]{3}$/.test(b)) return Errors.value('Use 3-letter currency codes like "USD", "EUR", "GBP", "JPY".');
  if (a === b) return 1;
  let path = 'latest';
  if (d !== undefined) {
    const ds = argNum(ctx, d);
    if (isError(ds)) return ds;
    const p = serialToParts(ds);
    path = `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`;
  }
  const res: Resource<number> = {
    key: `fx:${a}:${b}:${path}`,
    url: `https://api.frankfurter.dev/v1/${path}?base=${a}&symbols=${b}`,
    maxAgeSec: path === 'latest' ? 3600 : 365 * 86400,
    label: `the ${a}→${b} exchange rate`,
    parse: (text) => {
      const j = JSON.parse(text) as { rates?: Record<string, number> };
      const r = j.rates?.[b];
      if (typeof r !== 'number') throw new Error('No rate');
      return r;
    },
  };
  const got = external.get(res);
  if (!got.ok) return got.error.code === '#BUSY!' ? got.error : Errors.na(`No exchange rate for ${a}→${b}. Supported currencies are those published by the European Central Bank.`);
  return got.value;
}

define(
  'FXRATE',
  'Personal finance',
  'from_currency, to_currency, [date]',
  'Exchange rate between two currencies (European Central Bank reference rates, updated daily). Example: =FXRATE("USD", "EUR").',
  (ctx, [f, t, d]) => fxRate(ctx, f, t, d),
  { external: true, lift: [0, 1, 2], suggest: { 0: 'currency', 1: 'currency' } },
);

define(
  'FX',
  'Personal finance',
  'amount, from_currency, to_currency, [date]',
  'Converts money between currencies. Example: =FX(100, "USD", "EUR").',
  (ctx, [amt, f, t, d]) => {
    const a = argNum(ctx, amt);
    if (isError(a)) return a;
    const rate = fxRate(ctx, f, t, d);
    if (isError(rate)) return rate;
    return round2(a * rate);
  },
  { external: true, lift: [0, 1, 2, 3], suggest: { 1: 'currency', 2: 'currency' } },
);

export const STOCK_FIELDS = QUOTE_FIELDS;

