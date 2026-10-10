import { describe, expect, it } from 'vitest';
import { parseCell } from '../src/core/address';
import { dateToSerial } from '../src/core/dates';
import { ErrorValue } from '../src/core/values';
import { calc, makeDoc, set, val } from './helpers';

describe('operators', () => {
  it('follows Excel precedence', () => {
    const d = makeDoc();
    expect(calc(d, '=1+2*3')).toBe(7);
    expect(calc(d, '=-2^2')).toBe(4);
    expect(calc(d, '=2^3^2')).toBe(64);
    expect(calc(d, '=50%*2')).toBe(1);
    expect(calc(d, '="a"&1+1')).toBe('a2');
    expect(calc(d, '=1=1')).toBe(true);
    expect(calc(d, '="abc"="ABC"')).toBe(true);
    expect(calc(d, '=1<"a"')).toBe(true);
    expect(calc(d, '=0.1+0.2=0.3')).toBe(false);
  });

  it('handles references and blanks', () => {
    const d = makeDoc({ A1: '10', A2: '20', B1: '=A1+A2', B2: '=A3+1', B3: '=A3' });
    expect(val(d, 'B1')).toBe(30);
    expect(val(d, 'B2')).toBe(1);
    expect(val(d, 'B3')).toBe(0);
  });

  it('explains division by zero with the cell name', () => {
    const d = makeDoc({ A1: '5', C1: '=A1/B1' });
    const a = parseCell('C1')!;
    const v = d.engine.getValue(d.wb.activeSheet, a.row, a.col) as ErrorValue;
    expect(v).toBeInstanceOf(ErrorValue);
    expect(v.code).toBe('#DIV/0!');
    expect(v.message).toContain('B1');
    expect(v.message).toContain('empty');
  });

  it('propagates errors with their origin', () => {
    const d = makeDoc({ A1: '=1/0', B1: '=A1+1' });
    const b = d.engine.getValue(d.wb.activeSheet, 0, 1) as ErrorValue;
    expect(b.code).toBe('#DIV/0!');
    expect(b.origin).toEqual({ sheet: 'Sheet1', row: 0, col: 0 });
  });

  it('suggests a fix for misspelled functions', () => {
    const d = makeDoc({ A1: '=SUMM(1,2)' });
    const v = d.engine.getValue(d.wb.activeSheet, 0, 0) as ErrorValue;
    expect(v.code).toBe('#NAME?');
    expect(v.message).toContain('SUM');
  });

  it('detects circular references', () => {
    const d = makeDoc({ A1: '=B1+1', B1: '=A1+1' });
    expect(val(d, 'A1')).toBe('#CIRC!');
    expect(val(d, 'B1')).toBe('#CIRC!');
    set(d, { B1: '5' });
    expect(val(d, 'A1')).toBe(6);
  });

  it('recalculates dependents incrementally', () => {
    const d = makeDoc({ A1: '1', A2: '=A1*2', A3: '=A2*2', A4: '=SUM(A1:A3)' });
    expect(val(d, 'A4')).toBe(7);
    set(d, { A1: '10' });
    expect(val(d, 'A3')).toBe(40);
    expect(val(d, 'A4')).toBe(70);
  });

  it('handles long dependency chains without stack overflow', () => {
    const d = makeDoc();
    const cells: Record<string, string> = { A1: '1' };
    for (let i = 2; i <= 5000; i++) cells['A' + i] = `=A${i - 1}+1`;
    set(d, cells);
    expect(val(d, 'A5000')).toBe(5000);
    // Reverse-order chain (bottom depends on top-down evaluation order)
    const rev: Record<string, string> = { C5000: '1' };
    for (let i = 4999; i >= 1; i--) rev['C' + i] = `=C${i + 1}+1`;
    set(d, rev);
    expect(val(d, 'C1')).toBe(5000);
  });

  it('works across sheets', () => {
    const d = makeDoc({ A1: '7' });
    d.transact('add', (tx) => {
      const s = new (d.wb.sheets[0].constructor as any)('Data 2');
      tx.setProp(d.wb, 'sheets', [...d.wb.sheets, s], true);
    });
    const s2 = d.wb.sheets[1];
    set(d, { A1: '3' }, s2);
    set(d, { B1: "='Data 2'!A1*Sheet1!A1" });
    expect(val(d, 'B1')).toBe(21);
  });
});

describe('dynamic arrays', () => {
  it('spills results and reports #SPILL!', () => {
    const d = makeDoc({ A1: '=SEQUENCE(3)' });
    expect(val(d, 'A1')).toBe(1);
    expect(val(d, 'A3')).toBe(3);
    set(d, { A2: 'x' });
    expect(val(d, 'A1')).toBe('#SPILL!');
    set(d, { A2: '' });
    expect(val(d, 'A2')).toBe(2);
  });

  it('lets other formulas read spilled values', () => {
    const d = makeDoc({ B1: '=SUM(A1:A5)', A1: '=SEQUENCE(5)' });
    expect(val(d, 'B1')).toBe(15);
    set(d, { C1: '=SUM(A1#)' });
    expect(val(d, 'C1')).toBe(15);
  });

  it('evaluates elementwise array math', () => {
    const d = makeDoc({ A1: '1', A2: '2', A3: '3', B1: '=A1:A3*10' });
    expect(val(d, 'B3')).toBe(30);
    expect(calc(d, '=SUM(A1:A3*A1:A3)')).toBe(14);
  });

  it('supports FILTER, SORT, UNIQUE', () => {
    const d = makeDoc({ A1: 'b', A2: 'a', A3: 'b', A4: 'c', B1: '3', B2: '1', B3: '5', B4: '2' });
    set(d, { D1: '=UNIQUE(A1:A4)' });
    expect([val(d, 'D1'), val(d, 'D2'), val(d, 'D3')]).toEqual(['b', 'a', 'c']);
    set(d, { E1: '=SORT(B1:B4,1,-1)' });
    expect([val(d, 'E1'), val(d, 'E4')]).toEqual([5, 1]);
    set(d, { F1: '=FILTER(A1:A4,B1:B4>2)' });
    expect([val(d, 'F1'), val(d, 'F2')]).toEqual(['b', 'b']);
    set(d, { G1: '=FILTER(A1:A4,B1:B4>100,"none")' });
    expect(val(d, 'G1')).toBe('none');
  });
});

describe('functions', () => {
  const d = makeDoc({
    A1: 'Name',
    B1: 'Qty',
    C1: 'Price',
    A2: 'Apple',
    B2: '10',
    C2: '1.5',
    A3: 'Banana',
    B3: '5',
    C3: '0.25',
    A4: 'Cherry',
    B4: '20',
    C4: '3',
    A5: 'apple ',
    B5: '1',
    C5: '2',
  });

  it('math and stats', () => {
    expect(calc(d, '=SUM(B2:B5)')).toBe(36);
    expect(calc(d, '=SUM(B:B)')).toBe(36);
    expect(calc(d, '=AVERAGE(B2:B4)')).toBeCloseTo(35 / 3);
    expect(calc(d, '=COUNT(A1:C5)')).toBe(8);
    expect(calc(d, '=COUNTA(A1:A5)')).toBe(5);
    expect(calc(d, '=MAX(B2:B5)')).toBe(20);
    expect(calc(d, '=MEDIAN(1,2,3,4)')).toBe(2.5);
    expect(calc(d, '=ROUND(2.345,2)')).toBe(2.35);
    expect(calc(d, '=ROUND(1.005,2)')).toBe(1.01);
    expect(calc(d, '=ROUND(-2.5,0)')).toBe(-3);
    expect(calc(d, '=MOD(-3,2)')).toBe(1);
    expect(calc(d, '=SUMPRODUCT(B2:B4,C2:C4)')).toBe(76.25);
    expect(calc(d, '=STDEV.S(2,4,4,4,5,5,7,9)')).toBeCloseTo(2.138, 3);
    expect(calc(d, '=LARGE(B2:B5,2)')).toBe(10);
    expect(calc(d, '=RANK(10,B2:B5)')).toBe(2);
    expect(calc(d, '=PERCENTILE(B2:B5,0.5)')).toBe(7.5);
    expect(calc(d, '=CEILING(4.2,0.5)')).toBe(4.5);
    expect(calc(d, '=FLOOR(4.7,0.5)')).toBe(4.5);
  });

  it('conditional aggregates', () => {
    expect(calc(d, '=SUMIF(B2:B5,">5")')).toBe(30);
    expect(calc(d, '=SUMIF(A2:A5,"apple",B2:B5)')).toBe(10);
    expect(calc(d, '=COUNTIF(A2:A5,"*an*")')).toBe(1);
    expect(calc(d, '=COUNTIFS(B2:B5,">=5",C2:C5,"<2")')).toBe(2);
    expect(calc(d, '=AVERAGEIF(B2:B5,">4")')).toBeCloseTo(35 / 3);
    expect(calc(d, '=MAXIFS(B2:B5,C2:C5,">1")')).toBe(20);
  });

  it('lookups', () => {
    expect(calc(d, '=VLOOKUP("Banana",A2:C5,3,FALSE)')).toBe(0.25);
    expect(calc(d, '=VLOOKUP("banana",A2:C5,2,0)')).toBe(5);
    expect(calc(d, '=XLOOKUP("Cherry",A2:A5,C2:C5)')).toBe(3);
    expect(calc(d, '=XLOOKUP("Kiwi",A2:A5,C2:C5,"none")')).toBe('none');
    expect(calc(d, '=INDEX(B2:B5,MATCH("Cherry",A2:A5,0))')).toBe(20);
    expect(calc(d, '=MATCH(7,{1,5,10})')).toBe(2);
    expect(calc(d, '=HLOOKUP("Qty",A1:C5,3,FALSE)')).toBe(5);
    expect(calc(d, '=INDEX(A1:C5,4,3)')).toBe(3);
    expect(calc(d, '=SUM(OFFSET(B2,0,0,3,1))')).toBe(35);
    expect(calc(d, '=INDIRECT("B"&4)')).toBe(20);
    expect(calc(d, '=ROWS(A1:C5)*COLUMNS(A1:C5)')).toBe(15);
  });

  it('explains lookup misses', () => {
    set(d, { Z999: '=VLOOKUP("Kiwi",A2:C5,2,FALSE)' });
    const v = d.engine.getValue(d.wb.activeSheet, 998, 25) as ErrorValue;
    expect(v.code).toBe('#N/A');
    expect(v.message).toContain('Kiwi');
    set(d, { Z998: '=XLOOKUP("Apple ",A3:A4,B3:B4)' });
    set(d, { Z997: '=VLOOKUP("apple",A3:B5,2,FALSE)' });
    expect(val(d, 'Z997')).toBe('#N/A');
    const w = d.engine.getValue(d.wb.activeSheet, 996, 25) as ErrorValue;
    expect(w.message).toContain('extra spaces');
  });

  it('text functions', () => {
    expect(calc(d, '=LEFT("Hello",2)&RIGHT("World",3)')).toBe('Herld');
    expect(calc(d, '=MID("abcdef",2,3)')).toBe('bcd');
    expect(calc(d, '=UPPER(A2)')).toBe('APPLE');
    expect(calc(d, '=PROPER("hello wORLD")')).toBe('Hello World');
    expect(calc(d, '=TRIM("  a   b  ")')).toBe('a b');
    expect(calc(d, '=SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(calc(d, '=SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c');
    expect(calc(d, '=FIND("c","abcabc",4)')).toBe(6);
    expect(calc(d, '=SEARCH("B*","aabcd")')).toBe(3);
    expect(calc(d, '=TEXT(1234.5,"$#,##0.00")')).toBe('$1,234.50');
    expect(calc(d, '=TEXTJOIN(", ",TRUE,A2:A4)')).toBe('Apple, Banana, Cherry');
    expect(calc(d, '=VALUE("$1,000")')).toBe(1000);
    expect(calc(d, '=LEN(A3)')).toBe(6);
    expect(calc(d, '=TEXTBEFORE("john@x.com","@")')).toBe('john');
    expect(calc(d, '=TEXTAFTER("a,b,c",",",-1)')).toBe('c');
    expect(calc(d, '=CONCAT(A2:A3)')).toBe('AppleBanana');
    expect(calc(d, '=1/3&""')).toBe('0.333333333333333');
  });

  it('logical functions', () => {
    expect(calc(d, '=IF(B2>5,"big","small")')).toBe('big');
    expect(calc(d, '=IF(FALSE,1)')).toBe(false);
    expect(calc(d, '=IFERROR(1/0,"oops")')).toBe('oops');
    expect(calc(d, '=IFNA(NA(),"na")')).toBe('na');
    expect(calc(d, '=AND(TRUE,1,B2>1)')).toBe(true);
    expect(calc(d, '=OR(FALSE,0)')).toBe(false);
    expect(calc(d, '=IFS(B3>10,"a",B3>4,"b",TRUE,"c")')).toBe('b');
    expect(calc(d, '=SWITCH(2,1,"one",2,"two","other")')).toBe('two');
    expect(calc(d, '=LET(x,5,y,x*2,x+y)')).toBe(15);
    expect(calc(d, '=CHOOSE(2,"a","b","c")')).toBe('b');
    expect(calc(d, '=SUM(CHOOSE(1,B2:B3,C2:C3))')).toBe(15);
  });

  it('date functions', () => {
    const s = dateToSerial(2025, 1, 31);
    expect(calc(d, '=DATE(2025,1,31)')).toBe(s);
    expect(calc(d, '=DATE(2025,14,1)')).toBe(dateToSerial(2026, 2, 1));
    expect(calc(d, `=EDATE(${s},1)`)).toBe(dateToSerial(2025, 2, 28));
    expect(calc(d, `=EOMONTH(${s},1)`)).toBe(dateToSerial(2025, 2, 28));
    expect(calc(d, `=YEAR(${s})*100+MONTH(${s})`)).toBe(202501);
    expect(calc(d, `=WEEKDAY(${s})`)).toBe(6);
    expect(calc(d, '=DATEDIF(DATE(2020,5,15),DATE(2025,3,1),"Y")')).toBe(4);
    expect(calc(d, '=NETWORKDAYS(DATE(2025,3,3),DATE(2025,3,14))')).toBe(10);
    expect(calc(d, '=WORKDAY(DATE(2025,3,7),1)')).toBe(dateToSerial(2025, 3, 10));
    expect(calc(d, '=DATEVALUE("2025-03-14")')).toBe(dateToSerial(2025, 3, 14));
    expect(calc(d, '=HOUR(TIME(13,45,0))')).toBe(13);
    expect(calc(d, '=ISOWEEKNUM(DATE(2025,1,1))')).toBe(1);
  });

  it('financial functions', () => {
    expect(calc(d, '=PMT(0.05/12,360,200000)')).toBeCloseTo(-1073.64, 2);
    expect(calc(d, '=FV(0.05,10,-100)')).toBeCloseTo(1257.79, 2);
    expect(calc(d, '=PV(0.05,10,-100)')).toBeCloseTo(772.17, 2);
    expect(calc(d, '=NPV(0.1,-1000,300,400,500)')).toBeCloseTo(-19.12, 2);
    expect(calc(d, '=IRR({-1000,300,400,500})')).toBeCloseTo(0.0889633, 5);
    expect(calc(d, '=RATE(36,-300,10000)')).toBeCloseTo(0.0042207, 6);
    expect(calc(d, '=NPER(0.01,-100,1000)')).toBeCloseTo(10.588, 2);
  });

  it('information functions', () => {
    expect(calc(d, '=ISBLANK(Z1)')).toBe(true);
    expect(calc(d, '=ISNUMBER(B2)')).toBe(true);
    expect(calc(d, '=ISTEXT(A2)')).toBe(true);
    expect(calc(d, '=ISERROR(1/0)')).toBe(true);
    expect(calc(d, '=ISFORMULA(Z1000)')).toBe(true);
    expect(calc(d, '=TYPE("a")')).toBe(2);
  });

  it('subtotal ignores filtered rows', () => {
    const doc = makeDoc({ A1: '1', A2: '2', A3: '3', A4: '=SUBTOTAL(9,A1:A3)' });
    expect(val(doc, 'A4')).toBe(6);
    doc.transact('filter', (tx) => tx.setProp(doc.wb.activeSheet, 'filterHidden', new Set([1]), true));
    expect(val(doc, 'A4')).toBe(4);
  });
});

describe('syntax errors', () => {
  it('reports friendly messages', async () => {
    const { parseFormula, autoFixFormula } = await import('../src/core/formula/parser');
    expect(parseFormula('SUM(A1').error?.message).toContain('closing parenthesis');
    expect(parseFormula('1+').error?.message).toContain('missing');
    expect(autoFixFormula('SUM(A1')).toBe('SUM(A1)');
    expect(autoFixFormula('"abc')).toBe('"abc"');
  });
});
