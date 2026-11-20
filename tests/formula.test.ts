import { describe, expect, it } from 'vitest';
import { Workbook } from '../src/core/model';
import { applyInput } from '../src/core/ops';
import { parseCell } from '../src/core/addr';
import { FErr, Value } from '../src/core/values';
import { parseFormula, shiftFormula, tokenize } from '../src/core/parser';

function mk(cells: Record<string, string>, sheetName = 'Sheet1') {
  const wb = new Workbook();
  wb.sheets[0].name = sheetName;
  wb.transact('init', () => {
    for (const [k, v] of Object.entries(cells)) {
      const p = parseCell(k)!;
      applyInput(wb, wb.sheets[0], p.r, p.c, v);
    }
  });
  return wb;
}
function val(wb: Workbook, ref: string, sheetIdx = 0): Value {
  const p = parseCell(ref)!;
  return wb.getValue(wb.sheets[sheetIdx], p.r, p.c);
}
function ev(formula: string, cells: Record<string, string> = {}): Value {
  const wb = mk({ ...cells, Z99: '=' + formula });
  return val(wb, 'Z99');
}
const err = (v: Value) => (v instanceof FErr ? v.code : v);

describe('parser', () => {
  it('tokenizes references and ranges', () => {
    const t = tokenize('SUM(A1:B2,$C$3)+Sheet2!D4');
    expect(t.filter((x) => x.kind === 'ref').map((x) => x.text)).toEqual(['A1:B2', '$C$3', 'Sheet2!D4']);
  });
  it('treats LOG10( as a function, not a cell', () => {
    expect(tokenize('LOG10(100)')[0].kind).toBe('func');
  });
  it('respects operator precedence', () => {
    expect(ev('2+3*4')).toBe(14);
    expect(ev('(2+3)*4')).toBe(20);
    expect(ev('2^3^2')).toBe(64);
    expect(ev('-2^2')).toBe(4);
    expect(ev('10-4-3')).toBe(3);
    expect(ev('50%')).toBe(0.5);
    expect(ev('2*50%')).toBe(1);
  });
  it('rejects broken formulas', () => {
    expect(() => parseFormula('1+')).toThrow();
    expect(() => parseFormula('SUM(1,2')).toThrow();
  });
  it('shifts relative references only', () => {
    expect(shiftFormula('A1+$B$2+C$3+$D4', 1, 1)).toBe('B2+$B$2+D$3+$D5');
    expect(shiftFormula('SUM(A1:B2)', 2, 0)).toBe('SUM(A3:B4)');
    expect(shiftFormula('A1', -1, 0)).toBe('#REF!');
    expect(shiftFormula("'My Sheet'!A1", 1, 0)).toBe("'My Sheet'!A2");
  });
});

describe('operators', () => {
  it('arithmetic and text', () => {
    expect(ev('1+2')).toBe(3);
    expect(ev('"a"&"b"')).toBe('ab');
    expect(ev('1&2')).toBe('12');
    expect(ev('0.1+0.2-0.3')).toBe(0);
  });
  it('comparisons', () => {
    expect(ev('1<2')).toBe(true);
    expect(ev('"a"="A"')).toBe(true);
    expect(ev('1<>1')).toBe(false);
    expect(ev('2>=2')).toBe(true);
    expect(ev('"b">"a"')).toBe(true);
    expect(ev('1<"a"')).toBe(true);
  });
  it('errors', () => {
    expect(err(ev('1/0'))).toBe('#DIV/0!');
    expect(err(ev('"x"+1'))).toBe('#VALUE!');
    expect(err(ev('NOPE(1)'))).toBe('#NAME?');
    expect(err(ev('foo'))).toBe('#NAME?');
    expect(err(ev('NA()'))).toBe('#N/A');
    expect(err(ev('1+#REF!'))).toBe('#REF!');
    expect(err(ev('SQRT(-1)'))).toBe('#NUM!');
  });
});

describe('math & stats functions', () => {
  const data = { A1: '1', A2: '2', A3: '3', A4: '4', A5: 'x', A6: '' };
  it('SUM / AVERAGE / MIN / MAX / COUNT / COUNTA', () => {
    expect(ev('SUM(A1:A6)', data)).toBe(10);
    expect(ev('SUM(1,2,"3")', data)).toBe(6);
    expect(ev('AVERAGE(A1:A4)', data)).toBe(2.5);
    expect(ev('MIN(A1:A5)', data)).toBe(1);
    expect(ev('MAX(A1:A5)', data)).toBe(4);
    expect(ev('COUNT(A1:A6)', data)).toBe(4);
    expect(ev('COUNTA(A1:A6)', data)).toBe(5);
    expect(err(ev('AVERAGE(A5:A6)', data))).toBe('#DIV/0!');
  });
  it('whole-column and cross references', () => {
    expect(ev('SUM(A:A)', data)).toBe(10);
    expect(ev('SUM(1:1)', { A1: '5', B1: '6' })).toBe(11);
  });
  it('ROUND family, ABS, INT, MOD, POWER, SQRT', () => {
    expect(ev('ROUND(2.5,0)')).toBe(3);
    expect(ev('ROUND(-2.5,0)')).toBe(-3);
    expect(ev('ROUND(1.005,2)')).toBe(1.01);
    expect(ev('ROUND(1234.567,-2)')).toBe(1200);
    expect(ev('ROUNDUP(1.21,1)')).toBe(1.3);
    expect(ev('ROUNDDOWN(1.29,1)')).toBe(1.2);
    expect(ev('ABS(-3)')).toBe(3);
    expect(ev('INT(-1.5)')).toBe(-2);
    expect(ev('MOD(-3,2)')).toBe(1);
    expect(err(ev('MOD(1,0)'))).toBe('#DIV/0!');
    expect(ev('POWER(2,10)')).toBe(1024);
    expect(ev('SQRT(16)')).toBe(4);
  });
  it('MEDIAN, LARGE, SMALL, STDEV', () => {
    expect(ev('MEDIAN(1,3,2,10)')).toBe(2.5);
    expect(ev('LARGE(A1:A4,2)', data)).toBe(3);
    expect(ev('SMALL(A1:A4,2)', data)).toBe(2);
    expect(ev('STDEV.P(2,4,4,4,5,5,7,9)')).toBe(2);
  });
  it('SUMPRODUCT with arrays', () => {
    expect(ev('SUMPRODUCT(A1:A3,A1:A3)', data)).toBe(14);
    expect(ev('SUMPRODUCT((A1:A4>1)*(A1:A4))', data)).toBe(9);
  });
  it('RAND and RANDBETWEEN are in range', () => {
    const r = ev('RAND()') as number;
    expect(r).toBeGreaterThanOrEqual(0);
    expect(r).toBeLessThan(1);
    const b = ev('RANDBETWEEN(5,7)') as number;
    expect([5, 6, 7]).toContain(b);
  });
});

describe('conditional aggregates', () => {
  const data = { A1: 'apple', A2: 'banana', A3: 'apple', A4: 'cherry', B1: '10', B2: '20', B3: '30', B4: '40' };
  it('COUNTIF / SUMIF / AVERAGEIF', () => {
    expect(ev('COUNTIF(A1:A4,"apple")', data)).toBe(2);
    expect(ev('COUNTIF(A1:A4,"a*")', data)).toBe(2);
    expect(ev('COUNTIF(B1:B4,">15")', data)).toBe(3);
    expect(ev('COUNTIF(B1:B4,"<>20")', data)).toBe(3);
    expect(ev('SUMIF(A1:A4,"apple",B1:B4)', data)).toBe(40);
    expect(ev('SUMIF(B1:B4,">=20")', data)).toBe(90);
    expect(ev('AVERAGEIF(A1:A4,"apple",B1:B4)', data)).toBe(20);
  });
  it('COUNTIFS / SUMIFS', () => {
    expect(ev('COUNTIFS(A1:A4,"apple",B1:B4,">15")', data)).toBe(1);
    expect(ev('SUMIFS(B1:B4,A1:A4,"apple",B1:B4,"<=30")', data)).toBe(40);
    expect(ev('MAXIFS(B1:B4,A1:A4,"apple")', data)).toBe(30);
    expect(err(ev('SUMIFS(B1:B4,A1:A3,"apple")', data))).toBe('#VALUE!');
  });
  it('COUNTIF blanks and cell criteria', () => {
    expect(ev('COUNTIF(A1:A6,"")', data)).toBe(2);
    expect(ev('COUNTIF(A1:A4,A2)', data)).toBe(1);
  });
});

describe('logical', () => {
  it('IF family', () => {
    expect(ev('IF(1>2,"y","n")')).toBe('n');
    expect(ev('IF(TRUE,1)')).toBe(1);
    expect(ev('IF(FALSE,1)')).toBe(false);
    expect(ev('IFS(1>2,"a",2>1,"b")')).toBe('b');
    expect(err(ev('IFS(1>2,"a")'))).toBe('#N/A');
    expect(ev('IFERROR(1/0,"bad")')).toBe('bad');
    expect(ev('IFERROR(5,"bad")')).toBe(5);
    expect(ev('IF(1,2,1/0)')).toBe(2); // lazy branches
    expect(ev('CHOOSE(2,"a","b","c")')).toBe('b');
    expect(ev('SWITCH(2,1,"one",2,"two","other")')).toBe('two');
  });
  it('AND OR NOT XOR', () => {
    expect(ev('AND(TRUE,1,2>1)')).toBe(true);
    expect(ev('AND(TRUE,0)')).toBe(false);
    expect(ev('OR(FALSE,0,1)')).toBe(true);
    expect(ev('NOT(TRUE)')).toBe(false);
    expect(ev('XOR(TRUE,TRUE)')).toBe(false);
  });
});

describe('lookup', () => {
  const t = { A1: 'id', B1: 'name', A2: '1', B2: 'ann', A3: '2', B3: 'bob', A4: '3', B4: 'cy', D1: '10', D2: '20', D3: '30', E1: 'x', E2: 'y', E3: 'z' };
  it('VLOOKUP exact and approximate', () => {
    expect(ev('VLOOKUP(2,A2:B4,2,FALSE)', t)).toBe('bob');
    expect(err(ev('VLOOKUP(9,A2:B4,2,FALSE)', t))).toBe('#N/A');
    expect(ev('VLOOKUP(2.5,A2:B4,2)', t)).toBe('bob');
    expect(err(ev('VLOOKUP(2,A2:B4,3,FALSE)', t))).toBe('#REF!');
    expect(ev('VLOOKUP("BOB",B2:B4,1,FALSE)', t)).toBe('bob');
  });
  it('HLOOKUP', () => {
    expect(ev('HLOOKUP("name",A1:B4,3,FALSE)', t)).toBe('bob');
  });
  it('INDEX / MATCH', () => {
    expect(ev('INDEX(B2:B4,2)', t)).toBe('bob');
    expect(ev('INDEX(A2:B4,3,2)', t)).toBe('cy');
    expect(ev('MATCH("bob",B2:B4,0)', t)).toBe(2);
    expect(ev('MATCH(25,D1:D3,1)', t)).toBe(2);
    expect(ev('INDEX(B2:B4,MATCH(3,A2:A4,0))', t)).toBe('cy');
    expect(err(ev('MATCH("q",B2:B4,0)', t))).toBe('#N/A');
  });
  it('XLOOKUP', () => {
    expect(ev('XLOOKUP(20,D1:D3,E1:E3)', t)).toBe('y');
    expect(ev('XLOOKUP(99,D1:D3,E1:E3,"none")', t)).toBe('none');
    expect(ev('XLOOKUP(25,D1:D3,E1:E3,,-1)', t)).toBe('y');
    expect(ev('XLOOKUP(25,D1:D3,E1:E3,,1)', t)).toBe('z');
  });
  it('ROW / COLUMN / ROWS / COLUMNS', () => {
    expect(ev('ROW(B7)')).toBe(7);
    expect(ev('COLUMN(C1)')).toBe(3);
    expect(ev('ROWS(A1:A5)')).toBe(5);
    expect(ev('COLUMNS(A1:C1)')).toBe(3);
  });
});

describe('text', () => {
  it('basic text functions', () => {
    expect(ev('LEFT("hello",2)')).toBe('he');
    expect(ev('RIGHT("hello",3)')).toBe('llo');
    expect(ev('MID("hello",2,3)')).toBe('ell');
    expect(ev('LEN("hello")')).toBe(5);
    expect(ev('UPPER("abC")')).toBe('ABC');
    expect(ev('LOWER("AbC")')).toBe('abc');
    expect(ev('TRIM("  a   b  ")')).toBe('a b');
    expect(ev('CONCAT("a",1,TRUE)')).toBe('a1TRUE');
    expect(ev('TEXTJOIN(", ",TRUE,"a","","b")')).toBe('a, b');
    expect(ev('TEXTJOIN("-",FALSE,A1:A3)', { A1: 'x', A2: '', A3: 'z' })).toBe('x--z');
    expect(ev('SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(ev('FIND("l","hello")')).toBe(3);
    expect(ev('SEARCH("L","hello")')).toBe(3);
    expect(ev('PROPER("hello wORLD")')).toBe('Hello World');
    expect(ev('REPT("ab",3)')).toBe('ababab');
  });
  it('TEXT formats', () => {
    expect(ev('TEXT(1234.567,"#,##0.00")')).toBe('1,234.57');
    expect(ev('TEXT(0.256,"0%")')).toBe('26%');
    expect(ev('TEXT(DATE(2024,3,5),"yyyy-mm-dd")')).toBe('2024-03-05');
    expect(ev('TEXT(DATE(2024,3,5),"dddd d mmmm")')).toBe('Tuesday 5 March');
    expect(ev('TEXT(5,"000")')).toBe('005');
  });
  it('VALUE', () => {
    expect(ev('VALUE("12.5")')).toBe(12.5);
    expect(err(ev('VALUE("abc")'))).toBe('#VALUE!');
  });
});

describe('dates', () => {
  it('DATE and parts', () => {
    expect(ev('DATE(2000,1,1)')).toBe(36526);
    expect(ev('DATE(1900,3,1)')).toBe(61);
    expect(ev('YEAR(DATE(2024,2,29))')).toBe(2024);
    expect(ev('MONTH(DATE(2024,2,29))')).toBe(2);
    expect(ev('DAY(DATE(2024,2,29))')).toBe(29);
    expect(ev('DATE(2024,14,1)')).toBe(ev('DATE(2025,2,1)'));
    expect(ev('WEEKDAY(DATE(2024,3,5))')).toBe(3); // Tuesday
  });
  it('TODAY / NOW', () => {
    const t = ev('TODAY()') as number;
    expect(t).toBeGreaterThan(45000);
    expect(ev('NOW()') as number).toBeGreaterThanOrEqual(t);
  });
  it('EDATE EOMONTH DAYS DATEDIF', () => {
    expect(ev('EDATE(DATE(2024,1,31),1)')).toBe(ev('DATE(2024,2,29)'));
    expect(ev('EOMONTH(DATE(2024,1,15),0)')).toBe(ev('DATE(2024,1,31)'));
    expect(ev('DAYS(DATE(2024,3,1),DATE(2024,2,1))')).toBe(29);
    expect(ev('DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"Y")')).toBe(4);
  });
});

describe('references and recalculation', () => {
  it('updates dependents when a precedent changes', () => {
    const wb = mk({ A1: '1', B1: '=A1*2', C1: '=B1+1' });
    expect(val(wb, 'C1')).toBe(3);
    wb.transact('e', () => applyInput(wb, wb.sheets[0], 0, 0, '10'));
    expect(val(wb, 'B1')).toBe(20);
    expect(val(wb, 'C1')).toBe(21);
  });
  it('range dependents recalc', () => {
    const wb = mk({ A1: '1', A2: '2', A3: '3', B1: '=SUM(A1:A3)' });
    wb.transact('e', () => applyInput(wb, wb.sheets[0], 1, 0, '20'));
    expect(val(wb, 'B1')).toBe(24);
  });
  it('detects circular references', () => {
    const wb = mk({ A1: '=B1', B1: '=A1', C1: '=C1+1', D1: '=A1+1' });
    expect(err(val(wb, 'A1'))).toBe('#CIRC!');
    expect(err(val(wb, 'B1'))).toBe('#CIRC!');
    expect(err(val(wb, 'C1'))).toBe('#CIRC!');
    expect(err(val(wb, 'D1'))).toBe('#CIRC!');
  });
  it('recovers when a cycle is broken', () => {
    const wb = mk({ A1: '=B1', B1: '=A1' });
    wb.transact('e', () => applyInput(wb, wb.sheets[0], 0, 1, '5'));
    expect(val(wb, 'A1')).toBe(5);
    expect(val(wb, 'B1')).toBe(5);
  });
  it('cross-sheet references', () => {
    const wb = mk({ A1: '7' });
    const s2 = wb.addSheet('Sheet2');
    wb.transact('e', () => {
      applyInput(wb, s2, 0, 0, '=Sheet1!A1*2');
      applyInput(wb, s2, 1, 0, "=SUM('Sheet1'!A1:A3)");
    });
    expect(wb.getValue(s2, 0, 0)).toBe(14);
    wb.transact('e', () => applyInput(wb, wb.sheets[0], 0, 0, '9'));
    expect(wb.getValue(s2, 0, 0)).toBe(18);
    expect(wb.getValue(s2, 1, 0)).toBe(9);
  });
  it('missing sheet gives #REF!', () => {
    expect(err(ev('Nope!A1'))).toBe('#REF!');
  });
  it('named ranges', () => {
    const wb = mk({ A1: '5', A2: '6', B1: '=Total*2', B2: '=Rate' });
    wb.setName('Total', 'Sheet1!$A$1:$A$2');
    wb.setName('Rate', '0.5');
    wb.rebuildDeps();
    wb.recalcAll();
    expect(val(wb, 'B2')).toBe(0.5);
    // Total is a range: implicit intersection on row 1 gives A1
    expect(val(wb, 'B1')).toBe(10);
    wb.transact('e', () => applyInput(wb, wb.sheets[0], 0, 0, '8'));
    expect(val(wb, 'B1')).toBe(16);
  });
  it('empty references act as zero', () => {
    expect(ev('A1+1')).toBe(1);
    expect(ev('A1&"x"')).toBe('x');
  });
});

describe('performance', () => {
  it('recalculates 10,000 formulas in under 200 ms', () => {
    const wb = new Workbook();
    const s = wb.sheets[0];
    wb.transact('fill', () => {
      for (let r = 0; r < 10000; r++) {
        applyInput(wb, s, r, 0, String(r));
        applyInput(wb, s, r, 1, `=A${r + 1}*2+1`);
      }
    });
    const t0 = performance.now();
    wb.transact('edit', () => applyInput(wb, s, 0, 0, '100'));
    wb.recalcAll();
    const dt = performance.now() - t0;
    expect(wb.getValue(s, 9999, 1)).toBe(19999);
    expect(dt).toBeLessThan(200);
  });
  it('handles a long dependency chain without overflowing the stack', () => {
    const wb = new Workbook();
    const s = wb.sheets[0];
    wb.transact('chain', () => {
      applyInput(wb, s, 0, 0, '1');
      for (let r = 1; r < 20000; r++) applyInput(wb, s, r, 0, `=A${r}+1`);
    });
    expect(wb.getValue(s, 19999, 0)).toBe(20000);
    wb.transact('e', () => applyInput(wb, s, 0, 0, '5'));
    expect(wb.getValue(s, 19999, 0)).toBe(20004);
  });
});
