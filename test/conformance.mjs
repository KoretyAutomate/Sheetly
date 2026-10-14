import { _electron as electron } from 'playwright'
import { writeFileSync } from 'fs'
const app = await electron.launch({ args: ['.', '--no-sandbox'] })
const win = await app.firstWindow()
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })

// [formula, expected per Excel docs]
const F = [
  ['=SUM(1,2,3)', 6], ['=AVERAGE(2,4,6)', 4], ['=ROUND(2.567,2)', 2.57], ['=ROUNDUP(3.2,0)', 4],
  ['=MOD(-3,2)', 1], ['=POWER(2,10)', 1024], ['=SQRT(16)', 4], ['=ABS(-5)', 5], ['=INT(-1.5)', -2],
  ['=SUMPRODUCT({1,2,3},{4,5,6})', 32], ['=COUNTIF(J1:J4,">2")', 2], ['=SUMIF(J1:J4,">2")', 7],
  ['=SUMIFS(J1:J4,J1:J4,">1",J1:J4,"<4")', 5], ['=AVERAGEIF(J1:J4,">2")', 3.5],
  ['=IF(1>2,"a","b")', 'b'], ['=IFERROR(1/0,"err")', 'err'], ['=IFS(1>2,"a",2>1,"b")', 'b'],
  ['=SWITCH(2,1,"one",2,"two")', 'two'], ['=AND(TRUE,FALSE)', false], ['=XOR(TRUE,FALSE)', true],
  ['=LEFT("Excel",2)', 'Ex'], ['=RIGHT("Excel",3)', 'cel'], ['=MID("Excel",2,3)', 'xce'], ['=LEN("hello")', 5],
  ['=UPPER("abc")', 'ABC'], ['=PROPER("hello world")', 'Hello World'], ['=TRIM("  a   b ")', 'a b'],
  ['=SUBSTITUTE("a-b-c","-","+")', 'a+b+c'], ['=CONCAT("a","b")', 'ab'], ['=TEXTJOIN(",",TRUE,"a","","b")', 'a,b'],
  ['=TEXT(1234.5,"#,##0.00")', '1,234.50'], ['=TEXT(0.256,"0.0%")', '25.6%'], ['=FIND("c","abc")', 3],
  ['=SEARCH("C","abc")', 3], ['=REPT("ab",3)', 'ababab'], ['=VALUE("12")', 12], ['=EXACT("a","A")', false],
  ['=VLOOKUP(2,{1,"a";2,"b";3,"c"},2,FALSE)', 'b'], ['=HLOOKUP(2,{1,2,3;"a","b","c"},2,FALSE)', 'b'],
  ['=INDEX({1,2;3,4},2,1)', 3], ['=MATCH(3,{1,2,3},0)', 3], ['=XLOOKUP(2,{1,2,3},{"a","b","c"})', 'b'],
  ['=XMATCH(3,{1,2,3})', 3], ['=CHOOSE(2,"a","b","c")', 'b'], ['=LET(x,5,x*2)', 10],
  ['=ROWS({1;2;3})', 3], ['=COLUMNS({1,2,3})', 3], ['=LARGE({1,5,3},2)', 3], ['=SMALL({1,5,3},1)', 1],
  ['=MAX(1,9,3)', 9], ['=MIN(4,2,8)', 2], ['=MEDIAN(1,3,2)', 2], ['=MODE(1,2,2,3)', 2], ['=STDEV.S(1,2,3,4)', 1.2909944487358056],
  ['=RANK(3,{1,2,3})', 1], ['=N(PMT(0.05/12,60,-10000))', 188.71233566], ['=N(FV(0.05,10,-100))', 1257.789253555],
  ['=N(NPV(0.1,100,200))', 256.1983471074380], ['=N(IRR({-100,60,60}))', 0.13066238629], ['=N(DATE(2024,2,29))', 45351],
  ['=YEAR(DATE(2024,2,29))', 2024], ['=MONTH(DATE(2024,2,29))', 2], ['=N(EDATE(DATE(2024,1,31),1))', 45351],
  ['=EOMONTH(DATE(2024,1,15),0)', 45322], ['=DATEDIF(DATE(2020,1,1),DATE(2024,1,1),"Y")', 4],
  ['=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))', 5], ['=WEEKDAY(DATE(2024,1,1))', 2], ['=DAYS(DATE(2024,1,10),DATE(2024,1,1))', 9],
  ['=ISNUMBER(1)', true], ['=ISBLANK(Z99)', true], ['=ISERROR(1/0)', true], ['=NA()', '#N/A'],
    ['=SUBTOTAL(9,J1:J4)', 10], ['=AGGREGATE(9,6,J1:J4)', 10], ['=ROMAN(14)', 'XIV'], ['=GCD(12,18)', 6], ['=FACT(5)', 120], ['=SUM(FILTER(J1:J4,J1:J4>2))', 7], ['=SUM(SORT(J1:J4))', 10], ['=ROWS(UNIQUE(J1:J4))', 4], ['=SUM(SEQUENCE(3))', 6],
  ['=BIN2DEC("1010")', 10], ['=CONVERT(1,"km","m")', 1000],
]
const res = await win.evaluate(async (F) => {
  const ws = window.univerAPI.getActiveWorkbook().getActiveSheet()
  ;[1,2,3,4].forEach((v,i)=>ws.getRange(i,9).setValue(v))
  F.forEach(([f], i) => ws.getRange(i, 0).setFormula(f))
  await new Promise((r) => setTimeout(r, 4000))
  return F.map(([f, exp], i) => ({ f, exp, got: (() => { const c = ws.getRange(i, 0).getCellData(); return c?.t === 3 ? !!c.v : (c?.v ?? ws.getRange(i, 0).getValue()) })() }))
}, F)
const eq = (a, b) => (typeof b === 'number' && typeof a === 'number' ? Math.abs(a - b) < 1e-6 : a === b)
const bad = res.filter((r) => !eq(r.got, r.exp))
const out = [`FORMULA CONFORMANCE ${new Date().toISOString()}`, `pass ${res.length - bad.length}/${res.length}`,
  ...bad.map((r) => `FAIL ${r.f}  expected=${JSON.stringify(r.exp)} got=${JSON.stringify(r.got)}`)]
console.log(out.join('\n')); writeFileSync('test_results/formulas_2026-10-14.txt', out.join('\n') + '\n')
await app.close()
