import { _electron as electron } from 'playwright'
import { readFileSync, writeFileSync } from 'fs'
const cases = readFileSync('test/cases.txt', 'utf8').split('\n').filter((l) => l.startsWith('=')).map((l) => {
  const i = l.lastIndexOf(' => '); return { f: l.slice(0, i), exp: JSON.parse(l.slice(i + 4)), raw: l.slice(i + 4) } })
const app = await electron.launch({ args: ['.', '--no-sandbox'] })
const win = await app.firstWindow()
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
const res = await win.evaluate(async (cases) => {
  const ws = window.univerAPI.getActiveWorkbook().getActiveSheet()
  const put = (a, rows) => rows.forEach((r, i) => r.forEach((v, j) => ws.getRange(i, a + j).setValue(v)))
  put(9, [[1], [2], [3], [4]]); put(10, [[10], [20], [30], [40]]); put(11, [['a'], ['b'], ['c'], ['d']])
  put(13, [[2], [4], [4], [4], [5], [5], [7], [9]]); put(15, [[3], [1], [4], [1]])
  put(17, [['k', 'v'], ['a', 5], ['b', 9], ['c', 1]]); put(19, [['k'], ['a']]); put(20, [['v'], ['>0']])
  cases.forEach((c, i) => ws.getRange(i, 0).setFormula(c.f))
  await new Promise((r) => setTimeout(r, 12000))
  ws.getRange(0, 0, cases.length, 1).setNumberFormat('General')
  await new Promise((r) => setTimeout(r, 1500))
  return cases.map((c, i) => { const r = ws.getRange(i, 0), d = r.getCellData(); return { ...c, got: d?.t === 3 ? !!d.v : (r.getRawValue() ?? d?.v ?? r.getValue()) } })
}, cases)
const dec = (s) => (/^-?\d+\.(\d+)$/.exec(s) || [])[1]?.length ?? 0
const ok = (c) => {
  if (typeof c.exp === 'number' && typeof c.got === 'number') return Math.abs(c.got - c.exp) <= Math.max(0.5 * 10 ** -dec(c.raw) * 1.01, 1e-9 * Math.max(1, Math.abs(c.exp)))
  return c.got === c.exp
}
const bad = res.filter((c) => !ok(c))
const out = [`FULL FORMULA VALUE TEST ${new Date().toISOString()}`, `pass ${res.length - bad.length}/${res.length}`,
  ...bad.map((c) => `FAIL ${c.f}  expected=${c.raw} got=${JSON.stringify(c.got)}`)].join('\n')
console.log(out); writeFileSync('test_results/formulas_full_2026-10-14.txt', out + '\n'); await app.close()
