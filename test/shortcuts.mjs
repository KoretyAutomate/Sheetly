import { _electron as electron } from 'playwright'
import { writeFileSync } from 'fs'
import { xlsxToSnapshot, snapshotToXlsx } from '../src/xlsx.js'
const app = await electron.launch({ args: ['.', '--no-sandbox'] })
const win = await app.firstWindow()
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
const kb = win.keyboard, out = []
const sel = () => win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getActiveRange()?.getA1Notation())
const val = (a) => win.evaluate((a) => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange(a).getValue(), a)
const sleep = (ms) => win.waitForTimeout(ms)
const check = (name, ok, got) => out.push(`${ok ? 'PASS' : 'FAIL'} ${name}${ok ? '' : '  got=' + JSON.stringify(got)}`)
await win.mouse.click(120, 140); await sleep(400)             // select a cell in grid
await win.evaluate(() => { const s = window.univerAPI.getActiveWorkbook().getActiveSheet(); s.getRange('A1').activate()
  for (let r = 0; r < 5; r++) s.getRange(r, 0).setValue(r + 1) }); await sleep(300)
await kb.press('Control+ArrowDown'); await sleep(200); check('Ctrl+Down jumps to end of data (A5)', (await sel()) === 'A5', await sel())
await kb.press('Control+Home'); await sleep(200); check('Ctrl+Home -> A1', (await sel()) === 'A1', await sel())
await kb.press('Control+End'); await sleep(200); check('Ctrl+End -> last used cell (A5)', (await sel()) === 'A5', await sel())
await kb.press('Control+Home'); await kb.press('ArrowRight'); await kb.type('x'); await kb.press('Enter'); await sleep(300)
check('typing + Enter commits, moves down', (await val('B1')) === 'x' && (await sel()) === 'B2', [await val('B1'), await sel()])
await kb.press('Control+z'); await sleep(300); check('Ctrl+Z undo', (await val('B1')) === '' || (await val('B1')) == null, await val('B1'))
await kb.press('Control+y'); await sleep(300); check('Ctrl+Y redo', (await val('B1')) === 'x', await val('B1'))
await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('C1').activate()); await sleep(200)
await kb.type('=A1'); await kb.press('F4'); await sleep(200); await kb.press('Enter'); await sleep(300)
const f = await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('C1').getFormula())
check('F4 cycles to absolute ref ($A$1)', f === '=$A$1', f)
await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A1').activate()); await sleep(200)
await kb.press('F2'); await sleep(200); await kb.type('9'); await kb.press('Enter'); await sleep(300)
check('F2 edits in-cell (appends)', String(await val('A1')) === '19', await val('A1'))
await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A1:A3').activate()); await sleep(200)
await kb.press('Control+Shift+ArrowRight'); await sleep(200)
check('Ctrl+Shift+Right extends selection', /^A1:/.test(await sel() || ''), await sel())
await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A6').activate()); await sleep(200)
await kb.press('Alt+='); await sleep(400); const a6 = await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A6').getFormula())
check('Alt+= AutoSum', /SUM\(/i.test(a6 || ''), a6)
await win.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A7').activate()); await sleep(200)
await kb.press('Control+;'); await sleep(300); const d = await val('A7')
check('Ctrl+; inserts today', d !== '' && d != null, d)
// round trip
const wbSnap = await win.evaluate(() => window.univerAPI.getActiveWorkbook().save())
const buf = await snapshotToXlsx(wbSnap); const back = await xlsxToSnapshot(buf)
const s0 = back.sheets[back.sheetOrder[0]]
check('xlsx roundtrip keeps values+formulas', s0.cellData[0][1]?.v === 'x' && String(s0.cellData[2][2]?.f ?? s0.cellData[0][2]?.f).startsWith('='), [s0.cellData[0][1], s0.cellData[0][2]])
const txt = `SHORTCUTS + XLSX ${new Date().toISOString()}\n` + out.join('\n') + '\n'
console.log(txt); writeFileSync('test_results/shortcuts_2026-10-14.txt', txt); await app.close()
