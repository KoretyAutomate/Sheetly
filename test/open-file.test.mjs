// Opening workbooks named on the command line, and a second launch handing its file to the running Sheetly
// (a new window opens for it; a workbook that is already open is brought forward instead).
import { _electron as electron } from 'playwright'
import { mkdtempSync, writeFileSync } from 'fs'
import { spawnSync } from 'child_process'
import { tmpdir } from 'os'
import path from 'path'
import { snapshotToXlsx } from '../src/xlsx.js'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const dir = mkdtempSync(path.join(tmpdir(), 'sheetly-open-'))
const book = async (name, text) => {
  const p = path.join(dir, name)
  writeFileSync(p, await snapshotToXlsx({ sheetOrder: ['s1'], sheets: { s1: { id: 's1', name: 'Data', cellData: { 0: { 0: { v: text, t: 1 } } } } } }))
  return p
}
const a = await book('first.xlsx', 'from first'), b = await book('second.xlsx', 'from second')
const env = { ...process.env, SHEETLY_USER_DATA: path.join(dir, 'profile') }
delete env.SHEETLY_TEST_NO_DIALOGS
const value = (w) => w.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getRange('A1').getValue())
const ready = (w) => w.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.()?.getActiveSheet().getRange('A1').getValue(), null, { timeout: 60000 })

const app = await electron.launch({ args: ['.', '--no-sandbox', a], env })
const w1 = await app.firstWindow()
await ready(w1)
check('a workbook named on the command line opens at start', (await value(w1)) === 'from first', String(await value(w1)))
check('its sheet name comes from the file', (await w1.evaluate(() => window.univerAPI.getActiveWorkbook().getActiveSheet().getSheetName())) === 'Data')

// second launch with another file: this process exits, the running one opens a window for the file
const second = spawnSync('npx', ['electron', '.', '--no-sandbox', b], { env, timeout: 60000 })
check('the second launch exits (hands over to the running Sheetly)', second.status === 0, `status ${second.status} ${second.stderr}`)
const w2 = await app.waitForEvent('window', { timeout: 30000 })
await ready(w2)
check('the handed-over workbook opens in a new window', (await value(w2)) === 'from second', String(await value(w2)))
// launching the first file again focuses its window, no third window
spawnSync('npx', ['electron', '.', '--no-sandbox', a], { env, timeout: 60000 })
await new Promise((r) => setTimeout(r, 2000))
check('reopening an open workbook brings its window forward (no new window)', app.windows().length === 2, String(app.windows().length))
await app.close()
// a file that cannot be read: the window says so and does not take the file (Save would otherwise overwrite it)
const bad = path.join(dir, 'broken.xlsx'); writeFileSync(bad, 'not a workbook')
const env2 = { ...env, SHEETLY_USER_DATA: path.join(dir, 'profile2') }
const app2 = await electron.launch({ args: ['.', '--no-sandbox', bad], env: env2 })
const w3 = await app2.firstWindow()
await w3.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
await w3.waitForTimeout(2000)
const msg = await w3.evaluate(() => document.querySelector('.sx-dialog p')?.textContent ?? '')
const where = await w3.evaluate(() => window.sheetly.directory())
check('a file that cannot be read: a message, and the window does not take it as its file', msg.startsWith("Sheetly cannot open the file 'broken.xlsx'") && !where.startsWith(dir), `${msg} | ${where}`)
await app2.close()
const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`open file from the command line  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
