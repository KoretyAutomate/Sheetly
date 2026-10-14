import { _electron as electron } from 'playwright'
const app = await electron.launch({ args: ['.', '--no-sandbox'] })
const win = await app.firstWindow()
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
const r = await win.evaluate(async () => {
  const ws = window.univerAPI.getActiveWorkbook().getActiveSheet()
  ws.getRange('A1').setValue(5); ws.getRange('A2').setValue(7)
  ws.getRange('B1').setFormula('=SUM(A1:A2)')
  ws.getRange('B2').setFormula('=IF(A1>3,"big","small")')
  await new Promise((r) => setTimeout(r, 1500))
  return [ws.getRange('B1').getValue(), ws.getRange('B2').getValue()]
})
console.log('RESULT', JSON.stringify(r))
await win.screenshot({ path: 'test_results/smoke.png' })
await app.close()
