// Drives real key presses in the Electron app and checks the Excel-defined effect of each shortcut.
import { _electron as electron } from 'playwright'
import { writeFileSync } from 'fs'
const app = await electron.launch({ args: ['.', '--no-sandbox'] })
const win = await app.firstWindow()
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
await win.mouse.click(300, 300)

// Page-side helpers, injected as a string so every test snippet can use them.
const H = `
const api = window.univerAPI, wb = () => api.getActiveWorkbook(), S = () => wb().getActiveSheet()
const sel = () => S().getActiveRange()?.getA1Notation()
const v = (a) => { const r = S().getRange(a), x = r.getValue(); if (x != null && x !== '') return x
  const p = r.getCellData()?.p?.body?.dataStream; return p ? p.replace(/\\r\\n$/, '').replace(/\\r/g, '\\n') : x }
const f = (a) => S().getRange(a).getFormula()
const act = (a) => { S().getRange(a).activate(); return true }
const st = (a) => S().getRange(a).getCellStyleData() || {}
const dlg = () => document.querySelectorAll('[role="dialog"],[class*="dialog"],[class*="modal"],[class*="sidebar"],[class*="find-replace"],[class*="popup"]').length
`
const run = (code) => win.evaluate(`(async()=>{${H}\n${code}})()`)
const reset = () => run(`
  const w = wb(); w.getSheets().slice(1).forEach((s) => { try { w.deleteSheet(s) } catch (e) {} })
  w.setActiveSheet(w.getSheets()[0]); const ws = S()
  try { ws.getFilter()?.remove() } catch (e) {}
  try { ws.showColumns(0, 12); ws.showRows(0, 60) } catch (e) {}
  ws.getRange(0, 0, 60, 12).clear()
  for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) ws.getRange(r, c).setValue(r * 4 + c + 1)
  ws.getRange('A1').activate(); window.__dlg0 = dlg()`)
const out = [], tests = []
// T(name, setup, keys, check, [opts]) ; keys: 'Control+Home' = press, 'T:abc' = type text
const T = (name, setup, keys, check) => tests.push({ name, setup, keys, check })

// ---- Navigation
T('Arrow Right', `act('A1')`, ['ArrowRight'], `sel()==='B1'`)
T('Arrow Down', `act('A1')`, ['ArrowDown'], `sel()==='A2'`)
T('Tab moves right', `act('A1')`, ['Tab'], `sel()==='B1'`)
T('Shift+Tab moves left', `act('B1')`, ['Shift+Tab'], `sel()==='A1'`)
T('Enter moves down', `act('A1')`, ['Enter'], `sel()==='A2'`)
T('Shift+Enter moves up', `act('A2')`, ['Shift+Enter'], `sel()==='A1'`)
T('Ctrl+Right to edge of data', `act('A1')`, ['Control+ArrowRight'], `sel()==='D1'`)
T('Ctrl+Down to edge of data', `act('A1')`, ['Control+ArrowDown'], `sel()==='A4'`)
T('Ctrl+Left', `act('D1')`, ['Control+ArrowLeft'], `sel()==='A1'`)
T('Ctrl+Up', `act('A4')`, ['Control+ArrowUp'], `sel()==='A1'`)
T('Home -> start of row', `act('C3')`, ['Home'], `sel()==='A3'`)
T('Ctrl+Home -> A1', `act('C3')`, ['Control+Home'], `sel()==='A1'`)
T('Ctrl+End -> last used cell', `act('A1')`, ['Control+End'], `sel()==='D4'`)
T('PageDown moves down a screen', `act('A1')`, ['PageDown'], `S().getActiveRange().getRow()>5`)
T('PageUp moves up a screen', `act('A60')`, ['PageUp'], `S().getActiveRange().getRow()<59`)
T('Alt+PageDown moves right a screen', `act('A1')`, ['Alt+PageDown'], `S().getActiveRange().getColumn()>3`)
T('Ctrl+PageDown next sheet', `wb().create('S2',10,10); wb().setActiveSheet(wb().getSheets()[0]); act('A1')`, ['Control+PageDown'], `wb().getActiveSheet().getSheetName()==='S2'`)
T('Ctrl+PageUp previous sheet', `wb().create('S2',10,10); wb().setActiveSheet(wb().getSheets()[1]); act('A1')`, ['Control+PageUp'], `wb().getActiveSheet().getSheetName()!=='S2'`)
T('F5 opens Go To', `act('A1')`, ['F5'], `dlg()>window.__dlg0`)
T('Ctrl+G opens Go To', `act('A1')`, ['Control+g'], `dlg()>window.__dlg0`)
T('Ctrl+F opens Find', `act('A1')`, ['Control+f'], `dlg()>window.__dlg0`)
T('Ctrl+H opens Replace', `act('A1')`, ['Control+h'], `dlg()>window.__dlg0`)
// ---- Selection
T('Ctrl+A selects current region', `act('B2')`, ['Control+a'], `sel()==='A1:D4'`)
T('Ctrl+Space selects column', `act('B2')`, ['Control+Space'], `/^B1:B\\d{4,}$/.test(sel())||sel()==='B:B'`)
T('Shift+Space selects row', `act('B2')`, ['Shift+Space'], `/^A2:[A-Z]{1,3}2$/.test(sel())||sel()==='2:2'`)
T('Shift+Right extends', `act('A1')`, ['Shift+ArrowRight'], `sel()==='A1:B1'`)
T('Shift+Down extends', `act('A1')`, ['Shift+ArrowDown'], `sel()==='A1:A2'`)
T('Ctrl+Shift+Right to edge', `act('A1')`, ['Control+Shift+ArrowRight'], `sel()==='A1:D1'`)
T('Ctrl+Shift+Down to edge', `act('A1')`, ['Control+Shift+ArrowDown'], `sel()==='A1:A4'`)
T('Ctrl+Shift+End to last cell', `act('A1')`, ['Control+Shift+End'], `sel()==='A1:D4'`)
T('Ctrl+Shift+Home to start', `act('C3')`, ['Control+Shift+Home'], `sel()==='A1:C3'`)
T('Ctrl+Shift+* current region', `act('B2')`, ['Control+Shift+8'], `sel()==='A1:D4'`)
T('Shift+Backspace collapses to active cell', `S().getRange('A1:B2').activate()`, ['Shift+Backspace'], `sel()==='A1'`)
T('F8 extend mode', `act('A1')`, ['F8','ArrowRight','ArrowRight','F8'], `sel()==='A1:C1'`)
// ---- Editing
T('F2 edit cell', `act('A1')`, ['F2','T:9','Enter'], `String(v('A1'))==='19'`)
T('Esc cancels edit', `act('A1')`, ['T:99','Escape'], `String(v('A1'))==='1'`)
T('Typing + Enter commits', `act('F1')`, ['T:x','Enter'], `v('F1')==='x'`)
T('Typing + Tab commits, moves right', `act('F1')`, ['T:x','Tab'], `v('F1')==='x'&&sel()==='G1'`)
T('Alt+Enter new line in cell', `act('F1')`, ['T:a','Alt+Enter','T:b','Enter'], `v('F1')==='a\\nb'`)
T('Ctrl+Enter fills selection', `S().getRange('F1:F3').activate()`, ['T:7','Control+Enter'], `v('F1')==7&&v('F2')==7&&v('F3')==7`)
T('Ctrl+Shift+Enter array formula', `act('F1')`, ['T:=ROW(A1:A2)','Control+Shift+Enter'], `v('F2')==2`)
T('Delete clears cell', `act('A1')`, ['Delete'], `v('A1')===''||v('A1')==null`)
T('Backspace clears cell', `act('A1')`, ['Backspace','Enter'], `v('A1')===''||v('A1')==null`)
T('Ctrl+Z undo', `act('F1')`, ['T:x','Enter','Control+z'], `v('F1')===''||v('F1')==null`)
T('Ctrl+Y redo', `act('F1')`, ['T:x','Enter','Control+z','Control+y'], `v('F1')==='x'`)
T('F4 repeats last action', `act('A1')`, ['Control+b','ArrowRight','F4'], `st('B1').bl===1`)
T('F4 toggles absolute ref while editing', `act('C1')`, ['T:=A1','F4','Enter'], `f('C1')==='=$A$1'`)
T('Ctrl+D fill down', `S().getRange('A1:A3').activate()`, ['Control+d'], `v('A2')==1&&v('A3')==1`)
T('Ctrl+R fill right', `S().getRange('A1:C1').activate()`, ['Control+r'], `v('B1')==1&&v('C1')==1`)
T('Ctrl+; inserts date', `act('F1')`, ['Control+;'], `v('F1')!==''&&v('F1')!=null`)
T('Ctrl+Shift+; inserts time', `act('F1')`, ['Control+Shift+;'], `v('F1')!==''&&v('F1')!=null`)
T("Ctrl+' copies formula from above", `S().getRange('F1').setFormula('=A1+1'); act('F2')`, ["Control+'",'Enter'], `f('F2')==='=A1+1'`)
T('Ctrl+Shift+" copies value from above', `S().getRange('F1').setValue(42); act('F2')`, ['Control+Shift+"','Enter'], `v('F2')==42`)
T('Ctrl+C / Ctrl+V', `act('A1')`, ['Control+c','ArrowRight','ArrowRight','ArrowRight','ArrowRight','ArrowRight','Control+v'], `v('F1')==1`)
T('Ctrl+X / Ctrl+V', `act('A1')`, ['Control+x','ArrowRight','ArrowRight','ArrowRight','ArrowRight','ArrowRight','Control+v'], `v('F1')==1&&(v('A1')===''||v('A1')==null)`)
T('Ctrl+Alt+V paste special dialog', `act('A1')`, ['Control+c','Control+Alt+v'], `dlg()>window.__dlg0`)
T('Ctrl+Shift+= inserts row', `act('A2')`, ['Shift+Space','Control+Shift+='], `(v('A2')===''||v('A2')==null)&&v('A3')==5`)
T('Ctrl+- deletes row', `act('A2')`, ['Shift+Space','Control+-'], `v('A2')==9`)
T('Ctrl+9 hides row', `act('A2')`, ['Control+9'], `!S().getSheet().getRowVisible(1)`)
T('Ctrl+0 hides column', `act('B1')`, ['Control+0'], `!S().getSheet().getColVisible(1)`)
T('Ctrl+Shift+9 unhides rows', `S().hideRows(1,1); S().getRange('A1:A3').activate()`, ['Control+Shift+9'], `S().getSheet().getRowVisible(1)`)
T('Ctrl+Shift+0 unhides columns', `S().hideColumns(1,1); S().getRange('A1:C1').activate()`, ['Control+Shift+0'], `S().getSheet().getColVisible(1)`)
T('Alt+= AutoSum', `act('A5')`, ['Alt+='], `/SUM/i.test(f('A5'))`)
T('Shift+F2 insert comment', `act('A1')`, ['Shift+F2'], `dlg()>window.__dlg0`)
T('Ctrl+Shift+F3 create names', `S().getRange('A1:B2').activate()`, ['Control+Shift+F3'], `dlg()>window.__dlg0`)
T('Ctrl+K hyperlink', `act('A1')`, ['Control+k'], `dlg()>window.__dlg0`)
T('F11 chart', `act('A1')`, ['F11'], `dlg()>window.__dlg0||wb().getSheets().length>1`)
T('Ctrl+T create table', `S().getRange('A1:D4').activate()`, ['Control+t'], `dlg()>window.__dlg0`)
T('Ctrl+Shift+L toggle filter', `act('A1')`, ['Control+Shift+l'], `!!S().getFilter()`)
T('Shift+F11 new sheet', `act('A1')`, ['Shift+F11'], `wb().getSheets().length===2`)
T('Ctrl+P print', `act('A1')`, ['Control+p'], `dlg()>window.__dlg0`)
T('F1 help', `act('A1')`, ['F1'], `dlg()>window.__dlg0`)
// ---- In-cell editing keys
T('Home in edit', `S().getRange('F1').setValue('abc'); act('F1')`, ['F2','Home','T:x','Enter'], `v('F1')==='xabc'`)
T('Ctrl+Left word jump in edit', `S().getRange('F1').setValue('abc def'); act('F1')`, ['F2','Control+ArrowLeft','T:X','Enter'], `v('F1')==='abc Xdef'`)
T('Ctrl+Backspace deletes word in edit', `S().getRange('F1').setValue('abc def'); act('F1')`, ['F2','Control+Backspace','Enter'], `v('F1')==='abc '`)
T('Ctrl+Delete deletes to word end in edit', `S().getRange('F1').setValue('abc def'); act('F1')`, ['F2','Home','Control+Delete','Enter'], `v('F1')===' def'`)
// ---- Formatting
T('Ctrl+B bold', `act('A1')`, ['Control+b'], `st('A1').bl===1`)
T('Ctrl+I italic', `act('A1')`, ['Control+i'], `st('A1').it===1`)
T('Ctrl+U underline', `act('A1')`, ['Control+u'], `st('A1').ul?.s===1`)
T('Ctrl+5 strikethrough', `act('A1')`, ['Control+5'], `st('A1').st?.s===1`)
T('Ctrl+Shift+~ General', `S().getRange('A1').setNumberFormat('0.00'); act('A1')`, ['Control+Shift+`'], `S().getRange('A1').getNumberFormat()==='General'||S().getRange('A1').getNumberFormat()===''`)
T('Ctrl+Shift+! number #,##0.00', `act('A1')`, ['Control+Shift+1'], `/#,##0\\.00/.test(S().getRange('A1').getNumberFormat())`)
T('Ctrl+Shift+@ time', `act('A1')`, ['Control+Shift+2'], `/h:mm/.test(S().getRange('A1').getNumberFormat())`)
T('Ctrl+Shift+# date', `act('A1')`, ['Control+Shift+3'], `/d-mmm-yy/.test(S().getRange('A1').getNumberFormat())`)
T('Ctrl+Shift+$ currency', `act('A1')`, ['Control+Shift+4'], `/\\$#,##0\\.00/.test(S().getRange('A1').getNumberFormat())`)
T('Ctrl+Shift+% percent', `act('A1')`, ['Control+Shift+5'], `S().getRange('A1').getNumberFormat()==='0%'`)
T('Ctrl+Shift+^ scientific', `act('A1')`, ['Control+Shift+6'], `/E\\+/.test(S().getRange('A1').getNumberFormat())`)
T('Ctrl+Shift+& outline border', `act('A1')`, ['Control+Shift+7'], `!!st('A1').bd`)
T('Ctrl+Shift+_ remove border', `S().getRange('A1').setBorder?.('ALL','thin','#000000'); act('A1')`, ['Control+Shift+-'], `!st('A1').bd||Object.values(st('A1').bd).every((x)=>!x)`)
T('Ctrl+1 Format Cells', `act('A1')`, ['Control+1'], `dlg()>window.__dlg0`)
T('Ctrl+Shift+> font size up', `act('A1')`, ['Control+Shift+.'], `(st('A1').fs??11)>11`)
T('Ctrl+Shift+< font size down', `act('A1')`, ['Control+Shift+,'], `(st('A1').fs??11)<11`)

const t0 = Date.now()
for (const t of tests) {
  let ok = false, note = ''
  try {
    await reset(); await win.keyboard.press('Escape'); await win.waitForTimeout(120)
    await run(t.setup); await win.waitForTimeout(150)
    for (const k of t.keys) { if (k.startsWith('T:')) await win.keyboard.type(k.slice(2), { delay: 15 }); else await win.keyboard.press(k); await win.waitForTimeout(90) }
    await win.waitForTimeout(350)
    ok = !!(await run('return (' + t.check + ')'))
  } catch (e) { note = ' (error: ' + String(e.message).split('\n')[0].slice(0, 100) + ')' }
  out.push(`${ok ? 'PASS' : 'FAIL'}  ${t.name}${note}`)
  await win.keyboard.press('Escape'); await win.keyboard.press('Escape')
}
// File-menu accelerators (cannot be pressed in a test: they open native dialogs) — inspect the menu instead.
const accel = await app.evaluate(({ Menu }) => { const a = []; const walk = (m) => m.items.forEach((i) => { if (i.accelerator) a.push(i.accelerator); if (i.submenu) walk(i.submenu) }); walk(Menu.getApplicationMenu()); return a })
for (const [name, a] of [['Ctrl+O open', 'CmdOrCtrl+O'], ['Ctrl+S save', 'CmdOrCtrl+S'], ['F12 save as', 'F12'], ['Ctrl+N new workbook', 'CmdOrCtrl+N'], ['Ctrl+W close workbook', 'CmdOrCtrl+W']])
  out.push(`${accel.includes(a) ? 'PASS' : 'FAIL'}  ${name} (menu accelerator present; dialog not pressed in test)`)
const pass = out.filter((l) => l.startsWith('PASS')).length
const txt = [`SHORTCUT TEST (all) ${new Date().toISOString()}`, `pass ${pass}/${out.length}  (${((Date.now() - t0) / 1000) | 0}s)`, ...out,
  '', 'NOT AUTOMATED (no observable state to assert): Ctrl+` show formulas, F9 recalc, Ctrl+Backspace scroll-to-active, Ctrl+E flash fill, Alt+H ribbon key tips, F7 spelling, Alt+Down dropdown.'].join('\n')
console.log(txt); writeFileSync('test_results/shortcuts_full_2026-10-14.txt', txt + '\n'); await app.close()
