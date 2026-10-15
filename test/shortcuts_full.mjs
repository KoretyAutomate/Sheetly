// Drives real key presses in the Electron app and checks the Excel-defined effect of each shortcut.
import { _electron as electron } from 'playwright'
import { writeFileSync } from 'fs'
const app = await electron.launch({ args: ['.', '--no-sandbox'], env: { ...process.env, SHEETLY_TEST_NO_DIALOGS: '1' } })
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
const dlg = () => document.querySelectorAll('[role="dialog"],[class*="dialog"],[class*="modal"],[class*="sidebar"],aside,[class*="find-replace"],[class*="popup"]').length
`
const run = (code) => win.evaluate(`(async()=>{${H}\n${code}})()`)
const reset = () => run(`
  try { await wb().endEditingAsync?.(false) } catch (e) {}
  // Keep the original sheet (by id): new sheets may have been inserted before it.
  const w = wb(); window.__home ??= w.getSheets()[0].getSheetId()
  const home = w.getSheetBySheetId(window.__home) ?? w.getSheets()[0]
  w.setActiveSheet(home)
  w.getSheets().filter((s) => s.getSheetId() !== home.getSheetId()).forEach((s) => { try { w.deleteSheet(s) } catch (e) {} })
  const ws = S()
  try { ws.getFilter()?.remove() } catch (e) {}
  // Workbook-level objects that clear() does not remove: tables, names, notes, charts.
  for (const t of w.getTableList?.() ?? []) { try { await w.removeTable(t.id) } catch (e) {} }
  for (const n of w.getDefinedNames?.() ?? []) { try { w.deleteDefinedName(n.getName()) } catch (e) {} }
  for (const n of ws.getNotes?.() ?? []) { try { ws.getRange(n.row, n.col).deleteNote() } catch (e) {} }
  document.querySelectorAll('.sx-chart').forEach((c) => c.closest('[class*="float"]')?.remove?.())
  if (window.__sheetly.extras.isShowingFormulas()) window.__sheetly.extras.toggleShowFormulas()
  try { ws.showColumns(0, 12); ws.showRows(0, 60) } catch (e) {}
  ws.getRange(0, 0, 60, 12).clear()
  for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) ws.getRange(r, c).setValue(r * 4 + c + 1)
  ws.getRange('A1').activate(); window.__dlg0 = dlg()`)
const out = [], tests = []
// T(name, setup, keys, check, [opts]) ; keys: 'Control+Home' = press, 'T:abc' = type text
const T = (name, setup, keys, check, opts = {}) => tests.push({ name, setup, keys, check, ...opts })

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
T('F5 Go To C3', `act('A1')`, ['F5','T:C3','Enter'], `sel()==='C3'`)
T('Ctrl+G Go To B2:C3', `act('A1')`, ['Control+g','T:B2:C3','Enter'], `sel()==='B2:C3'`)
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
T('Ctrl+Shift+@ time displays en-US AM/PM', `S().getRange('F1').setValue(0.75); act('F1')`, ['Control+Shift+2'], `/^6:00 PM$/.test(S().getRange('F1').getDisplayValue())`)
T("Ctrl+' copies formula from above", `S().getRange('F1').setFormula('=A1+1'); act('F2')`, ["Control+'",'Enter'], `f('F2')==='=A1+1'`)
T('Ctrl+Shift+" copies value from above', `S().getRange('F1').setValue(42); act('F2')`, ['Control+Shift+"','Enter'], `v('F2')==42`)
T('Ctrl+C / Ctrl+V', `act('A1')`, ['Control+c','ArrowRight','ArrowRight','ArrowRight','ArrowRight','ArrowRight','Control+v'], `v('F1')==1`)
T('Ctrl+X / Ctrl+V', `act('A1')`, ['Control+x','ArrowRight','ArrowRight','ArrowRight','ArrowRight','ArrowRight','Control+v'], `v('F1')==1&&(v('A1')===''||v('A1')==null)`)
T('Ctrl+Alt+V paste special (All)', `act('A1')`, ['Control+c','ArrowRight','ArrowRight','ArrowRight','ArrowRight','ArrowRight','Control+Alt+v','Enter'], `v('F1')==1`)
T('Ctrl+Shift+= inserts row', `act('A2')`, ['Shift+Space','Control+Shift+='], `(v('A2')===''||v('A2')==null)&&v('A3')==5`)
T('Ctrl+- deletes row', `act('A2')`, ['Shift+Space','Control+-'], `v('A2')==9`)
T('Ctrl+9 hides row', `act('A2')`, ['Control+9'], `!S().getSheet().getRowVisible(1)`)
T('Ctrl+0 hides column', `act('B1')`, ['Control+0'], `!S().getSheet().getColVisible(1)`)
T('Ctrl+Shift+9 unhides rows', `S().hideRows(1,1); S().getRange('A1:A3').activate()`, ['Control+Shift+9'], `S().getSheet().getRowVisible(1)`)
T('Ctrl+Shift+0 unhides columns', `S().hideColumns(1,1); S().getRange('A1:C1').activate()`, ['Control+Shift+0'], `S().getSheet().getColVisible(1)`)
T('Alt+= AutoSum', `act('A5')`, ['Alt+='], `/SUM/i.test(f('A5'))`)
T('Shift+F2 insert note', `act('A1')`, ['Shift+F2'], `[...document.querySelectorAll('section textarea')].length>0`)
T('Ctrl+Shift+F3 create names', `S().getRange('A1').setValue('qty'); S().getRange('B1').setValue('cost'); S().getRange('A1:B4').activate()`, ['Control+Shift+F3','Enter'], `!!wb().getDefinedName('qty')&&!!wb().getDefinedName('cost')`)
T('Ctrl+K hyperlink', `act('A1')`, ['Control+k'], `[...document.querySelectorAll('section input')].length>0`)
T('F11 chart sheet', `act('A1')`, ['F11'], `wb().getActiveSheet().getSheetName()==='Chart1'&&!!document.querySelector('.sx-chart canvas')`)
T('Ctrl+T create table', `act('B2')`, ['Control+t','Enter'], `(wb().getTableList?.()??[]).length===1`)
T('Ctrl+Shift+L toggle filter', `act('A1')`, ['Control+Shift+l'], `!!S().getFilter()`)
T('Shift+F11 new sheet', `act('A1')`, ['Shift+F11'], `wb().getSheets().length===2&&S().getMaxRows()===1048576&&S().getMaxColumns()===16384`)
T('Ctrl+P print preview', `act('A1')`, ['Control+p'], `!!document.querySelector('[aria-label="Print preview"] table td')`)
T('F1 help', `act('A1')`, ['F1'], `!!document.querySelector('[aria-label="Help"] td')`)
// ---- added: dialogs with outcomes, modes, selection helpers
T('Ctrl+Shift+= shift cells down', `act('A1')`, ['Control+Shift+=','Enter'], `(v('A1')===''||v('A1')==null)&&v('A2')==1`)
T('Ctrl+- shift cells up', `act('A1')`, ['Control+-','Enter'], `v('A1')==5`)
T('Ctrl+L create table', `act('B2')`, ['Control+l','Enter'], `(wb().getTableList?.()??[]).length===1`)
T('Alt+F1 embedded chart', `act('B2')`, ['Alt+F1'], `!!document.querySelector('.sx-chart canvas')`)
T('Alt+Shift+F1 new sheet', `act('A1')`, ['Alt+Shift+F1'], `wb().getSheets().length===2`)
T('Ctrl+F2 print preview', `act('A1')`, ['Control+F2'], `!!document.querySelector('[aria-label="Print preview"]')`)
T('Ctrl+F3 name manager', `act('A1')`, ['Control+F3'], `[...document.querySelectorAll('aside')].some((a)=>a.innerText.includes('Defined names'))`)
T('Ctrl+Shift+F Format Cells font tab', `act('A1')`, ['Control+Shift+f'], `document.querySelector('[role=dialog][aria-label="Format Cells"] [role=tab][aria-selected="true"]')?.textContent==='Font'`)
T('Ctrl+Shift+P Format Cells font tab', `act('A1')`, ['Control+Shift+p'], `document.querySelector('[role=dialog][aria-label="Format Cells"] [role=tab][aria-selected="true"]')?.textContent==='Font'`)
T('Ctrl+Shift+Space select all', `act('B2')`, ['Control+Shift+Space'], `S().getActiveRange().getRange().endRow>=1048575&&S().getActiveRange().getRange().endColumn>=16383`)
T('Shift+Home select to row start', `act('C2')`, ['Shift+Home'], `sel()==='A2:C2'`)
T('Ctrl+. next corner', `S().getRange('A1:B2').activate()`, ['Control+.'], `(()=>{const c=S().getSelection().getCurrentCell();return c.actualRow===0&&c.actualColumn===1})()`)
T('End then Right', `act('A1')`, ['End','ArrowRight'], `sel()==='D1'`)
T('F8 then Esc leaves extend mode', `act('A1')`, ['F8','Escape','ArrowRight'], `sel()==='B1'`)
T('Ctrl+Shift+O select notes', `S().getRange('B2').createOrUpdateNote?.({note:'hi'}); act('A1')`, ['Control+Shift+o'], `sel()==='B2'`)
T('Alt+; visible cells', `S().getRange('A1:D4').activate()`, ['Alt+;'], `sel()==='A1:D4'`)
T('Ctrl+[ precedents', `S().getRange('F1').setFormula('=B2+C3'); act('F1')`, ['Control+['], `S().getSelection().getActiveRangeList?.().length===2||sel()==='B2'`)
T('Ctrl+] dependents', `S().getRange('F1').setFormula('=A1*2'); act('A1')`, ['Control+]'], `sel()==='F1'`)
T('Ctrl+\\ row differences', `S().getRange('B1').setValue(1); S().getRange('A1:C1').activate()`, ['Control+\\'], `sel()==='C1'`)
T('F9 recalculates', `S().getRange('F1').setFormula('=RAND()'); await new Promise((r)=>setTimeout(r,400)); window.__r=S().getRange('F1').getValue()`, ['F9'], `typeof window.__r==='number'&&S().getRange('F1').getValue()!==window.__r`, { timeout: 10000 })
T('Alt+Down pick from list', `S().getRange('F1').setValue('x'); S().getRange('F2').setValue('y'); act('F3')`, ['Alt+ArrowDown'], `[...document.querySelectorAll('div')].some((n)=>n.children.length===0&&n.textContent==='y')`)
T('Ctrl+Alt+= zoom in', `S().zoom?.(1)`, ['Control+Alt+='], `S().getZoom()>1`)
T("Ctrl+' then Esc restores the cell", `S().getRange('F1').setFormula('=A1+1'); S().getRange('F2').setValue('keep'); act('F2')`, ["Control+'",'Escape'], `v('F2')==='keep'`)
T("Ctrl+' formula from above (edit mode)", `S().getRange('F1').setFormula('=A1+1'); act('F2')`, ["Control+'",'Enter'], `f('F2')==='=A1+1'`)
// ---- added: remaining Excel keys
T('Ctrl+` show formulas', `S().getRange('F1').setFormula('=A1+1'); act('A1')`, ['Control+`'], `window.__sheetly.extras.isShowingFormulas()&&S().getRange('F1').getDisplayValue()==='=A1+1'`)
T('Shift+F3 insert function', `act('F1')`, ['Shift+F3','T:SUM','Enter','W:600','T:A1:A4)','Enter'], `f('F1')==='=SUM(A1:A4)'&&v('F1')==28`)
T('F3 paste name', `wb().insertDefinedName('qty','Sheet1!$A$1:$A$4'); act('F1')`, ['F3','Enter','W:600','Enter'], `f('F1')==='=qty'`)
T('Ctrl+Shift+A argument names', `act('F1')`, ['T:=ROUND','Control+Shift+a','W:400'], `window.__sheetly.extras.editorText().startsWith('=ROUND(')`)
T('Ctrl+E flash fill', `S().getRange('H1').setValue('John Smith'); S().getRange('H2').setValue('Mary Jones'); S().getRange('H3').setValue('Ann Lee'); S().getRange('I1').setValue('John'); act('I2')`, ['Control+e'], `v('I2')==='Mary'&&v('I3')==='Ann'`)
T('F7 spelling', `S().getRange('F1').setValue('speling mistake'); act('A1')`, ['F7'], `[...document.querySelectorAll('[role=dialog] input')].some((i)=>i.value==='speling')`)
T('Shift+F10 context menu', `act('B2')`, ['Shift+F10'], `[...document.querySelectorAll('*')].some((n)=>n.children.length===0&&/^Paste Special$/.test(n.textContent.trim())&&!n.closest('[role=toolbar],header'))`)
T('Ctrl+Shift+U expands formula bar', `window.__fb0=document.querySelector('[data-u-comp="formula-bar"]').getBoundingClientRect().height; act('A1')`, ['Control+Shift+u'], `document.querySelector('[data-u-comp="formula-bar"]').getBoundingClientRect().height>window.__fb0`)
T('Ctrl+6 hide objects', `act('B2')`, ['Alt+F1','W:300','Control+6'], `getComputedStyle(document.querySelector('.sx-chart')).visibility==='hidden'`)
// ---- In-cell editing keys
T('Home in edit', `S().getRange('F1').setValue('abc'); act('F1')`, ['F2','Home','T:x','Enter'], `v('F1')==='xabc'`)
T('Ctrl+Left word jump in edit', `S().getRange('F1').setValue('abc def'); act('F1')`, ['F2','Control+ArrowLeft','T:X','Enter'], `v('F1')==='abc Xdef'`)
T('Ctrl+Backspace deletes word in edit', `S().getRange('F1').setValue('abc def'); act('F1')`, ['F2','Control+Backspace','Enter'], `v('F1')==='abc '`)
T('Ctrl+Delete deletes to end of line in edit', `S().getRange('F1').setValue('abc def'); act('F1')`, ['F2','Control+ArrowLeft','Control+Delete','Enter'], `v('F1')==='abc '`)
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
T('Ctrl+Shift+& outline border', `S().getRange('A1:B2').activate()`, ['Control+Shift+7'], `!!st('A1').bd?.t&&!!st('A1').bd?.l&&!!st('B2').bd?.b&&!!st('B2').bd?.r`)
T('Ctrl+Shift+_ remove border', `S().getRange('A1').setBorder?.('ALL','thin','#000000'); act('A1')`, ['Control+Shift+-'], `!st('A1').bd||Object.values(st('A1').bd).every((x)=>!x)`)
T('Ctrl+1 Format Cells', `act('A1')`, ['Control+1'], `!!document.querySelector('[role=dialog][aria-label="Format Cells"] [role=tab]')`)
T('Ctrl+Shift+> font size up', `act('A1')`, ['Control+Shift+.'], `(st('A1').fs??11)>11`)
T('Ctrl+Shift+< font size down', `act('A1')`, ['Control+Shift+,'], `(st('A1').fs??11)<11`)

// ONLY=<regex> runs a subset (and prints renderer errors) when debugging.
const only = process.env.ONLY ? new RegExp(process.env.ONLY) : null
if (only) { win.on('pageerror', (e) => console.log('PAGEERROR', e.message)); win.on('console', (m) => { if (m.type() === 'error') console.log('CONSOLE', m.text().slice(0, 200)) }) }
const t0 = Date.now()
for (const t of tests.filter((x) => !only || only.test(x.name))) {
  let ok = false, note = ''
  try {
    // Close any panel/dialog a previous test opened and give the grid keyboard focus again.
    for (let i = 0; i < 3; i++) await win.keyboard.press('Escape')
    await win.evaluate(() => document.querySelectorAll('[class*="dialog"] [class*="close"],[class*="sidebar"] [class*="close"]').forEach((b) => b.click()))
    await win.mouse.click(700, 560); await win.waitForTimeout(80)
    await reset(); await win.waitForTimeout(120)
    await run(t.setup); await win.waitForTimeout(150)
    for (const k of t.keys) { if (k.startsWith('T:')) await win.keyboard.type(k.slice(2), { delay: 15 }); else if (k.startsWith('W:')) await win.waitForTimeout(+k.slice(2)); else await win.keyboard.press(k); await win.waitForTimeout(90) }
    await win.waitForTimeout(350)
    // Poll the check for up to 2 s: some effects (recalculation, spills, panels) land asynchronously.
    let waited = 0
    for (; waited <= (t.timeout ?? 2000) && !ok; waited += 150) {
      ok = !!(await run('return (' + t.check + ')'))
      if (!ok) await win.waitForTimeout(150)
    }
  } catch (e) { note = ' (error: ' + String(e.message).split('\n')[0].slice(0, 100) + ')' }
  if (!ok) console.log('DIAG', t.name, await run(`return JSON.stringify({ sel: sel(), range: S().getActiveRange().getRange(), editing: wb().isCellEditing(), F1: S().getRange('F1').getCellData(), F2: S().getRange('F2').getCellData(), A1: S().getRange('A1').getCellData() })`))
  if (t.timeout) note += ` (${ok ? 'took ~' + waited + ' ms' : 'no change in ' + t.timeout + ' ms'})`
  out.push(`${ok ? 'PASS' : 'FAIL'}  ${t.name}${note}`)
  await win.keyboard.press('Escape'); await win.keyboard.press('Escape')
}
// File / window keys, pressed for real. Each must reach its main-process action; Ctrl+N and Ctrl+W must
// actually open and close a window. (Open/Save native dialogs are suppressed in test mode.)
if (!only) {
  const actions = () => app.evaluate(() => global.__appActions.slice())
  const count = () => app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)
  for (const [name, key, action] of [['Ctrl+O open', 'Control+o', 'open'], ['Ctrl+F12 open', 'Control+F12', 'open'], ['Ctrl+S save', 'Control+s', 'save'],
    ['Shift+F12 save', 'Shift+F12', 'save'], ['F12 save as', 'F12', 'saveAs'], ['Ctrl+Tab next window', 'Control+Tab', 'nextWindow'],
    ['Ctrl+F6 next window', 'Control+F6', 'nextWindow'], ['Ctrl+Shift+Tab previous window', 'Control+Shift+Tab', 'prevWindow'],
    ['Ctrl+F10 maximize', 'Control+F10', 'maximize']]) {
    for (let i = 0; i < 3; i++) await win.keyboard.press('Escape')
    await win.mouse.click(700, 560); await reset()
    const before = (await actions()).length
    await win.keyboard.press(key); await win.waitForTimeout(400)
    const got = (await actions()).slice(before)
    out.push(`${got.includes(action) ? 'PASS' : 'FAIL'}  ${name}`)
  }
  // Ctrl+N opens a window; Ctrl+W and Ctrl+F4 each close a (secondary) window. Closing the page the
  // key was sent to makes the press itself reject, so that rejection is expected.
  const closeWith = async (key) => {
    await win.mouse.click(700, 560); await reset()
    const n0 = await count()
    await win.keyboard.press('Control+n'); await win.waitForTimeout(2000)
    const n1 = await count()
    const w2 = app.windows().find((w) => w !== win && !w.isClosed())
    if (!w2) return [n0, n1, n1]
    await w2.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
    await w2.mouse.click(700, 560); await w2.waitForTimeout(300)
    await w2.keyboard.press(key).catch(() => {})
    await win.waitForTimeout(1500)
    return [n0, n1, await count()]
  }
  const [a0, a1, a2] = await closeWith('Control+w')
  out.push(`${a1 === a0 + 1 ? 'PASS' : 'FAIL'}  Ctrl+N opens a new workbook window (${a0} -> ${a1})`)
  out.push(`${a2 === a0 ? 'PASS' : 'FAIL'}  Ctrl+W closes the workbook window (${a1} -> ${a2})`)
  const [b0, b1, b2] = await closeWith('Control+F4')
  out.push(`${b1 === b0 + 1 && b2 === b0 ? 'PASS' : 'FAIL'}  Ctrl+F4 closes the workbook window (${b1} -> ${b2})`)
}
const pass = out.filter((l) => l.startsWith('PASS')).length
const txt = [`SHORTCUT TEST (all) ${new Date().toISOString()}`, `pass ${pass}/${out.length}  (${((Date.now() - t0) / 1000) | 0}s)`, ...out,
  '', 'NOT AUTOMATED: Ctrl+F9 minimize (window-manager dependent).'].join('\n')
console.log(txt); writeFileSync('test_results/shortcuts_full_' + new Date().toISOString().slice(0,10) + '.txt', txt + '\n'); await app.close()
