// PivotTables in the app: Alt N V T and the Create PivotTable dialog, the fields pane (check boxes, areas, menus),
// the report in cells, Value Field Settings, a report filter, Refresh (Alt+F5), GETPIVOTDATA, saved definitions.
import { _electron as electron } from 'playwright'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const app = await electron.launch({ args: ['.', '--no-sandbox'], env: { ...process.env, SHEETLY_TEST_NO_DIALOGS: '1' } })
const win = await app.firstWindow()
const errors = []
win.on('pageerror', (e) => errors.push(e.message))
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
await win.waitForTimeout(1500)
await win.mouse.click(700, 500)
const ev = (code) => win.evaluate(`(async()=>{const api=window.univerAPI,W=()=>api.getActiveWorkbook(),S=()=>W().getActiveSheet(),X=window.__sheetly,P=X.pivots;${code}})()`)
const settle = (ms = 400) => win.waitForTimeout(ms)
const key = async (k, n = 1) => { for (let i = 0; i < n; i++) await win.keyboard.press(k); await settle(150) }
const grid = (a1) => ev(`const r=S().getRange('${a1}'); return r.getValues().map(row=>row.map(v=>v==null?'':v).join('|')).join('\\n')`)

await ev(`S().setName?.('Data'); S().getRange('A1:C6').setValues([['Region','Product','Sales'],['East','A',10],['East','B',20],['West','A',5],['West','A',7],['North','B',3]]); S().getRange('B3').activate()`)
await settle()
// ---- Alt N V T opens Create PivotTable with the current region
await key('Alt'); await key('n'); await key('v'); await key('t'); await settle()
const dlg = await ev(`return [document.querySelector('.sx-dialog .sx-title span')?.textContent, document.querySelector('.sx-dialog input[type=text]')?.value].join('|')`)
check('Alt N V T opens Create PivotTable with the current region', dlg === 'Create PivotTable|Data!$A$1:$C$6' || dlg === "Create PivotTable|Sheet1!$A$1:$C$6", dlg)
await key('Enter'); await settle(800)
const where = await ev(`return [W().getSheets().map(s=>s.getSheetName()).join(','), S().getSheetName(), S().getActiveRange().getA1Notation(), !document.querySelector('.sx-pivot-pane').hidden].join('|')`)
check('the report goes on a new sheet to the left, at A3, and the fields pane opens', /^(Sheet\d),Data\|\1\|A3\|true$/.test(where), where)
const fields = await ev(`return [...document.querySelectorAll('.sx-pivot-field')].map(l=>l.textContent.trim()).join(',')`)
check('the pane lists the source fields', fields === 'Region,Product,Sales', fields)

// ---- check boxes: text to Rows, numbers to Values
await ev(`[...document.querySelectorAll('.sx-pivot-field')].find(l=>l.textContent.trim()==='Region').querySelector('input').click()`); await settle(600)
await ev(`[...document.querySelectorAll('.sx-pivot-field')].find(l=>l.textContent.trim()==='Sales').querySelector('input').click()`); await settle(600)
let g = await grid('A3:B7')
check('checking Region and Sales: Row Labels / Sum of Sales, sorted, Grand Total', g === 'Row Labels|Sum of Sales\nEast|30\nNorth|3\nWest|12\nGrand Total|45', '\n' + g)
check('header row bold on the light blue band', (await ev(`const s=S().getRange('A3').getCellStyleData(); return s?.bl===1 && /ddebf7/i.test(s?.bg?.rgb??'')`)) === true, JSON.stringify(await ev(`return S().getRange('A3').getCellStyleData()`)))
// ---- Product to Columns through its menu
await ev(`[...document.querySelectorAll('.sx-pivot-field')].find(l=>l.textContent.trim()==='Product').querySelector('input').click()`); await settle(600)
await ev(`const s=[...document.querySelectorAll('.sx-pivot-area[data-area=rows] .sx-pivot-item')].find(x=>x.textContent.startsWith('Product')).querySelector('select'); s.value='cols'; s.dispatchEvent(new Event('change'))`); await settle(700)
g = await grid('A3:D8')
check('Move to Column Labels: Column Labels header, items A, B, Grand Total column', g === 'Sum of Sales|Column Labels||\nRow Labels|A|B|Grand Total\nEast|10|20|30\nNorth||3|3\nWest|12||12\nGrand Total|22|23|45', '\n' + g)
// ---- GETPIVOTDATA on another cell
await ev(`S().getRange('G1').setFormula('=GETPIVOTDATA("Sales",$A$3,"Region","East","Product","B")'); S().getRange('G2').setFormula('=GETPIVOTDATA("Sum of Sales",A3)'); S().getRange('G3').setFormula('=GETPIVOTDATA("Sales",A3,"Region","South")')`); await settle(800)
g = await ev(`return [S().getRange('G1').getRawValue(), S().getRange('G2').getRawValue(), S().getRange('G3').getRawValue()].join('|')`)
check('GETPIVOTDATA reads the report (20, grand total 45, #REF! for an item not shown)', g === '20|45|#REF!', g)
// ---- Value Field Settings: Average
await ev(`S().getRange('B5').activate()`); await settle()
await ev(`P.valueFieldSettings(0)`); await settle()
await ev(`const l=document.querySelector('.sx-dialog select'); l.value='average'; l.dispatchEvent(new Event('change'))`)
await key('Enter'); await settle(800)
g = await ev(`return [S().getRange('A3').getRawValue(), S().getRange('D5').getRawValue(), S().getRange('D8').getRawValue()].join('|')`)
check('Value Field Settings > Average: caption "Average of Sales", East 15, total 9', g === 'Average of Sales|15|9', g)
await settle(500)
g = await ev(`return [S().getRange('G1').getRawValue(), S().getRange('G2').getRawValue()].join('|')`)
check('GETPIVOTDATA after the change: "Sales" still finds the field (East/B = 20), the old caption "Sum of Sales" is #REF!', g === '20|#REF!', g)
// ---- report filter: Region to Filters, pick West
await ev(`const s=[...document.querySelectorAll('.sx-pivot-area[data-area=rows] .sx-pivot-item')].find(x=>x.textContent.startsWith('Region')).querySelector('select'); s.value='filters'; s.dispatchEvent(new Event('change'))`); await settle(700)
await ev(`const s=document.querySelector('.sx-pivot-area[data-area=filters] select[aria-label="Region filter"]'); s.value=[...s.options].find(o=>o.dataset.item==='West').value; s.dispatchEvent(new Event('change'))`); await settle(700)
g = await grid('A1:B1') + ' / ' + await grid('A3:D5')
check('report filter: Region / West above the report; only West is summarised', g === 'Region|West / Average of Sales|Column Labels||\n|A|Grand Total|\n|6|6|', g)
// ---- Refresh: change the source, Alt+F5
await ev(`P.setFilter(0, null)`); await settle(600)
await ev(`W().getSheets()[1].getRange('C2').setValue(100)`); await settle(300)
const before = await ev(`return S().getRange('B5').getRawValue()`)
await ev(`S().getRange('H1').setFormula('=GETPIVOTDATA("Sales",A3,"Product","A")')`); await settle(500)
const gpdBefore = await ev(`return S().getRange('H1').getRawValue()`)
await ev(`S().getRange('A4').activate()`); await settle()
await key('Alt+F5'); await settle(800)
const after = await ev(`return S().getRange('B5').getRawValue()`)
check('the report does not follow the source until Alt+F5 refreshes it', before !== after && after === (100 + 5 + 7) / 3, `${before} -> ${after}`)
check('until then GETPIVOTDATA answers from the report as shown, not the changed source', gpdBefore === before, `${gpdBefore} vs ${before}`)
// ---- saved with the workbook
const saved = await ev(`const s=W().save(); return JSON.stringify(s.sheets[S().getSheetId()].custom?.sheetlyPivots?.[0] ?? null)`)
check('the PivotTable definition is saved with the sheet', saved.includes('"name":"PivotTable1"') && saved.includes('"fn":"average"'), saved.slice(0, 200))
// ---- the pane follows the active cell
await ev(`S().getRange('J20').activate()`); await settle(400)
const hidden = await ev(`return document.querySelector('.sx-pivot-pane').hidden`)
await ev(`S().getRange('A4').activate()`); await settle(400)
const shown = await ev(`return document.querySelector('.sx-pivot-pane').hidden`)
check('the fields pane closes outside the PivotTable and comes back inside it', hidden === true && shown === false, `${hidden} ${shown} ${JSON.stringify(await ev(`return P.list().map(p=>[p.sheet,p.extent])`))} ${await ev(`return S().getSheetName()`)}`)
// ---- an existing-sheet destination: unrelated cells are never cleared; growing over data asks first
await ev(`W().setActiveSheet(W().getSheets()[1]); S().getRange('H20').setValue('keep me'); S().getRange('G5').setValue('in the way')`); await settle()
await ev(`P.create(S().getSheetName() + '!$A$1:$C$6', 'existing', S().getSheetName() + '!F3')`); await settle(600)
await ev(`P.toggleField('Region', true)`); await settle(600)
check('adding a field does not clear cells the PivotTable never wrote (H20 kept)', (await ev(`return S().getRange('H20').getRawValue()`)) === 'keep me')
await ev(`P.toggleField('Sales', true)`); await settle(600)
const asked = await ev(`return document.querySelector('.sx-dialog p')?.textContent ?? ''`)
check('growing over other data asks "There\'s already data in [...]. Do you want to replace it?"', /^There's already data in \[.+\]\. Do you want to replace it\?$/.test(asked), asked)
await key('Escape'); await settle(400)
check('Cancel leaves the data and the report as they were', (await ev(`return S().getRange('G5').getRawValue()`)) === 'in the way' && (await ev(`return S().getRange('G3').getRawValue()`)) == null)
// ---- inserted rows: source and report move, Refresh still reads the moved data
await ev(`S().getRange('G5').clear(); S().getRange('F4').activate(); P.toggleField('Sales', true)`); await settle(800)
const total0 = await ev(`return S().getRange('G7').getRawValue()`)
await ev(`S().insertRowsBefore(0, 2)`); await settle(600)
await ev(`S().getRange('F6').activate()`); await settle(400)
const inside = await ev(`return P.active()?.name ?? null`)
await key('Alt+F5'); await settle(800)
const total1 = await ev(`return S().getRange('G9').getRawValue()`)
check('after inserting rows above source and report, the report is found and Refresh reads the moved source', inside !== null && total0 === total1 && typeof total1 === 'number', `${inside} ${total0} ${total1}`)

// ---- deleting a row inside the report shrinks what it owns: data moved up into its old area is safe
await ev(`S().getRange('F10').setValue('below')`); await settle()
await ev(`S().deleteRows(6, 1)`); await settle(600)
await ev(`S().getRange('F6').activate()`); await settle(300)
await key('Alt+F5'); await settle(800)
const asked2 = await ev(`return document.querySelector('.sx-dialog p')?.textContent ?? ''`)
check('after a row inside the report is deleted, Refresh asks before overwriting the cell that moved up', asked2.startsWith("There's already data in") && (await ev(`return S().getRange('F9').getRawValue()`)) === 'below', `${asked2} | ${await ev(`return S().getRange('F9').getRawValue()`)}`)
await key('Escape'); await settle(300)
// ---- Copy Sheet: the copy's PivotTable is found on the copy
const ids0 = await ev(`return W().getSheets().map(s=>s.getSheetId()).join(',')`)
await ev(`await api._injector.get(window.__ICommandService).executeCommand('sheet.command.copy-sheet', { unitId: W().getId(), subUnitId: S().getSheetId() })`); await settle(800)
const copy = await ev(`const before='${ids0}'.split(','); const c=W().getSheets().find(s=>!before.includes(s.getSheetId())); W().setActiveSheet(c); await new Promise(r=>setTimeout(r,300)); S().getRange('F6').activate(); await new Promise(r=>setTimeout(r,300)); return [S().getSheetName(), P.active()?.name ?? null].join('|')`)
check('a copied sheet\'s PivotTable is found on the copy', /\|PivotTable\d+$/.test(copy) && !copy.startsWith('Sheet1|'), copy)

// ---- one undo step per change: Ctrl+Z puts back the report and its definition together
await ev(`W().setActiveSheet(W().getSheets()[0]); S().getRange('A4').activate()`); await settle(400)
const def0 = await ev(`return JSON.stringify(P.active() && [P.active().rows, P.active().values.map(v=>v.fn)])`), a30 = await grid('A3:B4')
await ev(`P.toggleField('Product', true)`); await settle(700)
const changed = await grid('A3:B4')
await ev(`await api._injector.get(window.__ICommandService).executeCommand('univer.command.undo')`); await settle(700)
const def1 = await ev(`S().getRange('A4').activate(); await new Promise(r=>setTimeout(r,200)); return JSON.stringify(P.active() && [P.active().rows, P.active().values.map(v=>v.fn)])`)
check('Ctrl+Z undoes a field change in one step (cells and definition)', changed !== a30 && (await grid('A3:B4')) === a30 && def1 === def0, `${def0} ${def1}`)
// ---- a table name works as the source
await ev(`W().setActiveSheet(W().getSheets()[W().getSheets().length-1]); S().getRange('K1:L3').setValues([['Item','Qty'],['x',1],['y',2]]); await S().addTable('Stock', { startRow: 0, endRow: 2, startColumn: 10, endColumn: 11 })`); await settle(600)
const okTable = await ev(`return P.create('Stock', 'new', '') && P.active()?.name`)
check('Create PivotTable accepts a table name ("Stock")', typeof okTable === 'string' && okTable.startsWith('PivotTable'), String(okTable))
await settle(400)

// ---- report filters own two columns even when the report is one column wide
await ev(`W().setActiveSheet(W().getSheets()[W().getSheets().length-1]); S().getRange('Q1').setValue('neighbour')`); await settle()
await ev(`P.create(W().getSheets()[W().getSheets().length-1].getSheetName() + '!$K$1:$L$3', 'existing', W().getSheets()[W().getSheets().length-1].getSheetName() + '!P3')`); await settle(600)
await ev(`P.place('Qty', 'values')`); await settle(500)
await ev(`P.place('Item', 'filters')`); await settle(600)
const asked3 = await ev(`return document.querySelector('.sx-dialog p')?.textContent ?? ''`)
check('a filter that would overwrite the cell beside it asks first', asked3.startsWith("There's already data in") && (await ev(`return S().getRange('Q1').getRawValue()`)) === 'neighbour', asked3)
await key('Escape'); await settle(300)

// ---- invalid source
await ev(`W().setActiveSheet(W().getSheets()[1]); S().getRange('E1:F3').setValues([['X',null],[1,2],[3,4]]); S().getRange('E1:F3').activate()`); await settle()
await ev(`P.createDialog()`); await settle(); await key('Enter'); await settle()
check('a source with a blank header is refused with Excel\'s message', (await ev(`return [...document.querySelectorAll('.sx-dialog p')].some(p=>p.textContent.startsWith('The PivotTable field name is not valid'))`)) === true)
await key('Escape'); await key('Escape')

check('no page errors', errors.length === 0, errors.join(' | '))
await app.close()
const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`pivot tables  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
