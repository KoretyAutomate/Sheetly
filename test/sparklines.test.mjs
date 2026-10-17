// Sparklines in the app: Alt N S L / O / W and the Create Sparklines dialog, drawing in the cells (pixels on the
// grid canvas change), redraw on data changes, Quick Analysis > Sparklines, Clear, and the .xlsx round trip.
import { _electron as electron } from 'playwright'
import JSZip from 'jszip'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const app = await electron.launch({ args: ['.', '--no-sandbox'], env: { ...process.env, SHEETLY_TEST_NO_DIALOGS: '1' } })
const win = await app.firstWindow()
const errors = []
win.on('pageerror', (e) => errors.push(e.message))
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
await win.waitForTimeout(1500)
await win.mouse.click(700, 500)
const ev = (code) => win.evaluate(`(async()=>{const api=window.univerAPI,W=()=>api.getActiveWorkbook(),S=()=>W().getActiveSheet(),X=window.__sheetly,SP=X.sparklines;${code}})()`)
const settle = (ms = 400) => win.waitForTimeout(ms)
const key = async (k, n = 1) => { for (let i = 0; i < n; i++) await win.keyboard.press(k); await settle(150) }
// A cell as it appears on screen (a screenshot of its box); equal images = same drawing.
const shot = async (a1) => {
  const box = await ev(`const canvas=[...document.querySelectorAll('canvas')].sort((a,b)=>b.width*b.height-a.width*a.height)[0], rc=canvas.getBoundingClientRect()
    const sk=api._injector.get(window.__IRenderManagerService).getRenderUnitById(W().getId()).with(window.__SheetSkeletonManagerService).getCurrentSkeleton()
    const g=S().getRange('${a1}').getRange(), cell=sk.getCellWithCoordByIndex(g.startRow,g.startColumn,false)
    return { x: rc.left+sk.rowHeaderWidth+cell.startX+2, y: rc.top+sk.columnHeaderHeight+cell.startY+2, width: cell.endX-cell.startX-4, height: cell.endY-cell.startY-4 }`)
  await settle(300)
  return (await win.screenshot({ clip: box })).toString('base64')
}
await ev(`S().getRange('A2:E4').setValues([[1,3,2,5,4],[5,4,3,2,1],[1,-1,1,-1,1]]); S().getRange('A2:E4').activate()`); await settle()
const blankF2 = await shot('F2'), blankF4 = await shot('F4'), blankF3 = await shot('F3')
await key('Alt'); await key('n'); await key('s'); await key('l'); await settle()
const title = await ev(`return document.querySelector('.sx-dialog .sx-title span')?.textContent + '|' + document.querySelector('.sx-dialog input[type=text]').value`)
check('Alt N S L opens Create Sparklines with the selection as the Data Range', title === 'Create Sparklines|A2:E4', title)
await win.keyboard.type('F2:F4'); await key('Enter'); await settle(700)
const groups = await ev(`return JSON.stringify(SP.groups().map(g=>[g.type, g.items.map(i=>[i.row,i.col,i.source.startRow,i.source.endColumn])]))`)
check('one line sparkline per row in F2:F4', groups === '[["line",[[1,5,1,4],[2,5,2,4],[3,5,3,4]]]]', groups)
const a = await shot('F2')
check('the sparkline is drawn in F2 (its cell looks different from before)', a !== blankF2)
await ev(`S().getRange('A2:E2').setValues([[5,5,5,5,5]])`); await settle(600)
const flat = await shot('F2')
check('changing the data redraws it (a flat line looks different, and is still drawn)', flat !== a && flat !== blankF2)
// mismatched shapes are refused with Excel's message
await ev(`SP.dialog('column')`); await settle()
await ev(`document.querySelectorAll('.sx-dialog input[type=text]')[1].value='G2:H3'`); await key('Enter'); await settle()
check('a location that is not one row or column is refused', (await ev(`return [...document.querySelectorAll('.sx-dialog p')].some(p=>p.textContent.startsWith('Location reference is not valid'))`)) === true)
await key('Escape'); await key('Escape')
// column and win/loss replace the line sparklines in their cells
await ev(`SP.add('column','A2:E2','F2'); SP.add('winloss','A4:E4','F4')`); await settle(600)
check('column and win/loss sparklines replace the line ones in their cells', (await ev(`return SP.groups().map(g=>g.type+':'+g.items.map(i=>i.row).join('')).join(',')`)) === 'line:2,column:1,winloss:3', await ev(`return SP.groups().map(g=>g.type+':'+g.items.map(i=>i.row).join('')).join(',')`))
check('win/loss is drawn', (await shot('F4')) !== blankF4)
// Quick Analysis > Sparklines
await ev(`S().getRange('A7:C8').setValues([[1,2,3],[3,2,1]]); S().getRange('A7:C8').activate()`); await settle()
await key('Control+q'); await key('s'); await key('Enter'); await settle(600)
check('Quick Analysis > Sparklines > Line puts them in the column to the right (D7:D8)', (await ev(`return SP.groups().some(g=>g.type==='line'&&g.items.some(i=>i.row===6&&i.col===3)&&g.items.some(i=>i.row===7&&i.col===3))`)) === true)
// saved in .xlsx and read back
const bytes = Buffer.from(await ev(`return Array.from(await X.xlsx.save())`))
const sheetXml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string')
check('saved as x14 sparkline groups', (sheetXml.match(/<x14:sparklineGroup\b/g) ?? []).length === 4 && /<x14:sparklineGroup type="column"/.test(sheetXml), (sheetXml.match(/<x14:sparklineGroup[^>]*>/g) ?? []).join(' '))
await ev(`await X.xlsx.load(new Uint8Array(${JSON.stringify([...bytes])}), 'spark.xlsx')`); await settle(1500)
check('opened again: the sparklines are back and drawn', (await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)) === 5 && (await shot('F3')) !== blankF3)
// structural edits move sparklines with their data
await ev(`S().insertRowsBefore(0, 1)`); await settle(500)
const moved = await ev(`return JSON.stringify(SP.groups().flatMap(g=>g.items.map(i=>[i.row,i.col,i.source.startRow])).filter(x=>x[1]===5).sort((a,b)=>a[0]-b[0]))`)
await ev(`S().insertColumnsBefore(0, 1)`); await settle(500)
const moved2 = await ev(`return JSON.stringify(SP.groups().flatMap(g=>g.items.map(i=>[i.row,i.col,i.source.startColumn,i.source.endColumn])).filter(x=>x[1]===6).sort((a,b)=>a[0]-b[0]))`)
check('a row inserted above moves sparklines and their data down; a column inserted left moves them right', moved === '[[2,5,2],[3,5,3],[4,5,4]]' && moved2 === '[[2,6,1,5],[3,6,1,5],[4,6,1,5]]', `${moved} ${moved2}`)
await ev(`S().deleteColumns(0, 1); S().deleteRows(0, 1)`); await settle(500)

// deleting the row of a sparkline and undoing brings it back
const n0 = await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)
await ev(`S().deleteRows(2, 1)`); await settle(500)
const n1 = await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)
await ev(`await api._injector.get(window.__ICommandService).executeCommand('univer.command.undo')`); await settle(600)
const n2 = await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)
check('a deleted sparkline row is gone, and Undo brings the sparkline back', n1 === n0 - 1 && n2 === n0, `${n0} ${n1} ${n2}`)

// Add and Clear are undo steps
const before = await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)
await ev(`SP.add('line','A2:E2','H2')`); await settle(400)
await ev(`await api._injector.get(window.__ICommandService).executeCommand('univer.command.undo')`); await settle(400)
const afterUndo = await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)
await ev(`await api._injector.get(window.__ICommandService).executeCommand('univer.command.redo')`); await settle(400)
const afterRedo = await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)
check('adding sparklines is one undo step (undo removes, redo puts back)', afterUndo === before && afterRedo === before + 1, `${before} ${afterUndo} ${afterRedo}`)
await ev(`S().getRange('H2').activate(); SP.clear()`); await settle(400)
await ev(`await api._injector.get(window.__ICommandService).executeCommand('univer.command.undo')`); await settle(400)
check('Clear can be undone', (await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)) === before + 1)
await ev(`S().getRange('H2').activate(); SP.clear()`); await settle(400)
// Copy Sheet keeps the sparklines on the copy, reading the copy's data
await ev(`await api._injector.get(window.__ICommandService).executeCommand('sheet.command.copy-sheet', { unitId: W().getId(), subUnitId: S().getSheetId() })`); await settle(800)
const copied = await ev(`const c=W().getSheets().find(s=>s.getSheetId()!==W().getSheets()[0].getSheetId()); W().setActiveSheet(c); await new Promise(r=>setTimeout(r,400)); return JSON.stringify([SP.groups().reduce((n,g)=>n+g.items.length,0), SP.groups().every(g=>g.items.every(i=>i.sheetId===S().getSheetId()))])`)
check('a copied sheet keeps its sparklines, reading its own data', copied === `[${before},true]`, copied)
await ev(`W().setActiveSheet(W().getSheets()[0])`); await settle(400)

// Clear
await ev(`S().getRange('F2:F4').activate(); SP.clear()`); await settle(600)
await ev(`S().getRange('J30').activate()`)
check('Clear removes the selected sparklines', (await shot('F3')) === blankF3 && (await ev(`return SP.groups().reduce((n,g)=>n+g.items.length,0)`)) === 2)

check('no page errors', errors.length === 0, errors.join(' | '))
await app.close()
const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`sparklines  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
