// Outline (row/column grouping): mouse on the +/- buttons and level buttons, nested collapse, Hide/Show Detail,
// undo, inserting inside a group, the active cell leaving hidden rows, and the .xlsx round-trip.
import { _electron as electron } from 'playwright'
import { writeFileSync, mkdirSync } from 'fs'
import { xlsxToSnapshot, snapshotToXlsx } from '../src/xlsx.js'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)

// ---- .xlsx round-trip (Node side, no app needed)
{
  const snap = { sheetOrder: ['s1'], sheets: { s1: { id: 's1', name: 'Sheet1', cellData: { 0: { 0: { v: 1, t: 2 } } },
    rowData: { 2: { custom: { outlineLevel: 1 }, hd: 1 }, 3: { custom: { outlineLevel: 2 }, hd: 1 }, 4: { custom: { outlineLevel: 1 }, hd: 1 } },
    columnData: { 1: { custom: { outlineLevel: 1 } }, 2: { custom: { outlineLevel: 1 } } } } } }
  const back = await xlsxToSnapshot(await snapshotToXlsx(snap))
  const s = back.sheets[back.sheetOrder[0]]
  check('xlsx: row levels survive', [2, 3, 4].map((r) => s.rowData[r]?.custom?.outlineLevel).join() === '1,2,1', JSON.stringify(s.rowData))
  check('xlsx: hidden rows survive', [2, 3, 4].every((r) => s.rowData[r]?.hd === 1))
  check('xlsx: column levels survive', [1, 2].every((c) => s.columnData[c]?.custom?.outlineLevel === 1), JSON.stringify(s.columnData))
  check('xlsx: fully hidden group loads collapsed', (s.rowData[5]?.custom?.outlineCollapsed & 2) === 2 && (s.rowData[4]?.custom?.outlineCollapsed & 4) === 4, JSON.stringify(s.rowData))
}

const app = await electron.launch({ args: ['.', '--no-sandbox'], env: { ...process.env, SHEETLY_TEST_NO_DIALOGS: '1' } })
const win = await app.firstWindow()
const errors = []
win.on('pageerror', (e) => errors.push(e.message))
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
await win.waitForTimeout(1500)
await win.mouse.click(700, 500)
const ev = (code) => win.evaluate(`(async()=>{const api=window.univerAPI,S=()=>api.getActiveWorkbook().getActiveSheet(),O=window.__sheetly.outline,
  C=api._injector.get(window.__ICommandService),hd=(r)=>S().getSheet().getRowManager().getRow(r)?.hd===1;${code}})()`)
const settle = () => win.waitForTimeout(350)
const reset = async () => { await ev(`O.clearOutline(); O.showSymbols(); S().showRows(0, 40); S().showColumns(0, 10); S().getRange('A1').activate()`); await settle() }
const hiddenRows = (from, to) => ev(`return Array.from({length:${to - from + 1}},(_, i)=>hd(${from}+i)?1:0).join('')`)
// Screen position of the +/- box for a row group: gutter slot of its level, on its summary row.
const rowButton = (level, summaryRow) => ev(`
  const canvas=[...document.querySelectorAll('canvas')].sort((a,b)=>b.width*b.height-a.width*a.height)[0], r=canvas.getBoundingClientRect()
  const sk=api._injector.get(window.__IRenderManagerService).getRenderUnitById(api.getActiveWorkbook().getId()).with(window.__SheetSkeletonManagerService).getCurrentSkeleton()
  const top=(i)=>i>0?sk.rowHeightAccumulation[i-1]:0
  return { x: r.left + (${level}-1)*16 + 9, y: r.top + sk.columnHeaderHeight + (top(${summaryRow}) + sk.rowHeightAccumulation[${summaryRow}])/2 }`)
const colButton = (level, summaryCol) => ev(`
  const canvas=[...document.querySelectorAll('canvas')].sort((a,b)=>b.width*b.height-a.width*a.height)[0], r=canvas.getBoundingClientRect()
  const sk=api._injector.get(window.__IRenderManagerService).getRenderUnitById(api.getActiveWorkbook().getId()).with(window.__SheetSkeletonManagerService).getCurrentSkeleton()
  const left=(i)=>i>0?sk.columnWidthAccumulation[i-1]:0
  return { x: r.left + sk.rowHeaderWidth + (left(${summaryCol}) + sk.columnWidthAccumulation[${summaryCol}])/2, y: r.top + (${level}-1)*16 + 9 }`)
const click = async (p) => { await win.mouse.click(p.x, p.y); await settle() }

// Rows 2-7 (index 1..6) level 1, rows 3-4 (index 2..3) level 2.
const build = async () => { await reset(); await ev(`S().getRange('2:7').activate(); O.group(); S().getRange('3:4').activate(); O.group(); S().getRange('A1').activate()`); await settle() }

await build()
let info = await ev(`return O.info()`)
check('two levels, gutter shown, header widened', info.rowLevels === 2 && info.gutter.row > 0 && info.rowHeaderWidth === 56 + info.gutter.row, JSON.stringify(info))
check('level buttons 1..3 in the corner', (await ev(`return [...document.querySelectorAll('.sx-outline-level[data-axis=row]')].map(b=>b.textContent).join('')`)) === '123')

await click(await rowButton(2, 4))
check('click inner minus collapses rows 3-4', (await hiddenRows(1, 7)) === '0110000', await hiddenRows(1, 7))
await click(await rowButton(2, 4))
check('click inner plus expands rows 3-4', (await hiddenRows(1, 7)) === '0000000', await hiddenRows(1, 7))

await click(await rowButton(2, 4)) // inner collapsed again
await click(await rowButton(1, 7))
check('click outer minus collapses rows 2-7', (await hiddenRows(1, 7)) === '1111110', await hiddenRows(1, 7))
await click(await rowButton(1, 7))
check('expanding outer keeps inner collapsed', (await hiddenRows(1, 7)) === '0110000', await hiddenRows(1, 7))

await win.click('.sx-outline-level[data-axis=row][data-level="1"]'); await settle()
check('level button 1 hides all detail', (await hiddenRows(1, 7)) === '1111110', await hiddenRows(1, 7))
await win.click('.sx-outline-level[data-axis=row][data-level="3"]'); await settle()
check('level button 3 shows everything', (await hiddenRows(1, 7)) === '0000000', await hiddenRows(1, 7))
await win.click('.sx-outline-level[data-axis=row][data-level="2"]'); await settle()
check('level button 2 hides only level-2 rows', (await hiddenRows(1, 7)) === '0110000', await hiddenRows(1, 7))
info = await ev(`return O.info()`)
check('collapsed state reported per group', JSON.stringify(info.rowsCollapsed) === '[false,true]', JSON.stringify(info.rowsCollapsed))

await win.keyboard.press('Control+z'); await settle()
check('Ctrl+Z undoes the level change', (await hiddenRows(1, 7)) === '0000000', await hiddenRows(1, 7))
await win.keyboard.press('Control+y'); await settle()
check('Ctrl+Y redoes it', (await hiddenRows(1, 7)) === '0110000', await hiddenRows(1, 7))

// Hide / Show Detail on the active cell.
await build()
await ev(`S().getRange('A3').activate(); C.executeCommand('sheetly.outline.hide-detail')`); await settle()
check('Hide Detail collapses the innermost group', (await hiddenRows(1, 7)) === '0110000', await hiddenRows(1, 7))
check('active cell moves to the summary row', (await ev(`return S().getActiveRange().getA1Notation()`)) === 'A5', await ev(`return S().getActiveRange().getA1Notation()`))
await ev(`C.executeCommand('sheetly.outline.show-detail')`); await settle()
check('Show Detail on the summary row expands it', (await hiddenRows(1, 7)) === '0000000', await hiddenRows(1, 7))
check("Univer's unhide arrows suppressed for collapsed groups", await ev(`
  await C.executeCommand('sheetly.outline.toggle',{axis:'row',start:2,end:3,level:2}); await new Promise(r=>setTimeout(r,300))
  const ctl=api._injector.get(window.__IRenderManagerService).getRenderUnitById(api.getActiveWorkbook().getId()).with(window.__HeaderUnhideRenderController)
  return ctl._shapes.rows.length===0`))

// Inserting a row inside a group joins it.
await build()
await ev(`S().insertRowsBefore(3, 1)`); await settle()
check('row inserted inside a group gets its level', (await ev(`return [1,2,3,4,5,6,7,8].map(r=>O.levelOf('row',r)).join('')`)) === '12221110',
  await ev(`return [1,2,3,4,5,6,7,8].map(r=>O.levelOf('row',r)).join('')`))
await ev(`S().insertRowsBefore(12, 1)`); await settle()
check('row inserted outside stays ungrouped', (await ev(`return O.levelOf('row',12)`)) === 0)

// Columns: C:D grouped, click its minus (on E), then level button 2.
await reset()
await ev(`S().getRange('C:D').activate(); O.group(); S().getRange('A1').activate()`); await settle()
await click(await colButton(1, 4))
check('click column minus hides C:D', (await ev(`return [2,3].every(c=>S().getSheet().getColumnManager().getColumn(c)?.hd===1)`)))
await win.click('.sx-outline-level[data-axis=col][data-level="2"]'); await settle()
check('column level button 2 shows them', (await ev(`return [2,3].every(c=>!S().getSheet().getColumnManager().getColumn(c)?.hd)`)))

// Header width must not creep with repeated refreshes; it returns to normal when the outline is cleared.
await ev(`for (let i=0;i<5;i++){ O.refresh(); await new Promise(r=>setTimeout(r,50)) }`)
info = await ev(`return O.info()`)
check('header size stable after refreshes', info.rowHeaderWidth === 56 && (await ev(`return S().getSheet().getConfig().columnHeader.height`)) === 20 + info.gutter.col, JSON.stringify(info))
await ev(`O.clearOutline()`); await settle()
check('clear outline restores header size', (await ev(`return S().getSheet().getConfig().columnHeader.height`)) === 20)

check('no page errors', errors.length === 0, errors.join(' | '))
await app.close()

const pass = out.filter((l) => l.startsWith('PASS')).length
const report = `outline.test.mjs  ${new Date().toISOString()}\npass ${pass}/${out.length}\n${out.join('\n')}\n`
mkdirSync('test_results', { recursive: true })
writeFileSync(`test_results/outline_${new Date().toISOString().slice(0, 10)}.txt`, report)
console.log(report)
process.exit(pass === out.length ? 0 : 1)
