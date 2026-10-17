// Round-2 gaps, driven with real key presses in the app: F6 / Shift+F6 focus cycling, Ctrl+F1 ribbon, Ctrl+Q Quick
// Analysis, Shift+F8 Add to Selection, outline Settings (summary above/left), Ctrl+8 saved, Auto Outline, and a
// UI-styled cell surviving the .xlsx writer.
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
const ev = (code) => win.evaluate(`(async()=>{const api=window.univerAPI,S=()=>api.getActiveWorkbook().getActiveSheet(),X=window.__sheetly,O=X.outline;${code}})()`)
const settle = (ms = 300) => win.waitForTimeout(ms)
const key = async (k, n = 1) => { for (let i = 0; i < n; i++) await win.keyboard.press(k); await settle(150) }
const focusArea = () => ev(`const a=document.activeElement; return a?.closest?.('header[data-u-comp="headerbar"]')?'ribbon':a?.closest?.('footer')?'statusbar':a?.closest?.('[data-u-comp="right-sidebar"]')?'taskpane':'worksheet'`)
const grid = async () => { await win.mouse.click(700, 500); await settle() }

// ---- Ctrl+F1
{
  await grid()
  const top0 = await ev(`return [...document.querySelectorAll('canvas')].sort((a,b)=>b.width*b.height-a.width*a.height)[0].getBoundingClientRect().top`)
  await key('Control+F1'); await settle(400)
  const st = await ev(`return { c: document.body.classList.contains('sx-ribbon-collapsed'), tool: document.querySelector('[data-u-comp="ribbon-toolbar"]').getBoundingClientRect().height,
    tabs: document.querySelector('[data-u-comp="ribbon-header-menu"]').getBoundingClientRect().height,
    top: [...document.querySelectorAll('canvas')].sort((a,b)=>b.width*b.height-a.width*a.height)[0].getBoundingClientRect().top }`)
  check('Ctrl+F1 collapses the ribbon to its tabs', st.c && st.tool === 0 && st.tabs > 0, JSON.stringify(st))
  check('Ctrl+F1: the grid moves up into the space', st.top < top0, `${top0} -> ${st.top}`)
  await key('Control+F1'); await settle(400)
  const st2 = await ev(`return { c: document.body.classList.contains('sx-ribbon-collapsed'), tool: document.querySelector('[data-u-comp="ribbon-toolbar"]').getBoundingClientRect().height }`)
  check('Ctrl+F1 again expands it', !st2.c && st2.tool > 0, JSON.stringify(st2))
}

// ---- F6 / Shift+F6 / Esc
{
  await grid()
  await key('F6'); const a = await focusArea()
  const tab = await ev(`return document.activeElement?.getAttribute('role')`)
  await key('F6'); const b = await focusArea()
  await key('F6'); const c = await focusArea()
  check('F6 cycles worksheet -> ribbon -> status bar -> worksheet', a === 'ribbon' && tab === 'tab' && b === 'statusbar' && c === 'worksheet', `${a}(${tab}) ${b} ${c}`)
  await key('Shift+F6'); const d = await focusArea()
  await key('Shift+F6'); const e = await focusArea()
  check('Shift+F6 goes the other way', d === 'statusbar' && e === 'ribbon', `${d} ${e}`)
  await key('Escape'); const f = await focusArea()
  check('Esc in the ribbon returns to the worksheet', f === 'worksheet', f)
  // typing now reaches the grid
  await ev(`S().getRange('H1').activate()`); await settle()
  await key('F6'); await key('F6'); await key('F6')
  await win.keyboard.type('77'); await key('Enter'); await settle()
  check('after a full F6 cycle typing goes into the cell', (await ev(`return S().getRange('H1').getValue()`)) === 77, String(await ev(`return S().getRange('H1').getValue()`)))
  await ev(`S().getRange('H1').clear()`)
}

// ---- Ctrl+Q Quick Analysis
{
  await grid()
  await ev(`S().getRange('A1:B3').setValues([[1,2],[3,4],[5,6]]); S().getRange('A1:B3').activate()`); await settle()
  await key('Control+q'); await settle()
  const tabs = await ev(`return [...document.querySelectorAll('.sx-qa-tab')].map(b=>b.textContent).join(',')`)
  check('Ctrl+Q opens Quick Analysis with Excel\'s tabs', tabs === 'FORMATTING,CHARTS,TOTALS,TABLES,SPARKLINES', tabs)
  const focusIn = await ev(`return document.activeElement?.textContent`)
  check('first option (Data Bars) has the focus', focusIn === 'Data Bars', focusIn)
  await key('Enter'); await settle()
  const rules = await ev(`return S().getConditionalFormattingRules().map(r=>r.rule.type).join(',')`)
  check('Enter applies Data Bars to the selection', rules === 'dataBar' && !(await ev(`return X.quickAnalysis.isOpen()`)), rules)
  await key('Control+q'); await key('ArrowRight'); await key('Enter'); await settle()
  check('Right then Enter applies Color Scale', (await ev(`return S().getConditionalFormattingRules().map(r=>r.rule.type).sort().join(',')`)) === 'colorScale,dataBar', await ev(`return S().getConditionalFormattingRules().map(r=>r.rule.type).join(',')`))
  await key('Control+q'); await key('ArrowLeft'); await key('Enter'); await settle()
  check('Left wraps to Clear Format, which clears the rules', (await ev(`return S().getConditionalFormattingRules().length`)) === 0)
  await key('Control+q'); await key('o'); await settle()
  const totals = await ev(`return document.querySelector('.sx-qa-tab[aria-selected=true]')?.textContent + ':' + [...document.querySelectorAll('.sx-qa-item')].slice(0,5).map(b=>b.textContent).join('|')`)
  check('O jumps to the Totals tab', totals === 'TOTALS:Sum|Average|Count|% Total|Running Total', totals)
  await key('Enter'); await settle(500)
  const sums = await ev(`return [S().getRange('A4').getFormula(), S().getRange('B4').getFormula(), S().getRange('A4').getValue(), S().getRange('B4').getValue()].join('|')`)
  check('Totals > Sum puts SUM formulas in the row below', sums === '=SUM(A1:A3)|=SUM(B1:B3)|9|12', sums)
  await ev(`S().getRange('A4:B4').clear(); S().getRange('A1:B3').activate()`); await settle()
  await key('Control+q'); await key('Tab', 2); await key('ArrowRight', 9); await key('Enter'); await settle(500)
  const run = await ev(`return [S().getRange('C1').getFormula(), S().getRange('C3').getFormula(), S().getRange('C3').getValue()].join('|')`)
  check('Totals > Running Total (columns) puts cumulative sums to the right', run === '=SUM($A$1:B1)|=SUM($A$1:B3)|21', run)
  await ev(`S().getRange('C1:C3').clear()`)
  await key('Control+q'); await key('Escape'); await settle()
  check('Esc closes Quick Analysis', !(await ev(`return X.quickAnalysis.isOpen()`)))
  await ev(`S().getRange('J20').activate()`); await settle()
  await key('Control+q'); await settle()
  check('Ctrl+Q on an empty cell does nothing', !(await ev(`return X.quickAnalysis.isOpen()`)))
  await ev(`S().getRange('A2').activate()`); await settle()
  await key('Control+q'); await settle()
  const region = await ev(`return S().getActiveRange().getA1Notation()`)
  check('Ctrl+Q on one cell of data selects its current region', region === 'A1:B3' && (await ev(`return X.quickAnalysis.isOpen()`)), region)
  await key('Escape')
  await ev(`S().getRange('A1:B3').clear(); S().getRange('A1:A4').setValues([['a'],['b'],['a'],['c']]); S().getRange('A1:A4').activate()`); await settle()
  await key('Control+q'); await settle()
  const textItems = await ev(`return [...document.querySelectorAll('.sx-qa-item')].map(b=>b.textContent).join('|')`)
  check('text selection: Formatting offers the text options', textItems === 'Text Contains|Duplicate Values|Unique Values|Equal To|Clear Format', textItems)
  await key('ArrowRight'); await key('Enter'); await settle()
  check('Duplicate Values applies a duplicate rule', (await ev(`return S().getConditionalFormattingRules().map(r=>r.rule.type+':'+r.rule.subType).join(',')`)) === 'highlightCell:duplicateValues', await ev(`return JSON.stringify(S().getConditionalFormattingRules().map(r=>r.rule))`))
  await ev(`S().getRange('A1:A4').clearConditionalFormatRules(); S().getRange('A1:A4').clear()`)
}

// ---- a plain arrow with a range selected moves from the active cell (Univer moved from the range edge)
{
  await grid()
  await ev(`S().getRange('A1:B2').activate()`); await settle()
  await key('ArrowDown')
  const a = await ev(`return S().getActiveRange().getA1Notation()`)
  await ev(`S().getRange('C3:D5').activate()`); await settle()
  await key('ArrowRight')
  const b = await ev(`return S().getActiveRange().getA1Notation()`)
  check('Down from A1:B2 (A1 active) goes to A2, Right from C3:D5 goes to D3', a === 'A2' && b === 'D3', `${a} ${b}`)
}

// ---- Shift+F8 Add to Selection
{
  await grid()
  const areas = () => ev(`return (S().getSelection().getActiveRangeList()??[]).map(r=>r.getA1Notation()).join(',')`)
  const status = () => ev(`return [...document.querySelectorAll('[aria-live=polite]')].map(n=>n.textContent).join('')`)
  await ev(`S().getRange('A1:B2').activate()`); await settle()
  await key('Shift+F8')
  check('Shift+F8 shows "Add to Selection"', (await status()) === 'Add to Selection', await status())
  await key('ArrowDown', 3)
  await key('Shift+ArrowRight')
  const a = await areas()
  check('arrows then Shift+arrow add a second area, the first stays', a === 'A1:B2,A4:B4', a)
  await key('ArrowDown', 2); await key('Shift+ArrowDown')
  const b = await areas()
  check('a third area keeps both earlier ones', b === 'A1:B2,A4:B4,A6:A7', b)
  await key('Shift+F8')
  check('Shift+F8 again leaves the mode', (await status()) === '', await status())
  await key('ArrowDown')
  check('after leaving, an arrow moves from the active cell (A6 -> A7) to one cell', (await areas()) === 'A7', await areas())
  await key('Shift+F8'); await key('Escape')
  check('Esc also leaves Add to Selection', (await status()) === '')
}

// ---- outline settings: summary rows above
{
  await grid()
  await ev(`O.clearOutline(); O.showSymbols(); O.setDirection({summaryBelow:true,summaryRight:true}); S().showRows(0,40); S().getRange('3:5').activate(); O.group(); S().getRange('A1').activate()`); await settle(500)
  await ev(`O.toggle('row', {start:2,end:4,level:1})`); await settle()
  await ev(`O.setDirection({summaryBelow:false,summaryRight:true})`); await settle(500)
  const moved = await ev(`const r=(i)=>S().getSheet().getRowManager().getRow(i)?.custom?.outlineCollapsed??0; return [r(1), r(5)].join(',')`)
  check('switching to summary above moves the collapse flag to the row above the group', moved === '2,0', moved)
  await ev(`O.toggle('row', {start:2,end:4,level:1})`); await settle()
  // click the + / - on the summary row above (index 1)
  const p = await ev(`const canvas=[...document.querySelectorAll('canvas')].sort((a,b)=>b.width*b.height-a.width*a.height)[0], r=canvas.getBoundingClientRect()
    const sk=api._injector.get(window.__IRenderManagerService).getRenderUnitById(api.getActiveWorkbook().getId()).with(window.__SheetSkeletonManagerService).getCurrentSkeleton()
    const top=(i)=>i>0?sk.rowHeightAccumulation[i-1]:0
    return { x: r.left + 9, y: r.top + sk.columnHeaderHeight + (top(1) + sk.rowHeightAccumulation[1])/2 }`)
  await win.mouse.click(p.x, p.y); await settle()
  const hid = await ev(`return [1,2,3,4,5].map(i=>S().getSheet().getRowManager().getRow(i)?.hd===1?1:0).join('')`)
  check('summary above: clicking - on the row above collapses the group', hid === '01110', hid)
  await win.mouse.click(p.x, p.y); await settle()
  check('and + expands it', (await ev(`return [2,3,4].every(i=>S().getSheet().getRowManager().getRow(i)?.hd!==1)`)))
  // Settings dialog with access keys: B toggles "Summary rows below detail", Enter = OK
  await ev(`O.settingsDialog()`); await settle()
  const labels = await ev(`return [...document.querySelectorAll('.sx-dialog label')].map(l=>l.textContent+':'+l.querySelector('input').checked).join('|')`)
  check('Settings dialog shows the direction boxes and Automatic styles', labels === 'Summary rows below detail:false|Summary columns to right of detail:true|Automatic styles:false', labels)
  await key('b'); await key('Enter'); await settle()
  check('B + Enter sets summary rows below again', (await ev(`return O.settings().summaryBelow`)) === true)
  await ev(`O.clearOutline()`)
}

// ---- Ctrl+8 is saved with the sheet
{
  await grid()
  await ev(`S().getRange('2:3').activate(); O.group(); S().getRange('A1').activate()`); await settle(400)
  await key('Control+8'); await settle(400)
  const saved = await ev(`const s=api.getActiveWorkbook().save(); return JSON.stringify(s.sheets[s.sheetOrder[0]].custom)`)
  check('Ctrl+8 is stored in the sheet (saved with the workbook)', saved.includes('"hideSymbols":true'), saved)
  const x = await ev(`const s=api.getActiveWorkbook().save(); const b=await X.xlsx.snapshotToXlsx(s); const back=await X.xlsx.xlsxToSnapshot(b); return JSON.stringify(back.sheets[back.sheetOrder[0]].custom)`)
  check('Ctrl+8 survives .xlsx (showOutlineSymbols)', x.includes('"hideSymbols":true'), x)
  await key('Control+8'); await settle(300)
  check('Ctrl+8 again shows them and clears the setting', (await ev(`return O.info().gutter.row`)) > 0 && !(await ev(`const s=api.getActiveWorkbook().save(); return JSON.stringify(s.sheets[s.sheetOrder[0]].custom??{})`)).includes('hideSymbols'))
  await ev(`O.clearOutline()`)
}

// ---- Auto Outline
{
  await grid()
  await ev(`S().getRange('A1:A9').setValues([[1],[2],[3],[null],[5],[6],[null],[null],[null]])
    S().getRange('A4').setFormula('=SUM(A1:A3)'); S().getRange('A7').setFormula('=SUM(A5:A6)'); S().getRange('A8').setFormula('=A4+A7'); S().getRange('A9').setFormula('=SUM(A1:A8)')
    S().getRange('A1').activate()`); await settle(500)
  await ev(`O.autoOutline()`); await settle(400)
  const lv = await ev(`return Array.from({length:9},(_, i)=>O.levelOf('row', i)).join('')`)
  check('Auto Outline: subtotals, a grand total over subtotals (=A4+A7) and SUM(A1:A8) nest', lv === '333233210', lv)
  await ev(`O.clearOutline(); S().getRange('A1:A4').activate(); O.autoOutline(); S().getRange('A1').activate()`); await settle(400)
  const scoped = await ev(`return Array.from({length:9},(_, i)=>O.levelOf('row', i)).join('')`)
  check('Auto Outline with A1:A4 selected outlines only that range', scoped === '111000000', scoped)
  await ev(`O.autoOutline()`); await settle()
  const ask = await ev(`return document.querySelector('.sx-dialog p')?.textContent`)
  check('Auto Outline over an outline asks "Modify existing outline?"', ask === 'Modify existing outline?', ask)
  await key('Escape')
  await ev(`O.clearOutline(); S().getRange('A1:A9').clear(); S().getRange('A1').activate()`); await settle()
  await ev(`O.autoOutline()`); await settle()
  const none = await ev(`return document.querySelector('.sx-dialog p')?.textContent`)
  check('no summary formulas -> "Cannot create an outline."', none === 'Cannot create an outline.', none)
  await key('Escape')
  await ev(`S().getRange('C1:C2').setValues([[1],[2]]); S().getRange('C3').setFormula('=C2*2'); S().getRange('A1').activate()`); await settle()
  await ev(`O.autoOutline()`); await settle()
  const single = await ev(`return document.querySelector('.sx-dialog p')?.textContent`)
  check('=C2*2 (one adjacent cell) is not a summary', single === 'Cannot create an outline.', single)
  await key('Escape'); await ev(`S().getRange('C1:C3').clear()`)
  // KeyTip Alt A G A
  await ev(`S().getRange('B1:B3').setValues([[1],[2],[3]]); S().getRange('B4').setFormula('=SUM(B1:B3)'); S().getRange('A1').activate()`); await settle()
  await key('Alt'); await key('a'); await key('g'); await key('a'); await settle(400)
  check('Alt A G A runs Auto Outline', (await ev(`return [0,1,2,3].map(i=>O.levelOf('row', i)).join('')`)) === '1110')
  await ev(`O.applyStyles()`); await settle()
  check('Apply Styles makes the level-1 summary row bold (RowLevel_1)', (await ev(`return S().getRange('B4').getCellStyleData()?.bl`)) === 1 && !(await ev(`return S().getRange('B3').getCellStyleData()?.bl`)))
  await ev(`S().getRange('B4').setFontWeight('normal')`)
  await ev(`O.clearOutline(); S().getRange('B1:B4').clear()`)
}

// ---- a UI-styled cell (Ctrl+B, Format via toolbar commands) reaches the .xlsx file
{
  await grid()
  await ev(`S().getRange('D5').setValue('styled'); S().getRange('D5').activate()`); await settle()
  await key('Control+b'); await key('Control+i'); await settle()
  await ev(`S().getRange('D5').setBackground('#ffff00').setFontColor('#ff0000').setFontSize(16).setHorizontalAlignment('center')`); await settle()
  const back = await ev(`const s=api.getActiveWorkbook().save(); const b=await X.xlsx.snapshotToXlsx(s); const r=await X.xlsx.xlsxToSnapshot(b); return JSON.stringify(r.sheets[r.sheetOrder[0]].cellData[4][3].s)`)
  const st = JSON.parse(back)
  check('UI styles survive save + reopen (bold, italic, fill, colour, size, centre)', st.bl === 1 && st.it === 1 && st.bg?.rgb === '#ffff00' && st.cl?.rgb === '#ff0000' && st.fs === 16 && st.ht === 2, back)
  await ev(`S().getRange('D5').clear()`)
}

// ---- filled-down formulas (Univer keeps them as shared-formula ids) reach the .xlsx file
{
  await grid()
  await ev(`S().getRange('E1:E3').setValues([[1],[2],[3]]); S().getRange('F1').setFormula('=E1*2'); S().getRange('F1:F3').activate()`); await settle()
  await key('Control+d'); await settle(500)
  const raw = await ev(`const s=api.getActiveWorkbook().save(); const c=s.sheets[s.sheetOrder[0]].cellData; return JSON.stringify([c[1]?.[5], c[2]?.[5]])`)
  const back = await ev(`const r=await X.xlsx.xlsxToSnapshot(await X.xlsx.save()); const c=r.sheets[r.sheetOrder[0]].cellData; return [c[1]?.[5]?.f, c[2]?.[5]?.f].join('|')`)
  check('Ctrl+D filled formulas are saved with their own formula text', back === '=E2*2|=E3*2', `${raw} -> ${back}`)
  await ev(`S().getRange('E1:F3').clear()`)
}

// ---- formatted numbers are still numbers (Univer's getValue() returns "$100" for a currency cell)
{
  await grid()
  await ev(`S().getRange('G1:G3').setValues([[100],[200],[300]]); S().getRange('G1:G3').setNumberFormat('$#,##0'); S().getRange('G4').activate()`); await settle()
  await key('Alt+='); await settle(300)
  const f = await ev(`return S().getRange('G4').getFormula()`)
  await key('Enter'); await settle(300)
  check('AutoSum finds currency-formatted numbers above', f === '=SUM(G1:G3)' && (await ev(`return S().getRange('G4').getRawValue()`)) === 600, `${f} ${await ev(`return S().getRange('G4').getRawValue()`)}`)
  await ev(`S().getRange('G5').activate()`); await settle()
  await key('Control+Shift+Quote'); await key('Enter'); await settle(300)
  check('Ctrl+Shift+" copies the value 600, not the text "$600"', (await ev(`return S().getRange('G5').getRawValue()`)) === 600, String(await ev(`return JSON.stringify(S().getRange('G5').getRawValue())`)))
  await ev(`S().getRange('G1:G5').clear()`)
}

check('no page errors', errors.length === 0, errors.join(' | '))
await app.close()
const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`round 2 (F6, Ctrl+F1, Ctrl+Q, Shift+F8, outline settings, Auto Outline, xlsx styles)  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
