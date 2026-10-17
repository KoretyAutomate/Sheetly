// Charts: linked to the source range (redrawn when values change), chart types, Switch Row/Column, chart elements,
// Insert tab KeyTips, the Insert Chart dialog, Quick Analysis recommendations, Chart Design KeyTips, saved with
// the workbook.
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
const ev = (code) => win.evaluate(`(async()=>{const api=window.univerAPI,W=()=>api.getActiveWorkbook(),S=()=>W().getActiveSheet(),X=window.__sheetly,CH=X.charts;
  const canv=()=>[...document.querySelectorAll('.sx-chart canvas')]; const cj=(i)=>canv()[i]?.__chart; ${code}})()`)
const settle = (ms = 400) => win.waitForTimeout(ms)
const key = async (k, n = 1) => { for (let i = 0; i < n; i++) await win.keyboard.press(k); await settle(150) }
const grid = async () => { await win.mouse.click(700, 500); await settle() }
const removeCharts = () => ev(`for (const c of CH.charts()) S().removeFloatDom(c.id)`)

await ev(`S().getRange('A1:C4').setValues([[null,'Q1','Q2'],['North',10,20],['South',30,40],['East',50,60]]); S().getRange('A1').activate()`)
await settle()
// ---- Alt+F1: clustered column of the current region, linked to it
await key('Alt+F1'); await settle(800)
let st = await ev(`const c=cj(0); return c && { type: c.config.type, sets: c.data.datasets.map(d=>d.label).join(','), labels: c.data.labels.join(','), v: c.data.datasets[0].data.join(',') }`)
check('Alt+F1: clustered column, series by columns with the header row as names', st?.type === 'bar' && st.sets === 'Q1,Q2' && st.labels === 'North,South,East' && st.v === '10,30,50', JSON.stringify(st))
await ev(`S().getRange('B2').setValue(99)`); await settle(600)
st = await ev(`return cj(0).data.datasets[0].data.join(',')`)
check('the chart follows its data: B2 = 99 redraws it', st === '99,30,50', st)
await ev(`S().getRange('B3').setFormula('=B2+1')`); await settle(800)
st = await ev(`return cj(0).data.datasets[0].data.join(',')`)
check('formula results feed the chart too', st === '99,100,50', st)
await ev(`S().getRange('B2:B3').setValues([[10],[30]])`); await settle(400)

// ---- change type, switch row/column, elements
await ev(`CH.setType('line')`); await settle(600)
check('Change Chart Type to Line', (await ev(`return cj(0).config.type`)) === 'line')
await ev(`CH.switchRowColumn()`); await settle(600)
st = await ev(`const c=cj(0); return c.data.datasets.map(d=>d.label).join(',') + '|' + c.data.labels.join(',')`)
check('Switch Row/Column: series become the rows', st === 'North,South,East|Q1,Q2', st)
await ev(`CH.toggle('legend')`); await settle(500)
check('Legend can be turned off', (await ev(`return cj(0).options.plugins.legend.display`)) === false)
await ev(`CH.toggle('dataLabels')`); await settle(500)
check('Data Labels can be turned on', (await ev(`return CH.selected().data.dataLabels`)) === true)
await ev(`CH.setType('columnStacked100')`); await settle(600)
st = await ev(`const c=cj(0); return c.options.scales.y.max + ':' + c.data.datasets[0].data.map(x=>Math.round(x)).join(',')`)
check('100% Stacked Column scales each category to 100% (North = 10 of 90 in Q1, 20 of 120 in Q2)', st === '100:11,17', st)
await removeCharts(); await settle()

// ---- Insert tab KeyTips: Alt N Q P = Pie
await ev(`S().getRange('A1:B4').activate()`); await settle()
await key('Alt'); await key('n'); await key('q'); await key('p'); await settle(800)
st = await ev(`const c=cj(0); return c && c.config.type + ':' + c.data.datasets[0].backgroundColor.length`)
check('Alt N Q P inserts a pie, one colour per slice', st === 'pie:3', st)
await removeCharts(); await settle()
await ev(`S().getRange('A1:C4').activate()`); await settle()
await key('Alt'); await key('n'); await key('c'); await key('1'); await key('b'); await settle(800)
check('Alt N C1 B inserts a clustered bar (horizontal)', (await ev(`return cj(0)?.options.indexAxis`)) === 'y')
await removeCharts(); await settle()

// ---- Insert Chart dialog (Alt N R)
await key('Alt'); await key('n'); await key('r'); await settle()
const types = await ev(`return document.querySelectorAll('.sx-dialog input[name=chart-type]').length`)
await ev(`[...document.querySelectorAll('.sx-dialog label')].find(l=>l.textContent==='Stacked Column').querySelector('input').click()`)
await key('Enter'); await settle(800)
check('Insert Chart lists the chart types and inserts the chosen one', types === 13 && (await ev(`return CH.charts()[0]?.data.type`)) === 'columnStacked', `${types} ${await ev(`return CH.charts()[0]?.data.type`)}`)
await removeCharts(); await settle()

// ---- Quick Analysis recommendations
await ev(`S().getRange('A1:B4').activate()`); await settle()
await key('Control+q'); await key('Tab'); await settle()
st = await ev(`return [...document.querySelectorAll('.sx-qa-item')].map(b=>b.textContent).join('|')`)
check('Quick Analysis Charts recommends for one series: Column, Bar, Line, Pie, More', st === 'Clustered Column|Clustered Bar|Line|Pie|More Charts', st)
await key('ArrowRight', 3); await key('Enter'); await settle(800)
check('choosing Pie from Quick Analysis inserts it', (await ev(`return cj(0)?.config.type`)) === 'pie')

// ---- Chart Design KeyTips appear only with a chart selected
await grid()
await key('Alt'); await settle()
const without = await ev(`return [...document.querySelectorAll('.sx-keytips-tab')].map(b=>b.textContent).some(t=>t.startsWith('Chart Design'))`)
await key('Escape')
await ev(`const c=CH.charts()[0]; api._injector.get(window.__IDrawingManagerService).focusDrawing([{unitId:api.getActiveWorkbook().getId(),subUnitId:S().getSheetId(),drawingId:c.id}])`); await settle()
await key('Alt'); await settle()
const withChart = await ev(`return [...document.querySelectorAll('.sx-keytips-tab')].map(b=>b.textContent).some(t=>t.startsWith('Chart Design'))`)
await key('j'); await key('c'); await key('c'); await settle()
const dlg = await ev(`return document.querySelector('.sx-dialog .sx-title span')?.textContent`)
check('Chart Design (JC) KeyTips only while a chart is selected; J C C opens Change Chart Type', !without && withChart && dlg === 'Change Chart Type', `${without} ${withChart} ${dlg}`)
await key('Escape')

// ---- series ranges follow inserted / deleted rows and columns; axis format and automatic minimum
await removeCharts(); await settle()
await ev(`S().getRange('A1:C4').setValues([[null,'Q1','Q2'],['North',10,20],['South',30,40],['East',50,60]]); S().getRange('A1').activate()`); await settle()
await key('Alt+F1'); await settle(800)
const rng = () => ev(`const g=CH.sourceOf(CH.charts()[0])?.range; return g ? [g.startRow,g.endRow,g.startColumn,g.endColumn].join(',') : 'REF'`)
await ev(`S().insertRowsBefore(0, 2)`); await settle(600)
const afterInsert = await rng()
await ev(`S().deleteRows(3, 1)`); await settle(600)
const afterDelete = await rng(), labels = await ev(`return cj(0).data.labels.join(',')`)
check('rows inserted above the data move the chart range; deleting a data row (North) shrinks it', afterInsert === '2,5,0,2' && afterDelete === '2,4,0,2' && labels === 'South,East', `${afterInsert} ${afterDelete} ${labels}`)
await key('Control+z'); await settle(600)
check('undo moves it back', (await rng()) === '2,5,0,2', await rng())
await ev(`S().insertColumnsBefore(1, 1)`); await settle(600)
check('a column inserted inside the data widens the range', (await rng()) === '2,5,0,3', await rng())
await removeCharts(); await ev(`S().deleteColumns(1, 1); S().deleteRows(0, 2); S().getRange('A1:H12').clear()`); await settle()
await ev(`S().getRange('A1:B4').setValues([['Item','Price'],['a',100],['b',101],['c',102]]); S().getRange('B2:B4').setNumberFormat('$#,##0'); S().getRange('A1').activate()`); await settle()
await key('Alt+F1'); await settle(800)
st = await ev(`const c=cj(canv().length-1); return c.options.scales.y.beginAtZero + '|' + c.options.scales.y.ticks.callback(100) + '|' + canv().length + '|' + c.data.datasets.map(d=>d.data.join(',')).join(';') + '|' + JSON.stringify(CH.charts()[0].data.source.range) + '|' + JSON.stringify(S().getRange('A1:B4').getValues())`)
check('value axis: Excel-style automatic minimum (data 100-102 do not start at 0) and the cells\' $ format', st.startsWith('false|$100'), st)
await removeCharts(); await ev(`S().getRange('A1:C4').clear()`); await settle()
await ev(`S().getRange('A1:C4').setValues([[null,'Q1','Q2'],['North',10,20],['South',30,40],['East',50,60]]); S().getRange('A1:B4').activate()`); await settle()
await key('Control+q'); await key('c'); await key('ArrowRight', 3); await key('Enter'); await settle(800)

// ---- two text columns are a two-level category axis, not a series
await removeCharts(); await ev(`S().getRange('A1:H20').clear(); S().getRange('A1:C6').setValues([['Region','Product','Sales'],['East','A',10],['East','B',20],['West','A',5],['West','B',7],['North','B',3]]); S().getRange('A1').activate()`); await settle()
await key('Alt+F1'); await settle(800)
st = await ev(`const c=cj(canv().length-1); return c.data.datasets.map(d=>d.label).join(',') + '|' + JSON.stringify(c.data.labels.slice(0,3)) + '|' + c.options.plugins.title.text`)
check('Region and Product (text) label the axis in two levels; only Sales is plotted', st === 'Sales|[["A","East"],["B",""],["A","West"]]|Sales', st)
const xml = await ev(`const b=await X.xlsx.save(); return Array.from(b)`).then(async (bytes) => { const z = await (await import('jszip')).default.loadAsync(Buffer.from(bytes)); const f = Object.keys(z.files).find((n) => /^xl\/charts\/chart\d+\.xml$/.test(n)); return f ? z.file(f).async('string') : '' })
check('saved chart: multi-level categories over both label columns', xml.includes('<c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$6</c:f>') && (xml.match(/<c:ser>/g) ?? []).length === 1, xml.slice(xml.indexOf('<c:cat>'), xml.indexOf('<c:cat>') + 160))
await ev(`S().getRange('A1').activate()`)
await key('Control+q'); await key('c'); await key('ArrowRight', 3); await key('Enter'); await settle(800)

// ---- a chart with series picked one by one (categories A, series B and D) plots just those
await ev(`S().getRange('A1:H20').clear(); S().getRange('A1:D4').setValues([['R','S1','skip','S2'],['x',1,100,4],['y',2,200,5],['z',3,300,6]])`); await settle()
await ev(`S().addFloatDomToPosition({ componentKey: 'sheetly-chart', allowTransform: true, initPosition: { startX: 400, endX: 800, startY: 0, endY: 250 },
  data: { v: 2, type: 'column', title: null, source: { sheetId: S().getSheetId(), range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 }, categories: { startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 },
    ref: X.refs.track(S().getSheetId(), { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 }), categoriesRef: X.refs.track(S().getSheetId(), { startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }),
    series: [{ name: [0, 1], nameRef: X.refs.track(S().getSheetId(), { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }), values: { startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }, valuesRef: X.refs.track(S().getSheetId(), { startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }) },
      { name: [0, 3], nameRef: X.refs.track(S().getSheetId(), { startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }), values: { startRow: 1, endRow: 3, startColumn: 3, endColumn: 3 }, valuesRef: X.refs.track(S().getSheetId(), { startRow: 1, endRow: 3, startColumn: 3, endColumn: 3 }) }] } } })`); await settle(800)
st = await ev(`const c=cj(canv().length-1); return c.data.datasets.map(d=>d.label+':'+d.data.join(',')).join('|') + '|' + c.data.labels.join(',')`)
check('series picked one by one: only those columns are plotted', st === 'S1:1,2,3|S2:4,5,6|x,y,z', st)
await ev(`S().hideRows(2, 1)`); await settle(700)
st = await ev(`const c=cj(canv().length-1); return c.data.datasets.map(d=>d.label+':'+d.data.join(',')).join('|') + '|' + c.data.labels.join(',')`)
check('a hidden row drops out of a chart with series picked one by one', st === 'S1:1,3|S2:4,6|x,z', st)
await ev(`S().showRows(2, 1)`); await settle(500)
await ev(`S().insertColumnsBefore(2, 1)`); await settle(700)
st = await ev(`const c=cj(canv().length-1); return c.data.datasets.map(d=>d.label+':'+d.data.join(',')).join('|')`)
check('and their references move with an inserted column', st === 'S1:1,2,3|S2:4,5,6', st)
await removeCharts(); await ev(`S().deleteColumns(2, 1); S().getRange('A1:H20').clear()`); await settle()
await ev(`S().getRange('A1:C6').setValues([['Region','Product','Sales'],['East','A',10],['East','B',20],['West','A',5],['West','B',7],['North','B',3]]); S().getRange('A1').activate()`); await settle()
await key('Control+q'); await key('c'); await key('ArrowRight', 3); await key('Enter'); await settle(800)

check('internal reference names are not listed among the workbook\'s names', (await ev(`return W().getDefinedNames().filter(n=>n.getName().startsWith('SheetlyRef_')).length`)) > 0
  && (await ev(`return document.body.innerText.includes('SheetlyRef_')`)) === false)
await key('Control+F3'); await settle(800)
check('Name Manager does not show the internal names', (await ev(`return document.body.innerText.includes('SheetlyRef_')`)) === false)
await key('Escape'); await ev(`document.querySelector('[data-u-comp=right-sidebar] button[aria-label=Close], .univer-sidebar button')?.click?.()`); await settle()
// ---- a scatter point at X = 0 stays at 0
await removeCharts(); await ev(`S().getRange('A1:H20').clear(); S().getRange('A1:B3').setValues([[0,5],[1,6],[2,7]])`); await settle()
await ev(`S().addFloatDomToPosition({ componentKey: 'sheetly-chart', allowTransform: true, initPosition: { startX: 400, endX: 800, startY: 0, endY: 250 },
  data: { v: 2, type: 'scatter', title: null, source: { sheetId: S().getSheetId(), range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }, categories: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    series: [{ name: null, nameText: 'Y', values: { startRow: 0, endRow: 2, startColumn: 1, endColumn: 1 } }] } } })`); await settle(800)
st = await ev(`return cj(canv().length-1).data.datasets[0].data.map(p=>p.x+':'+p.y).join(',')`)
check('scatter points keep X = 0', st === '0:5,1:6,2:7', st)
await removeCharts(); await ev(`S().getRange('A1:B3').clear()`)
await ev(`S().getRange('A1:C6').setValues([['Region','Product','Sales'],['East','A',10],['East','B',20],['West','A',5],['West','B',7],['North','B',3]]); S().getRange('A1').activate()`); await settle()
await key('Control+q'); await key('c'); await key('ArrowRight', 3); await key('Enter'); await settle(800)

// ---- Copy Sheet: the copy's chart reads the copy's data
await removeCharts(); await ev(`S().getRange('A1:H20').clear(); S().getRange('A1:B3').setValues([['k','v'],['a',1],['b',2]]); S().getRange('A1').activate()`); await settle()
await key('Alt+F1'); await settle(800)
const ids = await ev(`return W().getSheets().map(s=>s.getSheetId()).join(',')`)
await ev(`await api._injector.get(window.__ICommandService).executeCommand('sheet.command.copy-sheet', { unitId: W().getId(), subUnitId: S().getSheetId() })`); await settle(1000)
const orig = await ev(`return S().getSheetId()`)
await ev(`const before='${ids}'.split(','); const c=W().getSheets().find(s=>!before.includes(s.getSheetId())); W().setActiveSheet(c); await new Promise(r=>setTimeout(r,600)); S().getRange('B2').setValue(50); await new Promise(r=>setTimeout(r,800))`)
st = await ev(`return canv().map(c=>c.__chart?.data.datasets[0].data.join(',')).join('|') + '|' + (CH.sourceOf(CH.charts()[0])?.sheetId === S().getSheetId())`)
check('a copied sheet\'s chart follows the copy\'s data, not the original\'s', st.split('|').includes('50,2') && st.endsWith('true'), st)
await ev(`W().setActiveSheet(W().getSheetBySheetId('${orig}') ?? W().getSheets()[0])`); await settle(600)
await removeCharts(); await ev(`S().getRange('A1:B3').clear()`); await settle()
await ev(`S().getRange('A1:C6').setValues([['Region','Product','Sales'],['East','A',10],['East','B',20],['West','A',5],['West','B',7],['North','B',3]]); S().getRange('A1').activate()`); await settle()
await key('Control+q'); await key('c'); await key('ArrowRight', 3); await key('Enter'); await settle(800)

// ---- saved with the workbook
const snap = await ev(`const s=api.getActiveWorkbook().save(); const d=JSON.parse(s.resources.find(r=>r.name==='SHEET_DRAWING_PLUGIN').data); return Object.values(d).flatMap(v=>Object.values(v.data)).map(x=>x.data).find(x=>x?.type==='pie')`)
check('the chart definition is saved (type + source range)', snap?.type === 'pie' && snap?.source?.range?.endRow === 5, JSON.stringify(snap))
await removeCharts()
check('no page errors', errors.length === 0, errors.join(' | '))
await app.close()
const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`charts  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
