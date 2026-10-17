// Opening an .xlsx with every kind of object in the app, and saving it again: conditional formats, data
// validation, a table, a picture, a chart (linked to its range), a spilled dynamic array, hyperlinks, Excel's
// default sizes and font; File > Open keeps an edited workbook.
import { _electron as electron } from 'playwright'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { snapshotToXlsx } from '../src/xlsx.js'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const J = JSON.stringify
const R = (a, b, c, d) => ({ startRow: a, endRow: c ?? a, startColumn: b, endColumn: d ?? b })
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const link = (text, url) => ({ p: { id: 'd', body: { dataStream: text + '\r\n', paragraphs: [{ startIndex: text.length }], customRanges: [{ startIndex: 0, endIndex: text.length - 1, rangeType: 0, rangeId: 'r', properties: { url } }] } } })
const cells = {
  0: { 0: { v: 'Region', t: 1 }, 1: { v: 'Sales', t: 1 }, 3: link('site', 'https://example.com/'), 4: link('jump', '#gid=s2&range=C5') },
  1: { 0: { v: 'North', t: 1 }, 1: { v: 10, t: 2 } }, 2: { 0: { v: 'South', t: 1 }, 1: { v: 20, t: 2 } }, 3: { 0: { v: 'East', t: 1 }, 1: { v: 30, t: 2 } },
  9: { 0: { f: '=SEQUENCE(3)', v: 1, t: 2 } },
}
const snap = {
  sheetOrder: ['s1', 's2'],
  resources: [
    { name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: J({ s1: [{ cfId: 'a', ranges: [R(1, 1, 3, 1)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 15, style: { bg: { rgb: '#ffc7ce' } } } }] }) },
    { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: J({ s1: [{ uid: 'v', ranges: [R(1, 2, 3, 2)], type: 'list', formula1: '["Yes","No"]' }] }) },
    { name: 'SHEET_TABLE_PLUGIN', data: J({ s1: { tables: [{ id: 't', name: 'Sales', range: R(0, 0, 3, 1), options: { showHeader: true }, columns: [{ displayName: 'Region' }, { displayName: 'Sales' }] }] } }) },
    { name: 'SHEET_DRAWING_PLUGIN', data: J({ s1: { data: {
      p: { drawingId: 'p', drawingType: 0, imageSourceType: 'BASE64', source: `data:image/png;base64,${PNG}`, sheetTransform: { from: { column: 6, columnOffset: 0, row: 0, rowOffset: 0 }, to: { column: 7, columnOffset: 0, row: 3, rowOffset: 0 } } },
      c: { drawingId: 'c', drawingType: 8, componentKey: 'sheetly-chart', data: { v: 2, type: 'line', source: { sheetId: 's1', range: R(0, 0, 3, 1) }, title: null }, sheetTransform: { from: { column: 6, columnOffset: 0, row: 5, rowOffset: 0 }, to: { column: 13, columnOffset: 0, row: 19, rowOffset: 0 } } },
    }, order: ['p', 'c'] } }) },
  ],
  sheets: { s1: { id: 's1', name: 'Data', cellData: cells }, s2: { id: 's2', name: 'Other', cellData: {} } },
}
const bytes = Array.from(await snapshotToXlsx(snap, { arrayRangeOf: (id, r, c) => (id === 's1' && r === 9 && c === 0 ? R(9, 0, 11, 0) : null) }))

const app = await electron.launch({ args: ['.', '--no-sandbox'], env: { ...process.env, SHEETLY_TEST_NO_DIALOGS: '1' } })
const win = await app.firstWindow()
const errors = []
win.on('pageerror', (e) => errors.push(e.message))
await win.waitForFunction(() => window.univerAPI?.getActiveWorkbook?.(), null, { timeout: 60000 })
await win.waitForTimeout(1500)
const ev = (code, arg) => win.evaluate(new Function('arg', `return (async()=>{const api=window.univerAPI,W=()=>api.getActiveWorkbook(),S=()=>W().getActiveSheet(),X=window.__sheetly;${code}})()`), arg)

check('a new workbook has Excel\'s sizes: columns 64 px, rows 20 px', J(await ev(`return [S().getColumnWidth(0), S().getRowHeight(0)]`)) === '[64,20]')
check('Calibri (bundled Carlito metrics) is the sheet font and is loaded', (await ev(`return S().getSheet().getConfig().defaultStyle?.ff === 'Calibri' && document.fonts.check('11pt Calibri')`)) === true)
check('a new, untouched Book1 may be replaced by File > Open', (await ev(`return X.untouched()`)) === true)
await ev(`await X.xlsx.load(new Uint8Array(arg), 'objects.xlsx')`, bytes)
await win.waitForTimeout(2000)
check('opened: sheet names', (await ev(`return W().getSheets().map(s=>s.getSheetName()).join(',')`)) === 'Data,Other')
check('opened: the conditional format is active', (await ev(`return S().getConditionalFormattingRules().map(r=>r.rule.subType+':'+r.rule.value).join(',')`)) === 'number:15')
check('opened: the list validation is on C2:C4', (await ev(`return S().getRange('C3').getDataValidation()?.getCriteriaType?.() ?? S().getRange('C3').getDataValidation()?.rule?.type`)) === 'list')
check('opened: the table exists', (await ev(`return (S().getTableByCell?.(1,0) ?? null)?.name ?? null`)) !== null, String(await ev(`return J?1:0`).catch(() => '')))
check('opened: the picture is on the sheet', (await ev(`return S().getImages().length`)) === 1)
const chartType = await ev(`await new Promise(r=>setTimeout(r,500)); const c=document.querySelector('.sx-chart canvas'); return c?.__chart?.config.type + ':' + c?.__chart?.data.datasets[0]?.data.join(',')`)
check('opened: the chart is drawn, linked to B2:B4', chartType === 'line:10,20,30', chartType)
await ev(`S().getRange('B2').setValue(50)`); await win.waitForTimeout(600)
check('the opened chart follows its data', (await ev(`return document.querySelector('.sx-chart canvas').__chart.data.datasets[0].data.join(',')`)) === '50,20,30')
const spill = await ev(`return [S().getRange('A10').getValue(),S().getRange('A11').getValue(),S().getRange('A12').getValue()].join(',') + ' f=' + S().getRange('A10').getFormula() + ' raw=' + JSON.stringify(S().getRange('A11').getCellData())`)
check('opened: the dynamic array spills (A10:A12 = 1,2,3)', spill.startsWith('1,2,3 '), spill)
check('opened: hyperlinks (web and to Other!C5)', (await ev(`const u=(a)=>S().getRange(a).getCellData()?.p?.body?.customRanges?.[0]?.properties?.url; return u('D1')+'|'+u('E1')`)) === `https://example.com/|#gid=${await ev(`return W().getSheets()[1].getSheetId()`)}&range=C5`)
check('an opened file is not replaced by File > Open (it opens in a new window)', (await ev(`return X.untouched()`)) === false)

// ---- save from the app and look at the file
const saved = Buffer.from(await ev(`return Array.from(await X.xlsx.save())`))
const wb = new ExcelJS.Workbook(); await wb.xlsx.load(saved)
const ws = wb.getWorksheet('Data'), zip = await JSZip.loadAsync(saved)
const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string')
check('saved: conditional format and validation', (ws.conditionalFormattings ?? []).length === 1 && /<dataValidation [^>]*type="list"/.test(sheetXml), J(ws.conditionalFormattings))
check('saved: the table', (ws.getTables?.() ?? []).length === 1, J((ws.getTables?.() ?? []).map((t) => t.table?.ref)))
check('saved: the picture', (ws.getImages?.() ?? []).length === 1)
const chartPart = Object.keys(zip.files).find((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))
const chartXml = chartPart ? await zip.file(chartPart).async('string') : ''
check('saved: the chart part (line over Data!$B$2:$B$4, current value 50 cached)', /<c:lineChart>/.test(chartXml) && chartXml.includes('Data!$B$2:$B$4') && chartXml.includes('<c:v>50</c:v>'), chartXml.slice(0, 200))
check('saved: the spilling formula as a dynamic array', /<c r="A10"[^>]*cm="1"[^>]*><f t="array" ref="A10:A12">_xlfn.SEQUENCE\(3\)<\/f>/.test(sheetXml), sheetXml.match(/<c r="A10"[^>]*>.{0,60}/)?.[0])
check('saved: links (external and in-workbook)', ws.getCell('D1').value?.hyperlink === 'https://example.com/' && /location="(&apos;|')Other(&apos;|')!C5"/.test(sheetXml), J([ws.getCell('D1').value, sheetXml.match(/<hyperlink [^>]*>/g)]))

// legacy (Ctrl+Shift+Enter) array formulas, settled by calculating: one value fills the range, an array spills
{
  const w = new ExcelJS.Workbook(), t = w.addWorksheet('L')
  t.getCell('A1').value = 1; t.getCell('A2').value = 1; t.getCell('A3').value = 4
  t.getCell('C1').value = { formula: 'SUM(A1:A3)', result: 6, shareType: 'array', ref: 'C1:C3' }; t.getCell('C2').value = 6; t.getCell('C3').value = 6
  t.getCell('E1').value = { formula: 'A1:A2', result: 1, shareType: 'array', ref: 'E1:E2' }; t.getCell('E2').value = 1
  const b = Array.from(new Uint8Array(await w.xlsx.writeBuffer()))
  await ev(`await X.xlsx.load(new Uint8Array(arg), 'legacy.xlsx')`, b); await win.waitForTimeout(2000)
  const got = await ev(`const v=(a)=>S().getRange(a).getCalculatedValue(); return JSON.stringify([v('C1'),v('C2'),v('C3'),S().getRange('C3').getFormula(),v('E1'),v('E2'),v('E3')])`)
  check('legacy array formulas: a single result fills C1:C3 (6, with the formula); an array result spills into E1:E2 (1, 1, no #SPILL!)', got === JSON.stringify([6, 6, 6, '=SUM(A1:A3)', 1, 1, null]), got)
  check('opening the file left nothing to undo', (await ev(`return api._injector.get(window.__IUndoRedoService).getUndoRedoStatus(W().getId()).undos`)) === 0)
}

// a table named like a sheet (Excel allows it; Univer refuses it) survives open + save; opening leaves no undo
{
  const w = new ExcelJS.Workbook(), t = w.addWorksheet('Sales')
  t.addTable({ name: 'Sales', ref: 'A1', headerRow: true, columns: [{ name: 'Item' }, { name: 'Qty' }], rows: [['a', 1], ['b', 2]] })
  t.addTable({ name: 'Other', ref: 'D1', headerRow: true, columns: [{ name: 'K' }], rows: [['x']] })
  const b = Array.from(new Uint8Array(await w.xlsx.writeBuffer()))
  await ev(`await X.xlsx.load(new Uint8Array(arg), 'tables.xlsx')`, b); await win.waitForTimeout(1500)
  check('opening a workbook with tables leaves nothing to undo', (await ev(`return api._injector.get(window.__IUndoRedoService).getUndoRedoStatus(W().getId()).undos`)) === 0)
  const saved = Buffer.from(await ev(`return Array.from(await X.xlsx.save())`))
  const w2 = new ExcelJS.Workbook(); await w2.xlsx.load(saved)
  const names = w2.getWorksheet('Sales').getTables().map((x) => (x.table ?? x).name).sort().join(',')
  check('a table named like its sheet ("Sales") is kept on save, next to the other table', names === 'Other,Sales', names)
  await ev(`S().insertRowsBefore(0, 2)`); await win.waitForTimeout(500)
  const moved = Buffer.from(await ev(`return Array.from(await X.xlsx.save())`))
  const w3 = new ExcelJS.Workbook(); await w3.xlsx.load(moved)
  const ref = w3.getWorksheet('Sales').getTables().map((x) => x.table ?? x).find((x) => x.name === 'Sales')?.tableRef ?? w3.getWorksheet('Sales').getTables().map((x) => x.table ?? x).find((x) => x.name === 'Sales')?.ref
  check('the kept table moves with rows inserted above it (A1:B3 -> A3:B5)', ref === 'A3:B5', String(ref))
  // Copy Sheet: the copy's kept table follows the copy's cells, not the original's
  const ids = await ev(`return W().getSheets().map(s=>s.getSheetId()).join(',')`)
  await ev(`await api._injector.get(window.__ICommandService).executeCommand('sheet.command.copy-sheet', { unitId: W().getId(), subUnitId: S().getSheetId() })`); await win.waitForTimeout(900)
  const copyName = await ev(`const before='${ids}'.split(','); return W().getSheets().find(s=>!before.includes(s.getSheetId())).getSheetName()`)
  await ev(`W().getSheetByName('Sales').insertRowsBefore(0, 3)`); await win.waitForTimeout(500)
  const w4 = new ExcelJS.Workbook(); await w4.xlsx.load(Buffer.from(await ev(`return Array.from(await X.xlsx.save())`)))
  const refOf = (sheet) => { const t = w4.getWorksheet(sheet)?.getTables().map((x) => x.table ?? x).find((x) => /^Sales/.test(x.name)); return t?.tableRef ?? t?.ref }
  check('a copied sheet\'s kept table stays on the copy\'s cells when the original moves (copy A3:B5, original A6:B8)', refOf(copyName) === 'A3:B5' && refOf('Sales') === 'A6:B8', `${copyName} ${refOf(copyName)} ${refOf('Sales')}`)
}

// Ctrl+; in a 1904 workbook writes today's 1904 serial
{
  const w = new ExcelJS.Workbook(); w.properties.date1904 = true; w.addWorksheet('D').getCell('A1').value = 1
  const b = Array.from(new Uint8Array(await w.xlsx.writeBuffer()))
  await ev(`await X.xlsx.load(new Uint8Array(arg), 'd1904.xlsx')`, b); await win.waitForTimeout(1500)
  await ev(`S().getRange('B2').activate()`); await win.waitForTimeout(200)
  await win.keyboard.press('Control+;'); await win.waitForTimeout(300)
  const today = Math.floor((Date.now() - new Date().getTimezoneOffset() * 60000 - Date.UTC(1899, 11, 30)) / 864e5) - 1462
  const got = await ev(`return S().getRange('B2').getRawValue()`)
  check('Ctrl+; in a 1904 workbook writes the 1904 serial for today', got === today, `${got} vs ${today}`)
}

check('no page errors', errors.length === 0, errors.join(' | '))
await app.close()
const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`files (open / save objects in the app)  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
