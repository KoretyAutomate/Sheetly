// .xlsx fidelity (Node, no app): Sheetly's writer read back by ExcelJS, and workbooks built by ExcelJS read
// through xlsxToSnapshot. Covers styles (incl. snapshot style ids), dates, shared formulas, errors, row heights,
// freeze, tab colour, hidden sheets, defined names, notes and the outline settings.
import ExcelJS from 'exceljs'
import { xlsxToSnapshot, snapshotToXlsx } from '../src/xlsx.js'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const J = JSON.stringify

// ---- writer: a Univer snapshot -> .xlsx -> ExcelJS
{
  const snap = {
    sheetOrder: ['s1', 's2'],
    styles: { st1: { bl: 1, bg: { rgb: '#ff0000' } } },
    resources: [
      { name: 'SHEET_NOTE_PLUGIN', data: J({ s1: { 1: { 1: { id: 'x', row: 1, col: 1, note: 'hello' } } } }) },
      { name: 'SHEET_DEFINED_NAME_PLUGIN', data: J({ a: { id: 'a', name: 'MyName', formulaOrRefString: 'Sheet1!$A$1', localSheetId: 'AllDefaultWorkbook' } }) },
    ],
    sheets: {
      s1: { id: 's1', name: 'Sheet1', tabColor: '#00ff00', freeze: { startRow: 2, startColumn: -1, ySplit: 2, xSplit: 0 },
        custom: { sheetlyOutline: { summaryBelow: false, summaryRight: true, hideSymbols: true } },
        cellData: {
          0: { 0: { v: 'x', t: 1, s: 'st1' },
            1: { v: 3, t: 2, s: { it: 1, ul: { s: 1 }, st: { s: 1 }, fs: 14, ff: 'Arial', cl: { rgb: '#0000ff' } } },
            2: { v: 1, t: 2, s: { bd: { t: { s: 1, cl: { rgb: '#000000' } }, b: { s: 7, cl: { rgb: '#ff0000' } }, l: { s: 8, cl: { rgb: '#000000' } } } } },
            3: { v: 'w', t: 1, s: { ht: 2, vt: 1, tb: 3, tr: { a: -45, v: 0 } } },
            4: { v: 'v', t: 1, s: { tr: { a: 0, v: 1 } } } },
          1: { 1: { v: 5, t: 2 }, 2: { f: '=1/0', v: '#DIV/0!', t: 1 }, 3: { v: 45000, t: 2, s: { n: { pattern: 'm/d/yyyy' } } } },
        },
        rowData: { 3: { h: 40, ia: 0 }, 5: { h: 30, ia: 1 } } },
      s2: { id: 's2', name: 'Hidden', hidden: 1, cellData: {} },
    },
  }
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(await snapshotToXlsx(snap))
  const ws = wb.getWorksheet('Sheet1')
  const a1 = ws.getCell('A1'), b1 = ws.getCell('B1'), c1 = ws.getCell('C1'), d1 = ws.getCell('D1'), e1 = ws.getCell('E1')
  check('write: style id resolved (bold + fill)', a1.font?.bold === true && a1.fill?.fgColor?.argb === 'FFFF0000', J({ font: a1.font, fill: a1.fill }))
  check('write: font italic/underline/strike/size/name/colour', b1.font?.italic && b1.font?.underline && b1.font?.strike && b1.font?.size === 14 && b1.font?.name === 'Arial' && b1.font?.color?.argb === 'FF0000FF', J(b1.font))
  check('write: borders', c1.border?.top?.style === 'thin' && c1.border?.bottom?.style === 'double' && c1.border?.bottom?.color?.argb === 'FFFF0000' && c1.border?.left?.style === 'medium', J(c1.border))
  check('write: alignment (Univer angle-up -45 -> Excel 45)', d1.alignment?.horizontal === 'center' && d1.alignment?.vertical === 'top' && d1.alignment?.wrapText === true && d1.alignment?.textRotation === 45, J(d1.alignment))
  check('write: vertical text', e1.alignment?.textRotation === 'vertical', J(e1.alignment))
  check('write: formula error result', ws.getCell('C2').value?.formula === '1/0' && ws.getCell('C2').value?.result?.error === '#DIV/0!', J(ws.getCell('C2').value))
  check('write: number format', ws.getCell('D2').numFmt === 'm/d/yyyy')
  check('write: custom row height in points, auto height not written', ws.getRow(4).height === 30 && !ws.getRow(6).height, `${ws.getRow(4).height} ${ws.getRow(6).height}`)
  check('write: frozen panes', ws.views?.[0]?.state === 'frozen' && ws.views[0].ySplit === 2, J(ws.views))
  check('write: showOutlineSymbols=0', ws.views?.[0]?.showOutlineSymbols === false, J(ws.views))
  check('write: outlinePr summaryBelow=0', ws.properties?.outlineProperties?.summaryBelow === false && ws.properties?.outlineProperties?.summaryRight === true, J(ws.properties))
  check('write: tab colour', ws.properties?.tabColor?.argb === 'FF00FF00', J(ws.properties?.tabColor))
  check('write: hidden sheet', wb.getWorksheet('Hidden')?.state === 'hidden')
  check('write: defined name', J(wb.definedNames.model).includes('MyName') && J(wb.definedNames.model).includes('$A$1'), J(wb.definedNames.model))
  check('write: note', (typeof ws.getCell('B2').note === 'string' ? ws.getCell('B2').note : ws.getCell('B2').note?.texts?.map((t) => t.text).join('')) === 'hello', J(ws.getCell('B2').note))

  // and back through the reader
  const back = await xlsxToSnapshot(await snapshotToXlsx(snap))
  const s = back.sheets[back.sheetOrder[0]], c = s.cellData
  check('round-trip: bold + fill', c[0][0].s?.bl === 1 && c[0][0].s?.bg?.rgb === '#ff0000', J(c[0][0]))
  check('round-trip: font', J(c[0][1].s) === J({ it: 1, ul: { s: 1 }, st: { s: 1 }, fs: 14, ff: 'Arial', cl: { rgb: '#0000ff' } }), J(c[0][1].s))
  check('round-trip: borders', c[0][2].s?.bd?.t?.s === 1 && c[0][2].s?.bd?.b?.s === 7 && c[0][2].s?.bd?.b?.cl?.rgb === '#ff0000' && c[0][2].s?.bd?.l?.s === 8, J(c[0][2].s))
  check('round-trip: alignment + rotation (Univer -45 = Excel 45 counterclockwise)', c[0][3].s?.ht === 2 && c[0][3].s?.vt === 1 && c[0][3].s?.tb === 3 && c[0][3].s?.tr?.a === -45, J(c[0][3].s))
  check('round-trip: vertical text', c[0][4].s?.tr?.v === 1, J(c[0][4].s))
  check('round-trip: date serial stays a number', c[1][3].v === 45000 && c[1][3].t === 2, J(c[1][3]))
  check('round-trip: error result', c[1][2].f === '=1/0' && c[1][2].v === '#DIV/0!', J(c[1][2]))
  check('round-trip: row height', s.rowData[3]?.h === 40 && s.rowData[3]?.ia === 0, J(s.rowData))
  check('round-trip: freeze', s.freeze?.ySplit === 2 && s.freeze?.startRow === 2 && s.freeze?.xSplit === 0, J(s.freeze))
  check('round-trip: tab colour', s.tabColor === '#00ff00')
  check('round-trip: outline settings', J(s.custom?.sheetlyOutline) === J({ summaryBelow: false, summaryRight: true, hideSymbols: true }), J(s.custom))
  check('round-trip: hidden sheet', back.sheets[back.sheetOrder[1]].hidden === 1)
  const res = Object.fromEntries((back.resources ?? []).map((r) => [r.name, JSON.parse(r.data)]))
  check('round-trip: defined name', Object.values(res.SHEET_DEFINED_NAME_PLUGIN ?? {}).some((d) => d.name === 'MyName' && d.formulaOrRefString === 'Sheet1!$A$1'), J(res.SHEET_DEFINED_NAME_PLUGIN))
  check('round-trip: note', res.SHEET_NOTE_PLUGIN?.[back.sheetOrder[0]]?.[1]?.[1]?.note === 'hello', J(res.SHEET_NOTE_PLUGIN))
}

// ---- reader: workbook made by ExcelJS (stands in for one saved by Excel)
{
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Data')
  ws.getCell('A1').value = new Date(Date.UTC(2024, 0, 15))
  ws.getCell('A1').numFmt = 'yyyy-mm-dd'
  ws.getCell('B1').value = 1; ws.getCell('B2').value = 2; ws.getCell('B3').value = 3
  ws.getCell('C1').value = { formula: 'B1*2', result: 2, shareType: 'shared', ref: 'C1:C3' }
  ws.getCell('C2').value = { sharedFormula: 'C1', result: 4 }
  ws.getCell('C3').value = { sharedFormula: 'C1', result: 6 }
  ws.getCell('D1').value = { formula: 'NA()', result: { error: '#N/A' } }
  ws.getCell('E1').value = true
  ws.mergeCells('F1:G1'); ws.getCell('F1').value = 'merged'
  ws.getRow(1).outlineLevel = 1; ws.getRow(2).outlineLevel = 1
  ws.properties.outlineProperties = { summaryBelow: false, summaryRight: false }
  // Excel's "collapsed" flag on the summary, which is above the detail here: row 1's group would need a summary
  // at row 0, so use rows 3..4 with summary row 2.
  ws.getRow(1).outlineLevel = 0; ws.getRow(2).outlineLevel = 0
  ws.getRow(3).outlineLevel = 1; ws.getRow(4).outlineLevel = 1; ws.getRow(3).hidden = true; ws.getRow(4).hidden = true
  const snap = await xlsxToSnapshot(await wb.xlsx.writeBuffer())
  const s = snap.sheets[snap.sheetOrder[0]], c = s.cellData
  check('read: date becomes an Excel serial', c[0][0].v === 45306 && c[0][0].t === 2 && c[0][0].s?.n?.pattern === 'yyyy-mm-dd', J(c[0][0]))
  check('read: shared formula clones translated', c[1][2]?.f === '=B2*2' && c[2][2]?.f === '=B3*2' && c[2][2]?.v === 6, J([c[1][2], c[2][2]]))
  check('read: error result kept', c[0][3]?.f === '=NA()' && c[0][3]?.v === '#N/A', J(c[0][3]))
  check('read: boolean', c[0][4]?.v === 1 && c[0][4]?.t === 3, J(c[0][4]))
  check('read: merged-away cell left empty', c[0][6]?.v === undefined && J(s.mergeData) === J([{ startRow: 0, endRow: 0, startColumn: 5, endColumn: 6 }]), J([c[0][6], s.mergeData]))
  check('read: summary above -> collapse bit on the row before the group', (s.rowData[1]?.custom?.outlineCollapsed & 2) === 2 && !s.rowData[5]?.custom?.outlineCollapsed, J(s.rowData))
  check('read: summary direction', s.custom?.sheetlyOutline?.summaryBelow === false && s.custom?.sheetlyOutline?.summaryRight === false, J(s.custom))
}

// ---- reviewer findings (2026-10-16): colours, blank styled cells, row/column styles, shared formulas, _xlfn,
// very hidden sheets, scrolled freeze, sheet-scoped / constant names, auto-fit rows, format 14, error constants.
{
  const snap = {
    sheetOrder: ['s1', 's2'],
    styles: {},
    resources: [{ name: 'SHEET_DEFINED_NAME_PLUGIN', data: J({
      a: { id: 'a', name: 'Rate', formulaOrRefString: '0.08', localSheetId: 'AllDefaultWorkbook' },
      b: { id: 'b', name: 'Local', formulaOrRefString: 'Sheet1!$B$1', localSheetId: 's1' },
      c: { id: 'c', name: 'Local', formulaOrRefString: 'Sheet1!$C$1', localSheetId: 'AllDefaultWorkbook' } }) }],
    sheets: {
      s1: { id: 's1', name: 'Sheet1', freeze: { xSplit: 1, ySplit: 2, startRow: 9, startColumn: 1 },
        cellData: {
          0: { 0: { v: 1, t: 2, s: { bg: { rgb: '#f00' }, cl: { rgb: 'rgb(0, 128, 255)' } } }, 1: { s: { bd: { b: { s: 1, cl: { rgb: '#000' } } } } },
            2: { f: '=XLOOKUP(1,A1:A3,B1:B3)', v: 1, t: 2 }, 3: { f: '=FILTER(A1:A3,A1:A3>0)', v: 1, t: 2 }, 4: { v: '#N/A', t: 1 }, 5: { v: '#N/A' } },
          1: { 0: { v: 2, t: 2 }, 2: { si: 'x1', v: 2, t: 2 } },
          4: { 0: { v: 'in styled row', t: 1 } },
        },
        rowData: { 4: { s: { bl: 1 } } }, columnData: { 6: { s: { it: 1 }, w: 100 } } },
      s2: { id: 's2', name: 'Secret', hidden: 2, cellData: {} },
    },
  }
  const formulaOf = (id, r, c) => (id === 's1' && r === 1 && c === 2 ? '=A2*2' : null)
  const buf = await snapshotToXlsx(snap, { formulaOf })
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf)
  const ws = wb.getWorksheet('Sheet1')
  check('write: #rgb and rgb() colours normalised', ws.getCell('A1').fill?.fgColor?.argb === 'FFFF0000' && ws.getCell('A1').font?.color?.argb === 'FF0080FF' && ws.getCell('B1').border?.bottom?.color?.argb === 'FF000000', J([ws.getCell('A1').fill, ws.getCell('A1').font, ws.getCell('B1').border]))
  check('write: shared-formula cell gets its formula from the resolver', ws.getCell('C2').value?.formula === 'A2*2', J(ws.getCell('C2').value))
  check('write: _xlfn. / _xlfn._xlws. prefixes added', ws.getCell('C1').value?.formula === '_xlfn.XLOOKUP(1,A1:A3,B1:B3)' && ws.getCell('D1').value?.formula === '_xlfn._xlws.FILTER(A1:A3,A1:A3>0)', J([ws.getCell('C1').value, ws.getCell('D1').value]))
  check('write: text "#N/A" stays text, an error value is written as an error', ws.getCell('E1').value === '#N/A' && ws.getCell('F1').value?.error === '#N/A', J([ws.getCell('E1').value, ws.getCell('F1').value]))
  check('write: row style (bold row) and its cells', ws.getRow(5).font?.bold === true && ws.getCell('A5').font?.bold === true, J([ws.getRow(5).style, ws.getCell('A5').font]))
  check('write: column style (italic column G)', ws.getColumn(7).font?.italic === true, J(ws.getColumn(7).style))
  check('write: very hidden sheet', wb.getWorksheet('Secret')?.state === 'veryHidden')
  check('write: freeze keeps the scrolled top-left cell', ws.views[0]?.topLeftCell === 'B10' && ws.views[0]?.ySplit === 2 && ws.views[0]?.xSplit === 1, J(ws.views[0]))
  const back = await xlsxToSnapshot(buf), s = back.sheets[back.sheetOrder[0]]
  const names = Object.values(JSON.parse(back.resources.find((r) => r.name === 'SHEET_DEFINED_NAME_PLUGIN').data))
  check('round-trip: constant name', names.some((d) => d.name === 'Rate' && d.formulaOrRefString === '0.08'), J(names))
  check('round-trip: sheet-scoped and workbook names with the same name both kept', names.filter((d) => d.name === 'Local').map((d) => d.localSheetId).sort().join() === ['AllDefaultWorkbook', back.sheetOrder[0]].sort().join(), J(names))
  check('round-trip: blank cell with a border kept', s.cellData[0][1]?.s?.bd?.b?.s === 1 && s.cellData[0][1]?.v === undefined, J(s.cellData[0][1]))
  check('round-trip: _xlfn. removed on load', s.cellData[0][2].f === '=XLOOKUP(1,A1:A3,B1:B3)' && s.cellData[0][3].f === '=FILTER(A1:A3,A1:A3>0)', J([s.cellData[0][2].f, s.cellData[0][3].f]))
  check('round-trip: row and column styles', s.rowData[4]?.s?.bl === 1 && s.columnData[6]?.s?.it === 1, J([s.rowData[4], s.columnData[6]]))
  check('round-trip: very hidden', back.sheets[back.sheetOrder[1]].hidden === 2)
  check('round-trip: scrolled freeze', J(s.freeze) === J({ xSplit: 1, ySplit: 2, startRow: 9, startColumn: 1 }), J(s.freeze))
}
{
  // A workbook as Excel writes it: theme / indexed colours, an auto-fit row (ht, no customHeight), format 14.
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('T')
  ws.getCell('A1').value = 'theme'; ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4 } }
  ws.getCell('A2').value = 'tint'; ws.getCell('A2').font = { color: { theme: 4, tint: 0.3999755851924192 } }
  ws.getCell('A3').value = 'idx'; ws.getCell('A3').fill = { type: 'pattern', pattern: 'solid', fgColor: { indexed: 10 } }
  ws.getCell('A4').value = 45306; ws.getCell('A4').numFmt = 'mm-dd-yy'
  ws.getCell('A6').value = 'x'
  let buf = Buffer.from(await wb.xlsx.writeBuffer())
  // mark row 6 as auto-fit: give it ht without customHeight, as Excel does after autofit
  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(buf)
  let xml = await zip.file('xl/worksheets/sheet1.xml').async('string')
  xml = xml.replace(/<row r="6"([^>]*)>/, '<row r="6"$1 ht="30">').replace(/<row r="1"([^>]*)>/, '<row r="1"$1 ht="40" customHeight="1">')
  zip.file('xl/worksheets/sheet1.xml', xml)
  buf = await zip.generateAsync({ type: 'uint8array' })
  const snap = await xlsxToSnapshot(buf), c = snap.sheets[snap.sheetOrder[0]].cellData, rows = snap.sheets[snap.sheetOrder[0]].rowData
  check('read: theme colour (accent1 of the file\'s theme, Office 2007 here) resolved', c[0][0].s?.bg?.rgb === '#4f81bd', J(c[0][0].s))
  check('read: theme colour with tint (Accent 1, Lighter 40% = 95B3D7)', c[1][0].s?.cl?.rgb === '#95b3d7', J(c[1][0].s))
  check('read: indexed colour', c[2][0].s?.bg?.rgb === '#ff0000', J(c[2][0].s))
  check('read: built-in format 14 shows as m/d/yyyy', c[3][0].s?.n?.pattern === 'm/d/yyyy', J(c[3][0].s))
  check('read: auto-fit row stays automatic, custom row fixed', rows[5]?.h === undefined && rows[0]?.h === 53 && rows[0]?.ia === 0, J(rows))
}

const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`xlsx fidelity  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
