// .xlsx fidelity (Node, no app): Sheetly's writer read back by ExcelJS, and workbooks built by ExcelJS read
// through xlsxToSnapshot. Covers styles (incl. snapshot style ids), dates, shared formulas, errors, row heights,
// freeze, tab colour, hidden sheets, defined names, notes and the outline settings.
import ExcelJS from 'exceljs'
import { xlsxToSnapshot, snapshotToXlsx, formulaOut, formulaIn } from '../src/xlsx.js'
import { pivotXml, readPivot } from '../src/xlsx-pivots.js'
import { layout, readSource } from '../src/pivot-model.js'
import JSZip from 'jszip'

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
  check('round-trip: outline settings', J(s.custom?.sheetlyOutline) === J({ summaryBelow: false, summaryRight: true, hideSymbols: true, autoStyles: false }), J(s.custom))
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

// ---- hyperlinks: web and in-workbook
{
  const link = (text, url) => ({ p: { id: 'd', body: { dataStream: text + '\r\n', paragraphs: [{ startIndex: text.length }], customRanges: [{ startIndex: 0, endIndex: text.length - 1, rangeType: 0, rangeId: 'r', properties: { url } }] } } })
  const snap = { sheetOrder: ['s1', 's2', 's3'], sheets: {
    s1: { id: 's1', name: 'Sheet1', cellData: { 0: { 0: link('site', 'https://example.com/'), 1: link('jump', '#gid=s2&range=B3'), 2: link('obrien', '#gid=s3&range=A1') } } },
    s2: { id: 's2', name: 'My Data', cellData: {} }, s3: { id: 's3', name: "O'Brien", cellData: {} } } }
  const buf = await snapshotToXlsx(snap)
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf)
  const ws = wb.getWorksheet('Sheet1')
  check('write: web hyperlink', ws.getCell('A1').value?.hyperlink === 'https://example.com/' && ws.getCell('A1').value?.text === 'site', J(ws.getCell('A1').value))
  check('write: link into the workbook uses the sheet name', ws.getCell('B1').value?.hyperlink === "#'My Data'!B3", J(ws.getCell('B1').value))
  const back = await xlsxToSnapshot(buf), c = back.sheets[back.sheetOrder[0]].cellData
  const url = (x) => x?.p?.body?.customRanges?.[0]?.properties?.url
  check('round-trip: web hyperlink', url(c[0][0]) === 'https://example.com/' && c[0][0].p.body.dataStream === 'site\r\n', J(c[0][0]))
  check('round-trip: in-workbook link back to #gid form', url(c[0][1]) === `#gid=${back.sheetOrder[1]}&range=B3`, J(c[0][1]))
  check("round-trip: a link to a sheet named O'Brien", url(c[0][2]) === `#gid=${back.sheetOrder[2]}&range=A1`, J(c[0][2]))
}

// ---- LET / LAMBDA parameters are stored as _xlpm.name (Excel would read a bare name as a defined name)
{
  const cases = [['LET(x,2,y,x*3,x+y)', '_xlfn.LET(_xlpm.x,2,_xlpm.y,_xlpm.x*3,_xlpm.x+_xlpm.y)'], ['BYROW(A1:A3,LAMBDA(r,SUM(r)))', '_xlfn.BYROW(A1:A3,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))'],
    ['LET(x,"x",x&"x")', '_xlfn.LET(_xlpm.x,"x",_xlpm.x&"x")'], ['SUM(x1,A1)', 'SUM(x1,A1)'], ['LET(a1x,5,a1x+A1)', '_xlfn.LET(_xlpm.a1x,5,_xlpm.a1x+A1)'],
    ['IF(A1="LET(x,1,x)",1,0)', 'IF(A1="LET(x,1,x)",1,0)'], ['LET(x,1,SUM(Table1[x]))', '_xlfn.LET(_xlpm.x,1,SUM(Table1[x]))'], ['LET(x,1,x&"x""x")', '_xlfn.LET(_xlpm.x,1,_xlpm.x&"x""x")'],
    ['LET(fn,LAMBDA(x,x+1),fn(5))', '_xlfn.LET(_xlpm.fn,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.fn(5))'], ['LAMBDA(f,v,f(v))(SQRT,4)', '_xlfn.LAMBDA(_xlpm.f,_xlpm.v,_xlpm.f(_xlpm.v))(SQRT,4)'],
    ['LAMBDA(max,MAX(max,1))', '_xlfn.LAMBDA(_xlpm.max,MAX(_xlpm.max,1))'], ['LET(sum,5,SUM(sum,1))', '_xlfn.LET(_xlpm.sum,5,SUM(_xlpm.sum,1))'], ["LET(x,1,'Sales x Plan'!A1+x)", "_xlfn.LET(_xlpm.x,1,'Sales x Plan'!A1+_xlpm.x)"], ["LET(x,1,'It''s x'!A1+x)", "_xlfn.LET(_xlpm.x,1,'It''s x'!A1+_xlpm.x)"], ['LET(a,y,y,1,a+y)', '_xlfn.LET(_xlpm.a,y,_xlpm.y,1,_xlpm.a+_xlpm.y)'], ['LET(x,x+1,x)', '_xlfn.LET(_xlpm.x,x+1,_xlpm.x)'], ['LET(fn,LAMBDA(x,x+1),alias,fn,alias(2))', '_xlfn.LET(_xlpm.fn,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.alias,_xlpm.fn,_xlpm.alias(2))'],
    ['LET(sin,0,SIN(sin))', '_xlfn.LET(_xlpm.sin,0,SIN(_xlpm.sin))'], ['LET(sumproduct,2,SUMPRODUCT(sumproduct))', '_xlfn.LET(_xlpm.sumproduct,2,SUMPRODUCT(_xlpm.sumproduct))'], ['LET(x,1,税x+x)', '_xlfn.LET(_xlpm.x,1,税x+_xlpm.x)'], ['LET(x,2,X$1+x)', '_xlfn.LET(_xlpm.x,2,X$1+_xlpm.x)'], ['LET(sum,LAMBDA(x,x+1),LET(sum,5,SUM(sum,1)))', '_xlfn.LET(_xlpm.sum,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlfn.LET(_xlpm.sum,5,SUM(_xlpm.sum,1)))'], ['LET(sum,REDUCE(0,{1,2},LAMBDA(a,b,a+b)),SUM(sum,4))', '_xlfn.LET(_xlpm.sum,_xlfn.REDUCE(0,{1,2},_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)),SUM(_xlpm.sum,4))'], ['LET(x,2,$X$1+x)', '_xlfn.LET(_xlpm.x,2,$X$1+_xlpm.x)']]
  const bad = cases.filter(([f, want]) => formulaOut(f) !== want || formulaIn(formulaOut(f)) !== f)
  check('write: LET / LAMBDA parameters as _xlpm. (strings, cell refs and other names untouched), read back bare', !bad.length, J(bad.map(([f]) => [f, formulaOut(f)])))
}

// ---- PivotTables: Excel's cache and table parts, read back into the same report; what Sheetly cannot show stays values
{
  const cache = { fields: ['Region', 'Product', 'Sales'], records: [['East', 'A', 10], ['East', 'B', 20], ['West', 'A', 5], ['North', 'B', 3], ['East', 7, 4], ['west', null, 6], ['North', true, '#N/A']] }
  const spec = { name: 'PT', cache, rows: ['Product'], cols: [], values: [{ field: 'Sales', fn: 'sum' }, { field: 'Product', fn: 'count' }, { field: 'Sales', fn: 'stdDev', name: 'Spread' }],
    filters: [{ field: 'Region', items: ['East', 'WEST'] }], source: { sheetName: "It's", range: { startRow: 0, endRow: 7, startColumn: 0, endColumn: 2 } },
    location: { startRow: 3, endRow: 9, startColumn: 0, endColumn: 3 }, firstDataRow: 1, firstDataCol: 1 }
  const x = pivotXml(spec, 1)
  check('pivot cache: items listed for report fields (mixed, blank, logical) and fields with errors; plain numbers inline', /<cacheField name="Product" numFmtId="0"><sharedItems containsBlank="1" containsMixedTypes="1" containsNumber="1" containsInteger="1" minValue="7" maxValue="7" count="5"><n v="7"\/><s v="A"\/><s v="B"\/><b v="1"\/><m\/><\/sharedItems>/.test(x.cache)
    && /<r><x v="0"\/><x v="1"\/><x v="4"\/><\/r>/.test(x.records) && /<e v="#N\/A"\/><\/sharedItems>/.test(x.cache) && /sheet="It's"/.test(x.cache)
    && /<r><x v="0"\/><x v="1"\/><n v="10"\/><\/r>/.test(pivotXml({ ...spec, cache: { ...cache, records: cache.records.slice(0, 6) } }, 9).records), x.cache + x.records)
  check('pivot table: several values make a Values column field; a filter of two items hides the third; captions and functions', /<colFields count="1"><field x="-2"\/><\/colFields>/.test(x.table)
    && /<pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0"><items count="4"><item x="0"\/><item x="1" h="1"\/><item x="2"\/>/.test(x.table)
    && /<dataField name="Spread" fld="2" subtotal="stdDev"/.test(x.table) && /<dataField name="Sum of Sales" fld="2" baseField/.test(x.table) && /rowPageCount="1"/.test(x.table), x.table)
  const back = readPivot(x.table, x.cache, x.records)
  check('pivot parts read back: fields, values, filter items, records', J([back?.rows, back?.values, back?.filters, back?.cache.records, back?.source]) === J([['Product'], spec.values, [{ field: 'Region', items: ['East', 'West'] }], cache.records.map((r) => r.map((v) => (v === 'west' ? 'West' : v))), { range: spec.source.range, sheet: "It's" }]),
    J([back?.rows, back?.values, back?.filters, back?.cache.records, back?.source]))
  const one = pivotXml({ ...spec, filters: [{ field: 'Region', items: ['North'] }], values: [{ field: 'Sales', fn: 'sum' }] }, 2)
  check('a filter on one item is the page field\'s item; one value field has no Values field', /<pageField fld="0" item="1" hier="-1"\/>/.test(one.table) && !/<colFields/.test(one.table) && J(readPivot(one.table, one.cache, one.records)?.filters) === J([{ field: 'Region', items: ['North'] }]), one.table)
  const ws = { fields: ['Re\ngion', 'Sales'], records: [['A\nB', 1], ['A B', 2], ['C\tD\r', 3]] }
  const wx = pivotXml({ ...spec, cache: ws, rows: ['Re\ngion'], values: [{ field: 'Sales', fn: 'sum' }], filters: [] }, 3)
  check('line breaks and tabs in names and items are written as character references and read back', /name="Re&#10;gion"/.test(wx.cache) && /<s v="A&#10;B"\/>/.test(wx.cache) && /<s v="C&#9;D&#13;"\/>/.test(wx.cache)
    && J(readPivot(wx.table, wx.cache, wx.records)?.cache) === J(ws), wx.cache)
  const shownAs = x.table.replace('<dataField name="Sum of Sales" fld="2"', '<dataField name="Sum of Sales" fld="2" showDataAs="percentOfTotal"')
  const grouped = x.cache.replace('<cacheField name="Sales" numFmtId="0">', '<cacheField name="Sales" numFmtId="0"><fieldGroup base="2"/>')
  const external = x.cache.replace('<cacheSource type="worksheet">', '<cacheSource type="external" connectionId="1">')
  // review: filters laid out across, caches without records, t="data" items, 1904 dates
  const across = x.table.replace('<pivotTableDefinition ', '<pivotTableDefinition pageOverThenDown="1" ')
  const noData = x.cache.replace('refreshOnLoad="1"', 'refreshOnLoad="1" saveData="0"')
  check('filters laid out across, or a cache saved without its records, are not read', readPivot(across, x.cache, x.records) === null && readPivot(x.table, noData, x.records) === null && readPivot(x.table, x.cache, null) === null)
  const typed = one.table.replace(/<item x="(\d+)"\/>/g, '<item x="$1" t="data"/>')
  check('items typed t="data" are items (the filter item is found)', J(readPivot(typed, one.cache, one.records)?.filters) === J([{ field: 'Region', items: ['North'] }]), typed)
  const dated = x.cache.replace('<s v="East"/>', '<d v="2026-10-17T00:00:00"/>')
  const d1900 = readPivot(x.table, dated, x.records)?.cache.records[0][0], d1904 = readPivot(x.table, dated, x.records, { date1904: true })?.cache.records[0][0]
  check('cache dates become serials of the workbook\'s date system (46312 / 44850)', d1900 === 46312 && d1904 === 44850, `${d1900} ${d1904}`)
  const linked = x.cache.replace('<worksheetSource ref="A1:C8"', '<worksheetSource r:id="rId2" ref="A1:C8"')
  const calcItems = x.cache.replace('</cacheFields>', '</cacheFields><calculatedItems count="1"><calculatedItem formula="East+West"><pivotArea/></calculatedItem></calculatedItems>')
  const calcItem = x.table.replace('<item x="0"/>', '<item x="0"/><item f="1" x="3"/>')
  const ctl = { fields: ['K', 'Sales'], records: [['a\u0001b', 1], ['_x0041_', 2]] }
  const cxml = pivotXml({ ...spec, cache: ctl, rows: ['K'], values: [{ field: 'Sales', fn: 'sum', name: 'Sum of Sales ' }], filters: [] }, 4)
  const cback = readPivot(cxml.table, cxml.cache, cxml.records)
  check('control characters are written as _xHHHH_ (text like _x0041_ escaped) and read back; a caption keeps its trailing space',
    /<s v="a_x0001_b"\/>/.test(cxml.cache) && /<s v="_x005F_x0041_"\/>/.test(cxml.cache) && !/[\u0001]/.test(cxml.cache) && J(cback?.cache) === J(ctl) && cback?.values[0].name === 'Sum of Sales ', cxml.cache)
  const bl = { fields: ['K', 'N'], records: [[null, 1], ['(blank)', 2], ['x', 3]] }
  const bx = pivotXml({ ...spec, cache: bl, rows: [], filters: [{ field: 'K', items: ['(blank)'] }], values: [{ field: 'N', fn: 'sum' }] }, 5)
  check('a missing value and the text "(blank)" stay apart in the cache; the report\'s (blank) filter shows both',
    J(readPivot(bx.table, bx.cache, bx.records)?.cache.records) === J(bl.records) && /<item x="0" h="1"\/><item x="1"\/><item x="2"\/>/.test(bx.table) && !/item="/.test(bx.table), bx.cache + bx.table)
  const noGrand = x.table.replace('<pivotTableDefinition ', '<pivotTableDefinition rowGrandTotals="0" ')
  const noSub = x.table.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" defaultSubtotal="0"')
  const falsy = x.table.replace('<pivotTableDefinition ', '<pivotTableDefinition colGrandTotals="false" ')
  const across2 = x.table.replace('<pivotTableDefinition ', '<pivotTableDefinition pageOverThenDown="true" ')
  const renamed = x.table.replace('<pivotField axis="axisRow"', '<pivotField name="Item" axis="axisRow"')
  const renamedItem = x.table.replace('<item x="0"/>', '<item n="Seven" x="0"/>')
  const labelFilter = x.table.replace('<pivotTableStyleInfo', '<filters count="1"><filter fld="1" type="captionBeginsWith" evalOrder="-1" id="1" stringValue1="A"><autoFilter ref="A1"><filterColumn colId="0"><customFilters><customFilter val="A*"/></customFilters></filterColumn></autoFilter></filter></filters><pivotTableStyleInfo')
  check('label and value filters are not read (even when they hide nothing)', readPivot(labelFilter, x.cache, x.records) === null)
  check('XML booleans "true"/"false" count like "1"/"0"; renamed fields and items are not read',
    readPivot(falsy, x.cache, x.records) === null && readPivot(across2, x.cache, x.records) === null && readPivot(renamed, x.cache, x.records) === null && readPivot(renamedItem, x.cache, x.records) === null && readPivot(x.table, x.cache, x.records) !== null)
  check('reports with grand totals or subtotals turned off are not read', readPivot(noGrand, x.cache, x.records) === null && readPivot(noSub, x.cache, x.records) === null)
  check('sources in other workbooks and calculated items are not read', readPivot(x.table, linked, x.records) === null && readPivot(x.table, calcItems, x.records) === null && readPivot(calcItem, x.cache, x.records) === null)
  check('values shown as %, grouped fields and external sources are not read (the cells keep the values)', readPivot(shownAs, x.cache, x.records) === null && readPivot(x.table, grouped, x.records) === null && readPivot(x.table, external, x.records) === null)
}

// ---- writing: reports that cannot be saved as PivotTables stay values (and GETPIVOTDATA with them)
{
  const R = (r0, c0, r1, c1) => ({ startRow: r0, startColumn: c0, endRow: r1, endColumn: c1 })
  const data = [['Region', 'Sales'], ['East', 10], ['West', 5]]
  const cache = readSource((r, c) => data[r]?.[c] ?? null, R(0, 0, 2, 1))
  const book = (p, written) => {
    const L = layout(p, cache), fr = p.filters.length ? p.filters.length + 1 : 0
    const width = Math.max(...L.grid.map((r) => r.length), fr ? 2 : 0)
    p.extent = { written: true, dRow: 0, rows: L.grid.length, cols: width, filterRows: fr }
    return { sheetOrder: ['p', 'd'], resources: [{ name: 'SHEET_DEFINED_NAME_PLUGIN', data: J({ w: { id: 'w', name: 'SheetlyRef_w', formulaOrRefString: written, hidden: true }, s: { id: 's', name: 'SheetlyRef_s', formulaOrRefString: 'Data!$A$1:$B$3', hidden: true } }) }],
      sheets: { p: { id: 'p', name: 'Rep', cellData: { 0: { 9: { f: '=GETPIVOTDATA("Sales",$A$4)', v: 15, t: 2 } } }, custom: { sheetlyPivots: [p] } },
        d: { id: 'd', name: 'Data', cellData: Object.fromEntries(data.map((r, i) => [i, Object.fromEntries(r.map((v, c) => [c, { v, t: typeof v === 'number' ? 2 : 1 }]))])) } } }
  }
  const def = (filters) => ({ id: 'p1', name: 'PT', source: { sheetId: 'd', range: R(0, 0, 2, 1) }, anchor: { row: 3, col: 0 }, rows: ['Region'], cols: [], values: [{ field: 'Sales', fn: 'sum' }], filters, cache, sourceRef: 'SheetlyRef_s', writtenRef: 'SheetlyRef_w' })
  const parts = async (snap) => { const b = await snapshotToXlsx(snap), z = await JSZip.loadAsync(b), w = new ExcelJS.Workbook(); await w.xlsx.load(b); return [!!z.file('xl/pivotTables/pivotTable1.xml'), w.getWorksheet('Rep').getCell('J1').value] }
  const ok = await parts(book(def([{ field: 'Region', items: ['East'] }]), 'Rep!$A$2:$B$6'))
  const gone = await parts(book(def([{ field: 'Region', items: ['South'] }]), 'Rep!$A$2:$B$6'))
  const wide = await parts(book(def([]), 'Rep!$A$4:$C$7'))
  const clash = await parts(book({ ...def([]), values: [{ field: 'Sales', fn: 'sum', name: 'Region' }] }, 'Rep!$A$4:$B$7'))
  check('a value caption that is a field\'s name (Excel refuses it) leaves the report as values', !clash[0] && clash[1] === 15, J(clash))
  const twice = await parts(book({ ...def([]), values: [{ field: 'Sales', fn: 'sum' }, { field: 'Sales', fn: 'sum' }] }, 'Rep!$A$4:$C$7'))
  check('two value fields captioned alike leave the report as values', !twice[0] && twice[1] === 15, J(twice))
  // a copied sheet's report has the same name: the copy is saved under the next free name
  const two = book(def([]), 'Rep!$A$4:$B$7')
  two.sheetOrder = ['p', 'p2', 'd']
  two.sheets.p2 = { ...two.sheets.p, id: 'p2', name: 'Rep2', custom: { sheetlyPivots: [{ ...two.sheets.p.custom.sheetlyPivots[0], id: 'p2x', writtenRef: 'SheetlyRef_w2' }] } }
  two.resources[0].data = J({ ...JSON.parse(two.resources[0].data), w2: { id: 'w2', name: 'SheetlyRef_w2', formulaOrRefString: 'Rep2!$A$4:$B$7', hidden: true } })
  const z2 = await JSZip.loadAsync(await snapshotToXlsx(two))
  const names2 = await Promise.all(['xl/pivotTables/pivotTable1.xml', 'xl/pivotTables/pivotTable2.xml'].map(async (f) => /name="([^"]*)"/.exec((await z2.file(f)?.async('string')) ?? '')?.[1]))
  check('two reports named alike (a copied sheet) are saved under different names', J(names2) === J(['PT', 'PivotTable1']), J(names2))
  // a renamed source table whose old name another table now has: found by id
  const tb = book({ ...def([]), sourceTable: 'Old', sourceTableId: 'tA' }, 'Rep!$A$4:$B$7')
  tb.resources.push({ name: 'SHEET_TABLE_PLUGIN', data: J({ p: { tables: [{ id: 'tB', name: 'Old', range: R(10, 5, 12, 6), options: { showHeader: true } }] }, d: { tables: [{ id: 'tA', name: 'Renamed', range: R(0, 0, 2, 1), options: { showHeader: true } }] } }) })
  const zt = await JSZip.loadAsync(await snapshotToXlsx(tb))
  const srcXml = (await zt.file('xl/pivotCache/pivotCacheDefinition1.xml')?.async('string')) ?? ''
  check('a source table is found by id before name (the renamed table, not the one now called Old)', /<worksheetSource name="Renamed"\/>/.test(srcXml), srcXml.slice(0, 400))
  check('a report is saved as a PivotTable with GETPIVOTDATA as a formula; filtered on a vanished item, or widened by an inserted column, it stays values',
    ok[0] && ok[1]?.formula && !gone[0] && gone[1] === 15 && !wide[0] && wide[1] === 15, J([ok, gone, wide]))
  // a file with a report Sheetly cannot show: that one is counted, and GETPIVOTDATA is saved as its value
  const z = await JSZip.loadAsync(await snapshotToXlsx(book(def([]), 'Rep!$A$4:$B$7')))
  z.file('xl/pivotTables/pivotTable1.xml', (await z.file('xl/pivotTables/pivotTable1.xml').async('string')).replace('<dataField name="Sum of Sales" fld="1"', '<dataField name="Sum of Sales" fld="1" showDataAs="percentOfTotal"'))
  const r = await xlsxToSnapshot(await z.generateAsync({ type: 'uint8array' })), rep = r.sheets[r.sheetOrder[0]]
  const mixed = book(def([]), 'Rep!$A$4:$B$7')
  mixed.sheets.p.custom.sheetlyPivotValues = 1
  const kept = await parts(mixed)
  check('an unreadable report is counted on its sheet; with one there, GETPIVOTDATA is saved as its value', rep.custom.sheetlyPivotValues === 1 && !rep.custom.sheetlyPivots && kept[0] && kept[1] === 15, J([rep.custom, kept]))
  // numeric character references in cache text
  const cx = pivotXml({ name: 'PT', cache, rows: ['Region'], cols: [], values: [{ field: 'Sales', fn: 'sum' }], filters: [], source: { sheetName: 'Data', range: R(0, 0, 2, 1) }, location: R(3, 0, 6, 1), firstDataRow: 1, firstDataCol: 1 }, 1)
  const refs = readPivot(cx.table, cx.cache.replace('<s v="East"/>', '<s v="E&#97;s&#x74;"/>').replace('name="Sales"', 'name="Sa&#108;es"'), cx.records)
  check('numeric character references in cache text are decoded (E&#97;s&#x74; is East)', J(refs?.cache.records[0]) === J(['East', 10]) && refs?.values[0].field === 'Sales', J(refs?.cache))
}

const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`xlsx fidelity  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
