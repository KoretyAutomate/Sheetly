// .xlsx objects (Node, no app): conditional formats, data validation, tables, pictures, charts, dynamic arrays and
// the 1904 date system, written by Sheetly and read back (also through ExcelJS and the raw package XML).
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { SaxesParser } from 'saxes'
import { xlsxToSnapshot, snapshotToXlsx } from '../src/xlsx.js'
import { specFromXml, chartXml as writeChartXml } from '../src/xlsx-zip.js'
import { explicitSeries } from '../src/chart-model.js'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const J = JSON.stringify
const R = (a, b, c, d) => ({ startRow: a, endRow: c ?? a, startColumn: b, endColumn: d ?? b })
const res = (snap, name) => JSON.parse(snap.resources?.find((r) => r.name === name)?.data ?? '{}')
const wellFormed = (xml) => { try { const p = new SaxesParser(); p.write(xml).close(); return true } catch (e) { return e.message } }
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const red = { bg: { rgb: '#ffc7ce' }, cl: { rgb: '#9c0006' } }
const cfRules = [
  { cfId: 'a', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 5, style: red } },
  { cfId: 'b', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'number', operator: 'between', value: [2, 4], style: red } },
  { cfId: 'c', ranges: [R(0, 0, 9, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'containsText', value: 'ab', style: red } },
  { cfId: 'd', ranges: [R(0, 0, 9, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'beginsWith', value: 'x', style: red } },
  { cfId: 'e', ranges: [R(0, 0, 9, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'notContainsText', value: 'z', style: red } },
  { cfId: 'f', ranges: [R(0, 0, 9, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'equal', value: 'yes', style: red } },
  { cfId: 'g', ranges: [R(0, 2, 9, 2)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'timePeriod', operator: 'last7Days', style: red } },
  { cfId: 'h', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'rank', isBottom: true, isPercent: true, value: 10, style: red } },
  { cfId: 'i', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'average', operator: 'lessThanOrEqual', style: red } },
  { cfId: 'j', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'formula', value: '=B1>A1', style: red } },
  { cfId: 'k', ranges: [R(0, 0, 9, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'duplicateValues', style: red } },
  { cfId: 'l', ranges: [R(0, 0, 9, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'uniqueValues', style: red } },
  { cfId: 'm', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'dataBar', isShowValue: true, config: { min: { type: 'min' }, max: { type: 'max' }, isGradient: true, positiveColor: '#638EC6', nativeColor: '#FF0000' } } },
  { cfId: 'n', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'colorScale', config: [{ index: 0, color: '#F8696B', value: { type: 'min' } }, { index: 1, color: '#FFEB84', value: { type: 'percentile', value: 50 } }, { index: 2, color: '#63BE7B', value: { type: 'max' } }] } },
  { cfId: 'o', ranges: [R(0, 1, 9, 1)], stopIfTrue: false, rule: { type: 'iconSet', isShowValue: true, config: [
    { operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 }, iconType: '3Arrows', iconId: '0' },
    { operator: 'greaterThan', value: { type: 'percent', value: 33 }, iconType: '3Arrows', iconId: '1' },
    { operator: 'lessThanOrEqual', value: { type: 'percent', value: 33 }, iconType: '3Arrows', iconId: '2' }] } },
]
const dvRules = [
  { uid: 'v1', ranges: [R(0, 3, 9, 3)], type: 'list', formula1: '["Yes","No"]', showDropDown: true },
  { uid: 'v2', ranges: [R(0, 4)], type: 'list', formula1: '=$A$1:$A$5' },
  { uid: 'v3', ranges: [R(0, 5, 9, 5)], type: 'whole', operator: 'between', formula1: '1', formula2: '10', errorStyle: 2, error: 'One to ten', showErrorMessage: true },
  { uid: 'v4', ranges: [R(0, 6)], type: 'decimal', operator: 'greaterThan', formula1: '=B1' },
  { uid: 'v5', ranges: [R(0, 7)], type: 'date', operator: 'greaterThanOrEqual', formula1: '45306' },
  { uid: 'v6', ranges: [R(0, 8), R(5, 8)], type: 'custom', formula1: '=LEN(I1)<5', prompt: 'Short text', promptTitle: 'Tip', showInputMessage: true },
]
const cells = {}
for (let r = 0; r < 10; r++) cells[r] = { 0: { v: `t${r}`, t: 1 }, 1: { v: r + 1, t: 2 } }
cells[0][10] = { v: 'Region', t: 1 }; cells[0][11] = { v: 'Sales', t: 1 }
cells[1][10] = { v: 'North', t: 1 }; cells[1][11] = { v: 10, t: 2 }
cells[2][10] = { v: 'South', t: 1 }; cells[2][11] = { v: 20, t: 2 }
cells[3][10] = { v: 'East', t: 1 }; cells[3][11] = { v: 30, t: 2 }
cells[12] = { 0: { f: '=SEQUENCE(3)', v: 1, t: 2 } }
const chartSpec = { v: 2, type: 'column', source: { sheetId: 's1', range: R(0, 10, 3, 11) }, plotBy: null, title: null, legend: null, dataLabels: true, axisTitles: { x: 'Region', y: 'Sales' } }
const snap = {
  sheetOrder: ['s1'],
  resources: [
    { name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: J({ s1: cfRules }) },
    { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: J({ s1: dvRules }) },
    { name: 'SHEET_TABLE_PLUGIN', data: J({ s1: { tables: [{ id: 't1', name: 'Sales table', range: R(0, 10, 3, 11), options: { showHeader: true, showFooter: false }, columns: [{ displayName: 'Region' }, { displayName: 'Sales' }] }] } }) },
    { name: 'SHEET_DRAWING_PLUGIN', data: J({ s1: { data: {
      p1: { drawingId: 'p1', drawingType: 0, imageSourceType: 'BASE64', source: `data:image/png;base64,${PNG}`, sheetTransform: { from: { column: 13, columnOffset: 5, row: 1, rowOffset: 4 }, to: { column: 15, columnOffset: 0, row: 6, rowOffset: 0 } } },
      c1: { drawingId: 'c1', drawingType: 8, componentKey: 'sheetly-chart', data: chartSpec, sheetTransform: { from: { column: 13, columnOffset: 0, row: 8, rowOffset: 0 }, to: { column: 20, columnOffset: 0, row: 22, rowOffset: 0 } } },
    }, order: ['p1', 'c1'] } }) },
  ],
  sheets: { s1: { id: 's1', name: 'Data', cellData: cells } },
}
const arrayRangeOf = (id, r, c) => (id === 's1' && r === 12 && c === 0 ? R(12, 0, 14, 0) : null)
const buf = await snapshotToXlsx(snap, { arrayRangeOf })
const zip = await JSZip.loadAsync(buf)
const bad = []
for (const f of Object.keys(zip.files).filter((f) => /\.(xml|rels)$/.test(f))) { const w = wellFormed(await zip.file(f).async('string')); if (w !== true) bad.push(`${f}: ${w}`) }
check('every XML part in the package is well formed', bad.length === 0, bad.join(' | '))

// ---- what ExcelJS (standing in for Excel's reader) sees
const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf)
const ws = wb.getWorksheet('Data')
const types = (ws.conditionalFormattings ?? []).flatMap((c) => c.rules.map((x) => x.type + (x.operator ? ':' + x.operator : '')))
check('CF written: cellIs, text rules, timePeriod, top10, aboveAverage, expression, duplicate, unique, dataBar, colorScale, iconSet',
  ['cellIs:greaterThan', 'cellIs:between', 'containsText:containsText', 'containsText:beginsWith', 'containsText:notContainsText', 'cellIs:equal', 'timePeriod', 'top10', 'aboveAverage', 'expression', 'duplicateValues', 'uniqueValues', 'dataBar', 'colorScale', 'iconSet']
    .every((t) => types.includes(t)), J(types))
const sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string')
check('Excel attributes: text="ab", notContains, equalAverage, gte="0", dxf fill', /text="ab"/.test(sheetXml) && /operator="notContains"/.test(sheetXml) && /equalAverage="1"/.test(sheetXml) && /gte="0"/.test(sheetXml), sheetXml.match(/<cfRule[^>]*>/g)?.slice(0, 6).join(' '))
check('DV written with Excel forms ("Yes,No" list, range without =, warning style)', /<formula1>(&quot;|")Yes,No(&quot;|")<\/formula1>/.test(sheetXml) && /<formula1>\$A\$1:\$A\$5<\/formula1>/.test(sheetXml) && /errorStyle="warning"/.test(sheetXml) && /sqref="I1 I6"/.test(sheetXml), sheetXml.match(/<dataValidation [^>]*>/g)?.join(' '))
const tables = ws.getTables?.() ?? Object.values(ws.tables ?? {})
check('table written (name cleaned of spaces)', tables.length === 1 && J(tables[0]).includes('Sales_table'), J(tables.map((t) => t.table?.name ?? t.name)))
check('picture written', (ws.getImages?.() ?? []).length === 1 && wb.getImage(ws.getImages()[0].imageId)?.extension === 'png')
const chartFile = Object.keys(zip.files).find((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))
const chartXml = chartFile ? await zip.file(chartFile).async('string') : ''
check('chart part written: clustered column over Data!$L$2:$L$4 with names and categories', /<c:barDir val="col"\/>/.test(chartXml) && /<c:grouping val="clustered"\/>/.test(chartXml)
  && chartXml.includes('<c:f>Data!$L$2:$L$4</c:f>') && chartXml.includes('<c:f>Data!$K$2:$K$4</c:f>') && chartXml.includes('<c:f>Data!$L$1</c:f>'), chartXml.slice(0, 300))
const types2 = await zip.file('[Content_Types].xml').async('string'), drawingXml = await zip.file('xl/drawings/drawing1.xml')?.async('string')
check('chart and picture share the sheet\'s one drawing; content types listed', /graphicFrame/.test(drawingXml ?? '') && /xdr:pic/.test(drawingXml ?? '') && types2.includes('/xl/charts/chart1.xml'), (drawingXml ?? '').slice(0, 200))
check('dynamic array: t="array" ref="A13:A15" with cm="1" and xl/metadata.xml', /<c r="A13"[^>]*cm="1"[^>]*><f t="array" ref="A13:A15">_xlfn.SEQUENCE\(3\)<\/f>/.test(sheetXml) && !!zip.file('xl/metadata.xml') && types2.includes('/xl/metadata.xml'), sheetXml.match(/<c r="A13"[^>]*>.{0,80}/)?.[0])
check('Excel is asked to recalculate on open', /fullCalcOnLoad="1"/.test(await zip.file('xl/workbook.xml').async('string')))

// ---- and back into Sheetly
const back = await xlsxToSnapshot(buf), sid = back.sheetOrder[0]
const cf = res(back, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')[sid] ?? []
const key = (r) => [r.rule.type, r.rule.subType, r.rule.operator].filter(Boolean).join(':')
check('CF round trip keeps every rule', cf.length === cfRules.length && cfRules.every((x) => cf.some((y) => key(y) === key(x))), J(cf.map(key)))
const find = (sub, op) => cf.find((y) => y.rule.subType === sub && (!op || y.rule.operator === op))?.rule
check('CF values: >5, between [2,4], text "ab", rank bottom 10%, average <=, formula =B1>A1', find('number', 'greaterThan')?.value === 5 && J(find('number', 'between')?.value) === '[2,4]'
  && find('text', 'containsText')?.value === 'ab' && find('rank')?.isBottom === true && find('rank')?.isPercent === true && find('average')?.operator === 'lessThanOrEqual' && find('formula')?.value === '=B1>A1', J(cf.map((x) => x.rule)))
check('CF style: Light Red Fill with Dark Red Text', find('number', 'greaterThan')?.style?.bg?.rgb === '#ffc7ce' && find('number', 'greaterThan')?.style?.cl?.rgb === '#9c0006', J(find('number', 'greaterThan')?.style))
const icon = cf.find((y) => y.rule.type === 'iconSet')?.rule
check('icon set: bands, operators and icons back in Univer order', J(icon?.config?.map((c) => [c.operator, c.value.value, c.iconId])) === J([['greaterThanOrEqual', 67, '0'], ['greaterThan', 33, '1'], ['lessThanOrEqual', 33, '2']]), J(icon?.config))
const cs = cf.find((y) => y.rule.type === 'colorScale')?.rule
check('color scale: 3 stops with percentile 50', cs?.config?.length === 3 && cs.config[1].value.type === 'percentile' && cs.config[1].value.value === 50 && cs.config[2].color === '#63be7b', J(cs))
const dv = res(back, 'SHEET_DATA_VALIDATION_PLUGIN')[sid] ?? []
const dvBy = (c) => dv.find((x) => x.ranges.some((g) => g.startColumn === c))
check('DV round trip: list literal / list range / whole between / decimal ref / date / custom on two ranges',
  dvBy(3)?.formula1 === 'Yes,No' && dvBy(4)?.formula1 === '=$A$1:$A$5' && dvBy(5)?.type === 'whole' && dvBy(5)?.formula2 === '10' && dvBy(5)?.errorStyle === 2
  && dvBy(6)?.formula1 === '=B1' && dvBy(7)?.formula1 === '45306' && dvBy(8)?.ranges.length === 2 && dvBy(8)?.prompt === 'Short text', J(dv))
check('table read back for creation in the app', back.sheetlyTables?.length === 1 && back.sheetlyTables[0].showHeader === true && J(back.sheetlyTables[0].range) === J(R(0, 10, 3, 11)), J(back.sheetlyTables))
const dr = res(back, 'SHEET_DRAWING_PLUGIN')[sid]
const drawings = Object.values(dr?.data ?? {})
const pic = drawings.find((d) => d.drawingType === 0), chart = drawings.find((d) => d.drawingType === 8)
check('picture read back as a Univer drawing at its cell anchor', pic?.source?.startsWith('data:image/png;base64,') && pic.sheetTransform.from.column === 13 && pic.sheetTransform.from.columnOffset === 5 && pic.transform.width > 0, J(pic && { ...pic, source: pic.source.slice(0, 30) }))
check('chart read back as a linked Sheetly chart (type, range, data labels, axis titles)', chart?.componentKey === 'sheetly-chart' && chart.data.type === 'column' && J(chart.data.source.range) === J(R(0, 10, 3, 11))
  && chart.data.source.sheetId === sid && chart.data.dataLabels === true && chart.data.axisTitles?.y === 'Sales', J(chart?.data))
check('spilled cells are not kept as constants', back.sheets[sid].cellData[12][0].f === '=SEQUENCE(3)' && back.sheets[sid].cellData[13]?.[0]?.v === undefined, J([back.sheets[sid].cellData[12], back.sheets[sid].cellData[13]]))

// ---- chart types and titles
for (const [type, test] of [['bar', /<c:barDir val="bar"\/>/], ['columnStacked100', /percentStacked/], ['line', /<c:lineChart>[\s\S]*<c:symbol val="none"\/>/], ['lineMarkers', /<c:symbol val="circle"\/>/],
  ['pie', /<c:pieChart>/], ['doughnut', /<c:holeSize val="75"\/>/], ['area', /<c:areaChart>/], ['scatter', /<c:scatterChart>/]]) {
  const s2 = structuredClone(snap)
  const d = JSON.parse(s2.resources[3].data); d.s1.data.c1.data = { ...chartSpec, type, title: type === 'pie' ? '' : 'My chart' }; s2.resources[3].data = J(d)
  const b2 = await snapshotToXlsx(s2), z2 = await JSZip.loadAsync(b2), x2 = await z2.file(Object.keys(z2.files).find((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))).async('string')
  const read = Object.values(res(await xlsxToSnapshot(b2), 'SHEET_DRAWING_PLUGIN'))[0]
  const spec = Object.values(read.data).find((x) => x.drawingType === 8)?.data
  check(`chart type ${type} written and read back${type === 'pie' ? ' (no title)' : ' (title text)'}`, test.test(x2) && wellFormed(x2) === true && spec?.type === type && spec.title === (type === 'pie' ? '' : 'My chart'), `${spec?.type} ${J(spec?.title)}`)
}

// ---- sparklines, next to an x14 conditional format (one extLst per sheet)
{
  const s3 = structuredClone(snap)
  s3.sheets.s1.custom = { sheetlySparklines: [
    { id: 'g1', type: 'line', colors: {}, items: [{ row: 1, col: 12, sheetId: 's1', source: R(1, 1, 1, 5) }, { row: 2, col: 12, sheetId: 's1', source: R(2, 1, 2, 5) }] },
    { id: 'g2', type: 'winloss', colors: { color: '#00B050' }, items: [{ row: 3, col: 12, sheetId: 's1', source: R(3, 1, 3, 5) }] }] }
  const b3 = await snapshotToXlsx(s3), z3 = await JSZip.loadAsync(b3), x3 = await z3.file('xl/worksheets/sheet1.xml').async('string')
  const sheetExt = /<extLst>((?:(?!<extLst>)[\s\S])*)<\/extLst>\s*<\/worksheet>\s*$/.exec(x3)?.[1] ?? ''
  check('sparklines written as x14:sparklineGroups in the sheet\'s extLst, next to the x14 conditional formats', sheetExt.includes('{05C60535-1F16-4fd2-B633-F4F36F0B64E0}') && sheetExt.includes('{78C0D931-6437-407d-A8EE-F0AAD7539E65}')
    && !/<cfRule[^>]*>(?:(?!<\/cfRule>)[\s\S])*sparklineGroup/.test(x3) && /<x14:sparklineGroup displayEmptyCellsAs="gap">/.test(x3)
    && /<x14:sparklineGroup type="stacked"/.test(x3) && x3.includes('<xm:f>Data!B2:F2</xm:f><xm:sqref>M2</xm:sqref>') && wellFormed(x3) === true, (x3.match(/<extLst>[\s\S]{0,400}/) ?? [''])[0])
  const r3 = await xlsxToSnapshot(b3), sp = r3.sheets[r3.sheetOrder[0]].custom?.sheetlySparklines
  check('sparklines read back (types, cells, ranges, colour)', sp?.length === 2 && sp[0].type === 'line' && sp[0].items.length === 2 && J(sp[0].items[1].source) === J(R(2, 1, 2, 5)) && sp[0].items[1].col === 12
    && sp[1].type === 'winloss' && sp[1].colors.color === '#00B050', J(sp))
}

// ---- review findings (Codex, 2026-10-17)
{
  // Unicode table name, Stop If True, between with references
  const s4 = structuredClone(snap)
  const res4 = Object.fromEntries(s4.resources.map((r) => [r.name, JSON.parse(r.data)]))
  res4.SHEET_TABLE_PLUGIN.s1.tables[0].name = '売上'
  res4.SHEET_CONDITIONAL_FORMATTING_PLUGIN.s1 = [{ cfId: 'z', ranges: [R(0, 1, 9, 1)], stopIfTrue: true, rule: { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 5, style: red } }]
  s4.resources = Object.entries(res4).map(([name, d]) => ({ name, data: J(d) }))
  const b4 = await snapshotToXlsx(s4), w4 = new ExcelJS.Workbook(); await w4.xlsx.load(b4)
  const t4 = w4.getWorksheet('Data').getTables()
  check('a Japanese table name is kept (売上)', J(t4.map((t) => t.table?.name ?? t.name)).includes('売上'), J(t4.map((t) => t.table?.name ?? t.name)))
  const z4 = await JSZip.loadAsync(b4), x4 = await z4.file('xl/worksheets/sheet1.xml').async('string')
  const back4 = await xlsxToSnapshot(b4)
  check('Stop If True written and read back', /stopIfTrue="1"/.test(x4) && res(back4, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')[back4.sheetOrder[0]]?.[0]?.stopIfTrue === true, (x4.match(/<cfRule[^>]*>/) ?? [''])[0])
  const w5 = new ExcelJS.Workbook(), ws5 = w5.addWorksheet('S')
  ws5.addConditionalFormatting({ ref: 'A1:A5', rules: [{ type: 'cellIs', operator: 'between', formulae: ['$B$1', '$C$1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] })
  // a one-cell-anchored picture keeps its size
  const img = w5.addImage({ base64: PNG, extension: 'png' }); ws5.addImage(img, { tl: { col: 3, row: 1 }, ext: { width: 480, height: 288 }, editAs: 'oneCell' })
  const s5 = await xlsxToSnapshot(await w5.xlsx.writeBuffer())
  const cf5 = res(s5, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')[s5.sheetOrder[0]]?.[0]?.rule
  check('between with cell references becomes an equivalent formula rule (bounds kept)', cf5?.subType === 'formula' && cf5.value === '=AND(A1>=MIN($B$1,$C$1),A1<=MAX($B$1,$C$1))', J(cf5))
  const pic5 = Object.values(res(s5, 'SHEET_DRAWING_PLUGIN')[s5.sheetOrder[0]]?.data ?? {})[0]
  check('a one-cell-anchored picture keeps its size (480 x 288)', Math.round(pic5?.transform?.width) === 480 && Math.round(pic5?.transform?.height) === 288, J(pic5?.transform))
  // a chart whose series are not side by side (categories in K, series in L and N)
  const s6 = structuredClone(snap)
  for (let r = 0; r < 4; r++) s6.sheets.s1.cellData[r][13] = r ? { v: r * 100, t: 2 } : { v: 'Cost', t: 1 }
  const res6 = Object.fromEntries(s6.resources.map((r) => [r.name, JSON.parse(r.data)]))
  res6.SHEET_DRAWING_PLUGIN.s1.data.c1.data = { ...chartSpec, source: { sheetId: 's1', range: R(0, 10, 3, 13), categories: R(1, 10, 3, 10),
    series: [{ name: [0, 11], values: R(1, 11, 3, 11) }, { name: [0, 13], values: R(1, 13, 3, 13) }] } }
  s6.resources = Object.entries(res6).map(([name, d]) => ({ name, data: J(d) }))
  const b6 = await snapshotToXlsx(s6), z6 = await JSZip.loadAsync(b6)
  const c6 = await z6.file(Object.keys(z6.files).find((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))).async('string')
  check('non-adjacent series are written with their own references (no column M)', c6.includes('Data!$L$2:$L$4') && c6.includes('Data!$N$2:$N$4') && !c6.includes('$M$'), (c6.match(/<c:f>[^<]*<\/c:f>/g) ?? []).join(' '))
  const back6 = await xlsxToSnapshot(b6)
  const spec6 = Object.values(res(back6, 'SHEET_DRAWING_PLUGIN')[back6.sheetOrder[0]].data).find((d) => d.drawingType === 8)?.data
  check('and read back as two series (Sales, Cost), not three', spec6?.source?.series?.length === 2 && J(spec6.source.series[1].values) === J(R(1, 13, 3, 13)) && J(spec6.source.categories) === J(R(1, 10, 3, 10)), J(spec6?.source))
}

// ---- numeric categories (years) are not mistaken for a series when a chart is read back
{
  const s7 = structuredClone(snap)
  for (let r = 0; r < 4; r++) s7.sheets.s1.cellData[r][10] = r ? { v: 2019 + r, t: 2 } : { v: 'Year', t: 1 }
  const res7 = Object.fromEntries(s7.resources.map((r) => [r.name, JSON.parse(r.data)]))
  res7.SHEET_DRAWING_PLUGIN.s1.data.c1.data = { ...chartSpec, source: { sheetId: 's1', range: R(0, 10, 3, 11), categories: R(1, 10, 3, 10), series: [{ name: [0, 11], values: R(1, 11, 3, 11) }] } }
  s7.resources = Object.entries(res7).map(([name, d]) => ({ name, data: J(d) }))
  const back7 = await xlsxToSnapshot(await snapshotToXlsx(s7))
  const spec7 = Object.values(res(back7, 'SHEET_DRAWING_PLUGIN')[back7.sheetOrder[0]].data).find((d) => d.drawingType === 8)?.data
  check('years as categories stay categories (one series, categories K2:K4)', spec7?.source?.series?.length === 1 && J(spec7.source.categories) === J(R(1, 10, 3, 10)), J(spec7?.source))
  const names7 = Object.values(res(back7, 'SHEET_DEFINED_NAME_PLUGIN'))
  check('read back with hidden reference names (not written to the file as names)', names7.some((n) => n.name.startsWith('SheetlyRef_') && n.hidden) && !J((await (async () => { const w = new ExcelJS.Workbook(); await w.xlsx.load(await snapshotToXlsx(back7)); return w.rawDefinedNames })())).includes('SheetlyRef_'))
}

// ---- review round 3: totals rows, literal series names, multi-level and hidden explicit categories
{
  const s8 = structuredClone(snap)
  s8.sheets.s1.cellData[4] = { ...(s8.sheets.s1.cellData[4] ?? {}), 10: { v: 'My Total', t: 1 }, 11: { f: '=SUBTOTAL(109,L2:L4)', v: 60, t: 2 } }
  const res8 = Object.fromEntries(s8.resources.map((r) => [r.name, JSON.parse(r.data)]))
  res8.SHEET_TABLE_PLUGIN.s1.tables[0] = { ...res8.SHEET_TABLE_PLUGIN.s1.tables[0], range: R(0, 10, 4, 11), options: { showHeader: true, showFooter: true } }
  s8.resources = Object.entries(res8).map(([name, d]) => ({ name, data: J(d) }))
  const w8 = new ExcelJS.Workbook(); await w8.xlsx.load(await snapshotToXlsx(s8))
  const ws8 = w8.getWorksheet('Data')
  check('a table\'s totals row keeps its label and SUBTOTAL formula', ws8.getCell('K5').value === 'My Total' && ws8.getCell('L5').value?.formula === 'SUBTOTAL(109,L2:L4)' && ws8.getCell('L5').value?.result === 60, J([ws8.getCell('K5').value, ws8.getCell('L5').value]))
  const lit = specFromXml('<c:chartSpace><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:tx><c:v>Revenue</c:v></c:tx><c:cat><c:strRef><c:f>Data!$K$2:$K$4</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$L$2:$L$4</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>')
  check('a typed series name ("Revenue") is read as the name, not as a reference', lit?.source?.series?.[0]?.name === null && lit.source.series[0].nameText === 'Revenue' && J(lit.source.categories) === J(R(1, 10, 3, 10)), J(lit?.source))
  const grid = [['East', 'A', 1], ['East', 'B', 2], ['West', 'C', 3]]
  const ex = explicitSeries((r, c) => grid[r]?.[c] ?? null, { categories: R(0, 0, 2, 1), series: [{ name: null, nameText: 'S', values: R(0, 2, 2, 2) }] })
  check('two-column categories stay two levels (inner A,B,C; outer East, , West)', J(ex.labels) === J(['A', 'B', 'C']) && J(ex.levels) === J([['East', '', 'West']]) && ex.series[0].values.length === 3, J(ex))
  const hid = explicitSeries((r, c) => grid[r]?.[c] ?? null, { categories: R(0, 1, 2, 1), series: [{ name: null, nameText: 'S', values: R(0, 2, 2, 2) }] }, { hiddenRow: (r) => r === 1 })
  const two = explicitSeries((r, c) => [[1, 10, 100], [2, 20, 200]][r]?.[c] ?? null, { categories: R(0, 0, 1, 0), series: [{ name: null, nameText: 'B', values: R(0, 1, 1, 1) }, { name: null, nameText: 'C', values: R(0, 2, 1, 2) }] }, { hiddenCol: (c) => c === 2 })
  const twoB = explicitSeries((r, c) => [[1, 10, 100], [2, 20, 200]][r]?.[c] ?? null, { categories: R(0, 0, 1, 0), series: [{ name: null, nameText: 'B', values: R(0, 1, 1, 1) }, { name: null, nameText: 'C', values: R(0, 2, 1, 2) }] }, { hiddenCol: (c) => c === 1 })
  check('hiding a series column drops only that series (C hidden -> B stays; B hidden -> C keeps all points)', J(two.series.map((x) => [x.name, x.values])) === J([['B', [10, 20]]]) && J(twoB.series.map((x) => [x.name, x.values])) === J([['C', [100, 200]]]), J([two.series, twoB.series]))
  check('a hidden row is left out of explicit series and their categories', J(hid.labels) === J(['A', 'C']) && J(hid.series[0].values) === J([1, 3]), J(hid))
}

// ---- a table with its header row hidden keeps its column names
{
  const w = new ExcelJS.Workbook(), s9 = w.addWorksheet('T')
  s9.addTable({ name: 'Sales', ref: 'A1', headerRow: false, columns: [{ name: 'Amount' }, { name: 'Qty' }], rows: [[10, 1], [20, 2]] })
  const r9 = await xlsxToSnapshot(await w.xlsx.writeBuffer()), sh = r9.sheets[r9.sheetOrder[0]]
  check('a headerless table is not handed to Univer (which would name columns from the data)', !(r9.sheetlyTables ?? []).length && sh.custom?.sheetlyTablesNoHeader?.[0]?.columns?.join() === 'Amount,Qty', J([r9.sheetlyTables, sh.custom]))
  const w2 = new ExcelJS.Workbook(); await w2.xlsx.load(await snapshotToXlsx(r9))
  const t9 = w2.getWorksheet('T').getTables()[0]
  const m9 = t9?.table ?? t9
  check('and it is written back with its names and no header row', m9?.headerRow === false && J((m9.columns ?? []).map((c) => c.name)) === J(['Amount', 'Qty']), J(m9 && { headerRow: m9.headerRow, columns: m9.columns }))
}

// ---- scatter series keep their own X ranges
{
  const xml = '<c:chartSpace><c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="lineMarker"/>'
    + '<c:ser><c:idx val="0"/><c:tx><c:v>P</c:v></c:tx><c:xVal><c:numRef><c:f>Data!$A$2:$A$4</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Data!$B$2:$B$4</c:f></c:numRef></c:yVal></c:ser>'
    + '<c:ser><c:idx val="1"/><c:tx><c:v>Q</c:v></c:tx><c:xVal><c:numRef><c:f>Data!$C$2:$C$4</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Data!$D$2:$D$4</c:f></c:numRef></c:yVal></c:ser>'
    + '</c:scatterChart></c:plotArea></c:chart></c:chartSpace>'
  const sp = specFromXml(xml)
  check('scatter: each series keeps its X range on read', J(sp?.source?.series?.map((x) => x.xValues)) === J([R(1, 0, 3, 0), R(1, 2, 3, 2)]), J(sp?.source?.series))
  const s10 = { sheetOrder: ['s1'], resources: [{ name: 'SHEET_DRAWING_PLUGIN', data: J({ s1: { data: { c: { drawingId: 'c', drawingType: 8, componentKey: 'sheetly-chart',
    data: { v: 2, type: 'scatter', title: null, source: { sheetId: 's1', ...sp.source, sheetName: undefined } }, sheetTransform: { from: { column: 6, columnOffset: 0, row: 0, rowOffset: 0 }, to: { column: 12, columnOffset: 0, row: 12, rowOffset: 0 } } } }, order: ['c'] } }) }],
    sheets: { s1: { id: 's1', name: 'Data', cellData: { 1: { 0: { v: 1, t: 2 }, 1: { v: 2, t: 2 }, 2: { v: 10, t: 2 }, 3: { v: 20, t: 2 } }, 2: { 0: { v: 2, t: 2 }, 1: { v: 3, t: 2 }, 2: { v: 11, t: 2 }, 3: { v: 21, t: 2 } }, 3: { 0: { v: 3, t: 2 }, 1: { v: 4, t: 2 }, 2: { v: 12, t: 2 }, 3: { v: 22, t: 2 } } } } } }
  const z10 = await JSZip.loadAsync(await snapshotToXlsx(s10)), x10 = await z10.file(Object.keys(z10.files).find((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))).async('string')
  check('scatter: each series is written with its own X range', (x10.match(/<c:xVal><c:numRef><c:f>[^<]*<\/c:f>/g) ?? []).map((t) => t.replace(/.*<c:f>/, '').replace('</c:f>', '')).join(',') === 'Data!$A$2:$A$4,Data!$C$2:$C$4', (x10.match(/<c:xVal>[\s\S]{0,60}/g) ?? []).join(' '))
}

// ---- review round 6: legacy array formulas, validation flag defaults, chart prefixes, theme colours in CF
{
  const w = new ExcelJS.Workbook(), t = w.addWorksheet('L')
  t.getCell('A1').value = 1; t.getCell('A2').value = 2; t.getCell('A3').value = 3
  t.getCell('C1').value = { formula: 'SUM(A1:A3)', result: 6, shareType: 'array', ref: 'C1:C3' }; t.getCell('C2').value = 6; t.getCell('C3').value = 6
  t.getCell('E1').value = { formula: 'A1:A3*2', result: 2, shareType: 'array', ref: 'E1:E3' }; t.getCell('E2').value = 4; t.getCell('E3').value = 6
  t.dataValidations.model['G1'] = { type: 'whole', operator: 'between', formulae: [1, 9] }
  t.addConditionalFormatting({ ref: 'A1:A3', rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { theme: 4 } } } }] })
  const r = await xlsxToSnapshot(await w.xlsx.writeBuffer()), c = r.sheets[r.sheetOrder[0]].cellData
  const legacy = (r.sheetlyLegacyArrays ?? []).map((x) => `${x.formula}@${x.row},${x.col}:${x.range.endRow}`)
  check('legacy array formulas are listed for the app to settle by calculating (scalar vs array)', legacy.join('|') === '=SUM(A1:A3)@0,2:2|=A1:A3*2@0,4:2' && c[1][2]?.v === undefined && c[1][4]?.v === undefined, J([r.sheetlyLegacyArrays, c[1]]))
  const v = res(r, 'SHEET_DATA_VALIDATION_PLUGIN')[r.sheetOrder[0]]?.[0]
  check('a validation without allowBlank / showErrorMessage in the file has both off', v?.allowBlank === false && v?.showErrorMessage === false, J(v))
  const cf = res(r, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')[r.sheetOrder[0]]?.[0]?.rule
  check('a conditional format filled with a theme colour keeps the colour', /^#[0-9a-f]{6}$/.test(cf?.style?.bg?.rgb ?? ''), J(cf?.style))
  // the same chart with another namespace prefix
  const zc = await JSZip.loadAsync(await snapshotToXlsx(snap))
  const cp = Object.keys(zc.files).find((f) => /^xl\/charts\/chart\d+\.xml$/.test(f))
  zc.file(cp, (await zc.file(cp).async('string')).replace(/xmlns:c=/, 'xmlns:ch=').replace(/<(\/?)c:/g, '<$1ch:'))
  const back = await xlsxToSnapshot(await zc.generateAsync({ type: 'uint8array' }))
  const ch = Object.values(res(back, 'SHEET_DRAWING_PLUGIN')[back.sheetOrder[0]]?.data ?? {}).filter((d) => d.drawingType === 8)
  check('a chart written with another namespace prefix (ch:) is still read', ch.length === 1 && ch[0].data.type === 'column', String(ch.length))
}

// ---- review round 7: scatter axis titles, a linked table header
{
  const spec = { v: 2, type: 'scatter', title: null, axisTitles: { x: 'Time', y: 'Distance' }, source: { sheetId: 's', range: R(0, 0, 3, 1) } }
  const grid = [['t', 'd'], [1, 5], [2, 6], [3, 7]]
  const back = specFromXml(writeChartXml(spec, 'Data', (r, c) => grid[r]?.[c] ?? null))
  check('scatter: both axis titles survive (x Time, y Distance)', back?.axisTitles?.x === 'Time' && back.axisTitles.y === 'Distance', J(back?.axisTitles))
  const link = { p: { id: 'd', body: { dataStream: 'Site\r\n', paragraphs: [{ startIndex: 4 }], customRanges: [{ startIndex: 0, endIndex: 3, rangeType: 0, rangeId: 'r', properties: { url: 'https://example.com/' } }] } } }
  const s11 = { sheetOrder: ['s1'], resources: [{ name: 'SHEET_TABLE_PLUGIN', data: J({ s1: { tables: [{ id: 't', name: 'Links', range: R(0, 0, 2, 1), options: { showHeader: true } }] } }) }],
    sheets: { s1: { id: 's1', name: 'Data', cellData: { 0: { 0: link, 1: { v: 'Qty', t: 1 } }, 1: { 0: { v: 'a', t: 1 }, 1: { v: 1, t: 2 } }, 2: { 0: { v: 'b', t: 1 }, 1: { v: 2, t: 2 } } } } } }
  const w11 = new ExcelJS.Workbook(); await w11.xlsx.load(await snapshotToXlsx(s11))
  const ws11 = w11.getWorksheet('Data'), tb = ws11.getTables()[0], m = tb?.table ?? tb
  check('a table header that is a link keeps its text as the column name and stays a link', J((m?.columns ?? []).map((c) => c.name)) === J(['Site', 'Qty']) && ws11.getCell('A1').value?.hyperlink === 'https://example.com/', J([m?.columns?.map((c) => c.name), ws11.getCell('A1').value]))
}

// ---- review round 8: imports run side by side keep their own date systems; quotes in text rules
{
  const mk = async (d1904) => { const w = new ExcelJS.Workbook(); w.properties.date1904 = d1904; const t = w.addWorksheet('D'); t.getCell('A1').value = new Date(Date.UTC(2024, 0, 1)); t.getCell('A1').numFmt = 'yyyy-mm-dd'; return w.xlsx.writeBuffer() }
  const [a1904, a1900] = await Promise.all([mk(true), mk(false)])
  const [s1904, s1900] = await Promise.all([xlsxToSnapshot(a1904), xlsxToSnapshot(a1900)])
  const v = (x) => x.sheets[x.sheetOrder[0]].cellData[0][0].v
  check('two files read at the same time keep their own date systems (43830 / 45292)', v(s1904) === 43830 && v(s1900) === 45292, `${v(s1904)} ${v(s1900)}`)
  const s12 = { sheetOrder: ['s1'], resources: [{ name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: J({ s1: [{ cfId: 'q', ranges: [R(0, 0, 4, 0)], stopIfTrue: false, rule: { type: 'highlightCell', subType: 'text', operator: 'beginsWith', value: 'a"b', style: red } }] }) }],
    sheets: { s1: { id: 's1', name: 'Data', cellData: {} } } }
  const x12 = await (await JSZip.loadAsync(await snapshotToXlsx(s12))).file('xl/worksheets/sheet1.xml').async('string')
  const f12 = /<cfRule[^>]*>\s*<formula>([^<]*)<\/formula>/.exec(x12)?.[1]?.replace(/&quot;/g, '"')
  check('a quote inside a text rule is doubled in its formula', f12 === 'LEFT(A1,LEN("a""b"))="a""b"', f12)
}

// ---- review round 9: mixed series directions, a switched-off input message
{
  const g = [[10, null, null, null, null, null], [20, null, null, 40, 50, 60], [30]]
  const mixed = explicitSeries((r, c) => g[r]?.[c] ?? null, { categories: null, series: [{ name: null, nameText: 'A', values: R(0, 0, 2, 0) }, { name: null, nameText: 'B', values: R(1, 3, 1, 5) }] })
  const g2 = []; g2[0] = [1, null]; g2[1] = [2, null]; g2[2] = [3, null]; g2[4] = [null, 10]; g2[5] = [null, 20]; g2[6] = [null, 30]
  const vis = explicitSeries((r, c) => g2[r]?.[c] ?? null, { categories: null, series: [{ name: null, nameText: 'A', values: R(0, 0, 2, 0) }, { name: null, nameText: 'B', values: R(4, 1, 6, 1) }] }, { hiddenRow: (r) => r === 5 })
  check('each series hides the points in its own hidden rows (row 6 hides B6 only)', J(vis.series.map((x) => x.values)) === J([[1, 2, 3], [10, 30]]), J(vis.series))
  check('a column series and a row series are each read along their own direction', J(mixed.series.map((x) => x.values)) === J([[10, 20, 30], [40, 50, 60]]), J(mixed.series))
  const s13 = { sheetOrder: ['s1'], resources: [{ name: 'SHEET_DATA_VALIDATION_PLUGIN', data: J({ s1: [{ uid: 'p', ranges: [R(0, 0)], type: 'whole', operator: 'between', formula1: '1', formula2: '5', showInputMessage: false, prompt: 'Kept but hidden' }] }) }],
    sheets: { s1: { id: 's1', name: 'Data', cellData: {} } } }
  const b13 = await xlsxToSnapshot(await snapshotToXlsx(s13)), v13 = res(b13, 'SHEET_DATA_VALIDATION_PLUGIN')[b13.sheetOrder[0]]?.[0]
  check('an input message that is switched off stays off (its text kept)', !v13?.showInputMessage && v13?.prompt === 'Kept but hidden', J(v13))
}

// ---- 1904 date system
{
  const w = new ExcelJS.Workbook(); w.properties.date1904 = true
  const s = w.addWorksheet('D'); s.getCell('A1').value = new Date(Date.UTC(2024, 0, 15)); s.getCell('A1').numFmt = 'yyyy-mm-dd'
  const b = await xlsxToSnapshot(await w.xlsx.writeBuffer())
  check('1904 workbook: serials kept in the 1904 system and Univer told so', b.dateSystem === 'date1904' && b.sheets[b.sheetOrder[0]].cellData[0][0].v === 45306 - 1462, J([b.dateSystem, b.sheets[b.sheetOrder[0]].cellData[0][0]]))
  const again = new ExcelJS.Workbook(); await again.xlsx.load(await snapshotToXlsx(b))
  check('and written back as a 1904 workbook', again.properties.date1904 === true)
}

// ---- a whole-column validation stays one rule (ExcelJS used to expand it to a million cells)
{
  const w = new ExcelJS.Workbook(), s = w.addWorksheet('V')
  s.dataValidations.model['A1:A1048576'] = { type: 'list', formulae: ['"a,b"'], allowBlank: true }
  const b = await w.xlsx.writeBuffer()
  const t0 = Date.now(), snap2 = await xlsxToSnapshot(b), ms = Date.now() - t0
  const v = res(snap2, 'SHEET_DATA_VALIDATION_PLUGIN')[snap2.sheetOrder[0]] ?? []
  check('whole-column validation loads as one rule, quickly', v.length === 1 && v[0].ranges[0].endRow === 1048575 && ms < 3000, `${v.length} rules, ${ms} ms`)
}

const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`xlsx objects  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
