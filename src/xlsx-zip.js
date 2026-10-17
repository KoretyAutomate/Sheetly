// Parts of an .xlsx that ExcelJS does not write or read, handled on the zip package itself:
//  - charts: DrawingML chart parts (xl/charts/chartN.xml) anchored in the sheet's drawing, written from Sheetly's
//    chart specs and read back into them;
//  - dynamic arrays: formulas that spill carry cm="1" pointing at xl/metadata.xml (XLDAPR), as Excel 365 writes
//    them; without it Excel treats them as legacy array formulas.
import JSZip from 'jszip'
import { chartTitle, explicitSeries, seriesFrom } from './chart-model.js'
import { colName, parseRange } from './xlsx-objects.js'

const NS = {
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart', a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
}
const REL = {
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  metadata: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata',
}
const CT = { chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  metadata: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml' }
const EMU = 9525
const ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']
const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const unesc = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')
const sheetRef = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`)
const abs = (r, c) => `$${colName(c)}$${r + 1}`
const color = (i) => ACCENTS[i % 6]

// ---- chart XML
const rich = (text) => `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>`
const strRef = (f, values) => `<c:strRef><c:f>${esc(f)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache></c:strRef>`
const numRef = (f, values) => `<c:numRef><c:f>${esc(f)}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => (v == null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache></c:numRef>`
const dLbls = (on, pie) => (on ? `<c:dLbls>${pie ? '<c:dLblPos val="bestFit"/>' : ''}<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>` : '')

/** DrawingML for a Sheetly chart spec over a sheet (get(r, c) reads cell values). */
export function chartXml(spec, sheetName, get) {
  const g = spec.source.range, ex = !!spec.source.series, t = spec.type ?? 'column'
  const d = ex ? explicitSeries(get, spec.source) : seriesFrom(get, g, spec.plotBy)
  const sn = sheetRef(sheetName)
  const span = (x) => `${sn}!${abs(x.startRow, x.startColumn)}:${abs(x.endRow, x.endColumn)}`
  // References of each series: name, categories, values (a chart picked series by series keeps its own refs).
  const lineRef = (s, i) => (ex ? span(spec.source.series[i].values) : d.byCols ? `${sn}!${abs(d.r0, s.line)}:${abs(g.endRow, s.line)}` : `${sn}!${abs(s.line, d.c0)}:${abs(s.line, g.endColumn)}`)
  const catRef = ex ? (spec.source.categories ? span(spec.source.categories) : null) : d.byCols ? (d.labelCol ? `${sn}!${abs(d.r0, g.startColumn)}:${abs(g.endRow, d.c0 - 1)}` : null)
    : (d.headerRow ? `${sn}!${abs(g.startRow, d.c0)}:${abs(g.startRow, g.endColumn)}` : null)
  const nameRef = (s) => (s.nameCell ? `${sn}!${abs(s.nameCell[0], s.nameCell[1])}` : null)
  // A series name from a cell, or a typed one (<c:v>).
  const tx = (s) => (nameRef(s) ? `<c:tx>${strRef(nameRef(s), [s.name])}</c:tx>` : ex && s.name ? `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>` : '')
  // Several label columns: a multi-level category axis (innermost level first, outer labels once per group).
  const multi = () => `<c:multiLvlStrRef><c:f>${esc(catRef)}</c:f><c:multiLvlStrCache><c:ptCount val="${d.labels.length}"/>`
    + [d.labels, ...(d.levels ?? [])].map((lv) => `<c:lvl>${lv.map((x, i) => (x === '' ? '' : `<c:pt idx="${i}"><c:v>${esc(x)}</c:v></c:pt>`)).join('')}</c:lvl>`).join('') + '</c:multiLvlStrCache></c:multiLvlStrRef>'
  const cat = () => (catRef ? `<c:cat>${d.levels?.length ? multi() : strRef(catRef, d.labels)}</c:cat>` : '')
  const pie = t === 'pie' || t === 'doughnut', labels = !!spec.dataLabels
  let plot
  if (pie) {
    const s = d.series[0]
    const ser = s ? `<c:ser><c:idx val="0"/><c:order val="0"/>${tx(s)}${d.labels.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${color(i)}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('')}${dLbls(labels, true)}${cat()}<c:val>${numRef(lineRef(s, 0), s.values)}</c:val></c:ser>` : ''
    plot = t === 'pie' ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>`
  } else if (t === 'scatter') {
    const xs = ex ? null : d.series.length > 1 ? d.series[0] : null, ys = xs ? d.series.slice(1) : d.series
    // series picked one by one: each its own X values (else the shared categories)
    const xRef = (k) => { const own = ex ? spec.source.series[k]?.xValues : null; return own ? `<c:xVal>${numRef(span(own), d.series[k]?.xs ?? [])}</c:xVal>` : ex && catRef ? `<c:xVal>${numRef(catRef, d.labels.map(Number))}</c:xVal>` : '' }
    const sers = ys.map((s, i) => { const k = xs ? i + 1 : i; return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color(i)}"/></a:solidFill></c:spPr></c:marker>${dLbls(labels)}`
      + `${xs ? `<c:xVal>${numRef(lineRef(xs, 0), xs.values)}</c:xVal>` : xRef(k)}<c:yVal>${numRef(lineRef(s, k), s.values)}</c:yVal><c:smooth val="0"/></c:ser>` }).join('')
    plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}<c:axId val="500000001"/><c:axId val="500000002"/></c:scatterChart>`
  } else {
    const kind = t.startsWith('line') ? 'line' : t.startsWith('area') ? 'area' : 'bar'
    const grouping = /100$/.test(t) ? 'percentStacked' : /Stacked/.test(t) ? 'stacked' : kind === 'bar' ? 'clustered' : 'standard'
    const sers = d.series.map((s, i) => {
      const fill = kind === 'line' ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color(i)}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="${t === 'lineMarkers' ? 'circle' : 'none'}"/>${t === 'lineMarkers' ? '<c:size val="5"/>' : ''}</c:marker>`
        : `<c:spPr><a:solidFill><a:srgbClr val="${color(i)}"/></a:solidFill></c:spPr>${kind === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}`
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx(s)}${fill}${dLbls(labels)}${cat()}<c:val>${numRef(lineRef(s, i), s.values)}</c:val>${kind === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`
    }).join('')
    const axIds = '<c:axId val="500000001"/><c:axId val="500000002"/>'
    plot = kind === 'bar'
      ? `<c:barChart><c:barDir val="${t.startsWith('bar') ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}<c:gapWidth val="${grouping === 'clustered' ? 219 : 150}"/>${grouping === 'clustered' ? '<c:overlap val="-27"/>' : '<c:overlap val="100"/>'}${axIds}</c:barChart>`
      : kind === 'line' ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>`
        : `<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}${axIds}</c:areaChart>`
  }
  const axTitle = (s) => (s ? `<c:title>${rich(s)}<c:overlay val="0"/></c:title>` : '')
  const horizontal = t.startsWith('bar')
  const axes = pie ? '' : t === 'scatter'
    ? `<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${axTitle(spec.axisTitles?.x)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      + `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/>${axTitle(spec.axisTitles?.y)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'l' : 'b'}"/>${axTitle(spec.axisTitles?.x)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
      + `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horizontal ? 'b' : 'l'}"/><c:majorGridlines/>${axTitle(spec.axisTitles?.y)}<c:numFmt formatCode="${/100$/.test(t) ? '0%' : 'General'}" sourceLinked="${/100$/.test(t) ? 0 : 1}"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  // Title: none (''), automatic from the one series (no text: Excel uses the series name), or literal text.
  const title = spec.title === '' ? '' : spec.title == null && d.series.length === 1 && d.series[0].nameCell ? '<c:title><c:overlay val="0"/></c:title>'
    : `<c:title>${rich(chartTitle(spec, d))}<c:overlay val="0"/></c:title>`
  const legendPos = { bottom: 'b', right: 'r', top: 't', left: 'l' }[spec.legend ?? (pie || d.series.length > 1 ? 'bottom' : 'none')]
  const legend = legendPos ? `<c:legend><c:legendPos val="${legendPos}"/><c:overlay val="0"/></c:legend>` : ''
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:roundedCorners val="0"/>`
    + `<c:chart>${title}<c:autoTitleDeleted val="${spec.title === '' ? 1 : 0}"/><c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`
}

// ---- package helpers
const relsPath = (part) => { const i = part.lastIndexOf('/'); return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels` }
const resolve = (base, target) => {
  if (target.startsWith('/')) return target.slice(1)
  const parts = base.split('/').slice(0, -1)
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg) }
  return parts.join('/')
}
const relList = (xml) => [...(xml ?? '').matchAll(/<Relationship\b([^>]*)\/?>/g)].map((m) => Object.fromEntries([...m[1].matchAll(/(\w+)="([^"]*)"/g)].map((x) => [x[1], unesc(x[2])])))
const emptyRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rel}"></Relationships>`
function addRel(xml, type, target) {
  const ids = relList(xml).map((r) => Number(String(r.Id).replace(/\D/g, '')) || 0)
  const id = `rId${Math.max(0, ...ids) + 1}`
  return [xml.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`), id]
}
const addOverride = (xml, part, type) => (xml.includes(`PartName="/${part}"`) ? xml : xml.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`))
async function sheetParts(zip) {
  const wbXml = await zip.file('xl/workbook.xml').async('string'), rels = relList(await zip.file('xl/_rels/workbook.xml.rels').async('string'))
  return [...wbXml.matchAll(/<sheet\b([^>]*)\/>/g)].map((m) => {
    const name = unesc(/name="([^"]*)"/.exec(m[1])?.[1]), rid = /r:id="([^"]*)"/.exec(m[1])?.[1]
    return { name, part: resolve('xl/workbook.xml', rels.find((r) => r.Id === rid)?.Target ?? '') }
  })
}
const anchorXml = (a) => `<xdr:col>${a.column}</xdr:col><xdr:colOff>${Math.max(0, Math.round((a.columnOffset ?? 0) * EMU))}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${Math.max(0, Math.round((a.rowOffset ?? 0) * EMU))}</xdr:rowOff>`

/**
 * Adds chart parts and dynamic-array metadata to an ExcelJS-written package.
 * charts: [{sheetIndex, sheetName, spec, from, to, get}] ; spills: Set of `${sheetIndex}!A1` anchors that spill.
 */
export async function finishPackage(buf, { charts = [], spills = new Set(), sparklines = [] } = {}) {
  if (!charts.length && !spills.size && !sparklines.length) return buf
  const zip = await JSZip.loadAsync(buf)
  let types = await zip.file('[Content_Types].xml').async('string')
  const sheets = await sheetParts(zip)
  let chartN = Object.keys(zip.files).filter((f) => /^xl\/charts\/chart\d+\.xml$/.test(f)).length
  let drawingN = Object.keys(zip.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f)).length
  const bySheet = new Map()
  for (const c of charts) { if (!bySheet.has(c.sheetIndex)) bySheet.set(c.sheetIndex, []); bySheet.get(c.sheetIndex).push(c) }
  for (const [index, list] of bySheet) {
    const sheet = sheets[index]
    if (!sheet) continue
    const sRels = relsPath(sheet.part)
    let sheetRelsXml = (await zip.file(sRels)?.async('string')) ?? emptyRels
    let sheetXml = await zip.file(sheet.part).async('string')
    // Reuse the drawing ExcelJS wrote for pictures, or add one.
    let drawingPart = relList(sheetRelsXml).filter((r) => r.Type === REL.drawing).map((r) => resolve(sheet.part, r.Target))[0]
    let drawingXml
    if (drawingPart) drawingXml = await zip.file(drawingPart).async('string')
    else {
      drawingPart = `xl/drawings/drawing${++drawingN}.xml`
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`
      let id
      ;[sheetRelsXml, id] = addRel(sheetRelsXml, REL.drawing, `../drawings/${drawingPart.split('/').pop()}`)
      if (!/xmlns:r=/.test(sheetXml.slice(0, 600))) sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`)
      const before = /<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(sheetXml)
      sheetXml = before ? sheetXml.slice(0, before.index) + `<drawing r:id="${id}"/>` + sheetXml.slice(before.index) : sheetXml.replace('</worksheet>', `<drawing r:id="${id}"/></worksheet>`)
      types = addOverride(types, drawingPart, CT.drawing)
    }
    const dRels = relsPath(drawingPart)
    let drawingRelsXml = (await zip.file(dRels)?.async('string')) ?? emptyRels
    let shapeId = Math.max(1, ...[...drawingXml.matchAll(/cNvPr id="(\d+)"/g)].map((m) => +m[1])) + 1
    for (const c of list) {
      const part = `xl/charts/chart${++chartN}.xml`
      zip.file(part, chartXml(c.spec, c.sheetName ?? sheet.name, c.get))
      types = addOverride(types, part, CT.chart)
      let rid
      ;[drawingRelsXml, rid] = addRel(drawingRelsXml, REL.chart, `../charts/${part.split('/').pop()}`)
      const id = shapeId++
      drawingXml = drawingXml.replace('</xdr:wsDr>', `<xdr:twoCellAnchor><xdr:from>${anchorXml(c.from)}</xdr:from><xdr:to>${anchorXml(c.to)}</xdr:to>`
        + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id - 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
        + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData></a:graphic>`
        + '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')
    }
    zip.file(drawingPart, drawingXml); zip.file(dRels, drawingRelsXml); zip.file(sRels, sheetRelsXml); zip.file(sheet.part, sheetXml)
  }
  // Sparklines: one x14 extension per sheet, inside the sheet's extLst (ExcelJS may already have one for x14
  // conditional formats; a sheet has only one extLst).
  for (const { sheetIndex, groups } of sparklines) {
    const sheet = sheets[sheetIndex]
    if (!sheet || !groups.length) continue
    let xml = await zip.file(sheet.part).async('string')
    const ext = `<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}">${groups.map(sparkGroupXml).join('')}</x14:sparklineGroups></ext>`
    // The sheet's own extLst is its last child (rules have extLst elements of their own, e.g. x14 data bars).
    const tail = /<\/extLst>\s*<\/worksheet>\s*$/.exec(xml)
    xml = tail ? xml.slice(0, tail.index) + `${ext}</extLst></worksheet>` : xml.replace(/<\/worksheet>\s*$/, `<extLst>${ext}</extLst></worksheet>`)
    zip.file(sheet.part, xml)
  }
  if (spills.size) {
    for (const [i, sheet] of sheets.entries()) {
      let xml = await zip.file(sheet.part).async('string'), changed = false
      xml = xml.replace(/<c r="([A-Z]+\d+)"([^>]*)>(<f t="array")/g, (m, ref, attrs, f) => {
        if (!spills.has(`${i}!${ref}`) || /\bcm=/.test(attrs)) return m
        changed = true
        return `<c r="${ref}"${attrs} cm="1">${f}`
      })
      if (changed) zip.file(sheet.part, xml)
    }
    zip.file('xl/metadata.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
      + '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
      + '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
      + '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')
    let wbRels = await zip.file('xl/_rels/workbook.xml.rels').async('string')
    if (!wbRels.includes(REL.metadata)) [wbRels] = addRel(wbRels, REL.metadata, 'metadata.xml')
    zip.file('xl/_rels/workbook.xml.rels', wbRels)
    types = addOverride(types, 'xl/metadata.xml', CT.metadata)
  }
  zip.file('[Content_Types].xml', types)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** Formula cells marked as dynamic arrays (cm= on an array formula): Set of `${sheetIndex}!A1`. */
export async function readDynamicArrays(buf) {
  const zip = await JSZip.loadAsync(buf), out = new Set()
  for (const [i, sheet] of (await sheetParts(zip)).entries()) {
    const xml = await zip.file(sheet.part)?.async('string')
    if (!xml) continue
    for (const m of xml.matchAll(/<c\b([^>]*)>\s*<f\b[^>]*\bt="array"/g)) {
      const ref = /\br="([A-Z]+\d+)"/.exec(m[1])?.[1]
      if (ref && /\bcm="/.test(m[1])) out.add(`${i}!${ref}`)
    }
  }
  return out
}

// ---- sparklines (x14)
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main', XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const argbOf = (c) => `FF${String(c).replace('#', '').toUpperCase()}`
/** groups: [{ type, colors, items: [{row, col, source, sheetName}] }] */
function sparkGroupXml(g) {
  const type = g.type === 'column' ? ' type="column"' : g.type === 'winloss' ? ' type="stacked"' : ''
  const c = { color: '#376092', negative: '#D00000', markers: '#D00000', ...(g.colors ?? {}) }
  const colors = `<x14:colorSeries rgb="${argbOf(c.color)}"/><x14:colorNegative rgb="${argbOf(c.negative)}"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="${argbOf(c.markers)}"/>`
    + `<x14:colorFirst rgb="${argbOf(c.markers)}"/><x14:colorLast rgb="${argbOf(c.markers)}"/><x14:colorHigh rgb="${argbOf(c.markers)}"/><x14:colorLow rgb="${argbOf(c.markers)}"/>`
  const items = g.items.map((it) => `<x14:sparkline><xm:f>${esc(sheetRef(it.sheetName))}!${rangeRef(it.source)}</xm:f><xm:sqref>${colName(it.col)}${it.row + 1}</xm:sqref></x14:sparkline>`).join('')
  return `<x14:sparklineGroup${type} displayEmptyCellsAs="gap">${colors}<x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`
}
const rangeRef = (g) => `${colName(g.startColumn)}${g.startRow + 1}:${colName(g.endColumn)}${g.endRow + 1}`
/** Sparkline groups per sheet index: [{ sheetIndex, groups: [{ type, colors, items: [{ row, col, sheetName, source }] }] }]. */
export async function readSparklines(buf) {
  const zip = await JSZip.loadAsync(buf), out = []
  for (const [sheetIndex, sheet] of (await sheetParts(zip)).entries()) {
    const xml = await zip.file(sheet.part)?.async('string')
    if (!xml || !/sparklineGroup/.test(xml)) continue
    const groups = []
    for (const m of xml.matchAll(/<x14:sparklineGroup\b([^>]*)>([\s\S]*?)<\/x14:sparklineGroup>/g)) {
      const t = /\btype="(\w+)"/.exec(m[1])?.[1]
      const color = (tag) => { const v = new RegExp(`<x14:${tag}\\b[^>]*rgb="([0-9A-Fa-f]{8})"`).exec(m[2])?.[1]; return v ? `#${v.slice(2)}` : undefined }
      const items = [...m[2].matchAll(/<x14:sparkline>\s*<xm:f>([^<]*)<\/xm:f>\s*<xm:sqref>([^<]*)<\/xm:sqref>/g)].map((x) => {
        const f = unesc(x[1]), mm = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(f), at = parseRange(x[2])
        return mm && at ? { row: at.startRow, col: at.startColumn, sheetName: (mm[1] ?? mm[2]).replace(/''/g, "'"), source: parseRange(mm[3]) } : null
      }).filter((x) => x?.source)
      const colors = Object.fromEntries([['color', color('colorSeries')], ['negative', color('colorNegative')], ['markers', color('colorMarkers')]].filter(([, v]) => v))
      if (items.length) groups.push({ type: t === 'column' ? 'column' : t === 'stacked' ? 'winloss' : 'line', colors, items })
    }
    if (groups.length) out.push({ sheetIndex, groups })
  }
  return out
}

// ---- reading charts back
// Prefixes are the file's choice: rewrite the ones bound to the chart / drawing / DrawingML / relationship namespaces
// (and an unprefixed default namespace) to c:, xdr:, a:, r: before reading.
const CANON = { [NS.c]: 'c', [NS.xdr]: 'xdr', [NS.a]: 'a', [NS.r]: 'r' }
export function canonical(xml) {
  const map = {}
  for (const m of xml.matchAll(/\sxmlns:([\w.-]+)="([^"]*)"/g)) if (CANON[m[2]] && CANON[m[2]] !== m[1]) map[m[1]] = CANON[m[2]]
  let out = xml
  const keys = Object.keys(map)
  if (keys.length) {
    const re = new RegExp(`(<\\/?|\\s)(${keys.map((k) => k.replace(/[.-]/g, '\\$&')).join('|')}):`, 'g')
    out = out.replace(re, (m, lead, p) => `${lead}${map[p]}:`)
  }
  const def = /<[\w]+[^>]*\sxmlns="([^"]*)"/.exec(out)?.[1]
  if (def && CANON[def]) out = out.replace(/<(\/?)([A-Za-z][\w.-]*)(?=[\s>/])/g, (m, slash, name) => `<${slash}${CANON[def]}:${name}`)
  return out
}

const attr = (xml, tag, name) => new RegExp(`<${tag}\\b[^>]*\\b${name}="([^"]*)"`).exec(xml)?.[1]
const texts = (xml) => [...(xml ?? '').matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => unesc(m[1])).join('')
const anchorIn = (xml) => ({ column: +(/<xdr:col>(\d+)/.exec(xml)?.[1] ?? 0), columnOffset: Math.round(+(/<xdr:colOff>(-?\d+)/.exec(xml)?.[1] ?? 0) / EMU),
  row: +(/<xdr:row>(\d+)/.exec(xml)?.[1] ?? 0), rowOffset: Math.round(+(/<xdr:rowOff>(-?\d+)/.exec(xml)?.[1] ?? 0) / EMU) })

/** Chart specs per sheet index: [{sheetIndex, spec, from, to}] (spec.source.sheetName is resolved by the caller). */
export async function readCharts(buf) {
  const zip = await JSZip.loadAsync(buf)
  const out = []
  const sheets = await sheetParts(zip)
  for (const [sheetIndex, sheet] of sheets.entries()) {
    const sRels = relList(await zip.file(relsPath(sheet.part))?.async('string'))
    for (const dr of sRels.filter((r) => r.Type === REL.drawing)) {
      const drawingPart = resolve(sheet.part, dr.Target)
      const drawingRaw = await zip.file(drawingPart)?.async('string')
      if (!drawingRaw) continue
      const drawingXml = canonical(drawingRaw)
      const dRels = relList(await zip.file(relsPath(drawingPart))?.async('string'))
      for (const m of drawingXml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g)) {
        const body = m[2]
        const rid = /<c:chart\b[^>]*r:id="([^"]*)"/.exec(body)?.[1]
        if (!rid) continue
        const chartPart = resolve(drawingPart, dRels.find((r) => r.Id === rid)?.Target ?? '')
        const raw = await zip.file(chartPart)?.async('string')
        if (!raw) continue
        const spec = specFromXml(raw)
        if (!spec) continue
        const from = anchorIn(/<xdr:from>([\s\S]*?)<\/xdr:from>/.exec(body)?.[1] ?? '')
        const toXml = /<xdr:to>([\s\S]*?)<\/xdr:to>/.exec(body)?.[1]
        // A one-cell anchor stores the chart's size (xdr:ext, EMU) instead of an end cell.
        const ext = /<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(body)
        out.push({ sheetIndex, spec, from, to: toXml ? anchorIn(toXml) : null, size: { width: ext ? +ext[1] / EMU : 480, height: ext ? +ext[2] / EMU : 288 } })
      }
    }
  }
  return out
}

/** Sheetly chart spec from a DrawingML chart part (types Sheetly draws; 3-D charts read as their 2-D kind). */
export function specFromXml(raw) {
  const xml = canonical(raw)
  let type = null
  const plot = /<c:plotArea>([\s\S]*)<\/c:plotArea>/.exec(xml)?.[1] ?? ''
  const grouping = attr(plot, 'c:grouping', 'val')
  if (/<c:bar(3D)?Chart>/.test(plot)) {
    const dir = attr(plot, 'c:barDir', 'val') === 'bar' ? 'bar' : 'column'
    type = dir + (grouping === 'stacked' ? 'Stacked' : grouping === 'percentStacked' ? 'Stacked100' : '')
  } else if (/<c:line(3D)?Chart>/.test(plot)) type = /<c:marker>\s*<c:symbol val="none"\/>/.test(plot) || !/<c:marker val="1"\/>/.test(plot) ? 'line' : 'lineMarkers'
  else if (/<c:area(3D)?Chart>/.test(plot)) type = grouping === 'stacked' || grouping === 'percentStacked' ? 'areaStacked' : 'area'
  else if (/<c:pie(3D)?Chart>/.test(plot) || /<c:ofPieChart>/.test(plot)) type = 'pie'
  else if (/<c:doughnutChart>/.test(plot)) type = 'doughnut'
  else if (/<c:scatterChart>/.test(plot)) type = 'scatter'
  if (!type) return null
  // Series references -> one source block (the bounding box of names, categories and values on one sheet).
  const refs = [...plot.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => unesc(m[1]))
  const parsed = refs.map((f) => { const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(f); return m ? { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), range: parseRange(m[3]) } : null }).filter((x) => x?.range)
  if (!parsed.length || parsed.some((p) => p.sheet !== parsed[0].sheet)) return null
  const box = parsed.reduce((b, { range: g }) => ({ startRow: Math.min(b.startRow, g.startRow), endRow: Math.max(b.endRow, g.endRow), startColumn: Math.min(b.startColumn, g.startColumn), endColumn: Math.max(b.endColumn, g.endColumn) }),
    { startRow: Infinity, endRow: -1, startColumn: Infinity, endColumn: -1 })
  const valRefs = [...plot.matchAll(/<c:(val|yVal)>[\s\S]*?<c:f>([^<]*)<\/c:f>/g)].map((m) => parseRange(unesc(m[2]).replace(/^.*!/, ''))).filter(Boolean)
  const plotBy = valRefs.length && valRefs[0].startColumn === valRefs[0].endColumn && valRefs[0].startRow !== valRefs[0].endRow ? 'cols'
    : valRefs.length && valRefs[0].startRow === valRefs[0].endRow && valRefs[0].startColumn !== valRefs[0].endColumn ? 'rows' : null
  const titleXml = /<c:chart>\s*<c:title>([\s\S]*?)<\/c:title>/.exec(xml)?.[1]
  const deleted = attr(xml, 'c:autoTitleDeleted', 'val') === '1'
  const title = deleted && !titleXml ? '' : titleXml && /<c:tx>/.test(titleXml) ? texts(titleXml) : null
  const legendPos = attr(xml, 'c:legendPos', 'val')
  const legend = /<c:legend>/.test(xml) ? ({ b: 'bottom', r: 'right', t: 'top', l: 'left', tr: 'right' }[legendPos] ?? 'bottom') : 'none'
  const dataLabels = /<c:showVal val="1"\/>/.test(plot)
  // Axis titles: the category axis (or, in a scatter chart, the bottom value axis) is x; the other value axis is y.
  const axes = [...plot.matchAll(/<c:(catAx|valAx|dateAx)>([\s\S]*?)<\/c:\1>/g)].map((m) => ({ kind: m[1], pos: attr(m[2], 'c:axPos', 'val'), title: texts(/<c:title>([\s\S]*?)<\/c:title>/.exec(m[2])?.[1]) }))
  const horizontalAxis = (a) => a.pos === 'b' || a.pos === 't'
  const xAxis = axes.find((a) => a.kind !== 'valAx') ?? axes.find(horizontalAxis), yAxis = axes.find((a) => a !== xAxis && a.kind === 'valAx')
  const axisTitles = xAxis?.title || yAxis?.title ? { x: xAxis?.title || '', y: yAxis?.title || '' } : null
  const serXml = [...plot.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map((m) => m[1])
  // A reference inside one element only (<c:tx>, <c:cat>, <c:val>, ...), never one from the next element.
  const inner = (x, tag) => new RegExp(`<c:${tag}>([\\s\\S]*?)<\\/c:${tag}>`).exec(x)?.[1] ?? null
  const refIn = (x, tag) => { const f = /<c:f>([^<]*)<\/c:f>/.exec(inner(x, tag) ?? '')?.[1]; return f ? parseRange(unesc(f).replace(/^.*!/, '')) : null }
  // A literal series name: <c:tx><c:v>Revenue</c:v></c:tx> or rich text.
  const literalName = (x) => { const t = inner(x, 'tx'); if (!t || /<c:f>/.test(t)) return null; const v = /<c:v>([^<]*)<\/c:v>/.exec(t)?.[1]; return v != null ? unesc(v) : texts(t) || null }
  // (scatter series each have their own X values)
  const series = serXml.map((x) => ({ name: refIn(x, 'tx'), values: refIn(x, 'val') ?? refIn(x, 'yVal'), xValues: refIn(x, 'xVal'), nameText: literalName(x) }))
  const categories = serXml.length ? refIn(serXml[0], 'cat') ?? refIn(serXml[0], 'xVal') : null
  const source = { sheetName: parsed[0].sheet, range: box }
  // A chart from a file keeps its series exactly as referenced: re-reading the block by cell types could, for
  // example, take numeric category labels (years) for a series.
  if (series.every((x) => x.values)) {
    source.series = series.map((x) => ({ name: x.name ? [x.name.startRow, x.name.startColumn] : null, nameText: x.nameText, values: x.values, ...(x.xValues ? { xValues: x.xValues } : {}) }))
    source.categories = categories
  }
  return { v: 2, type, source, plotBy, title, legend, dataLabels, axisTitles }
}
