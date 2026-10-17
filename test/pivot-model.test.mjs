// PivotTable layout (Excel's Compact Form defaults) and GETPIVOTDATA, in Node.
import { layout, getPivotData, readSource, defaultFn, filterCells, aggregate } from '../src/pivot-model.js'

const out = []
const check = (name, ok, detail = '') => out.push(`${ok ? 'PASS' : 'FAIL'}  ${name}${ok || !detail ? '' : '  -> ' + detail}`)
const rows = [['Region', 'Product', 'Sales', 'Date'], ['East', 'A', 10, 45000], ['East', 'B', 20, 45001], ['West', 'A', 5, 45002], ['West', 'A', 7, null], ['', 'B', 3, 45003]]
const src = readSource((r, c) => rows[r]?.[c] ?? null, { startRow: 0, endRow: rows.length - 1, startColumn: 0, endColumn: 3 })
const show = (L) => L.grid.map((r) => r.map((c) => (c.v == null ? '' : (c.indent ? '  '.repeat(c.indent) : '') + c.v)).join(' | ')).join('\n')
const sum = [{ field: 'Sales', fn: 'sum' }]

let L = layout({ rows: ['Region'], cols: [], values: sum, filters: [] }, src)
check('one row field: Row Labels / Sum of Sales, items A to Z, (blank) last, Grand Total', show(L) === 'Row Labels | Sum of Sales\nEast | 30\nWest | 12\n(blank) | 3\nGrand Total | 45', '\n' + show(L))
L = layout({ rows: ['Region'], cols: ['Product'], values: sum, filters: [] }, src)
check('row and column fields: Column Labels header, empty cells where no data, Grand Total row and column',
  show(L) === 'Sum of Sales | Column Labels |  | \nRow Labels | A | B | Grand Total\nEast | 10 | 20 | 30\nWest | 12 |  | 12\n(blank) |  | 3 | 3\nGrand Total | 22 | 23 | 45', '\n' + show(L))
L = layout({ rows: ['Region', 'Product'], cols: [], values: sum, filters: [] }, src)
check('two row fields: compact form, subtotals on the group row, inner items indented',
  show(L) === 'Row Labels | Sum of Sales\nEast | 30\n  A | 10\n  B | 20\nWest | 12\n  A | 12\n(blank) | 3\n  B | 3\nGrand Total | 45', '\n' + show(L))
L = layout({ rows: ['Region'], cols: [], values: [{ field: 'Sales', fn: 'sum' }, { field: 'Product', fn: 'count' }], filters: [] }, src)
check('two value fields side by side', show(L) === 'Row Labels | Sum of Sales | Count of Product\nEast | 30 | 2\nWest | 12 | 2\n(blank) | 3 | 1\nGrand Total | 45 | 5', '\n' + show(L))
L = layout({ rows: [], cols: [], values: sum, filters: [] }, src)
check('values only: caption and total', show(L) === 'Sum of Sales\n45', '\n' + show(L))
L = layout({ rows: ['Region'], cols: [], values: [{ field: 'Sales', fn: 'average' }], filters: [{ field: 'Product', items: ['A'] }] }, src)
check('report filter Product = A and Average of Sales', show(L) === 'Row Labels | Average of Sales\nEast | 10\nWest | 6\nGrand Total | 7.333333333333333', '\n' + show(L))
check('filter cells: (All) / item / (Multiple Items)', JSON.stringify(filterCells({ filters: [{ field: 'A', items: null }, { field: 'B', items: ['x'] }, { field: 'C', items: ['x', 'y'] }] })) === '[["A","(All)"],["B","x"],["C","(Multiple Items)"]]')
check('a new value field: Sum for numbers, Count for text', defaultFn(src, 'Sales') === 'sum' && defaultFn(src, 'Region') === 'count')
check('aggregates: count counts non-blank, countNums numbers, stdDev / var', aggregate('count', ['a', null, 3]) === 2 && aggregate('countNums', ['a', null, 3]) === 1
  && Math.abs(aggregate('stdDev', [2, 4, 4, 4, 5, 5, 7, 9]) - 2.138089935) < 1e-6 && aggregate('varp', [2, 4, 4, 4, 5, 5, 7, 9]) === 4)
check('blank header -> not a valid PivotTable source', readSource((r, c) => [['A', null]][r]?.[c] ?? null, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }) === null)

L = layout({ rows: ['Region'], cols: ['Product'], values: [{ field: 'Sales', fn: 'sum' }, { field: 'Sales', fn: 'max' }], filters: [] }, src)
check('column field and two value fields: item row, then the value captions next to Row Labels',
  show(L).split('\n').slice(0, 3).join('\n') === ' | Column Labels |  |  |  |  | \n | A |  | B |  | Total Sum of Sales | Total Max of Sales\nRow Labels | Sum of Sales | Max of Sales | Sum of Sales | Max of Sales |  | ', '\n' + show(L))
{
  const rows2 = [['K', 'V'], [1, 10], ['1', 20]]
  const src2 = readSource((r, c) => rows2[r]?.[c] ?? null, { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 })
  const L2 = layout({ rows: ['K'], cols: [], values: [{ field: 'V', fn: 'sum' }], filters: [] }, src2)
  check('the number 1 and the text "1" are separate items, each counted once', show(L2) === 'Row Labels | Sum of V\n1 | 10\n1 | 20\nGrand Total | 30', '\n' + show(L2))
  const d2 = { rows: ['K'], cols: [], values: [{ field: 'V', fn: 'sum' }], filters: [] }
  const gp = (it) => { const x = getPivotData(d2, src2, 'V', [['K', it]]); return x.error ?? x.value }
  check('GETPIVOTDATA tells them apart: 1 -> 10, "1" -> 20', gp(1) === 10 && gp('1') === 20, `${gp(1)} ${gp('1')}`)
  const F2 = layout({ rows: [], cols: [], values: [{ field: 'V', fn: 'sum' }], filters: [{ field: 'K', items: ['1'] }] }, src2)
  check('a report filter on the text "1" keeps only that item', show(F2) === 'Sum of V\n20', '\n' + show(F2))
}
const def = { rows: ['Region'], cols: ['Product'], values: sum, filters: [] }
const g = (df, pairs) => { const x = getPivotData(def, src, df, pairs); return x.error ?? x.value }
check('GETPIVOTDATA("Sales", pt) = grand total 45', g('Sales', []) === 45)
check('GETPIVOTDATA("Sum of Sales", pt, "Region", "East") = 30', g('Sum of Sales', [['Region', 'East']]) === 30)
check('GETPIVOTDATA(..., "Region", "West", "Product", "A") = 12, names case-insensitive', g('sales', [['region', 'west'], ['product', 'a']]) === 12)
check('GETPIVOTDATA: item not in the report -> #REF!', g('Sales', [['Region', 'North']]) === '#REF!')
check('GETPIVOTDATA: field not on rows/columns -> #REF!', g('Sales', [['Date', 45000]]) === '#REF!')
check('GETPIVOTDATA: combination with no data -> #REF!', g('Sales', [['Region', 'West'], ['Product', 'B']]) === '#REF!')
check('GETPIVOTDATA: unknown data field -> #REF!', g('Profit', []) === '#REF!')

const pass = out.filter((l) => l.startsWith('PASS')).length
console.log(`pivot model  ${new Date().toISOString()}\n${out.join('\n')}\n\n${pass}/${out.length} passed`)
process.exitCode = pass === out.length ? 0 : 1
