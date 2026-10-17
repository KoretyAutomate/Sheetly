// Ranges that sheet objects point at (chart series, sparkline data and cells, PivotTable sources and places) are
// kept as hidden defined names, so Univer moves them with inserted / deleted rows and columns (and turns them into
// #REF! when their cells are deleted) inside the same undo step as the edit, exactly as it does for formulas.
// The names are internal: they are not shown in the Name Manager and not written to .xlsx (REF_PREFIX).
import { ICommandService } from '@univerjs/core'
import { IDefinedNamesService } from '@univerjs/engine-formula'
import { REF_PREFIX, parseRange, rangeA1 } from './xlsx-objects.js'

export { REF_PREFIX }
const quote = (n) => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`)
const absolute = (a1) => a1.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')

export function createRefs(api) {
  const commands = api._injector.get(ICommandService)
  const names = () => api._injector.get(IDefinedNamesService)
  const wb = () => api.getActiveWorkbook()
  let seq = 0
  const nextId = () => `${REF_PREFIX}${Date.now().toString(36)}${(++seq).toString(36)}`

  /** A new hidden name for a range on a sheet; returns its name. Not an undo step of its own. */
  function track(sheetId, range, book = wb()) {
    const sheet = book.getSheetBySheetId(sheetId)
    if (!sheet || !range) return null
    const name = nextId()
    commands.syncExecuteCommand('formula.mutation.set-defined-name', { unitId: book.getId(), id: name, name,
      formulaOrRefString: `${quote(sheet.getSheetName())}!${absolute(rangeA1(range))}`, hidden: true })
    return name
  }
  /** Where a name points now: { sheetId, range }, or null once its cells are gone (#REF!). */
  function resolve(name, book = wb()) {
    if (!name || !book) return null
    const v = names().getValueByName(book.getId(), name)
    const text = v?.formulaOrRefString
    if (!text || /#REF!/i.test(text)) return null
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(text)
    if (!m) return null
    const sheetName = (m[1] ?? m[2]).replace(/''/g, "'")
    const sheet = book.getSheets().find((s) => s.getSheetName() === sheetName)
    const range = parseRange(m[3])
    return sheet && range ? { sheetId: sheet.getSheetId(), range } : null
  }
  return { track, resolve }
}
