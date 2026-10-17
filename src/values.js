// FRange.getCalculatedValue(): a cell's value as formulas see it — the stored value, or for a cell inside a dynamic
// array's spill, the calculated result there — never the formatted text that Univer's getValue() returns for a
// formatted number ("$100"). Sheetly reads cell values through this.
import { FormulaDataModel } from '@univerjs/engine-formula'
import { FRange } from '@univerjs/sheets/facade'

/** The calculated spill value at a cell of a worksheet (null when none). */
export function spillValue(injector, unitId, sheetId, r, c) {
  const s = spillCell(injector, unitId, sheetId, r, c)
  return s ? (s.t === 3 ? !!s.v : s.v ?? null) : null
}
const spillCell = (injector, unitId, sheetId, r, c) => injector.get(FormulaDataModel).getArrayFormulaCellData?.()?.[unitId]?.[sheetId]?.[r]?.[c] ?? null

class SheetlyCalculatedValue extends FRange {
  getCalculatedValue() {
    const { startRow: r, startColumn: c } = this._range
    // Logical cells are stored as 1 / 0 with type 3: give TRUE / FALSE, as formulas see them.
    const raw = this._worksheet.getCellRaw?.(r, c)
    const v = this.getRawValue()
    if (v !== null && v !== undefined) return raw?.t === 3 ? !!v : v
    const s = spillCell(this._injector, this._workbook.getUnitId(), this._worksheet.getSheetId(), r, c)
    return s ? (s.t === 3 ? !!s.v : s.v ?? null) : null
  }
}
FRange.extend(SheetlyCalculatedValue)
