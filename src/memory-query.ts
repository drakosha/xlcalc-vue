// Queries of the memory source over a sheet's shown text: the display filter and sort, search, a column's values, totals.
import type { CellRect } from 'xlcalc/contract'
import { findCells, replaceCells, type CellText, type EachCell } from './find-match'
import { columnList as listOf, filterActive, visibleList as listUnder, type QuerySheet } from './sheet-query'
import type { FindRequest, FindResult, RangeStats, SheetFilter } from './source'
import { addCell, noStats } from './stats'

/** What the queries read of a sheet: the display filter's reading, and for search where its columns end and what each cell opens with. */
export interface MemorySheetQuery extends QuerySheet {
  maxCol: number
  /** The text a cell opens with: its formula, else its value as typed. */
  editAt(row: number, col: number): string
}

/** Hits a search gives at most; `total` counts all. */
const HITS_MAX = 1000

/** Text sorts as the runtime's locale orders it, digits by their numbers, in any case. */
const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' })

/** The sheet's rows in display order: the header first, then the rows that meet the conditions, sorted; null without conditions. */
export function visibleList(sheet: MemorySheetQuery, filter: SheetFilter | undefined): number[] | null {
  return filterActive(filter) ? listUnder(sheet, filter, collator) : null
}

/** A column's shown values with their counts, over the rows the other columns' conditions leave; '' is an empty cell. */
export function columnList(sheet: MemorySheetQuery, col: number, filter: SheetFilter): { values: { text: string; count: number }[]; more: number } {
  return listOf(sheet, col, filter, collator)
}

/** Every cell's text a search reads: what it opens with (Look in: Formulas, Excel's default), or what it shows. */
function cellsOf(sheet: MemorySheetQuery, lookIn: FindRequest['lookIn']): EachCell {
  const textAt = lookIn === 'values' ? sheet.textAt : sheet.editAt
  return (visit) => {
    for (let row = 1; row <= sheet.maxRow; row++) for (let col = 1; col <= sheet.maxCol; col++) visit(row, col, textAt(row, col))
  }
}
const shownRows = (sheet: MemorySheetQuery, filter: SheetFilter | undefined): Set<number> | null => {
  const list = visibleList(sheet, filter)
  return list ? new Set(list) : null
}

/** Excel's Find: hits in its order, only the shown rows under a filter. */
export function findHits(sheet: MemorySheetQuery, req: FindRequest): FindResult {
  return findCells(cellsOf(sheet, req.lookIn), req, shownRows(sheet, req.filter), HITS_MAX)
}

/** Excel's Replace All: the new text of each cell the request finds in formulas, as Excel replaces whatever Look in says. */
export function replaceTexts(sheet: MemorySheetQuery, req: FindRequest, by: string): CellText[] {
  return replaceCells(cellsOf(sheet, 'formulas'), req, by, shownRows(sheet, req.filter))
}

/** Totals of a rectangle as Excel's status line, over the shown rows under a filter. */
export function rangeStats(sheet: MemorySheetQuery, rect: CellRect, filter: SheetFilter | undefined): RangeStats {
  const shown = shownRows(sheet, filter)
  const out = noStats()
  for (let r = rect.top; r <= Math.min(rect.bottom, sheet.maxRow); r++) {
    if (shown && !shown.has(r)) continue
    for (let c = rect.left; c <= Math.min(rect.right, sheet.maxCol); c++) addCell(out, sheet.textAt(r, c), sheet.valueAt(r, c))
  }
  return out
}
