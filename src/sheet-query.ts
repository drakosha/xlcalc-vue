// The display filter, its sort and the funnel's list over a sheet's displayed texts and values, for memorySource and ./local:
// conditions compare the displayed text, header rows stay at the top, and a sort orders by value in Excel's order (`valueComparator`
// of `xlcalc/text`, the Sort command's: errors equal, text in the workbook locale's order) with empty cells last in either direction.
import type { CellValue } from 'xlcalc/contract'
import { valueComparator } from 'xlcalc/text'
import type { SheetFilter } from './contract/source'

/** What the display filter reads of a sheet: where its rows end, a cell's displayed text ('' when empty) and its value. */
export interface QuerySheet {
  readonly maxRow: number
  textAt(row: number, col: number): string
  valueAt(row: number, col: number): CellValue | undefined
}

/** Distinct values in the funnel's list; the rest are counted in `more`. */
const VALUES_MAX = 1000

/** Whether a filter requests anything: a sort, or a condition on a column (text of spaces requests nothing). */
export function filterActive(filter: SheetFilter | undefined): filter is SheetFilter {
  return !!filter && (!!filter.sort || Object.values(filter.cols ?? {}).some((c) => !!c.values || !!c.text?.trim()))
}

/** The columns a filter reads: those of its conditions and of its sort. */
export function filterColumns(filter: SheetFilter): number[] {
  const cols = new Set(Object.keys(filter.cols ?? {}).map(Number))
  if (filter.sort) cols.add(filter.sort.col)
  return [...cols]
}

const headOf = (filter: SheetFilter, maxRow: number) => Math.min(Math.max(0, Math.trunc(filter.head)), maxRow)

/** The column conditions as a row test, with one column excluded (its own funnel lists the rows the others keep). */
function conditions(sheet: QuerySheet, filter: SheetFilter, except?: number): (row: number) => boolean {
  const conds = Object.entries(filter.cols ?? {})
    .filter(([col]) => Number(col) !== except)
    .map(([col, c]) => ({ col: Number(col), values: c.values ? new Set(c.values) : null, text: c.text?.trim() ? c.text.trim().toLocaleLowerCase() : null }))
  return (row) =>
    conds.every((c) => {
      const text = sheet.textAt(row, c.col)
      return (!c.values || c.values.has(text)) && (!c.text || text.toLocaleLowerCase().includes(c.text))
    })
}

/** A cell's value as the sort takes it: an empty cell, and empty text (a formula's ""), is blank. */
const sortKey = (v: CellValue | undefined): CellValue => (v === undefined || v === '' ? null : v)

/** Excel's order of values, with text in the order of `locale` (a loaded locale's tag; en-US's without one). */
const orderOf = (locale: string | undefined, descending = false) => valueComparator({ descending, ...(locale ? { locale } : {}) })

/** The sheet's rows in display order: the header rows as they are, then the rows that match every condition, sorted when requested (stable, empty cells last), with text in `locale`'s order. */
export function visibleList(sheet: QuerySheet, filter: SheetFilter, locale?: string): number[] {
  const head = headOf(filter, sheet.maxRow)
  const meets = conditions(sheet, filter)
  const body: number[] = []
  for (let r = head + 1; r <= sheet.maxRow; r++) if (meets(r)) body.push(r)
  const sort = filter.sort
  let rows = body
  if (sort) {
    const compare = orderOf(locale, sort.dir === 'desc')
    const keys = body.map((r) => sortKey(sheet.valueAt(r, sort.col)))
    const order = body.map((_, i) => i)
    order.sort((i, j) => compare(keys[i]!, keys[j]!))
    rows = order.map((i) => body[i]!)
  }
  const out: number[] = []
  for (let r = 1; r <= head; r++) out.push(r)
  return out.concat(rows)
}

/** A column's funnel list: its displayed texts with their counts over the rows the other columns' conditions keep, '' for empty cells, in the sort's order. */
export function columnList(sheet: QuerySheet, col: number, filter: SheetFilter, locale?: string): { values: { text: string; count: number }[]; more: number } {
  const meets = conditions(sheet, filter, col)
  const counts = new Map<string, { count: number; value: CellValue }>()
  for (let r = headOf(filter, sheet.maxRow) + 1; r <= sheet.maxRow; r++) {
    if (!meets(r)) continue
    const text = sheet.textAt(r, col)
    const got = counts.get(text)
    if (got) got.count++
    else counts.set(text, { count: 1, value: sortKey(sheet.valueAt(r, col)) })
  }
  // values the order treats as equal (errors, a number displayed two ways) are ordered by their displayed text
  const compare = orderOf(locale)
  const all = [...counts.entries()].sort(([ta, a], [tb, b]) => compare(a.value, b.value) || compare(ta, tb))
  return { values: all.slice(0, VALUES_MAX).map(([text, { count }]) => ({ text, count })), more: Math.max(0, all.length - VALUES_MAX) }
}
