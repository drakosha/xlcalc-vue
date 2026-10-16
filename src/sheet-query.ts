// The display filter, its sort and the funnel's list over a sheet's shown texts and values, for memorySource and ./local:
// conditions compare the shown text, the head rows stay on top, a sort orders by value with empty cells last either way.
import type { CellValue } from 'xlcalc/contract'
import type { SheetFilter } from './source'

/** What the display filter reads of a sheet: where its rows end, a cell's shown text ('' when empty) and its value. */
export interface QuerySheet {
  readonly maxRow: number
  textAt(row: number, col: number): string
  valueAt(row: number, col: number): CellValue | undefined
}

/** Distinct values in the funnel's list; the rest are counted in `more`. */
const VALUES_MAX = 1000

/** Whether a filter asks for anything: a sort, or a condition on a column (a text of spaces asks for nothing). */
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

/** The column conditions as a test of a row, one column left out (its own funnel lists the rows the others keep). */
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

/** The sort's rank of a value: numbers, text, booleans, errors, and empty cells last. TODO(xlcalc): Excel's measured order of mixed values. */
function rank(v: CellValue | undefined): number {
  if (v === null || v === undefined || v === '') return 4
  if (typeof v === 'number') return 0
  if (typeof v === 'string') return 1
  if (typeof v === 'boolean') return 2
  return 3
}

/** Values in the sort's ascending order: by rank, then numbers by size, text by the collator, FALSE before TRUE. */
function compareValues(a: CellValue | undefined, b: CellValue | undefined, collator: Intl.Collator): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'string' && typeof b === 'string') return collator.compare(a, b)
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b)
  return 0
}

/** The sheet's rows in display order: the head as it is, then the rows that meet every condition, sorted when asked (stable, empty cells last). */
export function visibleList(sheet: QuerySheet, filter: SheetFilter, collator: Intl.Collator): number[] {
  const head = headOf(filter, sheet.maxRow)
  const meets = conditions(sheet, filter)
  const body: number[] = []
  for (let r = head + 1; r <= sheet.maxRow; r++) if (meets(r)) body.push(r)
  const sort = filter.sort
  let rows = body
  if (sort) {
    const dir = sort.dir === 'desc' ? -1 : 1
    const keys = body.map((r) => sheet.valueAt(r, sort.col))
    const order = body.map((_, i) => i)
    order.sort((i, j) => {
      const a = keys[i]
      const b = keys[j]
      const blank = (rank(a) === 4 ? 1 : 0) - (rank(b) === 4 ? 1 : 0)
      return blank || dir * compareValues(a, b, collator)
    })
    rows = order.map((i) => body[i]!)
  }
  const out: number[] = []
  for (let r = 1; r <= head; r++) out.push(r)
  return out.concat(rows)
}

/** The funnel's list of a column: its shown texts with their counts over the rows the other columns' conditions keep, '' for empty cells, in the sort's order. */
export function columnList(sheet: QuerySheet, col: number, filter: SheetFilter, collator: Intl.Collator): { values: { text: string; count: number }[]; more: number } {
  const meets = conditions(sheet, filter, col)
  const counts = new Map<string, { count: number; value: CellValue | undefined }>()
  for (let r = headOf(filter, sheet.maxRow) + 1; r <= sheet.maxRow; r++) {
    if (!meets(r)) continue
    const text = sheet.textAt(r, col)
    const got = counts.get(text)
    if (got) got.count++
    else counts.set(text, { count: 1, value: sheet.valueAt(r, col) })
  }
  const all = [...counts.entries()].sort(([ta, a], [tb, b]) => compareValues(a.value, b.value, collator) || collator.compare(ta, tb))
  return { values: all.slice(0, VALUES_MAX).map(([text, { count }]) => ({ text, count })), more: Math.max(0, all.length - VALUES_MAX) }
}
