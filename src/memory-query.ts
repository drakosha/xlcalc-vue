// Queries of the memory source over a sheet's shown text: the display filter and sort, search, a column's values, totals.
import type { CellRect, CellValue } from 'xlcalc/contract'
import type { RangeStats, SheetFilter } from './source'

/** What the queries read of a sheet. */
export interface QuerySheet {
  maxRow: number
  maxCol: number
  textAt(row: number, col: number): string
  valueAt(row: number, col: number): CellValue | undefined
}

/** Values a column's list gives at most; the rest is counted in `more`. */
const VALUES_MAX = 1000
/** Hits a search gives at most; `total` counts all. */
const HITS_MAX = 1000

const headOf = (sheet: QuerySheet, filter: SheetFilter) => Math.min(Math.max(0, filter.head), Math.max(0, sheet.maxRow - 1))

/** Whether a row meets the column conditions (one column left out: its own list shows the rows the others leave). */
function meets(sheet: QuerySheet, row: number, filter: SheetFilter, except?: number): boolean {
  for (const [key, cond] of Object.entries(filter.cols ?? {})) {
    const col = Number(key)
    if (col === except) continue
    const text = sheet.textAt(row, col)
    if (cond.values && !cond.values.includes(text)) return false
    if (cond.text && !text.toLocaleLowerCase().includes(cond.text.toLocaleLowerCase())) return false
  }
  return true
}

/** Excel's sort order: numbers, then text, then booleans, then errors; empty cells last whichever way. */
function rank(v: CellValue | undefined): number {
  if (v === null || v === undefined || v === '') return 4
  if (typeof v === 'number') return 0
  if (typeof v === 'string') return 1
  if (typeof v === 'boolean') return 2
  return 3
}
function compare(a: CellValue | undefined, b: CellValue | undefined): number {
  const ra = rank(a)
  const rb = rank(b)
  if (ra !== rb) return ra - rb
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b)
  return 0
}

/** The sheet's rows in display order: the header first, then the rows that meet the conditions, sorted; null without conditions. */
export function visibleList(sheet: QuerySheet, filter: SheetFilter | undefined): number[] | null {
  if (!filter || (!Object.keys(filter.cols ?? {}).length && !filter.sort)) return null
  const head = headOf(sheet, filter)
  const rows: number[] = []
  for (let r = head + 1; r <= sheet.maxRow; r++) if (meets(sheet, r, filter)) rows.push(r)
  const sort = filter.sort
  if (sort) {
    const dir = sort.dir === 'asc' ? 1 : -1
    rows.sort((a, b) => {
      const va = sheet.valueAt(a, sort.col)
      const vb = sheet.valueAt(b, sort.col)
      const blank = (rank(va) === 4 ? 1 : 0) - (rank(vb) === 4 ? 1 : 0)
      return blank || dir * compare(va, vb) || a - b
    })
  }
  return [...Array.from({ length: head }, (_, i) => i + 1), ...rows]
}

/** A substring in any case, hits in row order (only the shown rows under a filter). */
export function findHits(sheet: QuerySheet, query: string, filter: SheetFilter | undefined): { hits: { row: number; col: number; text: string }[]; total: number } {
  const q = query.trim().toLocaleLowerCase()
  const list = visibleList(sheet, filter)
  const rows = list ? [...list].sort((a, b) => a - b) : Array.from({ length: sheet.maxRow }, (_, i) => i + 1)
  const hits: { row: number; col: number; text: string }[] = []
  let total = 0
  if (!q) return { hits, total }
  for (const row of rows) {
    for (let col = 1; col <= sheet.maxCol; col++) {
      const text = sheet.textAt(row, col)
      if (!text.toLocaleLowerCase().includes(q)) continue
      total++
      if (hits.length < HITS_MAX) hits.push({ row, col, text })
    }
  }
  return { hits, total }
}

/** A column's shown values with their counts, over the rows the other columns' conditions leave; '' is an empty cell. */
export function columnList(sheet: QuerySheet, col: number, filter: SheetFilter): { values: { text: string; count: number }[]; more: number } {
  const counts = new Map<string, { count: number; value: CellValue | undefined }>()
  for (let r = headOf(sheet, filter) + 1; r <= sheet.maxRow; r++) {
    if (!meets(sheet, r, filter, col)) continue
    const text = sheet.textAt(r, col)
    const got = counts.get(text)
    if (got) got.count++
    else counts.set(text, { count: 1, value: sheet.valueAt(r, col) })
  }
  const all = [...counts.entries()].sort((a, b) => compare(a[1].value, b[1].value) || a[0].localeCompare(b[0]))
  return { values: all.slice(0, VALUES_MAX).map(([text, { count }]) => ({ text, count })), more: Math.max(0, all.length - VALUES_MAX) }
}

/** Totals of a rectangle as Excel's status line, over the shown rows under a filter. */
export function rangeStats(sheet: QuerySheet, rect: CellRect, filter: SheetFilter | undefined): RangeStats {
  const list = visibleList(sheet, filter)
  const shown = list ? new Set(list) : null
  const out: RangeStats = { cells: 0, numbers: 0, sum: 0, min: null, max: null }
  for (let r = rect.top; r <= Math.min(rect.bottom, sheet.maxRow); r++) {
    if (shown && !shown.has(r)) continue
    for (let c = rect.left; c <= Math.min(rect.right, sheet.maxCol); c++) {
      const v = sheet.valueAt(r, c)
      if (sheet.textAt(r, c) === '' && (v === null || v === undefined || v === '')) continue
      out.cells++
      if (typeof v !== 'number') continue
      out.numbers++
      out.sum += v
      out.min = out.min === null ? v : Math.min(out.min, v)
      out.max = out.max === null ? v : Math.max(out.max, v)
    }
  }
  return out
}
