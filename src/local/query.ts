// Search, the display filter and sort, the funnel's values and totals, by the the app's rules (`sheet-filter.ts`):
// conditions compare the shown text, the head rows stay on top, a sort puts numbers before text and empty cells last.
import type { CellValue, CellView } from 'xlcalc'
import { parseCell } from '../core/a1'
import type { RangeStats, SheetFilter } from '../source'

/** Hits of a search in one answer; the total counts them all. */
export const FIND_MAX_HITS = 500
/** Distinct values in the funnel's list; the rest are counted in `more`. */
export const COLUMN_VALUES_MAX = 1000

interface Entry {
  text: string
  /** The value when it is a number, else NaN. */
  num: number
}

const numOf = (value: CellValue | undefined): number => (typeof value === 'number' ? value : NaN)

/** A sheet's shown text and numbers by column and row, kept current by the changes of edits. */
export class SheetValues {
  readonly cols = new Map<number, Map<number, Entry>>()

  constructor(views: Record<string, CellView>) {
    for (const [addr, v] of Object.entries(views)) {
      const at = parseCell(addr)
      if (at) this.set(at.row, at.col, v.text ?? '', v.value)
    }
  }

  set(row: number, col: number, text: string, value: CellValue | undefined): void {
    let column = this.cols.get(col)
    if (text === '') {
      column?.delete(row)
      return
    }
    if (!column) this.cols.set(col, (column = new Map()))
    column.set(row, { text, num: numOf(value) })
  }
}

/** Whether a filter asks for anything: conditions on columns or a sort. */
export function filterActive(filter: SheetFilter | undefined): filter is SheetFilter {
  return !!filter && (!!filter.sort || Object.values(filter.cols ?? {}).some((c) => !!c.values || !!c.text?.trim()))
}

/** A stable key of a filter, for remembering its list of rows. */
export const filterKey = (f: SheetFilter): string =>
  JSON.stringify([f.head, Object.entries(f.cols ?? {}).sort(([a], [b]) => Number(a) - Number(b)), f.sort ?? null])

/** Text first by numbers' value, then text by the collator, the empty text last. */
function compareEntries(a: Entry | undefined, b: Entry | undefined, collator: Intl.Collator): number {
  if (!a && !b) return 0
  if (!a) return 1
  if (!b) return -1
  const an = !Number.isNaN(a.num)
  const bn = !Number.isNaN(b.num)
  if (an && bn) return a.num - b.num
  if (an !== bn) return an ? -1 : 1
  return collator.compare(a.text, b.text)
}

/** The sheet's rows in display order: the head as it is, then the rows that pass every condition, sorted when asked. */
export function visibleList(values: SheetValues, filter: SheetFilter, maxRow: number, collator: Intl.Collator): number[] {
  const head = Math.min(Math.max(0, Math.trunc(filter.head)), maxRow)
  const conds = Object.entries(filter.cols ?? {}).map(([col, c]) => ({
    column: values.cols.get(Number(col)),
    values: c.values ? new Set(c.values) : null,
    text: c.text?.trim() ? c.text.trim().toLocaleLowerCase() : null,
  }))
  const body: number[] = []
  for (let r = head + 1; r <= maxRow; r++) {
    const pass = conds.every((c) => {
      const text = c.column?.get(r)?.text ?? ''
      return (!c.values || c.values.has(text)) && (!c.text || text.toLocaleLowerCase().includes(c.text))
    })
    if (pass) body.push(r)
  }
  const sort = filter.sort
  if (sort) {
    const column = values.cols.get(sort.col)
    const sign = sort.dir === 'desc' ? -1 : 1
    // Empty cells stay last either way; equal keys keep the sheet's order (the sort is stable).
    body.sort((a, b) => {
      const ea = column?.get(a)
      const eb = column?.get(b)
      if (!ea || !eb) return compareEntries(ea, eb, collator)
      return sign * compareEntries(ea, eb, collator)
    })
  }
  const out: number[] = []
  for (let r = 1; r <= head; r++) out.push(r)
  return out.concat(body)
}

/** The funnel's list of a column: its shown texts with their counts over the rows the other columns' conditions keep. */
export function columnValuesOf(values: SheetValues, col: number, filter: SheetFilter, maxRow: number, collator: Intl.Collator): { values: { text: string; count: number }[]; more: number } {
  const head = Math.min(Math.max(0, Math.trunc(filter.head)), maxRow)
  const others = Object.fromEntries(Object.entries(filter.cols ?? {}).filter(([k]) => Number(k) !== col))
  const rows = visibleList(values, { head, cols: others }, maxRow, collator).slice(head)
  const column = values.cols.get(col)
  const counts = new Map<string, { count: number; entry: Entry | undefined }>()
  for (const r of rows) {
    const entry = column?.get(r)
    const text = entry?.text ?? ''
    const hit = counts.get(text)
    if (hit) hit.count++
    else counts.set(text, { count: 1, entry })
  }
  const all = [...counts.entries()].sort(([, a], [, b]) => compareEntries(a.entry, b.entry, collator))
  const shown = all.slice(0, COLUMN_VALUES_MAX)
  return { values: shown.map(([text, v]) => ({ text, count: v.count })), more: all.length - shown.length }
}

/** Cells whose shown text holds the query, without regard to case, in row order; only `visible` rows when given. */
export function findIn(values: SheetValues, query: string, visible: Set<number> | null): { hits: { row: number; col: number; text: string }[]; total: number } {
  const q = query.trim().toLocaleLowerCase()
  if (!q) return { hits: [], total: 0 }
  const found: { row: number; col: number; text: string }[] = []
  for (const [col, column] of values.cols) {
    for (const [row, e] of column) {
      if (visible && !visible.has(row)) continue
      if (e.text.toLocaleLowerCase().includes(q)) found.push({ row, col, text: e.text })
    }
  }
  found.sort((a, b) => a.row - b.row || a.col - b.col)
  return { hits: found.slice(0, FIND_MAX_HITS), total: found.length }
}

/** Totals of cells as Excel's status bar gives them: cells that show something, the numbers among them, their sum and ends. */
export function statsOf(views: Record<string, CellView>, visible: Set<number> | null): RangeStats {
  let cells = 0
  let numbers = 0
  let sum = 0
  let min = Infinity
  let max = -Infinity
  for (const [addr, v] of Object.entries(views)) {
    if ((v.text ?? '') === '') continue
    if (visible) {
      const at = parseCell(addr)
      if (!at || !visible.has(at.row)) continue
    }
    cells++
    if (typeof v.value !== 'number') continue
    numbers++
    sum += v.value
    if (v.value < min) min = v.value
    if (v.value > max) max = v.value
  }
  return { cells, numbers, sum, min: numbers ? min : null, max: numbers ? max : null }
}
