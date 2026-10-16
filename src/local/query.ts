// The display filter and sort, the funnel's values and totals by the display filter's rules: conditions
// compare the shown text, the head rows stay on top, a sort puts numbers before text and empty cells last.
import type { CellValue, CellView } from 'xlcalc'
import { parseCell } from 'xlcalc/a1'
import type { RangeStats, SheetFilter } from '../source'
import { addCell, noStats } from '../stats'

/** Distinct values in the funnel's list; the rest are counted in `more`. */
export const COLUMN_VALUES_MAX = 1000

interface Entry {
  text: string
  /** The value when it is a number, else NaN. */
  num: number
}

const numOf = (value: CellValue | undefined): number => (typeof value === 'number' ? value : NaN)

/** A sheet's shown texts by column and row with the numbers among the values, kept current by the changes of edits. */
export class SheetValues {
  readonly cols = new Map<number, Map<number, Entry>>()

  constructor(views: Record<string, CellView> = {}) {
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

/** Totals of the cells on the display filter's rows, as Excel's status bar gives them. */
export function statsOf(views: Record<string, CellView>, visible: Set<number>): RangeStats {
  const out = noStats()
  for (const [addr, v] of Object.entries(views)) {
    const at = parseCell(addr)
    if (at && visible.has(at.row)) addCell(out, v.text ?? '', v.value)
  }
  return out
}
