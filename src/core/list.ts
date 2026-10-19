// List mode: the host's list (a view, a report) on the sheet. Sortable column titles sort on a click through the host,
// drill-down cells and the first column open what they lead to, and a totals row stays at the bottom.
import { computed } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { SheetEmit } from './sheet'
import type { SheetState } from './state'
import type { Windows } from './windows'
import type { ListSort } from '../contract/view'

export function useList(s: SheetState, w: Windows, emit: SheetEmit, props: { sort?: ListSort | null; openRows?: boolean }) {
  const columnOf = (c: number) => s.sheet.value?.columns?.[c - 1]
  const sortable = (c: number) => !!columnOf(c)?.sortable
  /** The arrow a column's title shows: the host's sort direction on it, or none. */
  const sortOf = (c: number): ListSort['dir'] | null => (props.sort?.col === c ? props.sort.dir : null)
  /** A click on a sortable title: ascending first, descending on a column already sorted ascending. */
  function toggleSort(c: number): void {
    emit('sort', { col: c, dir: sortOf(c) === 'asc' ? 'desc' : 'asc' })
  }
  /** The title's alignment, as the host aligns the column's cells. */
  const alignOf = (c: number) => columnOf(c)?.align

  /** Rows a list opens: those after the workbook's frozen rows, which hold its header. */
  const dataRow = (r: number) => r > (s.sheet.value?.freeze?.rows ?? 0)
  /** What a cell leads to, sent to the host: its drill-down, otherwise its row when rows open from any cell (`fromAny`, Enter) or from this one. */
  function open(r: number, c: number, fromAny: boolean): boolean {
    if (w.drillOfRow(r, c)) {
      emit('cellDrill', { sheet: s.sheetName.value, cell: cellAddress(r, c) })
      return true
    }
    if (props.openRows && dataRow(r) && (fromAny || c === 1)) {
      emit('rowOpen', { sheet: s.sheetName.value, row: r })
      return true
    }
    return false
  }
  /** A click on a cell: its drill-down, or its row from the first column. */
  const click = (r: number, c: number): boolean => open(r, c, false)
  /** Enter on the grid: the active cell's drill-down, otherwise its row; true when handled, so the active cell stays. */
  function key(e: KeyboardEvent): boolean {
    const a = s.selected.value
    if (e.key !== 'Enter' || e.shiftKey || e.ctrlKey || e.metaKey || e.altKey || !a) return false
    if (!open(a.row, a.col, true)) return false
    e.preventDefault()
    return true
  }

  /** The totals row: the displayed text for each of the sheet's columns, null when the sheet has none. */
  const totals = computed(() => s.sheet.value?.totals ?? null)

  return { sortable, sortOf, toggleSort, alignOf, click, key, totals }
}

export type List = ReturnType<typeof useList>
