// Excel's Data > Filter on the sheet (Ctrl+Shift+L, the ribbon's toggle) through the source's `setAutoFilter`: the AutoFilter
// removed, or put over the range Excel picks from the cells around the selection; and Alt+Down opening a header button's list.
import { computed } from 'vue'
import { cellAddress, rangeAddress } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import type { EditOutcome } from '../contract/source'
import { cellRect } from '../rect'
import { autoFilterRange } from './autofilter-range'
import type { Commands } from './commands'
import type { Filter } from './filter'
import { shortcutLetter } from './keys'
import { readRange } from './range-read'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import { currentRegion } from './tables'

/** The data read around the selection at first, widened while the region found reaches the edge read. */
const REGION_ROWS = 200
const REGION_COLS = 30

export function useAutoFilter(s: SheetState, sel: Selection, cmd: Commands, filter: Filter, notify: (text: string) => void) {
  const read = useSheetReads(s)
  /** Data > Filter is offered: the source has `setAutoFilter` and the sheet is editable. */
  const available = computed(() => !!s.sheet.value && cmd.can('setAutoFilter'))

  /** The range Excel puts the AutoFilter over, from the cells read around the selection; null without data there, undefined when the cells could not be read. */
  async function guess(signal: AbortSignal): Promise<CellRect | null | undefined> {
    const a = s.selected.value!
    const sheet = s.sheetName.value
    const selection = sel.fullRect.value ?? cellRect(a.row, a.col)
    const data = { top: 1, left: 1, bottom: Math.max(1, s.dataRows.value), right: Math.max(1, s.dataCols.value) }
    const single = selection.top === selection.bottom && selection.left === selection.right
    // the cells read cover a single cell's neighbours, or the selection within the data; a selection outside it has none
    const start = single ? selection : { top: selection.top, left: selection.left, bottom: Math.min(selection.bottom, data.bottom), right: Math.min(selection.right, data.right) }
    if (start.top > start.bottom || start.left > start.right) return null
    for (let rows = REGION_ROWS, cols = REGION_COLS; ; rows *= 4, cols *= 4) {
      const area = { top: Math.max(1, start.top - rows), left: Math.max(1, start.left - cols), bottom: Math.min(data.bottom, start.bottom + rows), right: Math.min(data.right, start.right + cols) }
      const cells = await readRange(s.source(), sheet, area, signal)
      if (!cells) return undefined
      const filled = (row: number, col: number) => cells.has(cellAddress(row, col))
      const got = autoFilterRange(selection, area, filled)
      // a region that reaches the edge of what was read may go on beyond it
      const reach = got ? [currentRegion(got, area, filled), ...(single ? [currentRegion({ row: a.row, col: a.col }, area, filled)] : [])] : []
      const edge = reach.some((r) => (r.top === area.top && area.top > 1) || (r.left === area.left && area.left > 1) || (r.bottom === area.bottom && area.bottom < data.bottom) || (r.right === area.right && area.right < data.right))
      if (!edge) return got
    }
  }

  /**
   * Excel's Data > Filter: removes the sheet's AutoFilter (its conditions go when the sheet reloads without it), or puts one
   * over the range Excel picks; without data there, Excel's message and nothing sent.
   */
  async function toggle(): Promise<EditOutcome | null> {
    if (!available.value || !s.selected.value) return null
    const sheet = s.sheetName.value
    if (filter.auto.value) return cmd.setAutoFilter(sheet, null)
    const got = await guess(read())
    if (got === undefined || sheet !== s.sheetName.value) return null
    if (!got) {
      notify(s.messages.value.autoFilterNoRange)
      return null
    }
    // a new AutoFilter starts without conditions
    filter.clear()
    return cmd.setAutoFilter(sheet, rangeAddress(got))
  }

  /** Ctrl+Shift+L (Control on a Mac too, with any layout), and Alt+Down on a header cell of the AutoFilter, which opens its button's list; true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (e.ctrlKey && e.shiftKey && !e.altKey && !e.metaKey && shortcutLetter(e) === 'l' && available.value) {
      e.preventDefault()
      void toggle()
      return true
    }
    const a = s.selected.value
    const r = filter.auto.value
    if (e.key !== 'ArrowDown' || !e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || !a || !r || a.row !== r.top || !filter.canFunnel(a.col)) return false
    e.preventDefault()
    const button = s.scroller.value?.querySelector<HTMLElement>(`[data-xs-af="${a.col}"]`)
    void filter.openFunnel(a.col, (button ?? s.scroller.value)?.getBoundingClientRect())
    return true
  }

  return { available, toggle, key }
}

export type AutoFilter = ReturnType<typeof useAutoFilter>
