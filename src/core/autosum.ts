// Excel's AutoSum (Alt+=): the source's guess at the active cell opens the edit with `=SUM(range)` (or another function of
// AutoSum's list) and the range selected, so Enter enters it, typing replaces it and the arrows point at another; without a
// range the caret is between the parentheses.
import { computed } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { AutoSumFunction } from '../contract/source'
import type { Editing } from './editing'
import type { SheetState } from './state'

export interface AutoSumDeps {
  /** Whether a cell is being edited: AutoSum starts a new edit only from the grid. */
  editing: () => boolean
}

export function useAutoSum(s: SheetState, ed: Editing, deps: AutoSumDeps) {
  /** Whether AutoSum can run: the source has `autoSum` and the sheet is editable. */
  const available = computed(() => !!s.sourceMethod('autoSum') && s.editable.value)

  /**
   * Opens the active cell's edit with the source's AutoSum, under `fn` when given; false without the source's `autoSum`, with
   * a function it does not list, on a read-only sheet, while a cell is being edited or when the source guesses nothing.
   */
  async function run(fn: AutoSumFunction = 'sum'): Promise<boolean> {
    const autoSum = s.sourceMethod('autoSum')
    const a = s.selected.value
    if (!autoSum || !s.editable.value || !a || deps.editing()) return false
    if (fn !== 'sum' && !s.source().autoSumFunctions?.includes(fn)) return false
    const sheet = s.sheetName.value
    const stale = s.outdated()
    const cell = cellAddress(a.row, a.col)
    const got = await (fn === 'sum' ? autoSum(sheet, cell) : autoSum(sheet, cell, undefined, fn))
    if (!got || stale() || deps.editing() || s.sheetName.value !== sheet || s.selected.value?.row !== a.row || s.selected.value?.col !== a.col) return false
    const text = `=${got.formula}`
    const open = text.indexOf('(') + 1
    const close = text.lastIndexOf(')')
    return ed.startEdit({ text, mode: 'enter', caret: open, caretEnd: close > open ? close : open })
  }

  /** Alt+= (Option+= on a Mac, by the key's position on any layout): true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (!e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || e.code !== 'Equal' || !available.value) return false
    e.preventDefault()
    void run()
    return true
  }

  return { available, run, key }
}

export type AutoSum = ReturnType<typeof useAutoSum>
