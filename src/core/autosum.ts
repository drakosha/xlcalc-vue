// Excel's AutoSum (Alt+=): the source's guess at the active cell opens the edit with `=SUM(range)` and the range selected,
// so Enter enters it, typing replaces it and the arrows point at another; without a range the caret is between the parentheses.
import { cellAddress } from 'xlcalc/a1'
import type { Editing } from './editing'
import type { SheetState } from './state'

export interface AutoSumDeps {
  /** Whether a cell is being edited: AutoSum starts a new edit only from the grid. */
  editing: () => boolean
}

export function useAutoSum(s: SheetState, ed: Editing, deps: AutoSumDeps) {
  /** Opens the active cell's edit with the source's AutoSum; false without the source's `autoSum`, on a read-only sheet, while a cell is being edited or when the source guesses nothing. */
  async function run(): Promise<boolean> {
    const autoSum = s.sourceMethod('autoSum')
    const a = s.selected.value
    if (!autoSum || !s.editable.value || !a || deps.editing()) return false
    const sheet = s.sheetName.value
    const stale = s.outdated()
    const got = await autoSum(sheet, cellAddress(a.row, a.col))
    if (!got || stale() || deps.editing() || s.sheetName.value !== sheet || s.selected.value?.row !== a.row || s.selected.value?.col !== a.col) return false
    const text = `=${got.formula}`
    const open = text.indexOf('(') + 1
    const close = text.lastIndexOf(')')
    return ed.startEdit({ text, mode: 'enter', caret: open, caretEnd: close > open ? close : open })
  }

  /** Alt+= (Option+= on a Mac, by the key's position on any layout): true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (!e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || e.code !== 'Equal' || !s.sourceMethod('autoSum') || !s.editable.value) return false
    e.preventDefault()
    void run()
    return true
  }

  return { run, key }
}

export type AutoSum = ReturnType<typeof useAutoSum>
