// Excel's calculation options on the sheet: the source's mode and whether cells wait (the status bar's Calculate), F9 and
// Shift+F9, the host's Calculation Options, and Data > Refresh All (Ctrl+Alt+F5); the source keeps the values shown as of
// the last calculation.
import { computed, shallowRef } from 'vue'
import type { CalculationMode, EditOutcome, SheetCalculation } from '../contract/source'
import type { Editing } from './editing'
import type { SheetState } from './state'

export function useCalculation(s: SheetState, ed: Editing) {
  /** The source's calculation as last read or told; null for a source without `calculation`. */
  const state = shallowRef<SheetCalculation | null>(null)
  let readSeq = 0

  /** Reads the calculation with the sheets (a new workbook, a reload); a later read or event wins over an earlier answer. */
  async function read(): Promise<void> {
    const mine = ++readSeq
    const outdated = s.outdated()
    const call = s.sourceMethod('calculation')
    if (!call) {
      state.value = null
      return
    }
    try {
      const got = await call()
      if (mine === readSeq && !outdated()) state.value = got
    } catch {
      // a source that cannot tell its calculation leaves the status bar as it was
    }
  }
  /** The source's `calculation` event. */
  function told(next: SheetCalculation): void {
    readSeq++
    state.value = next
  }

  const pending = computed(() => state.value?.pending === true)
  /** The commands need an editable sheet, as the other changes do. */
  const method = <K extends 'calculate' | 'setCalculation' | 'refreshConnections'>(name: K) => (s.editable.value ? s.sourceMethod(name) : null)
  const canCalculate = computed(() => method('calculate') !== null)
  const canRefresh = computed(() => method('refreshConnections') !== null)

  /** Excel's Calculate Now, or Calculate Sheet on the active sheet; copy mode stays on. */
  async function calculate(scope: 'workbook' | 'sheet' = 'workbook'): Promise<EditOutcome | null> {
    const call = method('calculate')
    if (!call) return null
    const sheet = s.sheetName.value
    return ed.run(() => (scope === 'sheet' && sheet ? call(sheet) : call()), [], null, null, false)
  }
  async function setMode(mode: CalculationMode): Promise<EditOutcome | null> {
    const call = method('setCalculation')
    return call ? ed.run(() => call(mode), [], null, null, false) : null
  }

  /** Excel's Refresh All: every call outside the workbook asked again, not a step of undo; copy mode stays on. */
  async function refreshAll(): Promise<EditOutcome | null> {
    const call = method('refreshConnections')
    return call ? ed.run(() => call(), [], null, null, false) : null
  }

  /** F9 calculates the workbook and Shift+F9 the active sheet, as in Excel for Windows and for Mac; Ctrl+Alt+F5 refreshes all, as in Excel for Windows; true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (e.key === 'F5' && e.ctrlKey && e.altKey && !e.shiftKey && !e.metaKey && canRefresh.value) {
      e.preventDefault()
      void refreshAll()
      return true
    }
    if (e.key !== 'F9' || e.ctrlKey || e.altKey || e.metaKey || !canCalculate.value) return false
    e.preventDefault()
    void calculate(e.shiftKey ? 'sheet' : 'workbook')
    return true
  }

  return { state, pending, canCalculate, canRefresh, read, told, calculate, setMode, refreshAll, key }
}

export type Calculation = ReturnType<typeof useCalculation>
