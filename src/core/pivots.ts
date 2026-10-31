// PivotTables on the sheet: the PivotTable under the active cell (`SheetInfo.pivots`), Excel's Refresh (Alt+F5) of it
// through the source's `structure`, Excel's Field List (the PivotTable Fields pane shown while the active cell is in a
// PivotTable), the dialogs the pane asks for, and their items in the cell menu. What a PivotTable's cells take is the source's to say.
import { computed, ref } from 'vue'
import { parseRange, rectContains } from 'xlcalc/a1'
import type { EditOutcome } from '../contract/source'
import type { PivotDialogRequest } from '../contract/view'
import { cellRect } from '../rect'
import type { Commands } from './commands'
import type { MenuItem } from './menu'
import type { Protection } from './protection'
import type { SheetState } from './state'

export interface PivotsDeps {
  /** Sends a dialog request to the host (`SheetView`'s `pivotDialog` event). */
  dialog: (request: PivotDialogRequest) => void
}

export function usePivots(s: SheetState, cmd: Commands, protection: Protection, deps: PivotsDeps) {
  /** The PivotTable whose table or Filters area holds the active cell; null when there is none. */
  const active = computed(() => {
    const a = s.selected.value
    const pivots = s.sheet.value?.pivots
    if (!a || !pivots?.length) return null
    const here = cellRect(a.row, a.col)
    const holds = (range: string | undefined) => {
      const r = range ? parseRange(range) : null
      return !!r && rectContains(r, here)
    }
    return pivots.find((p) => holds(p.range) || holds(p.filters)) ?? null
  })
  /** Refresh needs the source's `structure` on an editable sheet; a protected sheet's refusal is the source's. */
  const canRefresh = computed(() => !!active.value && cmd.can('structure'))

  /** Excel's Field List: on by default, as Excel's; the pane shows while it is on and the active cell is in a PivotTable the source can read. */
  const fieldListOn = ref(true)
  const canRead = computed(() => !!s.sourceMethod('pivotTable'))
  const fieldListShown = computed(() => fieldListOn.value && !!active.value && canRead.value)
  function fieldList(show?: boolean): boolean {
    fieldListOn.value = show ?? !fieldListOn.value
    return fieldListOn.value
  }
  /** A dialog over a PivotTable's field that the pane asks for, sent to the host; false on a read-only sheet. */
  function dialog(request: PivotDialogRequest): boolean {
    if (!s.editable.value) return false
    deps.dialog(request)
    return true
  }

  /** Excel's Refresh: the PivotTable under the active cell, with every PivotTable over its cache, as one undo step. */
  async function refresh(): Promise<EditOutcome | null> {
    const p = active.value
    if (!p || !canRefresh.value) return null
    return cmd.structure({ op: 'refreshPivots', pivots: [{ sheet: s.sheetName.value, name: p.name }] })
  }

  /** Alt+F5 refreshes the PivotTable under the active cell, as in Excel for Windows; true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (e.key !== 'F5' || !e.altKey || e.ctrlKey || e.shiftKey || e.metaKey || !canRefresh.value) return false
    e.preventDefault()
    void refresh()
    return true
  }

  /** The cell menu's Refresh on a PivotTable's cell, disabled where the sheet's protection does not allow PivotTables. */
  function items(): MenuItem[] {
    if (!canRefresh.value) return []
    return [{ id: 'refreshPivot', label: s.messages.value.menuRefresh, disabled: !protection.can('pivotTables'), run: () => void refresh(), group: true }]
  }
  /** Show Field List or Hide Field List, which Excel lists last in the cell menu of a PivotTable's cell. */
  function lastItems(): MenuItem[] {
    if (!active.value || !canRead.value) return []
    const m = s.messages.value
    return [{ id: 'fieldList', label: fieldListOn.value ? m.menuHideFieldList : m.menuShowFieldList, run: () => void fieldList(), group: true }]
  }

  return { active, canRefresh, refresh, key, items, lastItems, fieldList, fieldListShown, dialog }
}

export type Pivots = ReturnType<typeof usePivots>
