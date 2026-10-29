// Excel's Create from Selection (Formulas > Defined Names, Ctrl+Shift+F3): the Create Names dialog with the boxes Excel for
// Mac checks, the names the source makes (`namesFromSelection`), Excel's question before a name's other definition is
// replaced, and one call to `setNames` for all of them, a single undo step.
import { computed, ref } from 'vue'
import { cellAddress, rangeAddress } from 'xlcalc/a1'
import type { CellRect, SelectionLabels } from 'xlcalc/contract'
import { onMac } from './keys'
import { nameKey, type Names } from './names'
import { readRange } from './range-read'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import type { AlertAnswer, SheetAlert } from './validation'

export type Labels = Required<SelectionLabels>

/** The dialog's state: the selection it was opened for, and its boxes. */
export interface CreateNamesDraft {
  sheet: string
  rect: CellRect
  labels: Labels
}

/**
 * The boxes Excel for Mac checked for a selection (17 ranges measured): Top row when the selection has two rows or more and
 * the cell right of its top-left one (the top-left one itself in a single column) holds text, a formula's included; Left
 * column likewise by the cell below the top-left one (the top-left one in a single row); never Bottom row or Right column.
 */
export function guessLabels(r: CellRect, isText: (row: number, col: number) => boolean): Labels {
  const rows = r.bottom - r.top + 1
  const cols = r.right - r.left + 1
  return { top: rows > 1 && isText(r.top, cols > 1 ? r.left + 1 : r.left), left: cols > 1 && isText(rows > 1 ? r.top + 1 : r.top, r.left), bottom: false, right: false }
}

/** Labels on these sides leave no cell past them: Excel says the selection is not valid. */
export function labelsAlone(r: CellRect, l: SelectionLabels): boolean {
  const rows = r.bottom - r.top + 1 - (l.top ? 1 : 0) - (l.bottom ? 1 : 0)
  const cols = r.right - r.left + 1 - (l.left ? 1 : 0) - (l.right ? 1 : 0)
  return rows < 1 || cols < 1
}

export interface CreateNamesDeps {
  editing: () => boolean
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
  notify: (text: string) => void
}

export function useCreateNames(s: SheetState, sel: Selection, names: Names, deps: CreateNamesDeps) {
  const read = useSheetReads(s)
  const draft = ref<CreateNamesDraft | null>(null)
  /** Create from Selection is offered: names can be changed (`names`, `setNames`, an editable sheet) and the source makes them. */
  const available = computed(() => names.editable.value && typeof s.source().namesFromSelection === 'function')

  /** Opens the dialog for the selection, once its two cells that decide the boxes are read; false when it is not offered or a cell is being edited. */
  function open(): boolean {
    const rect = sel.fullRect.value
    if (!available.value || !s.sheet.value || !rect || deps.editing()) return false
    const sheet = s.sheetName.value
    const probe = { top: rect.top, left: rect.left, bottom: Math.min(rect.bottom, rect.top + 1), right: Math.min(rect.right, rect.left + 1) }
    // A sheet switch (which aborts the read) or a new workbook while the cells are read: the dialog does not open.
    const signal = read()
    const outdated = s.outdated()
    void readRange(s.source(), sheet, probe, signal).then((cells) => {
      if (signal.aborted || outdated()) return
      const isText = (row: number, col: number) => {
        const v = cells?.get(cellAddress(row, col))?.value
        return typeof v === 'string' && v !== ''
      }
      draft.value = { sheet, rect, labels: guessLabels(rect, isText) }
    })
    return true
  }
  function close(): void {
    draft.value = null
    s.focusGrid()
  }

  /**
   * OK: the dialog closes and the names are made, as Excel makes them one by one: a name the workbook defines otherwise brings
   * Excel's question, Yes replacing it, No keeping it and Cancel ending there with what was decided before; a name with the
   * same definition is left as it is. Labels alone bring Excel's message instead. All go to `setNames` in one call.
   */
  async function submit(labels: Labels): Promise<void> {
    const d = draft.value
    const call = s.sourceMethod('namesFromSelection')
    close()
    if (!d || !call) return
    const m = s.messages.value
    if (labelsAlone(d.rect, labels)) {
      await deps.ask({ severity: 'notice', message: m.createNamesInvalid })
      return
    }
    let made
    try {
      made = await call(d.sheet, rangeAddress(d.rect), labels, read())
    } catch (e) {
      return deps.notify(e instanceof Error ? e.message : String(e))
    }
    const list = [...names.list.value]
    let changed = false
    for (const n of made) {
      const i = list.findIndex((x) => nameKey(x) === nameKey(n))
      if (i >= 0 && list[i]!.ref === n.ref) continue
      if (i >= 0) {
        const said = await deps.ask({ severity: 'question', message: m.createNamesReplace(n.name) })
        if (said === 'cancel') break
        if (said === 'retry') continue
        list[i] = n
      } else list.push(n)
      changed = true
    }
    if (!changed) return
    const out = await names.save(list)
    if (out && !out.ok) deps.notify(out.message || m.editFailed)
  }

  /** Ctrl+Shift+F3, and Cmd+Shift+F3 on a Mac as in Excel for Mac: true when handled. */
  function key(e: KeyboardEvent): boolean {
    const mod = (e.ctrlKey && !e.metaKey) || (onMac() && e.metaKey && !e.ctrlKey)
    if (e.key !== 'F3' || !e.shiftKey || e.altKey || !mod || !open()) return false
    e.preventDefault()
    return true
  }

  return { draft, available, open, close, submit, key }
}

export type CreateNames = ReturnType<typeof useCreateNames>
