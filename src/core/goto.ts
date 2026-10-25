// Excel's Go To (F5, Ctrl+G): the dialog's state, the locations it has left, newest first, and the jump with Excel's message for a bad reference.
import { computed, ref } from 'vue'
import { quoteSheetIfNeeded } from 'xlcalc/a1'
import { shortcutLetter } from './keys'
import type { AlertAnswer, SheetAlert } from './validation'
import type { Navigation } from './navigation'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** How many locations the Go To list keeps, as Excel's does. */
export const GOTO_PLACES = 4

/** A reference made absolute part by part, as Excel lists locations in Go To: `B2:D5` → `$B$2:$D$5`, `B:D` → `$B:$D`. */
export function absoluteRef(ref: string): string {
  return ref
    .split(':')
    .map((part) => part.replace(/^([A-Z]*)(\d*)$/i, (_, col: string, row: string) => `${col ? `$${col}` : ''}${row ? `$${row}` : ''}`))
    .join(':')
}

export interface GoToDeps {
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
}

export function useGoTo(s: SheetState, sel: Selection, nav: Navigation, deps: GoToDeps) {
  const open = ref(false)
  /** The locations jumps from Go To and from hyperlinks left, newest first: the first one fills the Reference box, so F5 and Enter go back. */
  const places = ref<string[]>([])
  /** The workbook's defined names, which the list shows after the locations, alphabetically; none until the source lists them. */
  const names = computed<string[]>(() => [])
  const items = computed(() => [...places.value, ...[...names.value].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }))])

  /** The selection as a location to come back to: `Sheet1!$B$2:$D$5`. */
  function here(): string | null {
    const ref = sel.selectionRef.value
    return s.sheet.value && ref ? `${quoteSheetIfNeeded(s.sheetName.value)}!${absoluteRef(ref)}` : null
  }
  /** Goes to a location as the Name Box does and, when it went there, records the location it left; false for text that is not one. */
  async function jump(text: string): Promise<boolean> {
    const from = here()
    if (!(await nav.goPlace(text))) return false
    if (from) places.value = [from, ...places.value.filter((p) => p.toUpperCase() !== from.toUpperCase())].slice(0, GOTO_PLACES)
    return true
  }

  /** Opens the dialog; false while no sheet is displayed. */
  function openGoTo(): boolean {
    if (!s.sheet.value) return false
    open.value = true
    return true
  }
  /** The dialog's OK: the jump, and the dialog closes; Excel's alert for text that is not a reference, after which the dialog stays. */
  async function go(text: string): Promise<boolean> {
    if (await jump(text)) {
      open.value = false
      return true
    }
    await deps.ask({ severity: 'notice', message: s.messages.value.goToBad })
    return false
  }
  function close(): void {
    open.value = false
    s.focusGrid()
  }
  /** Excel's F5 and Ctrl+G (Control+G on a Mac too, where Cmd+G finds again): true when handled. */
  function key(e: KeyboardEvent): boolean {
    const f5 = e.key === 'F5' && !e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey
    const g = e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey && shortcutLetter(e) === 'g'
    if (!(f5 || g)) return false
    e.preventDefault()
    openGoTo()
    return true
  }

  return { open, places, items, jump, openGoTo, go, close, key }
}

export type GoTo = ReturnType<typeof useGoTo>
