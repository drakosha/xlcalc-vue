// Jumps: the address box, references of the formula, links and `focus`; the history to go back, and the flash after a jump.
import { nextTick, ref, watch } from 'vue'
import { EXCEL_MAX_COLS, EXCEL_MAX_ROWS, parseCellRef, sheetRefName, type FormulaRef } from './formula-refs'
import type { Formula } from './formula'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Rect } from './a1'

/** A place to come back to: sheet, selection and scroll. */
export interface Place {
  sheet: number
  selected: SheetState['selected']['value']
  extent: SheetState['extent']['value']
  scrollTop: number
  scrollLeft: number
}

const HISTORY_MAX = 50
/** How long the range just jumped to flashes, so the eye finds it. */
const FLASH_MS = 2000

export interface NavigationDeps {
  pickSheet: (index: number) => Promise<void>
  notify: (text: string) => void
}

export function useNavigation(s: SheetState, sel: Selection, f: Formula, deps: NavigationDeps) {
  const history = ref<Place[]>([])
  const flash = ref<Rect | null>(null)
  let flashTimer: ReturnType<typeof setTimeout> | null = null

  /** Remember where we stand before a jump. */
  function rememberPlace(): void {
    history.value = [
      ...history.value.slice(-(HISTORY_MAX - 1)),
      {
        sheet: s.active.value,
        selected: s.selected.value,
        extent: s.extent.value,
        scrollTop: s.scroller.value?.scrollTop ?? 0,
        scrollLeft: s.scroller.value?.scrollLeft ?? 0,
      },
    ]
  }
  /** Back to where the last jump came from: sheet, selection and scroll as they were. */
  async function goBack(): Promise<void> {
    const place = history.value[history.value.length - 1]
    if (!place) return
    history.value = history.value.slice(0, -1)
    if (place.sheet !== s.active.value && s.sheets.value[place.sheet]) await deps.pickSheet(place.sheet)
    s.selected.value = place.selected
    s.extent.value = place.extent
    await nextTick()
    s.scroller.value?.scrollTo({ top: place.scrollTop, left: place.scrollLeft })
    s.scrollTop.value = place.scrollTop
    s.scrollLeft.value = place.scrollLeft
    s.scroller.value?.focus({ preventScroll: true })
  }

  function flashRange(range: Rect): void {
    if (flashTimer) clearTimeout(flashTimer)
    // Reset and set in different ticks, so a second jump to the same place restarts the animation.
    flash.value = null
    void nextTick(() => {
      flash.value = range
      flashTimer = setTimeout(() => (flash.value = null), FLASH_MS)
    })
  }
  /** Select a rectangle (cut to the shown grid), show its top-left corner and flash it. */
  function selectRange(range: Rect): void {
    // A whole column or row runs to Excel's limit: it stops at the grid, which an editable sheet stretches to the place.
    s.stretch(range.r2 >= EXCEL_MAX_ROWS ? range.r1 : range.r2, range.c2 >= EXCEL_MAX_COLS ? range.c1 : range.c2)
    const rows = Math.max(1, s.rowCount.value)
    const cols = Math.max(1, s.colCount.value)
    const r1 = Math.min(range.r1, rows)
    const c1 = Math.min(range.c1, cols)
    const r2 = Math.min(range.r2, rows)
    const c2 = Math.min(range.c2, cols)
    s.selected.value = { row: r1, col: c1 }
    s.extent.value = r1 === r2 && c1 === c2 ? null : { row: r2, col: c2 }
    flashRange({ r1, c1, r2, c2 })
    void nextTick(() => {
      sel.revealCell(r1, c1)
      s.scroller.value?.focus({ preventScroll: true })
    })
  }

  /** Go to a place: the sheet when named, and the range; false when the text is not an address or there is no such sheet. */
  async function goPlace(text: string): Promise<boolean> {
    const place = parseCellRef(text)
    if (!place) return false
    let index = s.active.value
    if (place.sheet !== null) {
      index = f.sheetIndexByName(place.sheet)
      if (index < 0) return false
    }
    // The book's first frame (A1 unscrolled, a link from outside) has nothing to go back to.
    const a = s.selected.value
    if (history.value.length || s.scrollTop.value || s.scrollLeft.value || s.extent.value || (a && (a.row !== 1 || a.col !== 1))) rememberPlace()
    if (index !== s.active.value) await deps.pickSheet(index)
    if (place.range) selectRange(place.range)
    return true
  }
  /** A reference clicked in the formula bar: select its range, switching sheets when it names another. */
  async function goRef(r: FormulaRef): Promise<void> {
    const index = f.refSheet(r)
    if (index < 0) return
    rememberPlace()
    if (index !== s.active.value) await deps.pickSheet(index)
    selectRange(r)
  }

  // ── the address box ──
  const addressDraft = ref('')
  const addressFocused = ref(false)
  watch(sel.selectionRef, (v) => {
    if (!addressFocused.value) addressDraft.value = v
  })
  async function goAddress(): Promise<boolean> {
    const ok = await goPlace(addressDraft.value)
    if (!ok) deps.notify(s.messages.value.addressBad)
    addressDraft.value = sel.selectionRef.value
    return ok
  }

  /** A place asked for from outside before the book loaded; taken once the sheets are in. */
  let pendingFocus: string | null = null
  watch(
    () => s.props.focus,
    (focus) => {
      if (!focus) return
      if (s.sheets.value.length && s.active.value >= 0) void goPlace(focus.ref)
      else pendingFocus = focus.ref
    },
    { immediate: true },
  )
  /** The focus waiting for the book, if any; it is taken once. */
  function takePendingFocus(): string | null {
    const f = pendingFocus
    pendingFocus = null
    return f
  }

  /** The host's page address of the selected place, to the clipboard. */
  async function copyCellLink(): Promise<void> {
    const url = s.props.cellUrl
    if (!url || !s.sheet.value || !sel.selectionRef.value) return
    const place = `${sheetRefName(s.sheet.value.name)}!${sel.selectionRef.value}`
    try {
      await navigator.clipboard.writeText(url(place))
      deps.notify(s.messages.value.cellLinkCopied(place))
    } catch {
      deps.notify(s.messages.value.clipboardDenied)
    }
  }

  return { history, flash, goBack, goPlace, goRef, selectRange, addressDraft, addressFocused, goAddress, takePendingFocus, copyCellLink }
}

export type Navigation = ReturnType<typeof useNavigation>
