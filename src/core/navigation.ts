// Jumps: the address box, references of the formula, links and `focus`; the history to go back, and the flash after a jump.
import { nextTick, onScopeDispose, ref, watch } from 'vue'
import { MAX_COLS, MAX_ROWS } from 'xlcalc/a1'
import type { CellRect, FormulaReference, GridReference } from 'xlcalc/contract'
import { quoteSheetIfNeeded, scanReferences } from 'xlcalc/text'
import { isGridRef } from '../grid-ref'
import type { Formula } from './formula'
import type { Selection } from './selection'
import type { SheetState } from './state'

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

/** A place from the address box, a link or `focus`: a range with its sheet (`B12`, `'Sheet 2'!C7:D9`, `B:B`) or a sheet alone; null for anything else (`B12+C1`). */
function parsePlace(text: string): { sheet?: string; range?: CellRect } | null {
  const raw = text.trim()
  if (!raw) return null
  const whole = (refs: FormulaReference[], end: number): GridReference | null => {
    const r = refs.length === 1 ? refs[0]! : null
    return r && isGridRef(r) && r.at === 0 && r.end === end && !r.spill ? r : null
  }
  const refs = scanReferences(raw)
  if (refs.length) {
    const r = whole(refs, raw.length)
    return r && { sheet: r.sheet, range: r }
  }
  // A sheet alone reads as the sheet of a reference to its first cell, its quotes as a reference has them.
  const sheet = whole(scanReferences(`${raw}!A1`), raw.length + 3)?.sheet
  return sheet === undefined ? null : { sheet }
}

export interface NavigationDeps {
  pickSheet: (index: number) => Promise<void>
  notify: (text: string) => void
}

export function useNavigation(s: SheetState, sel: Selection, f: Formula, deps: NavigationDeps) {
  const history = ref<Place[]>([])
  /** The range just jumped to, on its sheet: another sheet shown while it flashes has none. */
  const flash = ref<{ sheet: string; rect: CellRect } | null>(null)
  let flashTimer: ReturnType<typeof setTimeout> | null = null
  onScopeDispose(() => flashTimer && clearTimeout(flashTimer))

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
    s.scrollGridTo(place.scrollTop, place.scrollLeft)
    s.focusGrid()
  }

  function flashRange(range: CellRect): void {
    if (flashTimer) clearTimeout(flashTimer)
    // Reset and set in different ticks, so a second jump to the same place restarts the animation.
    flash.value = null
    const sheet = s.sheetName.value
    void nextTick(() => {
      flash.value = { sheet, rect: range }
      flashTimer = setTimeout(() => (flash.value = null), FLASH_MS)
    })
  }
  /** Select a rectangle cut to the shown grid (whole columns and rows stay whole), show its top-left corner and flash it. */
  function selectRange(range: CellRect): void {
    // A whole column or row runs to Excel's limit: it stops at the grid, which an editable sheet stretches to the place.
    s.stretch(range.bottom >= MAX_ROWS ? range.top : range.bottom, range.right >= MAX_COLS ? range.left : range.right)
    // Sheet rows: under a display filter the grid's count is of the shown rows, not the last row.
    const rows = Math.max(1, s.lastRow.value)
    const cols = Math.max(1, s.colCount.value)
    const top = Math.min(range.top, rows)
    const left = Math.min(range.left, cols)
    sel.selectRect({ top, left, bottom: range.bottom >= MAX_ROWS ? MAX_ROWS : Math.min(range.bottom, rows), right: range.right >= MAX_COLS ? MAX_COLS : Math.min(range.right, cols) })
    flashRange(sel.selRect.value!)
    void nextTick(() => {
      sel.revealCell(top, left)
      s.focusGrid()
    })
  }

  /** Go to a place: the sheet when named, and the range; false when the text is not an address or there is no such sheet. */
  async function goPlace(text: string): Promise<boolean> {
    const place = parsePlace(text)
    if (!place) return false
    let index = s.active.value
    if (place.sheet !== undefined) {
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
  async function goRef(r: GridReference): Promise<void> {
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
    const place = `${quoteSheetIfNeeded(s.sheet.value.name)}!${sel.selectionRef.value}`
    try {
      await navigator.clipboard.writeText(url(place))
      deps.notify(s.messages.value.cellLinkCopied(place))
    } catch {
      deps.notify(s.messages.value.clipboardDenied)
    }
  }

  return { history, flash, goBack, goPlace, goRef, addressDraft, addressFocused, goAddress, takePendingFocus, copyCellLink }
}

export type Navigation = ReturnType<typeof useNavigation>
