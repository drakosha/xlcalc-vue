// Jumps: the Name Box (addresses and defined names), the formula's references, hyperlinks and `focus`; the Back history, and the flash after a jump.
import { computed, nextTick, onScopeDispose, ref, watch } from 'vue'
import { MAX_COLS, MAX_ROWS, quoteSheetIfNeeded, rangeAddress } from 'xlcalc/a1'
import type { CellRect, GridReference } from 'xlcalc/contract'
import type { SheetName } from '../contract/workbook'
import type { Formula } from './formula'
import type { Names } from './names'
import { parsePlace, rangeRef } from './place'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A location to return to: sheet, selection and scroll. */
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
  /** The workbook's defined names, which the Name Box goes to, lists, shows and defines. */
  names: Names
}

export function useNavigation(s: SheetState, sel: Selection, f: Formula, deps: NavigationDeps) {
  const history = ref<Place[]>([])
  /** The range just jumped to, on its sheet: none while another sheet is displayed during the flash. */
  const flash = ref<{ sheet: string; rect: CellRect } | null>(null)
  let flashTimer: ReturnType<typeof setTimeout> | null = null
  onScopeDispose(() => flashTimer && clearTimeout(flashTimer))

  /** Records the current location before a jump. */
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
    // Reset and set in different ticks, so a second jump to the same location restarts the animation.
    flash.value = null
    const sheet = s.sheetName.value
    void nextTick(() => {
      flash.value = { sheet, rect: range }
      flashTimer = setTimeout(() => (flash.value = null), FLASH_MS)
    })
  }
  /** Selects a range clipped to the displayed grid (entire columns and rows stay entire), shows its top-left corner and flashes it. */
  function selectRange(range: CellRect): void {
    // An entire column or row extends to Excel's limit: it stops at the grid, which an editable sheet extends to the location.
    sel.stretchTo(range)
    // Sheet rows: under a display filter the grid's count is of the visible rows, not the last row.
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

  /** Goes to a location: the sheet when named, and the range; false when the text is not an address or there is no such sheet. */
  async function goPlace(text: string): Promise<boolean> {
    const place = parsePlace(text)
    if (!place) return false
    let index = s.active.value
    if (place.sheet !== undefined) {
      index = f.sheetIndexByName(place.sheet)
      if (index < 0) return false
    }
    // The workbook's first view (A1 unscrolled, a link from outside) has nothing to go back to.
    const a = s.selected.value
    if (history.value.length || s.scrollTop.value || s.scrollLeft.value || s.extent.value || (a && (a.row !== 1 || a.col !== 1))) rememberPlace()
    if (index !== s.active.value) await deps.pickSheet(index)
    if (place.range) selectRange(place.range)
    return true
  }
  /** A reference clicked in the formula bar: selects its range, switching sheets when it names another. */
  async function goRef(r: GridReference): Promise<void> {
    const index = f.refSheet(r)
    if (index < 0) return
    rememberPlace()
    if (index !== s.active.value) await deps.pickSheet(index)
    selectRange(r)
  }

  // ── the Name Box ──
  const names = deps.names
  /** What the Name Box shows: the name whose range is exactly the selection, as Excel's does, otherwise the selection's address. */
  const boxText = computed(() => names.nameOf(sel.fullRect.value) ?? sel.selectionRef.value)
  const addressDraft = ref('')
  const addressFocused = ref(false)
  watch(boxText, (v) => {
    if (!addressFocused.value) addressDraft.value = v
  })
  /** Goes to a defined name's range, as Excel's Name Box and Go To do; false for a name that does not refer to a range. */
  async function goName(name: SheetName): Promise<boolean> {
    const r = names.rangeOf(name)
    return !!r && goPlace(`${quoteSheetIfNeeded(r.sheet)}!${rangeAddress(r.rect)}`)
  }
  /**
   * Enter in the Name Box, as in Excel: a defined name goes to its range, an address to its location, and a new name with a
   * selection defines it over the selection (workbook scope) where names can be defined; otherwise the message line says why.
   */
  async function goAddress(): Promise<boolean> {
    const ok = await enterAddress(addressDraft.value.trim())
    addressDraft.value = boxText.value
    return ok
  }
  async function enterAddress(text: string): Promise<boolean> {
    const m = s.messages.value
    const named = text ? names.find(text, s.sheetName.value) : undefined
    if (named) {
      if (await goName(named)) return true
      deps.notify(m.nameNotRange)
      return false
    }
    if (await goPlace(text)) return true
    const rect = sel.fullRect.value
    if (!text || parsePlace(text)?.range || !names.editable.value || !rect) {
      deps.notify(m.addressBad)
      return false
    }
    const out = await names.save([...names.list.value, { name: text, ref: rangeRef(s.sheetName.value, rect) }])
    if (out && !out.ok) deps.notify(out.reason === 'refused' || out.reason === 'invalid' ? m.nameBoxBad : out.message || m.editFailed)
    return !!out?.ok
  }

  /** A location requested from outside before the workbook loaded; applied once the sheets have loaded. */
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
  /** The focus waiting for the workbook, if any; it is applied once. */
  function takePendingFocus(): string | null {
    const f = pendingFocus
    pendingFocus = null
    return f
  }

  /** Copies the host page's address of the selected location to the clipboard. */
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

  return { history, flash, goBack, goPlace, goRef, goName, boxText, addressDraft, addressFocused, goAddress, takePendingFocus, copyCellLink }
}

export type Navigation = ReturnType<typeof useNavigation>
