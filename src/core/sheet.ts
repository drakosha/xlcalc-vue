// The component's wiring: state, geometry, windows, selection, editing and jumps, the source's events and the frame's size.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import type { CellStyle } from 'xlcalc/contract'
import type { SheetWindow } from '../source'
import { cellAddress, colsAddress } from './a1'
import { useCellView } from './cell-view'
import { useEditing, type EditorState } from './editing'
import { useFormula } from './formula'
import { MAX_COL_PX, MIN_COL_PX, useGeometry, widthFromPx } from './geometry'
import { useNavigation } from './navigation'
import { useSelection } from './selection'
import { createState, type SheetViewProps } from './state'
import { useFilter } from './filter'
import { useFind } from './find'
import { useTotals } from './totals'
import { useViewport } from './viewport'
import { useWindows } from './windows'

/** What `state` reports: the sheet, the selection, and the active cell's formula and shown value. */
export interface SheetViewState {
  sheet: string
  ref: string
  formula?: string
  value?: string
  /** The style the active cell shows, for a toolbar; absent while its window is not loaded or it has none. */
  style?: CellStyle
  /** The active cell's number format code, from a source whose windows give them; absent for General. */
  numberFormat?: string
}

/** A link the component does not open itself, with the cell it is in. */
export type SheetLinkEvent = NonNullable<SheetWindow['links']>[string] & { sheet: string; cell: string }

export interface SheetEmit {
  (e: 'state', state: SheetViewState | null): void
  (e: 'openLink', link: SheetLinkEvent): void
}

/** How long a message stays. */
const MESSAGE_MS = 6000

export function useSheet(props: SheetViewProps, emit: SheetEmit, hostLinks: () => boolean) {
  const s = createState(props)
  const geo = useGeometry(s)
  const v = useViewport(s, geo)
  const w = useWindows(s, v, geo)
  const sel = useSelection(s, geo, v)
  const editor = ref<EditorState | null>(null)
  const f = useFormula(s, w, sel, computed(() => editor.value?.text ?? null))

  // ── messages: one line at the bottom, the last one wins ──
  const message = ref<string | null>(null)
  let messageTimer: ReturnType<typeof setTimeout> | null = null
  function notify(text: string): void {
    if (messageTimer) clearTimeout(messageTimer)
    message.value = text
    messageTimer = setTimeout(() => (message.value = null), MESSAGE_MS)
  }
  const dismiss = () => (message.value = null)

  const nav = useNavigation(s, sel, f, { pickSheet, notify })
  const cellInput = ref<HTMLTextAreaElement | null>(null)
  const barInput = ref<HTMLTextAreaElement | null>(null)
  const ed = useEditing(s, w, sel, f, editor, { notify, reload, goPlace: nav.goPlace, cellInput, barInput })
  const totals = useTotals(s, w, sel, notify)
  const filter = useFilter(s, w, notify)
  const find = useFind(s, sel, notify)
  const views = useCellView(s, geo, v, w, sel, f, nav, ed, { editor, hostLinks, find })

  // ── sheets ──
  /** Tabs: sheets hidden in the book are not shown, as in Excel; their indexes stay. */
  const tabs = computed(() => s.sheets.value.map((x, i) => ({ sheet: x, i })).filter(({ sheet }) => !sheet.hidden))
  const hasLayout = () => typeof props.source.layout === 'function'

  /** Make a sheet active: windows, sizes, selection and scroll start over. */
  function activate(index: number): void {
    w.reset()
    w.resetSizes()
    editor.value = null
    s.active.value = index
    s.grown.value = { rows: 0, cols: 0 }
    s.reach.value = { rows: 0, cols: 0 }
    s.selected.value = s.sheet.value ? { row: 1, col: 1 } : null
    s.extent.value = null
    s.userWidths.value = hasLayout() ? {} : (s.memory.recall<Record<number, number>>(`widths:${s.sheetName.value}`) ?? {})
    originalWidths.clear()
    v.readFreeze()
    filter.restore()
    s.scroller.value?.scrollTo({ top: 0, left: 0 })
    s.scrollTop.value = 0
    s.scrollLeft.value = 0
    w.ensure()
  }
  async function pickSheet(index: number): Promise<void> {
    if (index === s.active.value || !s.sheets.value[index]) return
    activate(index)
    await nextTick()
  }

  let initSeq = 0
  /** Read the book from the start: sheets, the first shown sheet, a focus waiting for it. */
  async function init(): Promise<void> {
    const mine = ++initSeq
    w.reset()
    s.failure.value = null
    s.sheets.value = []
    s.active.value = -1
    s.loading.value = true
    try {
      const list = await props.source.sheets()
      if (mine !== initSeq) return
      s.sheets.value = list
      const first = list.findIndex((x) => !x.hidden)
      activate(first >= 0 ? first : 0)
      const focus = nav.takePendingFocus()
      if (focus) await nav.goPlace(focus)
    } catch (e) {
      if (mine !== initSeq) return
      s.loading.value = false
      s.failure.value = { error: e, text: e instanceof Error && e.message ? e.message : s.messages.value.loadFailed }
    }
  }
  /** Read the sheets and windows again, keeping the sheet, selection and scroll (a reload event or a conflict). */
  async function reload(): Promise<void> {
    const name = s.sheetName.value
    try {
      const list = await props.source.sheets()
      s.sheets.value = list
      const index = list.findIndex((x) => x.name === name)
      s.active.value = index >= 0 ? index : Math.min(Math.max(0, s.active.value), list.length - 1)
    } catch (e) {
      s.failure.value = { error: e, text: e instanceof Error && e.message ? e.message : s.messages.value.loadFailed }
      return
    }
    s.grown.value = { rows: 0, cols: 0 }
    w.reset()
    w.resetSizes()
    if (hasLayout()) s.userWidths.value = {}
    if (filter.active.value) filter.refresh()
    w.ensure()
  }
  function retry(): void {
    void init()
  }

  // ── column width with the mouse ──
  /** Book widths before the first resize in this session, for a double click to restore them through the source. */
  const originalWidths = new Map<number, number>()
  let drag: { col: number; startX: number; startW: number } | null = null
  function startResize(e: PointerEvent, col: number): void {
    drag = { col, startX: e.clientX, startW: geo.widthOf(col) }
    if (!originalWidths.has(col)) originalWidths.set(col, s.bookWidths.value.get(col) ?? geo.widthOf(col))
    ;(e.target as HTMLElement).setPointerCapture?.(e.pointerId)
    e.preventDefault()
    e.stopPropagation()
  }
  function onResizeMove(e: PointerEvent): void {
    if (!drag) return
    const px = Math.min(MAX_COL_PX, Math.max(MIN_COL_PX, drag.startW + e.clientX - drag.startX))
    s.userWidths.value = { ...s.userWidths.value, [drag.col]: px }
  }
  function endResize(): void {
    if (!drag) return
    const { col } = drag
    drag = null
    const px = s.userWidths.value[col]
    if (px !== undefined) void saveWidth(col, px)
    w.ensure()
  }
  /** A double click on the edge: the book's width back. */
  function resetWidth(col: number): void {
    const next = { ...s.userWidths.value }
    delete next[col]
    if (!hasLayout()) {
      s.userWidths.value = next
      s.memory.remember(`widths:${s.sheetName.value}`, Object.keys(next).length ? next : null)
      return
    }
    const px = originalWidths.get(col)
    if (px === undefined) return
    s.userWidths.value = { ...next, [col]: px }
    void saveWidth(col, px)
  }
  /** With `layout` the width goes to the source; without it the browser remembers it. */
  async function saveWidth(col: number, px: number): Promise<void> {
    const layout = props.source.layout
    if (!layout) {
      s.memory.remember(`widths:${s.sheetName.value}`, s.userWidths.value)
      return
    }
    const sheet = s.sheetName.value
    let ok = false
    let text = ''
    try {
      const got = await layout.call(props.source, { sheet, columns: colsAddress(col, col), width: widthFromPx(px) })
      ok = got.ok
      if (!got.ok) text = got.message
    } catch (e) {
      text = e instanceof Error ? e.message : ''
    }
    if (ok || sheet !== s.sheetName.value) return
    const next = { ...s.userWidths.value }
    delete next[col]
    s.userWidths.value = next
    notify(text || s.messages.value.editFailed)
  }

  // ── links in cells ──
  /** A link clicked in a cell: a place in the book is the component's, anything else goes to the host if it listens. */
  function onLink(e: MouseEvent, link: NonNullable<SheetWindow['links']>[string], row: number, col: number): void {
    if (link.location) {
      e.preventDefault()
      void nav.goPlace(link.location.replace(/^#/, ''))
      return
    }
    if (!hostLinks()) return
    e.preventDefault()
    emit('openLink', { ...link, sheet: s.sheetName.value, cell: cellAddress(row, col) })
  }

  // ── keys on the grid: search and Esc first, then editing and moves ──
  function onGridKey(e: KeyboardEvent): void {
    if (e.target !== s.scroller.value) return
    if ((e.ctrlKey || e.metaKey) && /^[fFаА]$/.test(e.key) && find.available.value) {
      e.preventDefault()
      find.openFind()
      return
    }
    if (e.key === 'Escape' && (filter.open.value !== null || find.isOpen.value)) {
      e.preventDefault()
      if (filter.open.value !== null) filter.close()
      else find.closeFind()
      return
    }
    ed.onGridKey(e)
  }

  // ── the frame ──
  function onScroll(): void {
    const el = s.scroller.value
    if (!el) return
    s.scrollTop.value = el.scrollTop
    s.scrollLeft.value = el.scrollLeft
    w.ensure()
  }
  // The scroller comes and goes with the fallback: whichever is mounted is measured and observed.
  const observer = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => sel.measure()) : null
  watch(s.scroller, (el, old) => {
    if (old) observer?.unobserve(old)
    if (!el) return
    observer?.observe(el)
    sel.measure()
  })
  // What is drawn decides what is loaded: scroll, resize, freeze and new data all come here.
  watch(
    () => [v.firstRow.value, v.lastRow.value, v.firstCol.value, v.lastCol.value, v.frozenRows.value, v.frozenCols.value, s.dataRows.value, s.dataCols.value],
    () => w.ensure(),
  )
  onMounted(() => {
    sel.measure()
    subscribe()
    // The first window goes after the frame is measured, or it would be asked again at the real size.
    void init()
  })

  // ── the source's events ──
  let unsubscribe: (() => void) | null = null
  function subscribe(): void {
    unsubscribe?.()
    unsubscribe =
      props.source.subscribe?.((event) => {
        if (event.kind === 'changed') w.applyChanges(event.changes)
        else {
          if (event.message) notify(event.message)
          void reload()
        }
      }) ?? null
  }
  watch(
    () => props.source,
    () => {
      subscribe()
      void init()
    },
  )

  // The sheet and selection go out: formula, value and look come from the loaded window, so they are watched too.
  const activeLook = computed(() => (s.selected.value ? w.lookOfRow(s.selected.value.row, s.selected.value.col) : null))
  watch(
    [s.sheet, sel.selectionRef, f.selectedFormula, f.selectedValue, activeLook],
    () => {
      if (!s.sheet.value || !sel.selectionRef.value) return emit('state', null)
      const look = activeLook.value
      emit('state', {
        sheet: s.sheet.value.name,
        ref: sel.selectionRef.value,
        ...(f.selectedFormula.value ? { formula: f.selectedFormula.value } : {}),
        ...(f.selectedValue.value ? { value: f.selectedValue.value } : {}),
        ...(look?.style ? { style: look.style } : {}),
        ...(look?.numberFormat !== undefined ? { numberFormat: look.numberFormat } : {}),
      })
    },
    { immediate: true },
  )
  onBeforeUnmount(() => {
    observer?.disconnect()
    unsubscribe?.()
    if (messageTimer) clearTimeout(messageTimer)
    w.reset()
    emit('state', null)
  })

  return {
    ...s,
    geo,
    v,
    w,
    sel,
    f,
    nav,
    ed,
    totals,
    filter,
    find,
    views,
    editor,
    message,
    notify,
    dismiss,
    cellInput,
    barInput,
    tabs,
    pickSheet,
    reload,
    retry,
    startResize,
    onResizeMove,
    endResize,
    resetWidth,
    onLink,
    onScroll,
    onGridKey,
  }
}

export type SheetContext = ReturnType<typeof useSheet>
