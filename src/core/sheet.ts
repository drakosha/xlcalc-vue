// The component's wiring: state, geometry, windows, selection, editing, copy and fill, jumps, the source's events and the viewport's size.
import { computed, nextTick, onBeforeUnmount, onMounted, ref, triggerRef, watch } from 'vue'
import { cellAddress, linesKey, parseCell, rangeAddress, sameName } from 'xlcalc/a1'
import type { CellStyle } from 'xlcalc/contract'
import { errorText } from '../failure'
import type { SheetInfo, SheetWindow } from '../contract/source'
import { useCellView } from './cell-view'
import { useClip } from './clip'
import { useCommands, type SheetTarget } from './commands'
import type { ListSort, SheetLinkEvent, SheetViewApi, SheetViewProps, SheetViewState } from '../contract/view'
import { useEditing, type EditorState } from './editing'
import { useExplain } from './explain'
import { useFill } from './fill'
import { useFormula } from './formula'
import { usePoint } from './point'
import { colPx, MAX_COL_PX, MAX_ROW_PX, MIN_COL_PX, MIN_ROW_PX, ptFromPx, useGeometry, widthFromPx } from './geometry'
import { isShortcut, shortcutLetter } from './keys'
import { useMenu } from './menu'
import { useMerge } from './merge'
import { useNavigation } from './navigation'
import { useOutline } from './outline'
import { useSelection } from './selection'
import { useValidation, useValidationAlert, VALIDATION_DELAY_MS } from './validation'
import { useValidationDialog } from './validation-dialog'
import { createState } from './state'
import { useFilter } from './filter'
import { useFind } from './find'
import { useList } from './list'
import { useTotals } from './totals'
import { useViewport } from './viewport'
import { useWindows } from './windows'

export interface SheetEmit {
  (e: 'state', state: SheetViewState | null): void
  (e: 'openLink', link: SheetLinkEvent): void
  (e: 'sort', sort: ListSort): void
  (e: 'cellDrill', cell: { sheet: string; cell: string }): void
  (e: 'rowOpen', row: { sheet: string; row: number }): void
}

/** How long a message stays. */
const MESSAGE_MS = 6000

/**
 * The sheet to display once the sheets have been read again: the requested one; otherwise the same one (`keep`); otherwise,
 * as Excel moves on from a deleted or hidden sheet, the next visible sheet from its position, or the previous one at the
 * end. Names match case-insensitively, as Excel's do.
 */
function sheetAfter(list: SheetInfo[], name: string, at: number, target: SheetTarget | null): { index: number; keep: boolean } {
  const shown = (i: number) => i >= 0 && !list[i]!.hidden
  if (target) {
    const i = list.findIndex((x) => sameName(x.name, target.name))
    if (shown(i)) return { index: i, keep: target.keep }
  }
  const i = list.findIndex((x) => sameName(x.name, name))
  if (shown(i)) return { index: i, keep: true }
  const from = i >= 0 ? i + 1 : Math.max(0, at)
  for (let j = from; j < list.length; j++) if (shown(j)) return { index: j, keep: false }
  for (let j = Math.min(from, list.length) - 1; j >= 0; j--) if (shown(j)) return { index: j, keep: false }
  return { index: Math.min(Math.max(0, at), list.length - 1), keep: false }
}

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
  /** Excel's formula bar expanded to several lines; it stays so until toggled, whatever the cell. */
  const barOpen = ref(false)
  const toggleBar = () => (barOpen.value = !barOpen.value)
  /** Excel's Ctrl+Shift+U (on a Mac too, with any layout): expands or collapses the formula bar; true when handled. */
  function barKey(e: KeyboardEvent): boolean {
    if (!e.ctrlKey || !e.shiftKey || e.altKey || e.metaKey || shortcutLetter(e) !== 'u') return false
    e.preventDefault()
    toggleBar()
    return true
  }
  const alerts = useValidationAlert()
  const ed = useEditing(s, w, sel, f, editor, { notify, reload, pickSheet, cellInput, barInput, barKey, point: (e) => point.key(e), alert: alerts.ask })
  const point = usePoint(s, sel, v, editor, { input: ed.inputOf, show: ed.focusEditor, closeList: ed.ac.close, scan: f.scanOptions, commit: ed.commit })
  // Excel's data validation on the active cell: its dropdown list and input message; an item goes to the source's pick, which writes it as Excel does, otherwise in as typed, past the rule.
  const validation = useValidation({
    source: () => s.source(),
    sheet: () => s.sheetName.value,
    active: () => s.selected.value,
    text: () => (s.selected.value ? w.textOfRow(s.selected.value.row, s.selected.value.col) : ''),
    pick: (text, at, item) => {
      const a = parseCell(at.cell)
      if (!s.editable.value || !a || at.sheet !== s.sheetName.value) return
      const lock = ed.lockAt(a.row, a.col)
      if (lock !== null) return notify(s.messages.value.locked(lock))
      const source = s.source()
      const entry = [{ sheet: at.sheet, cell: at.cell, text }]
      // a rejected pick (the list changed meanwhile) requests the rule again, so the list shows what the source holds now
      if (source.pick) void ed.run(() => source.pick!(at.sheet, at.cell, item, text), entry, null, null).then((out) => out.ok || validation.refresh())
      else void ed.send(entry, null, { accept: true })
    },
  })
  const totals = useTotals(s, w, sel)
  const filter = useFilter(s, w, notify)
  const find = useFind(s, sel, notify)
  const explain = useExplain(s, f)
  const views = useCellView(s, geo, v, w, sel, f, nav, ed, { editor, hostLinks, find, openRows: () => !!props.openRows })
  const list = useList(s, w, emit, props)
  /** The active cell's formatting, from its loaded window: the `state` event and the font shortcuts use it. */
  const activeLook = computed(() => (s.selected.value ? w.lookOfRow(s.selected.value.row, s.selected.value.col) : null))
  const cmd = useCommands(s, sel, ed, {
    reload,
    reloads: () => reloads,
    repaints: () => repaints,
    target: (t) => (target = t),
    activeStyle: () => activeLook.value?.style,
    frozen: () => v.bookFreeze.value,
    middlePanes: () => v.middlePanes(),
    showFreeze: (sheet, on) => v.showFreeze(sheet, on),
    structured: () => originalWidths.clear(),
  })
  const clip = useClip(s, w, sel, v, ed, cmd, totals, { notify, editor, colShown: (col) => geo.widthOf(col) > 0 })
  const fill = useFill(s, geo, v, sel, ed, cmd, { notify, editor })
  const outline = useOutline(s, v, sel, cmd, { notify, ask: (op) => menu.askOutline(op) })
  const menu = useMenu(s, sel, cmd, { pickSheet, tabs: () => tabs.value, outline, clip, filter, explain, frozen: () => v.bookFreeze.value })
  const merge = useMerge(s, w, sel, cmd, { ask: alerts.ask })
  const rules = useValidationDialog(s, sel, w, { scan: f.scanOptions, ask: alerts.ask, notify, changed: () => validation.refresh() })

  /** What the host's template ref gets: the commands as the sheet's own keys and menus run them. */
  const api: SheetViewApi = {
    undo: () => ed.stepHistory('undo'),
    redo: () => ed.stepHistory('redo'),
    format: cmd.format,
    structure: (op) => cmd.structure(op),
    outline: cmd.outline,
    formatCells: cmd.openFormat,
    freezePanes: cmd.freeze,
    reapplyFilter: filter.reapply,
    explain: explain.show,
    merge: merge.run,
    dataValidation: rules.open,
  }

  // ── sheets ──
  /** Tabs: sheets hidden in the workbook get no tab, as in Excel; their indexes stay. */
  const tabs = computed(() => s.sheets.value.map((x, i) => ({ sheet: x, i })).filter(({ sheet }) => !sheet.hidden))
  /** Excel shows the tabs with a single sheet too; `hideLoneTab` shows them only from two. */
  const showTabs = computed(() => tabs.value.length > (props.hideLoneTab ? 1 : 0))
  const hasLayout = () => typeof s.source().layout === 'function'

  /** Makes a sheet active: windows, sizes, selection and scroll start over. */
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
    s.scrollGridTo(0, 0)
    w.ensure()
  }
  async function pickSheet(index: number): Promise<void> {
    if (index === s.active.value || !s.sheets.value[index]) return
    activate(index)
    await nextTick()
  }

  let initSeq = 0
  /** Reads the workbook from the start: sheets, the first visible sheet, and a pending focus. */
  async function init(): Promise<void> {
    const mine = ++initSeq
    const outdated = s.outdated()
    reloadSeq++
    clip.clear()
    w.reset()
    s.failure.value = null
    s.sheets.value = []
    s.active.value = -1
    s.loading.value = true
    try {
      const list = await s.source().sheets()
      if (mine !== initSeq || outdated()) return
      s.sheets.value = list
      const first = list.findIndex((x) => !x.hidden)
      activate(first >= 0 ? first : 0)
      const focus = nav.takePendingFocus()
      if (focus) await nav.goPlace(focus)
    } catch (e) {
      if (mine !== initSeq || outdated()) return
      s.loading.value = false
      s.failure.value = s.failureOf(e)
    }
  }
  let reloads = 0
  let repaints = 0
  let target: SheetTarget | null = null
  /** The latest reload: an earlier one answered later drops its list, as does one of a replaced workbook. */
  let reloadSeq = 0
  /** Reads the sheets and windows again, keeping the sheet, selection and scroll (a reload event or a conflict); a deleted or hidden sheet gives way to its neighbor. */
  async function reload(): Promise<void> {
    reloads++
    const mine = ++reloadSeq
    const outdated = s.outdated()
    const name = s.sheetName.value
    const at = s.active.value
    const asked = target
    try {
      const list = await s.source().sheets()
      if (mine !== reloadSeq || outdated()) return
      if (target === asked) target = null
      s.sheets.value = list
      const next = sheetAfter(list, name, at, asked)
      if (!next.keep) return activate(next.index)
      s.active.value = next.index
    } catch (e) {
      if (mine !== reloadSeq || outdated()) return
      if (target === asked) target = null
      s.failure.value = s.failureOf(e)
      return
    }
    s.grown.value = { rows: 0, cols: 0 }
    w.reset()
    w.resetSizes()
    if (hasLayout()) s.userWidths.value = {}
    if (filter.active.value) void filter.refresh()
    w.ensure()
  }
  function retry(): void {
    void init()
  }

  /** A grip captures the pointer for its drag: the page neither selects nor scrolls, and the grid does not select the cell under it. */
  function holdPointer(e: PointerEvent): void {
    ;(e.target as HTMLElement).setPointerCapture?.(e.pointerId)
    e.preventDefault()
    e.stopPropagation()
  }

  // ── column width with the mouse ──
  /** Workbook widths before the first resize in this session, so a double click can restore them through the source. */
  const originalWidths = new Map<number, number>()
  let drag: { col: number; startX: number; startW: number } | null = null
  function startResize(e: PointerEvent, col: number): void {
    drag = { col, startX: e.clientX, startW: geo.widthOf(col) }
    if (!originalWidths.has(col)) originalWidths.set(col, s.bookWidths.value.has(col) ? colPx(s.bookWidths.value.get(col)) : geo.widthOf(col))
    holdPointer(e)
  }
  function onResizeMove(e: PointerEvent): void {
    if (!drag) return
    const px = Math.min(MAX_COL_PX, Math.max(MIN_COL_PX, drag.startW + e.clientX - drag.startX))
    // In place: a drag sets one width many times a second.
    s.userWidths.value[drag.col] = px
    triggerRef(s.userWidths)
  }
  function endResize(): void {
    if (!drag) return
    const { col } = drag
    drag = null
    const px = s.userWidths.value[col]
    if (px !== undefined) void saveWidth(col, px)
    w.ensure()
  }
  /** A double click on a column edge fits the column, as Excel does, together with every selected entire column when it is one of them; without `layout` (or with `edgeRestoresWidth`) it restores the workbook's width. */
  const fitsOnEdge = computed(() => !props.edgeRestoresWidth && cmd.can('layout'))
  function onEdgeDouble(col: number): void {
    if (!fitsOnEdge.value) return resetWidth(col)
    const { first, last } = sel.linesOf(sel.wholeColumns.value, col)
    void cmd.layout({ sheet: s.sheetName.value, columns: linesKey(first, last, 'columns'), fit: true })
  }
  /** The widths set with the mouse, except a column's. */
  function widthsWithout(col: number): Record<number, number> {
    const next = { ...s.userWidths.value }
    delete next[col]
    return next
  }
  /** Restores the workbook's width. */
  function resetWidth(col: number): void {
    const next = widthsWithout(col)
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
  /** With `layout` the width goes to the source; without it the browser stores it. */
  async function saveWidth(col: number, px: number): Promise<void> {
    const layout = s.sourceMethod('layout')
    if (!layout) {
      s.memory.remember(`widths:${s.sheetName.value}`, s.userWidths.value)
      return
    }
    const sheet = s.sheetName.value
    let ok = false
    let text = ''
    try {
      const got = await layout({ sheet, columns: linesKey(col, col, 'columns'), width: widthFromPx(px) })
      ok = got.ok
      // cells whose display changes with the width (a date that no longer shows #####) come back as changes, as with an edit
      if (got.ok) w.applyChanges(got.changes)
      else text = got.message
    } catch (e) {
      text = errorText(e)
    }
    if (ok || sheet !== s.sheetName.value) return
    s.userWidths.value = widthsWithout(col)
    notify(text || s.messages.value.editFailed)
  }

  // ── row height with the mouse ──
  /** The grip on a row number's edge, where the source sets heights; not for touch, which would snap to it from the number (a long press opens the menu, which has Row Height). */
  const rowsResizable = computed(() => cmd.can('layout') && !v.touch.value)
  /** The dragged row is drawn at the new height meanwhile; the release sets it for every selected entire row it is one of. */
  let rowDrag: { first: number; last: number; row: number; startY: number; startH: number; px: number } | null = null
  function startRowResize(e: PointerEvent, row: number): void {
    if (!rowsResizable.value || e.button !== 0) return
    const { first, last } = sel.linesOf(sel.wholeRows.value, row)
    const startH = s.heights.value.get(row) ?? s.defaultRowPx.value
    rowDrag = { first, last, row, startY: e.clientY, startH, px: startH }
    holdPointer(e)
  }
  function onRowResizeMove(e: PointerEvent): void {
    if (!rowDrag) return
    rowDrag.px = Math.min(MAX_ROW_PX, Math.max(MIN_ROW_PX, Math.round(rowDrag.startH + e.clientY - rowDrag.startY)))
    geo.setHeights([[rowDrag.row, rowDrag.px]])
  }
  /** A double click on a row number's edge fits the row, and every selected entire row when it is one of them, as Excel's AutoFit does. */
  function onRowEdgeDouble(row: number): void {
    if (!rowsResizable.value) return
    const { first, last } = sel.linesOf(sel.wholeRows.value, row)
    void cmd.layout({ sheet: s.sheetName.value, rows: linesKey(first, last, 'rows'), fit: true })
  }
  async function endRowResize(): Promise<void> {
    const d = rowDrag
    rowDrag = null
    if (!d || d.px === d.startH) return
    const out = await cmd.layout({ sheet: s.sheetName.value, rows: linesKey(d.first, d.last, 'rows'), height: ptFromPx(d.px) })
    // A rejected height: the rows are read again as the source has them.
    if (!out?.ok) await reload()
    w.ensure()
  }

  // ── hyperlinks in cells ──
  /** A hyperlink clicked in a cell: a location in the workbook is handled by the component, anything else goes to the host if it listens. */
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

  // ── keys on the grid: search and Esc first (a popover, the Find bar, copy mode, the precedents), then font shortcuts, editing and navigation ──
  function onGridKey(e: KeyboardEvent): void {
    if (e.target !== s.scroller.value) return
    if (validation.onKey(e)) return
    if (!e.isComposing && (cmd.fontKey(e) || outline.key(e) || barKey(e))) return
    if (isShortcut(e) && shortcutLetter(e) === 'f' && find.available.value) {
      e.preventDefault()
      find.openFind()
      return
    }
    // Ctrl+H on a Mac too, as in Excel for Mac: Cmd+H hides the browser.
    if (e.ctrlKey && !e.altKey && shortcutLetter(e) === 'h' && find.available.value) {
      e.preventDefault()
      find.openReplace()
      return
    }
    // Ctrl+1 on a Mac too: the browser reserves Cmd+1 for its first tab.
    if (e.ctrlKey && !e.altKey && !e.shiftKey && (e.key === '1' || e.code === 'Digit1') && cmd.openFormat()) {
      e.preventDefault()
      return
    }
    if (e.key === 'Escape' && (filter.open.value !== null || find.isOpen.value)) {
      e.preventDefault()
      if (filter.open.value !== null) filter.close()
      else find.closeFind()
      return
    }
    // Esc ends Excel's copy mode: the moving border disappears.
    if (e.key === 'Escape' && clip.mark.value) {
      e.preventDefault()
      clip.clear()
      return
    }
    if (e.key === 'Escape' && explain.open.value) {
      e.preventDefault()
      explain.close()
      return
    }
    // Excel's Ctrl+Page Down and Ctrl+Page Up: the next or the previous visible sheet, stopping at the ends.
    if (e.ctrlKey && !e.altKey && !e.metaKey && !e.shiftKey && (e.key === 'PageDown' || e.key === 'PageUp')) {
      e.preventDefault()
      const next = tabs.value[tabs.value.findIndex((t) => t.i === s.active.value) + (e.key === 'PageDown' ? 1 : -1)]
      if (next) void pickSheet(next.i)
      return
    }
    // Excel's Reapply, Ctrl+Alt+L on a Mac too: Cmd+Option+L opens the browser's downloads.
    if (e.ctrlKey && e.altKey && !e.metaKey && !e.shiftKey && shortcutLetter(e) === 'l' && filter.active.value) {
      e.preventDefault()
      void filter.reapply()
      return
    }
    // In list mode Enter opens the active cell's drill-down or row instead of moving.
    if (list.key(e)) return
    ed.onGridKey(e)
  }

  // ── the viewport ──
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
  // What is drawn decides what is loaded: scrolling, resizing, freezing and new data all end up here.
  watch(
    () => [v.firstRow.value, v.lastRow.value, v.firstCol.value, v.lastCol.value, v.frozenRows.value, v.frozenCols.value, s.dataRows.value, s.dataCols.value],
    () => w.ensure(),
  )
  onMounted(() => {
    sel.measure()
    subscribe()
    // The first window is requested after the viewport is measured, or it would be requested again at the real size.
    void init()
  })

  // ── the source's events ──
  let unsubscribe: (() => void) | null = null
  function subscribe(): void {
    unsubscribe?.()
    unsubscribe =
      s.source().subscribe?.((event) => {
        if (event.kind === 'changed') w.applyChanges(event.changes)
        else if (event.kind === 'repaint') {
          repaints++
          w.repaint(event.ranges)
        }
        else {
          if (event.message) notify(event.message)
          void reload()
        }
        // A list of cells or the rules themselves may have changed, and the precedents' values too: requested once a burst of events ends.
        if (event.kind !== 'repaint') {
          validation.refresh(VALIDATION_DELAY_MS)
          explain.refresh()
        }
      }) ?? null
  }
  watch(
    () => props.source,
    () => {
      // A new workbook: the previous one's pending calls, marks and open alert end with it.
      s.endWorkbook()
      ed.forget()
      alerts.answer('cancel')
      rules.state.value = null
      subscribe()
      void init()
    },
  )

  // The sheet and selection are reported: formula, value and formatting come from the loaded window; the Format Cells dialog reads them too.
  const viewState = computed<SheetViewState | null>(() => {
    if (!s.sheet.value || !sel.selectionRef.value) return null
    const look = activeLook.value
    return {
      sheet: s.sheet.value.name,
      ref: sel.selectionRef.value,
      ...(f.selectedFormula.value ? { formula: f.selectedFormula.value } : {}),
      ...(f.selectedValue.value ? { value: f.selectedValue.value } : {}),
      ...(look?.style ? { style: look.style } : {}),
      ...(look?.numberFormat !== undefined ? { numberFormat: look.numberFormat } : {}),
      ...(merge.area.value ? { merged: rangeAddress(merge.area.value) } : {}),
      ...(merge.available.value ? { canMerge: true as const } : {}),
    }
  })
  watch(viewState, (state) => emit('state', state), { immediate: true })
  onBeforeUnmount(() => {
    s.endWorkbook()
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
    point,
    totals,
    filter,
    find,
    explain,
    views,
    list,
    cmd,
    api,
    clip,
    fill,
    viewState,
    validation,
    rules,
    alerts,
    outline,
    menu,
    editor,
    message,
    notify,
    dismiss,
    cellInput,
    barInput,
    barOpen,
    toggleBar,
    tabs,
    showTabs,
    pickSheet,
    reload,
    retry,
    startResize,
    onResizeMove,
    endResize,
    rowsResizable,
    startRowResize,
    onRowResizeMove,
    endRowResize,
    onRowEdgeDouble,
    fitsOnEdge,
    onEdgeDouble,
    onLink,
    onScroll,
    onGridKey,
  }
}

export type SheetContext = ReturnType<typeof useSheet>
