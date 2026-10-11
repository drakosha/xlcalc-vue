// Windows of data: three slots (body, frozen rows, frozen columns), loaded under the scroll and patched by edits.
// Rows of windows are display indexes (`d`): with a display filter they are places in its list, not the sheet's rows.
import { computed, ref, shallowRef, type ShallowRef } from 'vue'
import type { CellStyle, CellValue } from 'xlcalc/contract'
import type { CellUpdate, SheetWindow, WindowRequest } from '../source'
import { cellAddress, parseCell, parseRect, type Rect } from './a1'
import { colPx, DEFAULT_ROW_PT, rowPx, type Geometry } from './geometry'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** Fetch margins: a window is taken with spare rows and columns so an ordinary scroll needs no request. */
const FETCH_ROWS = 120
const FETCH_COLS = 12
const MAX_CELLS = 20_000
const MAX_COLS = 200
/** Delay before asking again for a window that is not ready yet, and how many times (over the the app's recompute limit). */
const PENDING_RETRY_MS = 1000
const PENDING_POLLS_MAX = 200

export type SlotKind = 'body' | 'top' | 'left'

/** One window: what is loaded, the request in flight and the one already answered (so neither repeats). */
export interface Slot {
  sheet: string
  /** The sheet and the filter the window was asked under: another filter means other rows. */
  view: string
  win: SheetWindow | null
  rows: number
  cols: number
  merges: Rect[]
  pendingKey: string
  servedKey: string
  ctrl: AbortController | null
}

interface Need {
  r1: number
  r2: number
  c1: number
  c2: number
}

const newSlot = (): Slot => ({ sheet: '', view: '', win: null, rows: 0, cols: 0, merges: [], pendingKey: '', servedKey: '', ctrl: null })
const isAbort = (e: unknown) => e instanceof DOMException && e.name === 'AbortError'

/** Everything a cell has in its window, for drawing it. */
export interface CellData {
  text: string
  value: CellValue | undefined
  style: CellStyle | undefined
  formula: string | undefined
  color: string | undefined
  link: NonNullable<SheetWindow['links']>[string] | undefined
  note: NonNullable<SheetWindow['notes']>[string] | undefined
  bar: NonNullable<SheetWindow['bars']>[string] | undefined
  loaded: boolean
}

export function useWindows(s: SheetState, v: Viewport, geo: Geometry) {
  const slots: Record<SlotKind, ShallowRef<Slot>> = { body: shallowRef(newSlot()), top: shallowRef(newSlot()), left: shallowRef(newSlot()) }
  const order: SlotKind[] = ['body', 'top', 'left']
  /** Generation of the book and sheet: an answer sent before a switch must not overwrite fresh data. */
  let epoch = 0

  /** The sheet and filter the windows must be of now. */
  const viewKey = () => `${s.sheetName.value}\u0000${s.visible.value?.key ?? ''}`
  /** The window holding a cell by display row: the body first, then the frozen ones (they meet at the corner). */
  function slotAt(d: number, col: number): Slot | null {
    const view = viewKey()
    for (const kind of order) {
      const slot = slots[kind].value
      const w = slot.win
      if (!w || slot.view !== view) continue
      if (d >= w.row && d < w.row + slot.rows && col >= w.col && col < w.col + slot.cols) return slot
    }
    return null
  }
  function textAt(d: number, col: number): string {
    const slot = slotAt(d, col)
    return slot ? (slot.win!.text[d - slot.win!.row]?.[col - slot.win!.col] ?? '') : ''
  }
  function valueAt(d: number, col: number): CellValue | undefined {
    const slot = slotAt(d, col)
    return slot ? slot.win!.values[d - slot.win!.row]?.[col - slot.win!.col] : undefined
  }
  /** A cell's style; records of a window are keyed by the cell's own address (`addr`, by its sheet row). */
  function styleAt(d: number, col: number, addr: string): CellStyle | undefined {
    const w = slotAt(d, col)?.win
    const id = w?.styleIds[addr]
    return id === undefined ? undefined : w!.styles[id]
  }
  function cellData(d: number, col: number, addr: string): CellData {
    const w = slotAt(d, col)?.win
    if (!w) return { text: '', value: undefined, style: undefined, formula: undefined, color: undefined, link: undefined, note: undefined, bar: undefined, loaded: false }
    const id = w.styleIds[addr]
    return {
      text: w.text[d - w.row]?.[col - w.col] ?? '',
      value: w.values[d - w.row]?.[col - w.col],
      style: id === undefined ? undefined : w.styles[id],
      formula: w.formulas[addr],
      color: w.colors[addr],
      link: w.links?.[addr],
      note: w.notes?.[addr],
      bar: w.bars?.[addr],
      loaded: true,
    }
  }
  /** Whether a cell is known to be empty: loaded, no text, no value. */
  function emptyAt(d: number, col: number): boolean {
    const slot = slotAt(d, col)
    if (!slot) return false
    const w = slot.win!
    const value = w.values[d - w.row]?.[col - w.col]
    return (w.text[d - w.row]?.[col - w.col] ?? '') === '' && (value === null || value === undefined || value === '')
  }
  /** The same by the sheet's row: a row the filter hides has nothing loaded. */
  const textOfRow = (r: number, col: number): string => {
    const d = s.indexOf(r)
    return d === undefined ? '' : textAt(d, col)
  }
  const valueOfRow = (r: number, col: number): CellValue | undefined => {
    const d = s.indexOf(r)
    return d === undefined ? undefined : valueAt(d, col)
  }
  const formulaOfRow = (r: number, col: number): string | undefined => {
    const d = s.indexOf(r)
    const w = d === undefined ? undefined : slotAt(d, col)?.win
    return w?.formulas[cellAddress(r, col)]
  }
  const emptyAtRow = (r: number, col: number): boolean => {
    const d = s.indexOf(r)
    return d !== undefined && emptyAt(d, col)
  }
  /** The text a cell opens with from its window's `editText` (a cell left out opens empty); undefined when the window has none. */
  const editTextOfRow = (r: number, col: number): string | undefined => {
    const d = s.indexOf(r)
    const texts = d === undefined ? undefined : slotAt(d, col)?.win?.editText
    return texts ? (texts[cellAddress(r, col)] ?? '') : undefined
  }

  const current = (kind: SlotKind) => (slots[kind].value.view === viewKey() ? slots[kind].value : null)
  /** Merged areas of the loaded windows (by sheet rows), each once. */
  const merges = computed(() => {
    const seen = new Set<string>()
    const out: Rect[] = []
    for (const kind of order) {
      for (const m of current(kind)?.merges ?? []) {
        const key = `${m.r1}:${m.c1}`
        if (!seen.has(key)) {
          seen.add(key)
          out.push(m)
        }
      }
    }
    return out
  })
  /** The sheet's tables and the notice: the same in every window of the sheet. */
  const tables = computed(() => {
    for (const kind of order) {
      const t = current(kind)?.win?.tables
      if (t) return t.flatMap((x) => { const rect = parseRect(x.range); return rect ? [{ ...x, rect }] : [] })
    }
    return []
  })
  const notice = computed(() => {
    for (const kind of order) {
      const n = current(kind)?.win?.notice
      if (n) return n
    }
    return ''
  })

  /** The rectangle a slot must hold: the body under the scroll, frozen rows across the body, frozen columns down it. */
  function needOf(kind: SlotKind): Need | null {
    if (kind === 'body') return { r1: v.firstRow.value, r2: v.lastRow.value, c1: v.firstCol.value, c2: v.lastCol.value }
    if (kind === 'top') return v.frozenRows.value ? { r1: 1, r2: v.frozenRows.value, c1: 1, c2: v.lastCol.value } : null
    return v.frozenCols.value ? { r1: v.firstRow.value, r2: v.lastRow.value, c1: 1, c2: v.frozenCols.value } : null
  }
  /** Where the source stops giving rows and columns (it cut a window there); until then the grid asks past the data, where line styles reach. */
  let edge = { rows: Infinity, cols: Infinity }
  /** The edge in force: data grown by edits past it is asked for again. */
  const edgeRows = () => Math.max(edge.rows, s.dataShown.value)
  const edgeCols = () => Math.max(edge.cols, s.dataCols.value)
  /** Whether a slot covers a need; a window cut at the source's edge or by the cell limit counts, or loading would loop. */
  function covered(slot: Slot, need: Need): boolean {
    const w = slot.win
    if (!w || slot.view !== viewKey()) return false
    const lastRow = w.row + slot.rows - 1
    const lastCol = w.col + slot.cols - 1
    const rowsOk = need.r1 >= w.row && (need.r2 <= lastRow || lastRow >= edgeRows())
    const colsOk = need.c1 >= w.col && (need.c2 <= lastCol || lastCol >= edgeCols())
    return rowsOk && colsOk
  }

  async function load(kind: SlotKind, need: Need): Promise<void> {
    const slot = slots[kind].value
    const maxCells = s.source().maxCells ?? MAX_CELLS
    // Margins only up to the grid's end; past the source's known edge there is nothing to ask for.
    const lastCol = Math.max(1, Math.min(s.colCount.value, edgeCols()))
    const lastRow = Math.max(1, Math.min(s.rowCount.value, edgeRows()))
    const col = kind === 'left' ? 1 : Math.max(1, Math.min(need.c1, lastCol) - FETCH_COLS)
    const cols = kind === 'left' ? Math.min(need.c2, lastCol) : Math.min(MAX_COLS, Math.min(need.c2 + FETCH_COLS, lastCol) - col + 1)
    const row = kind === 'top' ? 1 : Math.max(1, Math.min(need.r1, lastRow) - FETCH_ROWS)
    const rows = kind === 'top' ? Math.min(need.r2, lastRow) : Math.min(Math.min(need.r2 + FETCH_ROWS, lastRow) - row + 1, Math.max(1, Math.floor(maxCells / cols)))
    const filter = s.visible.value
    const req: WindowRequest = { sheet: s.sheetName.value, row, rows, col, cols, ...(filter ? { filter: filter.spec } : {}) }
    const view = viewKey()
    const key = `${view}:${row}:${rows}:${col}:${cols}`
    // One request per slot at a time: dragging the scrollbar would otherwise send dozens, of which only the last matters.
    if (slot.pendingKey || key === slot.servedKey) return
    slot.pendingKey = key
    slot.ctrl = new AbortController()
    const mine = epoch
    // An edit may have replaced the slot object meanwhile: release the request on whichever holds it.
    const release = () => {
      if (slots[kind].value.pendingKey === key) slots[kind].value.pendingKey = ''
    }
    s.loading.value = true
    try {
      const got = await s.source().window(req, slot.ctrl.signal)
      if (mine !== epoch) return
      if ('pending' in got) {
        if (!awaitPending(got.pending.elapsedMs)) return
        setTimeout(() => {
          if (mine !== epoch) return
          release()
          ensure()
        }, PENDING_RETRY_MS)
        return
      }
      stopWait()
      takeSizes(got)
      const width = Math.max(got.widths.length, got.text[0]?.length ?? 0)
      // A window shorter than asked, at or past the data's edge: the source stops there, the grid will not ask past it.
      if (width < cols && req.col + width - 1 >= s.dataCols.value) edge = { ...edge, cols: req.col + width - 1 }
      if (got.text.length < rows && req.row + got.text.length - 1 >= s.dataShown.value) edge = { ...edge, rows: req.row + got.text.length - 1 }
      const merged = got.merges.flatMap((m) => { const r = parseRect(m); return r ? [r] : [] })
      slots[kind].value = { sheet: req.sheet, view, win: got, rows: got.text.length, cols: width, merges: merged, pendingKey: '', servedKey: key, ctrl: null }
      s.failure.value = null
    } catch (e) {
      if (mine !== epoch || isAbort(e)) return
      release()
      stopWait()
      s.failure.value = { error: e, text: e instanceof Error && e.message ? e.message : s.messages.value.loadFailed }
    } finally {
      if (mine === epoch && !wait.value) {
        s.loading.value = order.some((k) => slots[k].value.pendingKey !== '')
        // The scroll may have moved on while the request was out.
        if (!s.failure.value) ensure()
      }
    }
  }

  // ── a book being recomputed: the source answers `pending`, the sheet waits with a stopwatch and asks again ──
  const wait = ref<{ since: number; elapsedMs: number } | null>(null)
  const waitSeconds = ref(0)
  let waitTimer: ReturnType<typeof setInterval> | null = null
  let polls = 0
  /** Start or go on waiting; false (and a failure) once the polls run out. */
  function awaitPending(elapsedMs: number): boolean {
    if (!wait.value) {
      wait.value = { since: Date.now(), elapsedMs }
      polls = 0
      const tick = () => {
        const w = wait.value
        waitSeconds.value = w ? Math.floor((w.elapsedMs + Date.now() - w.since) / 1000) : 0
      }
      tick()
      waitTimer = setInterval(tick, 250)
    }
    if (++polls > PENDING_POLLS_MAX) {
      stopWait()
      s.failure.value = { error: new Error('recompute timeout'), text: s.messages.value.waitTimeout }
      return false
    }
    return true
  }
  function stopWait(): void {
    if (waitTimer) clearInterval(waitTimer)
    waitTimer = null
    wait.value = null
    waitSeconds.value = 0
  }

  /** Widths, heights and the default height from a window; rows above the frame keep it still when their heights change. */
  function takeSizes(w: SheetWindow): void {
    const widths = new Map(s.bookWidths.value)
    let widthsChanged = false
    w.widths.forEach((width, i) => {
      const px = colPx(width)
      if (widths.get(w.col + i) !== px) {
        widths.set(w.col + i, px)
        widthsChanged = true
      }
    })
    if (widthsChanged) s.bookWidths.value = widths
    const heights = new Map(s.heights.value)
    let heightsChanged = false
    const own = new Set<number>()
    for (const [key, pt] of Object.entries(w.rowHeights)) {
      const r = Number(key)
      own.add(r)
      if (heights.get(r) !== rowPx(pt)) {
        heights.set(r, rowPx(pt))
        heightsChanged = true
      }
    }
    for (let d = w.row; d < w.row + w.text.length; d++) {
      const r = s.rowOf(d)
      if (!own.has(r) && heights.delete(r)) heightsChanged = true
    }
    const def = rowPx(w.defaultRowHeight ?? DEFAULT_ROW_PT)
    if (!heightsChanged && def === s.defaultRowPx.value) return
    const anchor = geo.rowAt(s.scrollTop.value + v.frozenH.value)
    const before = geo.rowOffsets.value[anchor - 1] ?? 0
    if (heightsChanged) s.heights.value = heights
    s.defaultRowPx.value = def
    const after = geo.rowOffsets.value[anchor - 1] ?? 0
    const el = s.scroller.value
    if (el && after !== before && s.scrollTop.value > 0) {
      el.scrollTop += after - before
      s.scrollTop.value = el.scrollTop
    }
  }

  /** Load what the drawn piece needs and the slots do not hold. */
  function ensure(): void {
    if (s.failure.value || s.filterWait.value || !s.sheet.value || s.rowCount.value === 0) return
    for (const kind of order) {
      const need = needOf(kind)
      if (need && !covered(slots[kind].value, need)) void load(kind, need)
    }
  }

  /** Drop every window: answers in flight are aborted and ignored. */
  function reset(): void {
    epoch++
    edge = { rows: Infinity, cols: Infinity }
    stopWait()
    for (const kind of order) {
      slots[kind].value.ctrl?.abort()
      slots[kind].value = newSlot()
    }
    s.loading.value = false
  }
  /** Forget the sheet's sizes too (another sheet, or a reload). */
  function resetSizes(): void {
    s.bookWidths.value = new Map()
    s.heights.value = new Map()
    s.defaultRowPx.value = rowPx(DEFAULT_ROW_PT)
  }

  /** New cell states into the loaded windows; the rest are dropped, an unloaded window comes fresh anyway. */
  function applyChanges(changes: CellUpdate[]): void {
    const touched = new Set<SlotKind>()
    for (const ch of changes) {
      if (ch.sheet !== s.sheetName.value) continue
      const at = parseCell(ch.cell)
      if (!at) continue
      const addr = cellAddress(at.row, at.col)
      if (ch.text !== '') s.grow(at.row, at.col)
      const d = s.indexOf(at.row)
      if (d === undefined) continue
      for (const kind of order) {
        const slot = slots[kind].value
        const w = slot.win
        if (!w || slot.view !== viewKey()) continue
        if (d < w.row || d >= w.row + slot.rows || at.col < w.col || at.col >= w.col + slot.cols) continue
        const line = w.text[d - w.row]
        const values = w.values[d - w.row]
        if (line) line[at.col - w.col] = ch.text
        if (values) values[at.col - w.col] = ch.value
        if (ch.color) w.colors[addr] = ch.color
        else delete w.colors[addr]
        if (ch.formula === null) delete w.formulas[addr]
        else if (ch.formula !== undefined) w.formulas[addr] = ch.formula
        if (ch.editText !== undefined && w.editText) w.editText[addr] = ch.editText
        touched.add(kind)
      }
    }
    for (const kind of touched) slots[kind].value = { ...slots[kind].value }
  }

  return {
    slots,
    viewKey,
    slotAt,
    textAt,
    valueAt,
    styleAt,
    cellData,
    emptyAt,
    textOfRow,
    valueOfRow,
    formulaOfRow,
    emptyAtRow,
    editTextOfRow,
    merges,
    tables,
    notice,
    wait,
    waitSeconds,
    ensure,
    reset,
    resetSizes,
    applyChanges,
  }
}

export type Windows = ReturnType<typeof useWindows>
