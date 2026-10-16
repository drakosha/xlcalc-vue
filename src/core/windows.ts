// Windows of data: three slots (body, frozen rows, frozen columns), loaded under the scroll, patched by edits, read again in place by a repaint.
// Rows of windows are display indexes (`d`): with a display filter they are places in its list, not the sheet's rows.
import { computed, nextTick, ref, shallowRef, triggerRef, type ShallowRef } from 'vue'
import { cellAddress, parseCell, parseRange } from 'xlcalc/a1'
import type { BarLook, CellRect, CellStyle, CellValue, ConditionalAppearance } from 'xlcalc/contract'
import type { CellUpdate, SheetWindow, WindowRequest } from '../source'
import { lookStyle } from './conditional'
import { DEFAULT_ROW_PT, rowPx, type Geometry } from './geometry'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** Fetch margins: a window is taken with spare rows and columns so an ordinary scroll needs no request. */
const FETCH_ROWS = 120
const FETCH_COLS = 12
const MAX_CELLS = 20_000
const MAX_COLS = 200
/** Delay before asking again for a window that is not ready yet, and how many times (over a host's recompute limit). */
const PENDING_RETRY_MS = 1000
const PENDING_POLLS_MAX = 200
/** A repaint that follows the last one closer than this waits as long for more before asking. */
const REPAINT_DELAY_MS = 80

export type SlotKind = 'body' | 'top' | 'left'

/** A rectangle of a window by display rows. */
interface Area {
  row: number
  rows: number
  col: number
  cols: number
}

/** The request in flight: its view and rectangle (a repaint may meet it), whether it reads a drawn window again in place, and the changes that came meanwhile. */
interface Asked extends Area {
  view: string
  /** A window on screen read again for a repaint: it stays drawn meanwhile, so there is no loading mark. */
  quiet: boolean
  /** Cell updates that came while the request was out: the answer may have been read before them, so they go over it. */
  missed: CellUpdate[]
}

/** One window: what is loaded, the request in flight and the one already answered (so neither repeats). */
export interface Slot {
  /** The sheet and the filter the window was asked under: another filter means other rows. */
  view: string
  win: SheetWindow | null
  rows: number
  cols: number
  merges: CellRect[]
  pendingKey: string
  servedKey: string
  ctrl: AbortController | null
  asked: Asked | null
  /** A repaint met the window: it is read again, drawn as it is until the answer comes. */
  stale: boolean
}

interface Need {
  r1: number
  r2: number
  c1: number
  c2: number
}

const newSlot = (): Slot => ({ view: '', win: null, rows: 0, cols: 0, merges: [], pendingKey: '', servedKey: '', ctrl: null, asked: null, stale: false })
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
  /** What conditional formats give the cell; a window's old `bars` come as a bar of it. */
  look: ConditionalAppearance | undefined
  loaded: boolean
}

/** A cell's look: its appearance, else its share of the old `bars` as a gradient bar from the left edge. */
function lookOf(w: SheetWindow, addr: string): ConditionalAppearance | undefined {
  const look = w.appearance?.[addr]
  if (look) return look
  const old = w.bars?.[addr]
  if (!old) return undefined
  const bar: BarLook = { start: 0, end: old.ratio, color: old.color, gradient: true }
  return { bar }
}
/** The style a window's cell shows: its own (or its row's, column's), with a conditional format's over it. */
function shownStyle(w: SheetWindow, addr: string, look: ConditionalAppearance | undefined): CellStyle | undefined {
  const id = w.styleIds[addr]
  return lookStyle(id === undefined ? undefined : w.styles[id], look?.style)
}

export function useWindows(s: SheetState, v: Viewport, geo: Geometry) {
  const slots: Record<SlotKind, ShallowRef<Slot>> = { body: shallowRef(newSlot()), top: shallowRef(newSlot()), left: shallowRef(newSlot()) }
  const order: SlotKind[] = ['body', 'top', 'left']
  /** Generation of the book and sheet: an answer sent before a switch must not overwrite fresh data. */
  let epoch = 0

  /** The sheet and filter the windows must be of now. */
  const viewKey = () => `${s.sheetName.value}\u0000${s.visible.value?.key ?? ''}`
  /**
   * Cells change in place inside their windows: `revision` counts every change, `drawnRevision` only those of drawn
   * cells and their neighbours. The grid reads cells under the second, so a change off screen does not draw it again.
   */
  const revision = shallowRef(0)
  const drawnRevision = shallowRef(0)
  /** The window holding a cell by display row: the body first, then the frozen ones (they meet at the corner). */
  function slotAt(d: number, col: number): Slot | null {
    void revision.value
    return holder(d, col)
  }
  /** The same for drawing a cell. */
  function drawnSlot(d: number, col: number): Slot | null {
    void drawnRevision.value
    return holder(d, col)
  }
  function holder(d: number, col: number): Slot | null {
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
  /** A cell's style as drawn; records of a window are keyed by the cell's own address (`addr`, by its sheet row). */
  function styleAt(d: number, col: number, addr: string): CellStyle | undefined {
    const w = drawnSlot(d, col)?.win
    return w ? shownStyle(w, addr, w.appearance?.[addr]) : undefined
  }
  function cellData(d: number, col: number, addr: string): CellData {
    const w = drawnSlot(d, col)?.win
    if (!w) return { text: '', value: undefined, style: undefined, formula: undefined, color: undefined, link: undefined, note: undefined, look: undefined, loaded: false }
    const look = lookOf(w, addr)
    return {
      text: w.text[d - w.row]?.[col - w.col] ?? '',
      value: w.values[d - w.row]?.[col - w.col],
      style: shownStyle(w, addr, look),
      formula: w.formulas[addr],
      color: w.colors[addr],
      link: w.links?.[addr],
      note: w.notes?.[addr],
      look,
      loaded: true,
    }
  }
  /** Whether a cell is known to be empty: loaded, no text, no value. */
  function emptyAt(d: number, col: number): boolean {
    const slot = drawnSlot(d, col)
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
  /** The style a cell shows and its number format code (none: General), from its loaded window; null when none holds it. */
  /** Whether a cell (a sheet row) is a drill-down of list mode: loaded, marked in its window's `drill` and not empty. */
  const drillOfRow = (r: number, col: number): boolean => {
    const d = s.indexOf(r)
    const slot = d === undefined ? null : drawnSlot(d, col)
    const w = slot?.win
    return !!w?.drill?.[cellAddress(r, col)] && (w.text[d! - w.row]?.[col - w.col] ?? '') !== ''
  }
  const lookOfRow = (r: number, col: number): { style?: CellStyle; numberFormat?: string } | null => {
    const d = s.indexOf(r)
    const w = d === undefined ? undefined : slotAt(d, col)?.win
    if (!w) return null
    const addr = cellAddress(r, col)
    const id = w.styleIds[addr]
    const style = id === undefined ? undefined : w.styles[id]
    const numberFormat = w.numberFormats?.[addr]
    return { ...(style ? { style } : {}), ...(numberFormat !== undefined ? { numberFormat } : {}) }
  }

  const current = (kind: SlotKind) => (slots[kind].value.view === viewKey() ? slots[kind].value : null)
  /** Merged areas of the loaded windows (by sheet rows), each once. */
  const merges = computed(() => {
    const seen = new Set<string>()
    const out: CellRect[] = []
    for (const kind of order) {
      for (const m of current(kind)?.merges ?? []) {
        const key = `${m.top}:${m.left}`
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
      if (t) return t.flatMap((x) => { const rect = parseRange(x.range); return rect ? [{ ...x, rect }] : [] })
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
    if (!w || slot.stale || slot.view !== viewKey()) return false
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
    const ctrl = new AbortController()
    const quiet = slot.stale && !!slot.win && slot.view === view
    slot.pendingKey = key
    slot.ctrl = ctrl
    slot.asked = { view, row, rows, col, cols, quiet, missed: [] }
    const mine = epoch
    // An edit may have replaced the slot object meanwhile, a repaint may have aborted the request for another: it is the slot's while its controller is.
    const current = () => slots[kind].value.ctrl === ctrl
    const release = () => {
      if (!current()) return
      const held = slots[kind].value
      held.pendingKey = ''
      held.ctrl = null
      held.asked = null
    }
    if (!quiet) s.loading.value = true
    try {
      const got = await s.source().window(req, ctrl.signal)
      if (mine !== epoch || !current()) return
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
      const merged = got.merges.flatMap((m) => { const r = parseRange(m); return r ? [r] : [] })
      const next: Slot = { view, win: got, rows: got.text.length, cols: width, merges: merged, pendingKey: '', servedKey: key, ctrl: null, asked: null, stale: false }
      for (const ch of slots[kind].value.asked?.missed ?? []) patch(next, ch)
      slots[kind].value = next
      s.failure.value = null
    } catch (e) {
      if (mine !== epoch || isAbort(e) || !current()) return
      release()
      stopWait()
      s.failure.value = { error: e, text: e instanceof Error && e.message ? e.message : s.messages.value.loadFailed }
    } finally {
      if (mine === epoch && !wait.value) {
        s.loading.value = order.some((k) => slots[k].value.pendingKey !== '' && !slots[k].value.asked?.quiet)
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

  /** Widths, heights and the default height from a window, changed in place; rows above the frame keep it still when their heights change. */
  function takeSizes(w: SheetWindow): void {
    const widths = s.bookWidths.value
    let widthsChanged = false
    w.widths.forEach((width, i) => {
      if (widths.get(w.col + i) !== width) {
        widths.set(w.col + i, width)
        widthsChanged = true
      }
    })
    if (widthsChanged) triggerRef(s.bookWidths)
    const heights = s.heights.value
    const own = new Set<number>()
    const next = new Map<number, number>()
    for (const [key, pt] of Object.entries(w.rowHeights)) {
      const r = Number(key)
      own.add(r)
      if (heights.get(r) !== rowPx(pt)) next.set(r, rowPx(pt))
    }
    const gone: number[] = []
    for (let d = w.row; d < w.row + w.text.length; d++) {
      const r = s.rowOf(d)
      if (!own.has(r) && heights.has(r)) gone.push(r)
    }
    const def = rowPx(w.defaultRowHeight ?? DEFAULT_ROW_PT)
    if (!next.size && !gone.length && def === s.defaultRowPx.value) return
    const anchor = geo.rowAt(s.scrollTop.value + v.frozenH.value)
    const before = geo.rowOffsets.value[anchor - 1] ?? 0
    for (const [r, px] of next) heights.set(r, px)
    for (const r of gone) heights.delete(r)
    if (next.size || gone.length) triggerRef(s.heights)
    s.defaultRowPx.value = def
    const after = geo.rowOffsets.value[anchor - 1] ?? 0
    const el = s.scroller.value
    if (el && after !== before && s.scrollTop.value > 0) {
      // The scroll the grid draws by moves now; the frame follows once the rows are drawn at their new heights, one write and no read.
      const top = Math.max(0, s.scrollTop.value + after - before)
      s.scrollTop.value = top
      void nextTick(() => {
        if (s.scrollTop.value === top) el.scrollTop = top
      })
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
    repaints = []
    if (repaintTimer) clearTimeout(repaintTimer)
    repaintTimer = null
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

  /** A cell's new state into a slot's window, in place; its display row when that changed anything, else null (the window does not hold the cell, or it was so already). */
  function patch(slot: Slot, ch: CellUpdate): number | null {
    const w = slot.win
    const at = parseCell(ch.cell)
    const d = at ? s.indexOf(at.row) : undefined
    if (!w || !at || d === undefined) return null
    if (d < w.row || d >= w.row + slot.rows || at.col < w.col || at.col >= w.col + slot.cols) return null
    const addr = cellAddress(at.row, at.col)
    const line = w.text[d - w.row]
    const values = w.values[d - w.row]
    const i = at.col - w.col
    let changed = false
    const set = <T>(rec: Record<string, T>, value: T | undefined) => {
      if (value === undefined ? !(addr in rec) : rec[addr] === value) return
      if (value === undefined) delete rec[addr]
      else rec[addr] = value
      changed = true
    }
    if (line && line[i] !== ch.text) {
      line[i] = ch.text
      changed = true
    }
    if (values && values[i] !== ch.value) {
      values[i] = ch.value
      changed = true
    }
    set(w.colors, ch.color || undefined)
    if (ch.formula !== undefined) set(w.formulas, ch.formula ?? undefined)
    if (ch.editText !== undefined && w.editText) set(w.editText, ch.editText)
    if (w.numberFormats && ch.numberFormat !== undefined) set(w.numberFormats, ch.numberFormat ?? undefined)
    return changed ? d : null
  }
  /** Whether a display row and a column are drawn, or next to what is (a cell's bottom and right lines are its neighbours' too). */
  function drawnNear(d: number, col: number): boolean {
    const rowOk = d <= v.frozenRows.value + 1 || (d >= v.firstRow.value - 1 && d <= v.lastRow.value + 1)
    return rowOk && (col <= v.frozenCols.value + 1 || (col >= v.firstCol.value - 1 && col <= v.lastCol.value + 1))
  }

  /** New cell states into the loaded windows and over the answers of requests out; the rest are dropped, an unloaded window comes fresh anyway. */
  function applyChanges(changes: CellUpdate[]): void {
    let changed = false
    let drawn = false
    const view = viewKey()
    for (const ch of changes) {
      if (ch.sheet !== s.sheetName.value) continue
      const at = parseCell(ch.cell)
      if (!at) continue
      if (ch.text !== '') s.grow(at.row, at.col)
      for (const kind of order) {
        const slot = slots[kind].value
        if (slot.asked?.view === view) slot.asked.missed.push(ch)
        const d = slot.view === view ? patch(slot, ch) : null
        if (d === null) continue
        changed = true
        drawn ||= drawnNear(d, at.col)
      }
    }
    if (changed) revision.value++
    if (drawn) drawnRevision.value++
  }

  /** Whether an area of display rows holds a cell of these rectangles (by sheet rows): under a filter, by the rows of its places. */
  function meets(area: Area, rects: CellRect[]): boolean {
    const last = area.row + area.rows - 1
    for (const r of rects) {
      if (r.right < area.col || r.left > area.col + area.cols - 1) continue
      if (!s.visible.value) {
        if (r.bottom >= area.row && r.top <= last) return true
        continue
      }
      for (let d = area.row; d <= last; d++) {
        const row = s.rowOf(d)
        if (row >= r.top && row <= r.bottom) return true
      }
    }
    return false
  }

  /** Ranges of repaints that came in the same task, read together. */
  let repaints: { sheet: string; range: string }[] = []
  /** When the last repaint was read, and the timer of a repaint put off. */
  let repaintAt = -Infinity
  let repaintTimer: ReturnType<typeof setTimeout> | null = null
  /**
   * Read again, in place, the windows that meet these ranges of the sheet shown (cells that may look otherwise after a change):
   * each stays drawn until its answer comes, under the same display rows; a request out that meets them is aborted and sent
   * again, as it may have been read before the change. Ranges of other sheets are read when their sheet is shown. Repaints
   * of one task are read together; one that follows another closely asks once they stop.
   */
  function repaint(ranges: { sheet: string; range: string }[]): void {
    if (!repaints.length) queueMicrotask(takeRepaints)
    repaints.push(...ranges)
  }
  function takeRepaints(): void {
    const ranges = repaints
    repaints = []
    const sheet = s.sheetName.value
    const rects = ranges.flatMap((x) => {
      const r = x.sheet === sheet ? parseRange(x.range) : null
      return r ? [r] : []
    })
    if (!rects.length) return
    const view = viewKey()
    let again = false
    for (const kind of order) {
      const slot = slots[kind].value
      const out = slot.asked?.view === view ? slot.asked : null
      const area = out ?? (slot.win && slot.view === view ? { row: slot.win.row, rows: slot.rows, col: slot.win.col, cols: slot.cols } : null)
      if (!area || !meets(area, rects)) continue
      if (out) {
        slot.ctrl?.abort()
        slot.ctrl = null
        slot.pendingKey = ''
        slot.asked = null
      }
      slot.stale = true
      slot.servedKey = ''
      again = true
    }
    if (!again) return
    const now = Date.now()
    const soon = now - repaintAt < REPAINT_DELAY_MS
    repaintAt = now
    if (repaintTimer) clearTimeout(repaintTimer)
    repaintTimer = null
    if (!soon) return ensure()
    repaintTimer = setTimeout(() => {
      repaintTimer = null
      ensure()
    }, REPAINT_DELAY_MS)
  }

  return {
    slots,
    viewKey,
    revision,
    slotAt,
    textAt,
    valueAt,
    styleAt,
    cellData,
    emptyAt,
    textOfRow,
    drillOfRow,
    valueOfRow,
    formulaOfRow,
    emptyAtRow,
    editTextOfRow,
    lookOfRow,
    merges,
    tables,
    notice,
    wait,
    waitSeconds,
    ensure,
    reset,
    resetSizes,
    applyChanges,
    repaint,
  }
}

export type Windows = ReturnType<typeof useWindows>
