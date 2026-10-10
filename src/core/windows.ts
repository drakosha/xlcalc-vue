// Windows of data: three slots (body, frozen rows, frozen columns), loaded under the scroll and patched by edits.
import { computed, shallowRef, type ShallowRef } from 'vue'
import type { CellStyle, CellValue } from 'xlcalc/contract'
import type { CellUpdate, SheetWindow, WindowRequest } from '../source'
import { parseCell, parseRect, type Rect } from './a1'
import { colPx, DEFAULT_ROW_PT, rowPx, type Geometry } from './geometry'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** Fetch margins: a window is taken with spare rows and columns so an ordinary scroll needs no request. */
const FETCH_ROWS = 120
const FETCH_COLS = 12
const MAX_CELLS = 20_000
const MAX_COLS = 200
/** Delay before asking again for a window that is not ready yet. */
const PENDING_RETRY_MS = 1000

export type SlotKind = 'body' | 'top' | 'left'

/** One window: what is loaded, the request in flight and the one already answered (so neither repeats). */
export interface Slot {
  sheet: string
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

const newSlot = (): Slot => ({ sheet: '', win: null, rows: 0, cols: 0, merges: [], pendingKey: '', servedKey: '', ctrl: null })
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

  /** The window holding a cell: the body first, then the frozen ones (they meet at the corner). */
  function slotAt(row: number, col: number): Slot | null {
    const name = s.sheetName.value
    for (const kind of order) {
      const slot = slots[kind].value
      const w = slot.win
      if (!w || slot.sheet !== name) continue
      if (row >= w.row && row < w.row + slot.rows && col >= w.col && col < w.col + slot.cols) return slot
    }
    return null
  }
  function textAt(row: number, col: number): string {
    const slot = slotAt(row, col)
    return slot ? (slot.win!.text[row - slot.win!.row]?.[col - slot.win!.col] ?? '') : ''
  }
  function valueAt(row: number, col: number): CellValue | undefined {
    const slot = slotAt(row, col)
    return slot ? slot.win!.values[row - slot.win!.row]?.[col - slot.win!.col] : undefined
  }
  function styleAt(row: number, col: number, addr: string): CellStyle | undefined {
    const w = slotAt(row, col)?.win
    const id = w?.styleIds[addr]
    return id === undefined ? undefined : w!.styles[id]
  }
  function cellData(row: number, col: number, addr: string): CellData {
    const w = slotAt(row, col)?.win
    if (!w) return { text: '', value: undefined, style: undefined, formula: undefined, color: undefined, link: undefined, note: undefined, bar: undefined, loaded: false }
    const id = w.styleIds[addr]
    return {
      text: w.text[row - w.row]?.[col - w.col] ?? '',
      value: w.values[row - w.row]?.[col - w.col],
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
  function emptyAt(row: number, col: number): boolean {
    const slot = slotAt(row, col)
    if (!slot) return false
    const w = slot.win!
    const value = w.values[row - w.row]?.[col - w.col]
    return (w.text[row - w.row]?.[col - w.col] ?? '') === '' && (value === null || value === undefined || value === '')
  }

  const current = (kind: SlotKind) => (slots[kind].value.sheet === s.sheetName.value ? slots[kind].value : null)
  /** Merged areas of the loaded windows, each once. */
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
  /** The need cut to the data: past it every cell is empty and there is nothing to ask for. */
  function clip(need: Need): Need {
    const rows = Math.max(1, s.dataRows.value)
    const cols = Math.max(1, s.dataCols.value)
    return { r1: Math.min(need.r1, rows), r2: Math.min(need.r2, rows), c1: Math.min(need.c1, cols), c2: Math.min(need.c2, cols) }
  }
  /** Whether a slot covers a need; a window cut at the data's end or by the cell limit counts, or loading would loop. */
  function covered(slot: Slot, need: Need): boolean {
    const w = slot.win
    if (!w || slot.sheet !== s.sheetName.value) return false
    const n = clip(need)
    const lastRow = w.row + slot.rows - 1
    const lastCol = w.col + slot.cols - 1
    const rowsOk = n.r1 >= w.row && (n.r2 <= lastRow || lastRow >= s.dataRows.value)
    const colsOk = n.c1 >= w.col && (n.c2 <= lastCol || lastCol >= s.dataCols.value)
    return rowsOk && colsOk
  }

  async function load(kind: SlotKind, need: Need): Promise<void> {
    const slot = slots[kind].value
    const n = clip(need)
    const maxCells = s.source().maxCells ?? MAX_CELLS
    const col = kind === 'left' ? 1 : Math.max(1, n.c1 - FETCH_COLS)
    const cols = kind === 'left' ? n.c2 : Math.min(MAX_COLS, n.c2 - col + 1 + FETCH_COLS)
    const row = kind === 'top' ? 1 : Math.max(1, n.r1 - FETCH_ROWS)
    const rows = kind === 'top' ? n.r2 : Math.min(n.r2 - row + 1 + FETCH_ROWS, Math.max(1, Math.floor(maxCells / cols)))
    const req: WindowRequest = { sheet: s.sheetName.value, row, rows, col, cols }
    const key = `${req.sheet}:${row}:${rows}:${col}:${cols}`
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
        setTimeout(() => {
          if (mine !== epoch) return
          release()
          ensure()
        }, PENDING_RETRY_MS)
        return
      }
      takeSizes(got)
      const width = Math.max(got.widths.length, got.text[0]?.length ?? 0)
      const merged = got.merges.flatMap((m) => { const r = parseRect(m); return r ? [r] : [] })
      slots[kind].value = { sheet: req.sheet, win: got, rows: got.text.length, cols: width, merges: merged, pendingKey: '', servedKey: key, ctrl: null }
      s.failure.value = null
    } catch (e) {
      if (mine !== epoch || isAbort(e)) return
      release()
      s.failure.value = { error: e, text: e instanceof Error && e.message ? e.message : s.messages.value.loadFailed }
    } finally {
      if (mine === epoch) {
        s.loading.value = order.some((k) => slots[k].value.pendingKey !== '')
        // The scroll may have moved on while the request was out.
        if (!s.failure.value) ensure()
      }
    }
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
    for (let r = w.row; r < w.row + w.text.length; r++) if (!own.has(r) && heights.delete(r)) heightsChanged = true
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
    if (s.failure.value || !s.sheet.value || s.rowCount.value === 0) return
    for (const kind of order) {
      const need = needOf(kind)
      if (need && !covered(slots[kind].value, need)) void load(kind, need)
    }
  }

  /** Drop every window: answers in flight are aborted and ignored. */
  function reset(): void {
    epoch++
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
      if (ch.text !== '') s.grow(at.row, at.col)
      for (const kind of order) {
        const slot = slots[kind].value
        const w = slot.win
        if (!w || slot.sheet !== ch.sheet) continue
        if (at.row < w.row || at.row >= w.row + slot.rows || at.col < w.col || at.col >= w.col + slot.cols) continue
        const line = w.text[at.row - w.row]
        const values = w.values[at.row - w.row]
        if (line) line[at.col - w.col] = ch.text
        if (values) values[at.col - w.col] = ch.value
        if (ch.color) w.colors[ch.cell] = ch.color
        else delete w.colors[ch.cell]
        if (ch.formula === null) delete w.formulas[ch.cell]
        else if (ch.formula !== undefined) w.formulas[ch.cell] = ch.formula
        touched.add(kind)
      }
    }
    for (const kind of touched) slots[kind].value = { ...slots[kind].value }
  }

  return { slots, slotAt, textAt, valueAt, styleAt, cellData, emptyAt, merges, tables, notice, ensure, reset, resetSizes, applyChanges }
}

export type Windows = ReturnType<typeof useWindows>
