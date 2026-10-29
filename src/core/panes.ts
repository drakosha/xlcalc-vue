// Excel's split window drawn (View > Split): two or four panes, each a scroller of its own over the whole grid. As in Excel,
// the panes of a row of panes scroll up and down together and those of a column of panes left and right together; the
// active pane is the sheet's scroller (keys, the editor, scrolling to the active cell), the others draw and scroll.
import { computed, ref, watch } from 'vue'
import { parseCell } from 'xlcalc/a1'
import { clamp } from '../rect'
import type { SheetSplit } from '../contract/source'
import type { Geometry } from './geometry'
import type { SheetState } from './state'
import { scrolledArea, type Viewport } from './viewport'
import type { Need, SlotKind } from './windows'

/** Excel's panes in its order (`Panes(1)` to `Panes(4)`): top left, top right, bottom left, bottom right. */
export type PaneId = 0 | 1 | 2 | 3
export const PANE_NAMES = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'] as const
/** A pane's row of panes (1 below the horizontal line) and column of panes (1 right of the vertical one). */
export const rowGroup = (p: PaneId): 0 | 1 => (p >> 1) as 0 | 1
export const colGroup = (p: PaneId): 0 | 1 => (p & 1) as 0 | 1
const paneAt = (rg: number, cg: number): PaneId => (rg * 2 + cg) as PaneId
/** The data window each pane reads into. */
const SLOTS: readonly SlotKind[] = ['corner', 'top', 'left', 'body']

/** Points in a CSS pixel: the file's split lines are points on the screen. */
export const PT = 0.75
/** The split bar's thickness on the screen. */
export const SPLIT_BAR_PX = 5

export interface Box {
  left: number
  top: number
  width: number
  height: number
}
export interface PaneBox {
  id: PaneId
  /** The pane in the split window, in screen pixels. */
  box: Box
  /**
   * Its scroller, relative to the pane: past the pane's top and left edges by the headings another pane shows (the letters
   * over a bottom pane, the numbers left of a right pane), and past its right and bottom by the scrollbars Excel draws only
   * at the window's right and bottom edges, so the pane shows neither.
   */
  scroller: Box
  /** The scroller's client area: its size less its scrollbars, which it always has. */
  client: { width: number; height: number }
}
export interface PaneLayout {
  panes: PaneBox[]
  /** The vertical and the horizontal split bar's left or top edge in the window; null without that line. */
  x: number | null
  y: number | null
}
export interface LayoutInput {
  width: number
  height: number
  /** The lines' distance from the window's left and top edges, as drawn on the screen; null for no line. */
  x: number | null
  y: number | null
  /** The headings on the screen: the column letters' height (the outline bar above them counted in), the row numbers' width. */
  headH: number
  leadW: number
  /** The scrollbars' thickness, 0 where they overlay the content. */
  barV: number
  barH: number
}

/** Where a split window's panes go: the lines at `x` and `y`, each followed by its bar, the panes' scrollers as Excel draws them. */
export function paneLayout(i: LayoutInput): PaneLayout {
  const x = i.x === null ? null : clamp(i.x, 0, Math.max(0, i.width - SPLIT_BAR_PX))
  const y = i.y === null ? null : clamp(i.y, 0, Math.max(0, i.height - SPLIT_BAR_PX))
  const spans = (line: number | null, size: number) => (line === null ? [{ start: 0, size }] : [{ start: 0, size: line }, { start: line + SPLIT_BAR_PX, size: Math.max(0, size - line - SPLIT_BAR_PX) }])
  const cols = spans(x, i.width)
  const rows = spans(y, i.height)
  const panes: PaneBox[] = []
  rows.forEach((r, rg) =>
    cols.forEach((c, cg) => {
      const hideTop = rg > 0 ? i.headH : 0
      const hideLeft = cg > 0 ? i.leadW : 0
      const width = c.size + hideLeft + (cg === cols.length - 1 ? 0 : i.barV)
      const height = r.size + hideTop + (rg === rows.length - 1 ? 0 : i.barH)
      panes.push({
        id: paneAt(rg, cg),
        box: { left: c.start, top: r.start, width: c.size, height: r.size },
        scroller: { left: hideLeft ? -hideLeft : 0, top: hideTop ? -hideTop : 0, width, height },
        client: { width: width - i.barV, height: height - i.barH },
      })
    }),
  )
  return { panes, x, y }
}

/** The edge nearest to `at` among prefix sums, by the index of the item under it: the count of items before that edge. */
function nearestEdge(offsets: Float64Array, at: number, indexAt: (x: number) => number): number {
  const k = indexAt(at)
  const a = offsets[k - 1] ?? 0
  const b = offsets[k] ?? a
  return at - a <= b - at ? k - 1 : k
}

interface PaneDeps {
  /** The active pane's size measured again (the sheet's viewport). */
  measure: () => void
  /** The active pane scrolled by the user: the sheet's own scroll handling. */
  scrolled: () => void
  /** Loads what the panes now draw. */
  ensure: () => void
}

export function usePanes(s: SheetState, geo: Geometry, v: Viewport, deps: PaneDeps) {
  /** The sheet's split with a line, unless frozen panes take its place. */
  const split = computed<SheetSplit | null>(() => {
    const x = v.bookFreeze.value ? null : s.sheet.value?.split
    return x && (x.x || x.y) ? x : null
  })
  /** The lines while a bar is dragged, in screen pixels: they take the split's place until the drag ends. */
  const drag = ref<{ x: number | null; y: number | null } | null>(null)
  /** The split window's size on the screen and the scrollbars' thickness, measured by the grid. */
  const box = ref({ width: 0, height: 0 })
  const scrollbars = ref({ v: 0, h: 0 })
  const lines = computed(() => drag.value ?? (split.value ? { x: split.value.x ? split.value.x / PT : null, y: split.value.y ? split.value.y / PT : null } : null))
  const layout = computed<PaneLayout | null>(() => {
    const l = lines.value
    if (!l || (l.x === null && l.y === null)) return null
    const z = s.zoom.value
    // unmeasured (a test's DOM) as the sheet's viewport falls back
    const width = box.value.width || 800
    const height = box.value.height || 600
    return paneLayout({ width, height, x: l.x, y: l.y, headH: v.headH.value * z, leadW: v.leadW.value * z, barV: scrollbars.value.v, barH: scrollbars.value.h })
  })
  const on = computed(() => layout.value !== null)
  const has = (id: PaneId) => !!layout.value?.panes.some((p) => p.id === id)

  // ── the panes' scroll: the active pane's rows and columns of panes scroll by the sheet's own scroll, the others by their own ──
  const active = ref<PaneId>(0)
  const tops = [ref(0), ref(0)]
  const lefts = [ref(0), ref(0)]
  const topOf = (rg: number): number => (rg === rowGroup(active.value) ? s.scrollTop.value : tops[rg]!.value)
  const leftOf = (cg: number): number => (cg === colGroup(active.value) ? s.scrollLeft.value : lefts[cg]!.value)
  function setTop(rg: number, px: number): void {
    if (rg === rowGroup(active.value)) s.scrollTop.value = px
    else tops[rg]!.value = px
  }
  function setLeft(cg: number, px: number): void {
    if (cg === colGroup(active.value)) s.scrollLeft.value = px
    else lefts[cg]!.value = px
  }

  /** Each pane's viewport: the sheet's, with the rows and columns under the pane's own scroll and its own size. */
  const viewports = ([0, 1, 2, 3] as const).map((id): Viewport => {
    const client = computed(() => layout.value?.panes.find((p) => p.id === id)?.client ?? null)
    const area = scrolledArea(
      {
        rowCount: s.rowCount,
        colCount: s.colCount,
        scrollTop: computed(() => topOf(rowGroup(id))),
        scrollLeft: computed(() => leftOf(colGroup(id))),
        viewH: computed(() => (id === active.value ? s.viewH.value : (client.value?.height ?? 0) / s.zoom.value)),
        viewW: computed(() => (id === active.value ? s.viewW.value : (client.value?.width ?? 0) / s.zoom.value)),
      },
      geo,
      v,
    )
    return { ...v, ...area }
  })
  /** The range each pane's window must hold; null without a split, when the windows follow the sheet's viewport. */
  const needs = computed<Partial<Record<SlotKind, Need>> | null>(() => {
    const lay = layout.value
    if (!lay) return null
    const out: Partial<Record<SlotKind, Need>> = {}
    for (const { id } of lay.panes) {
      const a = viewports[id]!
      out[SLOTS[id]!] = { r1: a.firstRow.value, r2: a.lastRow.value, c1: a.firstCol.value, c2: a.lastCol.value }
    }
    return out
  })

  // ── the panes' scrollers: each one's scroll follows its rows and columns of panes ──
  const els: (HTMLElement | null)[] = [null, null, null, null]
  /** The scroll last set on a pane by the sheet: its scroll event is that echo, not the user's. */
  const echo = new WeakMap<HTMLElement, { top: number; left: number }>()
  function setEl(id: PaneId, el: HTMLElement | null): void {
    const was = els[id]
    if (el === was) return
    els[id] = el
    if (id === active.value) {
      if (el) s.scroller.value = el
      else if (s.scroller.value === was) s.scroller.value = null
    }
    if (el) sync()
  }
  /** Scrolls every pane to its rows' and columns' scroll. */
  function sync(): void {
    const lay = layout.value
    if (!lay) return
    const z = s.zoom.value
    for (const { id } of lay.panes) {
      const el = els[id]
      if (!el) continue
      const top = topOf(rowGroup(id)) * z
      const left = leftOf(colGroup(id)) * z
      if (Math.abs(el.scrollTop - top) < 1 && Math.abs(el.scrollLeft - left) < 1) continue
      el.scrollTop = top
      el.scrollLeft = left
      echo.set(el, { top: el.scrollTop, left: el.scrollLeft })
    }
  }
  // Without a split this reads nothing, so the single window's scrolling never runs it.
  watch(() => (layout.value ? [layout.value, s.zoom.value, active.value, s.scrollTop.value, s.scrollLeft.value, tops[0]!.value, tops[1]!.value, lefts[0]!.value, lefts[1]!.value] : null), sync, { flush: 'post' })
  /** A pane scrolled: the active one as the sheet scrolls; another moves its rows and columns of panes, which follow it. */
  function onScroll(id: PaneId): void {
    const el = els[id]
    if (!el) return
    const last = echo.get(el)
    if (last) {
      echo.delete(el)
      if (Math.abs(el.scrollTop - last.top) < 1 && Math.abs(el.scrollLeft - last.left) < 1) return
    }
    if (id === active.value) return deps.scrolled()
    setTop(rowGroup(id), el.scrollTop / s.zoom.value)
    setLeft(colGroup(id), el.scrollLeft / s.zoom.value)
    deps.ensure()
  }
  /** Makes a pane active, as a click in it does: it takes the sheet's scroll, keys and editor; the panes stay where they are. */
  function activate(id: PaneId): void {
    const was = active.value
    if (id === was || !has(id)) return
    const [r0, r1, c0, c1] = [rowGroup(was), rowGroup(id), colGroup(was), colGroup(id)]
    const top = r1 === r0 ? s.scrollTop.value : tops[r1]!.value
    const left = c1 === c0 ? s.scrollLeft.value : lefts[c1]!.value
    if (r1 !== r0) tops[r0]!.value = s.scrollTop.value
    if (c1 !== c0) lefts[c0]!.value = s.scrollLeft.value
    active.value = id
    s.scrollTop.value = top
    s.scrollLeft.value = left
    if (els[id]) s.scroller.value = els[id]
    deps.measure()
  }

  // ── the split appearing, changing and going: Excel's panes keep their scroll ──
  const rowOffset = (row: number) => geo.rowOffsets.value[(s.indexOf(row) ?? row) - 1] ?? 0
  const colOffset = (col: number) => geo.offsets.value[col - 1] ?? 0
  /** The pane a split names active, on the lines it has. */
  const named = (x: SheetSplit): PaneId => {
    const p = Math.max(0, PANE_NAMES.indexOf(x.activePane ?? 'topLeft')) as PaneId
    return paneAt(x.y ? rowGroup(p) : 0, x.x ? colGroup(p) : 0)
  }
  /** Sets the scroll of both rows and columns of panes and the active pane. */
  function place(rows: number[], cols: number[], id: PaneId): void {
    tops[0]!.value = rows[0]!
    tops[1]!.value = rows[1]!
    lefts[0]!.value = cols[0]!
    lefts[1]!.value = cols[1]!
    active.value = id
    s.scrollTop.value = rows[rowGroup(id)]!
    s.scrollLeft.value = cols[colGroup(id)]!
    if (els[id]) s.scroller.value = els[id]
  }
  /** The split the panes follow, and whether frozen panes had taken its place. */
  let shown: { split: SheetSplit | null; frozen: boolean } = { split: null, frozen: false }
  /** A sheet made active, at its top left corner: the second panes from the split's `topLeft`, its active pane active. */
  function enter(): void {
    const x = split.value
    shown = { split: x, frozen: v.bookFreeze.value }
    drag.value = null
    const at = parseCell(x?.topLeft ?? '')
    const rows = [0, x?.y && at ? rowOffset(at.row) : 0]
    const cols = [0, x?.x && at ? colOffset(at.col) : 0]
    place(rows, cols, x ? named(x) : 0)
  }
  /**
   * The split changed on the same sheet, as Excel's panes follow it (measured in Excel for Mac): a new line's second panes
   * scroll from the split's `topLeft`; a removed vertical line keeps the left panes' columns, a removed horizontal one the
   * top panes' rows, and with no line left the window keeps the top left pane's scroll. Frozen panes in its place start at
   * the top left, past the frozen rows and columns, and the split back from under them at the top left too, since they
   * were frozen from the sheet's first row and column. Another sheet made active has entered already (`enter`); the same
   * sheet renamed keeps its panes.
   */
  function follow(): void {
    const prev = shown.split
    const next = split.value
    const thawed = shown.frozen && !v.bookFreeze.value
    shown = { split: next, frozen: v.bookFreeze.value }
    if (prev === next) return
    const rows = [topOf(0), topOf(1)]
    const cols = [leftOf(0), leftOf(1)]
    let id = active.value
    const at = parseCell(next?.topLeft ?? '')
    if (next?.y && !prev?.y) rows[1] = at ? rowOffset(at.row) : rows[0]!
    if (next?.x && !prev?.x) cols[1] = at ? colOffset(at.col) : cols[0]!
    if (!prev && next) id = named(next)
    id = paneAt(next?.y ? rowGroup(id) : 0, next?.x ? colGroup(id) : 0)
    if ((!next && v.bookFreeze.value) || (!prev && thawed)) {
      rows[0] = 0
      cols[0] = 0
    }
    place(rows, cols, id)
    deps.measure()
  }
  watch(split, follow)

  // ── the bars: dragged to a row's or a column's edge, removed at the window's edge ──
  /** A line dragged to `px` on the screen, snapped to the nearest edge of a row or column in the panes before it. */
  function snap(axis: 'x' | 'y', px: number): number {
    const z = s.zoom.value
    const head = axis === 'x' ? v.leadW.value : v.headH.value
    const scroll = axis === 'x' ? leftOf(0) : topOf(0)
    const offsets = axis === 'x' ? geo.offsets.value : geo.rowOffsets.value
    const at = scroll + px / z - head
    const k = nearestEdge(offsets, at, axis === 'x' ? geo.colAt : geo.rowAt)
    return Math.max(0, ((offsets[k] ?? 0) - scroll + head) * z)
  }
  /** Whether a line at `px` leaves a pane without cells (on the headings, or at the far edge), where Excel removes it. */
  function offEdge(axis: 'x' | 'y', px: number): boolean {
    const z = s.zoom.value
    const head = (axis === 'x' ? v.leadW.value : v.headH.value) * z
    const size = axis === 'x' ? box.value.width || 800 : box.value.height || 600
    const bar = axis === 'x' ? scrollbars.value.v : scrollbars.value.h
    return px <= head + 1 || px >= size - SPLIT_BAR_PX - bar - 2
  }

  /**
   * Freeze Panes on a split window freezes at its lines, as in Excel: the rows above the horizontal line and the columns to
   * the left of the vertical one, counted from the sheet's first; null without a split.
   */
  function freezeLines(): { rows: number; cols: number } | null {
    const lay = layout.value
    if (!lay) return null
    const z = s.zoom.value
    return {
      rows: lay.y === null ? 0 : nearestEdge(geo.rowOffsets.value, topOf(0) + lay.y / z - v.headH.value, geo.rowAt),
      cols: lay.x === null ? 0 : nearestEdge(geo.offsets.value, leftOf(0) + lay.x / z - v.leadW.value, geo.colAt),
    }
  }
  /** Before row heights change: the row at the top of the rows of panes that are not active, to keep them steady afterwards. */
  function hold(): (() => void) | null {
    const lay = layout.value
    if (!lay || lay.y === null) return null
    const rg = 1 - rowGroup(active.value)
    const top = tops[rg]!.value
    if (top <= 0) return null
    const anchor = geo.rowAt(top)
    const before = geo.rowOffsets.value[anchor - 1] ?? 0
    return () => {
      const after = geo.rowOffsets.value[anchor - 1] ?? 0
      if (after !== before) tops[rg]!.value = Math.max(0, top + after - before)
    }
  }

  return { split, drag, box, scrollbars, layout, on, active, viewports, needs, topOf, leftOf, setEl, onScroll, activate, enter, snap, offEdge, freezeLines, hold }
}

export type Panes = ReturnType<typeof usePanes>
