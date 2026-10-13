// Selection as in Excel: a rectangle from the active cell to the other corner, by cells, rows and columns, with Shift and drag.
// It is kept in the sheet's rows; keys move along the display rows, so under a filter the next row is the next shown one.
// Whole columns and rows reach Excel's last row and column, as Excel selects them; what is drawn stops at the grid's end.
import { computed, nextTick, watch } from 'vue'
import { cellAddress, MAX_COLS, MAX_ROWS, parseCell, rangeAddress } from 'xlcalc/a1'
import type { CellPosition, CellRect } from 'xlcalc/contract'
import { colPx, type Geometry } from './geometry'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** The arrows' ways for Control and an arrow (`SheetSource.end`). */
const END_WAYS: Record<string, 'up' | 'down' | 'left' | 'right'> = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' }

/** Space kept from the frame's edge when scrolling to a cell: a cell glued to the edge reads as cut. */
const REVEAL_PAD = 24

export function useSelection(s: SheetState, geo: Geometry, v: Viewport) {
  /** The selection as Excel holds it: whole columns reach the last row, whole rows the last column. */
  const fullRect = computed<CellRect | null>(() => {
    const sa = s.selected.value
    if (!sa) return null
    const a = s.anchor.value ?? sa
    const b = s.extent.value ?? sa
    return { top: Math.min(a.row, b.row), left: Math.min(a.col, b.col), bottom: Math.max(a.row, b.row), right: Math.max(a.col, b.col) }
  })
  // Any other move of the active cell or the corner starts the rectangle from the active cell again.
  let inside = false
  watch([() => s.selected.value, () => s.extent.value], () => !inside && (s.anchor.value = null), { flush: 'sync' })
  /** The same cut to the grid's end, for what draws, counts and copies it. */
  const selRect = computed<CellRect | null>(() => {
    const r = fullRect.value
    if (!r) return null
    return { ...r, bottom: r.bottom >= MAX_ROWS ? Math.max(r.top, s.lastRow.value) : r.bottom, right: r.right >= MAX_COLS ? Math.max(r.left, s.colCount.value) : r.right }
  })
  /** Whole columns (`B:D`) or rows (`2:5`) the selection is, as their headers select them; null for a range of cells. */
  const wholeColumns = computed(() => (fullRect.value && fullRect.value.top === 1 && fullRect.value.bottom >= MAX_ROWS ? { first: fullRect.value.left, last: fullRect.value.right } : null))
  const wholeRows = computed(() => (fullRect.value && fullRect.value.left === 1 && fullRect.value.right >= MAX_COLS ? { first: fullRect.value.top, last: fullRect.value.bottom } : null))
  const multi = computed(() => {
    const r = selRect.value
    return !!r && (r.top !== r.bottom || r.left !== r.right)
  })
  /** `B12`, `B12:D15`, or `B:D` and `2:9` over every row or column, as Excel writes a range. */
  const selectionRef = computed(() => (fullRect.value ? rangeAddress(fullRect.value) : ''))
  /** The selection cut to the grid, as a range: what a source counts. */
  const gridRef = computed(() => (selRect.value ? rangeAddress(selRect.value) : ''))
  const isActive = (row: number, col: number) => s.selected.value?.row === row && s.selected.value?.col === col

  function pickCell(row: number, col: number, extend = false): void {
    if (extend && s.selected.value) s.extent.value = { row, col }
    else {
      s.selected.value = { row, col }
      s.extent.value = null
    }
  }
  /** Whole columns (a click on a letter, Shift from the active cell's column) or rows (on a number). */
  function pickColumn(col: number, extend = false): void {
    const from = extend && s.selected.value ? s.selected.value.col : col
    s.selected.value = { row: 1, col: from }
    s.extent.value = { row: MAX_ROWS, col }
  }
  function pickRow(row: number, extend = false): void {
    const from = extend && s.selected.value ? s.selected.value.row : row
    s.selected.value = { row: from, col: 1 }
    s.extent.value = { row, col: MAX_COLS }
  }
  function selectAll(): void {
    s.selected.value = { row: 1, col: 1 }
    s.extent.value = { row: MAX_ROWS, col: MAX_COLS }
  }
  /** Select a rectangle; one over every row or column stays whole, the grid stretches to its finite corner. */
  function selectRect(r: CellRect): void {
    const bottom = r.bottom >= MAX_ROWS ? MAX_ROWS : r.bottom
    const right = r.right >= MAX_COLS ? MAX_COLS : r.right
    s.stretch(bottom >= MAX_ROWS ? r.top : bottom, right >= MAX_COLS ? r.left : right)
    s.selected.value = { row: r.top, col: r.left }
    s.extent.value = r.top === bottom && r.left === right ? null : { row: bottom, col: right }
  }
  /** A corner of the selection cut to the grid: keys move from where the grid shows it. */
  const onGrid = (p: CellPosition): CellPosition => ({ row: p.row >= MAX_ROWS ? s.lastRow.value : p.row, col: p.col >= MAX_COLS ? Math.max(1, s.colCount.value) : p.col })

  // A mouse drags the corner; a finger scrolls from the same pointerdown, so touch picks by a tap (click) and never drags.
  let dragging = false
  let touching = false
  /** A tap's time and cell, to tell a double tap on touch screens. */
  let lastTap = { at: 0, key: '' }
  function onCellDown(e: PointerEvent, row: number, col: number): void {
    if (e.button !== 0 || (e.target as HTMLElement).closest('a,button')) return
    touching = e.pointerType === 'touch'
    if (touching) return
    pickCell(row, col, e.shiftKey)
    dragging = true
    s.scroller.value?.focus({ preventScroll: true })
  }
  /** A tap on touch: picks the cell; true when it is the second tap on the same cell. */
  function onCellTap(e: MouseEvent, row: number, col: number): boolean {
    if (!touching || (e.target as HTMLElement).closest('a,button')) return false
    pickCell(row, col, e.shiftKey)
    const key = `${row}:${col}`
    const double = lastTap.key === key && e.timeStamp - lastTap.at < 400
    lastTap = { at: double ? 0 : e.timeStamp, key }
    return double
  }
  function onGridMove(e: PointerEvent): void {
    if (!dragging || typeof document.elementFromPoint !== 'function') return
    const el = document.elementFromPoint(e.clientX, e.clientY)?.closest('[data-cell]')
    if (!el) return
    const [row, col] = (el.getAttribute('data-cell') ?? '').split(':').map(Number)
    if (row && col && (s.extent.value?.row !== row || s.extent.value?.col !== col)) s.extent.value = { row, col }
  }
  function onGridUp(): void {
    dragging = false
  }

  /** Scroll so the cell (a sheet row) is in the frame with some space; a row the filter hides drops the filter first. */
  function revealCell(row: number, col: number): void {
    const el = s.scroller.value
    if (!el) return
    if (s.indexOf(row) === undefined) s.hooks.unfilter(row)
    const d = s.indexOf(row) ?? row
    measure()
    const off = geo.rowOffsets.value
    const top = off[d - 1] ?? 0
    const bottom = off[d] ?? top + s.defaultRowPx.value
    const left = geo.offsets.value[col - 1] ?? 0
    const right = geo.offsets.value[col] ?? left + colPx(null)
    let y = el.scrollTop
    let x = el.scrollLeft
    // Frozen rows and columns are always in the frame.
    if (d > v.frozenRows.value) {
      if (top - v.frozenH.value < y) y = top - v.frozenH.value - REVEAL_PAD
      else if (bottom > y + s.viewH.value - v.headH.value) y = bottom - s.viewH.value + v.headH.value + REVEAL_PAD
    }
    if (col > v.frozenCols.value) {
      if (left - v.frozenW.value < x) x = left - v.frozenW.value - REVEAL_PAD
      else if (right > x + s.viewW.value - v.leadW.value) x = right - s.viewW.value + v.leadW.value + REVEAL_PAD
    }
    if (y !== el.scrollTop || x !== el.scrollLeft) {
      el.scrollTo({ top: Math.max(0, y), left: Math.max(0, x) })
      s.scrollTop.value = Math.max(0, y)
      s.scrollLeft.value = Math.max(0, x)
    }
  }
  /** The frame's size, read again where it matters: a scrollbar may have appeared since the last read. */
  function measure(): void {
    const el = s.scroller.value
    if (!el) return
    s.viewH.value = el.clientHeight || 600
    s.viewW.value = el.clientWidth || 800
  }

  /** Arrow keys and the like: move the active cell, or the corner with Shift; true when the key was one of them. */
  function navKey(e: KeyboardEvent): boolean {
    // Excel moves the active cell from itself, and only Shift moves the selection's far corner
    const corner = e.shiftKey ? (s.extent.value ?? s.selected.value) : s.selected.value
    if (!corner || s.rowCount.value === 0) return false
    const a = onGrid(corner)
    const way = END_WAYS[e.key]
    // Control and an arrow go where the source's Range.End says; under the display filter its rows are not the sheet's
    if (way && (e.ctrlKey || e.metaKey) && !e.altKey && s.source().end && !s.visible.value) {
      e.preventDefault()
      void endKey(a, way, e.shiftKey)
      return true
    }
    let d = s.indexOf(a.row) ?? 1
    let col = a.col
    const page = Math.max(1, Math.floor((s.viewH.value - v.headH.value - v.frozenH.value) / s.defaultRowPx.value))
    const ctrl = e.ctrlKey || e.metaKey
    // Ctrl jumps to the data's edge first, then to the grid's (an editable grid goes on past the data).
    const lastRow = (x: number) => (x < s.dataShown.value ? s.dataShown.value : s.rowCount.value)
    const lastCol = (c: number) => (c < s.dataCols.value ? s.dataCols.value : s.colCount.value)
    switch (e.key) {
      case 'ArrowDown':
        d = ctrl ? lastRow(d) : geo.nextRow(d, 1)
        break
      case 'ArrowUp':
        d = ctrl ? 1 : geo.nextRow(d, -1)
        break
      case 'ArrowRight':
        col = ctrl ? lastCol(col) : geo.nextCol(col, 1)
        break
      case 'ArrowLeft':
        col = ctrl ? 1 : geo.nextCol(col, -1)
        break
      case 'PageDown':
        d = Math.min(s.rowCount.value, d + page)
        break
      case 'PageUp':
        d = Math.max(1, d - page)
        break
      case 'Home':
        col = 1
        if (ctrl) d = 1
        break
      case 'End':
        col = ctrl ? Math.max(1, s.dataCols.value) : lastCol(col)
        if (ctrl) d = Math.max(1, s.dataShown.value)
        break
      default:
        return false
    }
    e.preventDefault()
    // Moves went along display rows; the selection keeps sheet rows.
    const row = s.rowOf(d)
    if (e.shiftKey && s.selected.value) s.extent.value = { row, col }
    else pickCell(row, col)
    if (!s.visible.value) s.stretch(row, col)
    revealCell(row, col)
    return true
  }
  let endSeq = 0
  /** Control and an arrow by the source's `end`: the active cell there, or the corner with Shift, unless the selection moved meanwhile. */
  async function endKey(from: CellPosition, way: 'up' | 'down' | 'left' | 'right', extend: boolean): Promise<void> {
    const mine = ++endSeq
    const sheet = s.sheetName.value
    const was = [s.selected.value, s.extent.value]
    let to: CellPosition | null
    try {
      to = parseCell(await s.source().end!(sheet, cellAddress(from.row, from.col), way))
    } catch {
      return
    }
    if (!to || mine !== endSeq || sheet !== s.sheetName.value || was[0] !== s.selected.value || was[1] !== s.extent.value) return
    // a read-only grid ends with its data
    const row = s.editable.value ? to.row : Math.min(to.row, Math.max(1, s.lastRow.value))
    const col = s.editable.value ? to.col : Math.min(to.col, Math.max(1, s.colCount.value))
    if (extend && s.selected.value) s.extent.value = { row, col }
    else pickCell(row, col)
    s.stretch(row, col)
    revealCell(row, col)
  }

  /** Enter and Tab: the active cell one step down or right (back with Shift); a selection of one cell collapses to it, one of several keeps it. */
  function step(dr: -1 | 0 | 1, dc: -1 | 0 | 1): void {
    const a = s.selected.value
    if (!a) return
    const r = selRect.value
    // under a display filter the selection's sheet rows are not its rows on screen: it collapses as before
    if (multi.value && r && !s.visible.value) return stepInside(a, r, dr, dc)
    const d = s.indexOf(a.row) ?? 1
    const row = s.rowOf(dr ? geo.nextRow(d, dr) : d)
    const col = dc ? geo.nextCol(a.col, dc) : a.col
    pickCell(row, col)
    if (!s.visible.value) s.stretch(row, col)
    void nextTick(() => revealCell(row, col))
  }

  /**
   * Excel's Enter and Tab in a selection of several cells: the active cell moves inside it, down its column and on to the
   * next column's top (Tab: along its row, on to the next row's start), back with Shift, round past the end, hidden lines
   * skipped; the selection stays. Without a display filter a display index is the sheet's row.
   */
  function stepInside(a: CellPosition, r: CellRect, dr: -1 | 0 | 1, dc: -1 | 0 | 1): void {
    const firstRow = (dir: 1 | -1) => (geo.heightOf(dir > 0 ? r.top : r.bottom) > 0 ? (dir > 0 ? r.top : r.bottom) : geo.nextRow(dir > 0 ? r.top : r.bottom, dir))
    const firstCol = (dir: 1 | -1) => (geo.widthOf(dir > 0 ? r.left : r.right) > 0 ? (dir > 0 ? r.left : r.right) : geo.nextCol(dir > 0 ? r.left : r.right, dir))
    const inRows = (x: number, from: number) => x !== from && x >= r.top && x <= r.bottom
    const inCols = (x: number, from: number) => x !== from && x >= r.left && x <= r.right
    let { row, col } = a
    if (dr) {
      const n = geo.nextRow(row, dr)
      if (inRows(n, row)) row = n
      else {
        const c = geo.nextCol(col, dr)
        col = inCols(c, col) ? c : firstCol(dr)
        row = firstRow(dr)
      }
    } else if (dc) {
      const c = geo.nextCol(col, dc)
      if (inCols(c, col)) col = c
      else {
        const n = geo.nextRow(row, dc)
        row = inRows(n, row) ? n : firstRow(dc)
        col = firstCol(dc)
      }
    }
    const full = fullRect.value!
    const far = s.extent.value ?? a
    inside = true
    s.anchor.value ??= { row: far.row === full.top ? full.bottom : full.top, col: far.col === full.left ? full.right : full.left }
    s.selected.value = { row, col }
    inside = false
    void nextTick(() => revealCell(row, col))
  }

  return {
    fullRect,
    selRect,
    wholeColumns,
    wholeRows,
    multi,
    selectionRef,
    gridRef,
    isActive,
    pickCell,
    pickColumn,
    pickRow,
    selectAll,
    selectRect,
    onCellDown,
    onCellTap,
    onGridMove,
    onGridUp,
    revealCell,
    measure,
    navKey,
    step,
  }
}

export type Selection = ReturnType<typeof useSelection>
