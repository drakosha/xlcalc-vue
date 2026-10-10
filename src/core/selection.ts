// Selection as in Excel: a rectangle from the active cell to the other corner, by cells, rows and columns, with Shift and drag.
import { computed, nextTick } from 'vue'
import { rectAddress, type Rect } from './a1'
import { colPx, type Geometry } from './geometry'
import type { SheetState } from './state'
import { HEAD_H, type Viewport } from './viewport'

/** Space kept from the frame's edge when scrolling to a cell: a cell glued to the edge reads as cut. */
const REVEAL_PAD = 24

export function useSelection(s: SheetState, geo: Geometry, v: Viewport) {
  const selRect = computed<Rect | null>(() => {
    const a = s.selected.value
    if (!a) return null
    const b = s.extent.value ?? a
    return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) }
  })
  const multi = computed(() => {
    const r = selRect.value
    return !!r && (r.r1 !== r.r2 || r.c1 !== r.c2)
  })
  /** `B12` or `B12:D15`. */
  const selectionRef = computed(() => (selRect.value ? rectAddress(selRect.value) : ''))
  const isActive = (row: number, col: number) => s.selected.value?.row === row && s.selected.value?.col === col

  function pickCell(row: number, col: number, extend = false): void {
    if (extend && s.selected.value) s.extent.value = { row, col }
    else {
      s.selected.value = { row, col }
      s.extent.value = null
    }
  }
  /** A whole column (a click on its letter) or row (on its number). */
  function pickColumn(col: number, extend = false): void {
    const last = Math.max(1, s.rowCount.value)
    if (extend && s.selected.value) s.extent.value = { row: last, col }
    else {
      s.selected.value = { row: 1, col }
      s.extent.value = { row: last, col }
    }
  }
  function pickRow(row: number, extend = false): void {
    const last = Math.max(1, s.colCount.value)
    if (extend && s.selected.value) s.extent.value = { row, col: last }
    else {
      s.selected.value = { row, col: 1 }
      s.extent.value = { row, col: last }
    }
  }
  function selectAll(): void {
    s.selected.value = { row: 1, col: 1 }
    s.extent.value = { row: Math.max(1, s.rowCount.value), col: Math.max(1, s.colCount.value) }
  }

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

  /** Scroll so the cell is in the frame with some space (nothing moves when it already is). */
  function revealCell(row: number, col: number): void {
    const el = s.scroller.value
    if (!el) return
    measure()
    const off = geo.rowOffsets.value
    const top = off[row - 1] ?? 0
    const bottom = off[row] ?? top + s.defaultRowPx.value
    const left = geo.offsets.value[col - 1] ?? 0
    const right = geo.offsets.value[col] ?? left + colPx(null)
    let y = el.scrollTop
    let x = el.scrollLeft
    // Frozen rows and columns are always in the frame.
    if (row > v.frozenRows.value) {
      if (top - v.frozenH.value < y) y = top - v.frozenH.value - REVEAL_PAD
      else if (bottom > y + s.viewH.value - HEAD_H) y = bottom - s.viewH.value + HEAD_H + REVEAL_PAD
    }
    if (col > v.frozenCols.value) {
      if (left - v.frozenW.value < x) x = left - v.frozenW.value - REVEAL_PAD
      else if (right > x + s.viewW.value - v.numW.value) x = right - s.viewW.value + v.numW.value + REVEAL_PAD
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
    const a = s.extent.value ?? s.selected.value
    if (!a || s.rowCount.value === 0) return false
    let row = a.row
    let col = a.col
    const page = Math.max(1, Math.floor((s.viewH.value - HEAD_H - v.frozenH.value) / s.defaultRowPx.value))
    const ctrl = e.ctrlKey || e.metaKey
    // Ctrl jumps to the data's edge first, then to the grid's (an editable grid goes on past the data).
    const lastRow = (r: number) => (r < s.dataRows.value ? s.dataRows.value : s.rowCount.value)
    const lastCol = (c: number) => (c < s.dataCols.value ? s.dataCols.value : s.colCount.value)
    switch (e.key) {
      case 'ArrowDown':
        row = ctrl ? lastRow(row) : geo.nextRow(row, 1)
        break
      case 'ArrowUp':
        row = ctrl ? 1 : geo.nextRow(row, -1)
        break
      case 'ArrowRight':
        col = ctrl ? lastCol(col) : geo.nextCol(col, 1)
        break
      case 'ArrowLeft':
        col = ctrl ? 1 : geo.nextCol(col, -1)
        break
      case 'PageDown':
        row = Math.min(s.rowCount.value, row + page)
        break
      case 'PageUp':
        row = Math.max(1, row - page)
        break
      case 'Home':
        col = 1
        if (ctrl) row = 1
        break
      case 'End':
        col = ctrl ? Math.max(1, s.dataCols.value) : lastCol(col)
        if (ctrl) row = Math.max(1, s.dataRows.value)
        break
      default:
        return false
    }
    e.preventDefault()
    if (e.shiftKey && s.selected.value) s.extent.value = { row, col }
    else pickCell(row, col)
    s.stretch(row, col)
    revealCell(row, col)
    return true
  }
  /** Enter and Tab: the active cell one step down or right (back with Shift), the selection collapses to it. */
  function step(dr: -1 | 0 | 1, dc: -1 | 0 | 1): void {
    const a = s.selected.value
    if (!a) return
    const row = dr ? geo.nextRow(a.row, dr) : a.row
    const col = dc ? geo.nextCol(a.col, dc) : a.col
    pickCell(row, col)
    s.stretch(row, col)
    void nextTick(() => revealCell(row, col))
  }

  return { selRect, multi, selectionRef, isActive, pickCell, pickColumn, pickRow, selectAll, onCellDown, onCellTap, onGridMove, onGridUp, revealCell, measure, navKey, step }
}

export type Selection = ReturnType<typeof useSelection>
