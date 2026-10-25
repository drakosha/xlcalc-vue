// Selection as in Excel: a range from the active cell to the other corner, by cells, rows and columns, with Shift and drag.
// It is stored in sheet rows; keys move along display rows, so under a filter the next row is the next visible one.
// Entire columns and rows extend to Excel's last row and column, as Excel selects them; what is drawn stops at the grid's end.
import { computed, nextTick, watch } from 'vue'
import { cellAddress, MAX_COLS, MAX_ROWS, parseCell, rangeAddress } from 'xlcalc/a1'
import type { CellPosition, CellRect } from 'xlcalc/contract'
import { rectOf } from '../rect'
import { colPx, type Geometry } from './geometry'
import { isMod } from './keys'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** The arrow directions for Ctrl+arrow (`SheetSource.end`). */
const END_WAYS: Record<string, 'up' | 'down' | 'left' | 'right'> = { ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right' }

/** A drawn cell's position from its `data-cell` (`row:col`); null for anything else. */
export function cellOfData(data: string | null | undefined): CellPosition | null {
  const [row, col] = (data ?? '').split(':').map(Number)
  return row && col ? { row, col } : null
}

/** The margin kept from the viewport's edge when scrolling to a cell: a cell flush with the edge looks clipped. */
const REVEAL_PAD = 24

export function useSelection(s: SheetState, geo: Geometry, v: Viewport) {
  /** The selection as Excel holds it: entire columns extend to the last row, entire rows to the last column. */
  const fullRect = computed<CellRect | null>(() => {
    const sa = s.selected.value
    if (!sa) return null
    const a = s.anchor.value ?? sa
    const b = s.extent.value ?? sa
    return rectOf(a, b)
  })
  // Any other move of the active cell or the corner starts the range from the active cell again.
  let inside = false
  watch([() => s.selected.value, () => s.extent.value], () => !inside && (s.anchor.value = null), { flush: 'sync' })
  /** The same, clipped to the grid's end, for what draws, counts and copies it. */
  const selRect = computed<CellRect | null>(() => {
    const r = fullRect.value
    if (!r) return null
    return { ...r, bottom: r.bottom >= MAX_ROWS ? Math.max(r.top, s.lastRow.value) : r.bottom, right: r.right >= MAX_COLS ? Math.max(r.left, s.colCount.value) : r.right }
  })
  /** The entire columns (`B:D`) or rows (`2:5`) the selection consists of, as their headers select them; null for a cell range. */
  const wholeColumns = computed(() => (fullRect.value && fullRect.value.top === 1 && fullRect.value.bottom >= MAX_ROWS ? { first: fullRect.value.left, last: fullRect.value.right } : null))
  const wholeRows = computed(() => (fullRect.value && fullRect.value.left === 1 && fullRect.value.right >= MAX_COLS ? { first: fullRect.value.top, last: fullRect.value.bottom } : null))
  /** The selected entire rows or columns (`wholeRows`, `wholeColumns`) when `line` is one of them, otherwise that row or column alone: what a row number's or column letter's command acts on. */
  const linesOf = (whole: { first: number; last: number } | null, line: number): { first: number; last: number } =>
    whole && line >= whole.first && line <= whole.last ? whole : { first: line, last: line }
  const multi = computed(() => {
    const r = selRect.value
    return !!r && (r.top !== r.bottom || r.left !== r.right)
  })
  /** `B12`, `B12:D15`, or `B:D` and `2:9` over every row or column, as Excel writes a range. */
  const selectionRef = computed(() => (fullRect.value ? rangeAddress(fullRect.value) : ''))
  /** The selection clipped to the grid, as a range: what a source counts. */
  const gridRef = computed(() => (selRect.value ? rangeAddress(selRect.value) : ''))
  const isActive = (row: number, col: number) => s.selected.value?.row === row && s.selected.value?.col === col

  function pickCell(row: number, col: number, extend = false): void {
    if (extend && s.selected.value) s.extent.value = { row, col }
    else {
      s.selected.value = { row, col }
      s.extent.value = null
    }
  }
  /** Entire columns (a click on a letter, Shift from the active cell's column) or rows (on a number). */
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
  /** An editable grid extends to show a range's finite corner: for one over every row or column, to its first row or column there. */
  function stretchTo(r: CellRect): void {
    s.stretch(r.bottom >= MAX_ROWS ? r.top : r.bottom, r.right >= MAX_COLS ? r.left : r.right)
  }
  /** Selects a range; one over every row or column stays entire, and the grid extends to its finite corner. */
  function selectRect(r: CellRect): void {
    const bottom = r.bottom >= MAX_ROWS ? MAX_ROWS : r.bottom
    const right = r.right >= MAX_COLS ? MAX_COLS : r.right
    stretchTo(r)
    s.selected.value = { row: r.top, col: r.left }
    s.extent.value = r.top === bottom && r.left === right ? null : { row: bottom, col: right }
  }
  /** A selection corner clipped to the grid: keys move from where the grid displays it. */
  const onGrid = (p: CellPosition): CellPosition => ({ row: p.row >= MAX_ROWS ? s.lastRow.value : p.row, col: p.col >= MAX_COLS ? Math.max(1, s.colCount.value) : p.col })

  // A mouse drags the corner; a finger scrolls from the same pointerdown, so touch selects by a tap (click) and never drags.
  let dragging = false
  let touching = false
  /** A tap's time and cell, to detect a double tap on touch screens. */
  let lastTap = { at: 0, key: '' }
  function onCellDown(e: PointerEvent, row: number, col: number): void {
    if (e.button !== 0 || (e.target as HTMLElement).closest('a,button')) return
    touching = e.pointerType === 'touch'
    if (touching) return
    pickCell(row, col, e.shiftKey)
    dragging = true
    s.focusGrid()
  }
  /** A tap on a touch screen: selects the cell; true when it is the second tap on the same cell. */
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
    const at = cellOfData(document.elementFromPoint(e.clientX, e.clientY)?.closest('[data-cell]')?.getAttribute('data-cell'))
    if (at && (s.extent.value?.row !== at.row || s.extent.value?.col !== at.col)) s.extent.value = at
  }
  function onGridUp(): void {
    dragging = false
  }

  /** Scrolls so the cell (by sheet row) is in the viewport with a margin; a row hidden by the filter clears the filter first. */
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
    const z = s.zoom.value
    let y = el.scrollTop / z
    let x = el.scrollLeft / z
    // Frozen rows and columns are always in the viewport.
    if (d > v.frozenRows.value) {
      if (top - v.frozenH.value < y) y = top - v.frozenH.value - REVEAL_PAD
      else if (bottom > y + s.viewH.value - v.headH.value - v.footH.value) y = bottom - s.viewH.value + v.headH.value + v.footH.value + REVEAL_PAD
    }
    if (col > v.frozenCols.value) {
      if (left - v.frozenW.value < x) x = left - v.frozenW.value - REVEAL_PAD
      else if (right > x + s.viewW.value - v.leadW.value) x = right - s.viewW.value + v.leadW.value + REVEAL_PAD
    }
    if (y !== el.scrollTop / z || x !== el.scrollLeft / z) {
      s.scrollGridTo(Math.max(0, y), Math.max(0, x))
    }
  }
  /** The viewport's size in the grid's unzoomed pixels, measured again where it matters: a scrollbar may have appeared since the last measurement. */
  function measure(): void {
    const el = s.scroller.value
    if (!el) return
    s.viewH.value = (el.clientHeight || 600) / s.zoom.value
    s.viewW.value = (el.clientWidth || 800) / s.zoom.value
  }

  /**
   * Where a key moves from `a` (a cell on the grid), as the sheet navigates: the arrows, Ctrl+arrow (by the source's `end`
   * when it has one: a promise of the cell, null when it fails), Page Down and Page Up, Home, End; null for another key.
   */
  function keyTarget(a: CellPosition, e: KeyboardEvent): CellPosition | Promise<CellPosition | null> | null {
    if (s.rowCount.value === 0) return null
    const way = END_WAYS[e.key]
    // Ctrl+arrow goes where the source's Range.End points; under the display filter its rows are not the sheet's
    const ctrl = isMod(e)
    if (way && ctrl && !e.altKey && s.source().end && !s.visible.value) return endTarget(a, way)
    let d = s.indexOf(a.row) ?? 1
    let col = a.col
    const page = Math.max(1, Math.floor((s.viewH.value - v.headH.value - v.frozenH.value) / s.defaultRowPx.value))
    // Ctrl jumps to the data's edge first, then to the grid's (an editable grid continues beyond the data).
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
        return null
    }
    // Moves went along display rows; the selection stores sheet rows.
    return { row: s.rowOf(d), col }
  }
  /** The cell the source's `end` returns from `from`; a read-only grid ends at its data. */
  async function endTarget(from: CellPosition, way: 'up' | 'down' | 'left' | 'right'): Promise<CellPosition | null> {
    let to: CellPosition | null
    try {
      to = parseCell(await s.source().end!(s.sheetName.value, cellAddress(from.row, from.col), way))
    } catch {
      return null
    }
    if (!to) return null
    return s.editable.value ? to : { row: Math.min(to.row, Math.max(1, s.lastRow.value)), col: Math.min(to.col, Math.max(1, s.colCount.value)) }
  }

  /** Arrow keys and the like: move the active cell, or the corner with Shift; true when the key was one of them. */
  function navKey(e: KeyboardEvent): boolean {
    // Excel moves the active cell from itself, and only Shift moves the selection's far corner
    const corner = e.shiftKey ? (s.extent.value ?? s.selected.value) : s.selected.value
    const to = corner && keyTarget(onGrid(corner), e)
    if (!to) return false
    e.preventDefault()
    if (to instanceof Promise) void endKey(to, e.shiftKey)
    else moveTo(to, e.shiftKey)
    return true
  }
  /** Moves the active cell to `p`, or the corner with Shift; the grid extends to it, and the viewport scrolls to it. */
  function moveTo({ row, col }: CellPosition, extend: boolean): void {
    if (extend && s.selected.value) s.extent.value = { row, col }
    else pickCell(row, col)
    if (!s.visible.value) s.stretch(row, col)
    revealCell(row, col)
  }
  let endSeq = 0
  /** Ctrl+arrow by the source's `end`: the active cell moves there, or the corner with Shift, unless the selection moved meanwhile. */
  async function endKey(target: Promise<CellPosition | null>, extend: boolean): Promise<void> {
    const mine = ++endSeq
    const sheet = s.sheetName.value
    const was = [s.selected.value, s.extent.value]
    const to = await target
    if (!to || mine !== endSeq || sheet !== s.sheetName.value || was[0] !== s.selected.value || was[1] !== s.extent.value) return
    moveTo(to, extend)
  }

  /** Enter and Tab: the active cell moves one step down or right (back with Shift); a single-cell selection collapses to it, a multi-cell one keeps it. */
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
   * top of the next column (Tab: along its row and on to the start of the next row), back with Shift, wrapping around at the
   * end, skipping hidden rows and columns; the selection stays. Without a display filter a display index is the sheet row.
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
    linesOf,
    multi,
    selectionRef,
    gridRef,
    isActive,
    pickCell,
    pickColumn,
    pickRow,
    selectAll,
    stretchTo,
    selectRect,
    onCellDown,
    onCellTap,
    onGridMove,
    onGridUp,
    revealCell,
    measure,
    navKey,
    keyTarget,
    step,
  }
}

export type Selection = ReturnType<typeof useSelection>
