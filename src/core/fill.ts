// Excel's fill handle: the square at the selection's corner. Dragged down, up, right or left, its outline grows cell by
// cell along the axis in which the pointer moved farthest, and the release fills the outlined range from the selection
// through the source's `structure` (`fillRange`, Excel's AutoFill: series continue, other values are copied; Ctrl makes a
// single number continue as a series); a double click fills down as far as the data of the adjacent column extends.
// Excel's Fill Down (Ctrl+D) and Fill Right (Ctrl+R) copy the selection's first row or column over the rest of it (`copyRange`).
import { computed, getCurrentScope, onScopeDispose, shallowRef, type Ref } from 'vue'
import { cellAddress, MAX_COLS, MAX_ROWS, rangeAddress } from 'xlcalc/a1'
import type { CellRect, FillType } from 'xlcalc/contract'
import type { EditOutcome } from '../contract/source'
import { unionRect } from '../rect'
import { isBlank } from '../stats'
import { DEFAULT_MAX_CELLS } from '../window-size'
import type { Commands } from './commands'
import type { Editing, EditorState } from './editing'
import type { Geometry } from './geometry'
import { isShortcut, onMac, shortcutLetter } from './keys'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import type { Viewport } from './viewport'

/** Within this many pixels of the viewport's edge a drag scrolls the grid, by at most `EDGE_STEP_PX` per move. */
const EDGE_PX = 24
const EDGE_STEP_PX = 40

export interface FillDeps {
  notify: (text: string) => void
  editor: Ref<EditorState | null>
}

/** A drag of the handle: the selection it fills from, the cells added so far, the pointer. */
export interface FillDrag {
  from: CellRect
  to: CellRect | null
  pointer: number
  touch: boolean
}

export function useFill(s: SheetState, geo: Geometry, v: Viewport, sel: Selection, ed: Editing, cmd: Commands, deps: FillDeps) {
  /** The range the handle fills from: a finite selection on an editable sheet with `structure`, not under the display filter or while an editor is open. */
  const source = computed<CellRect | null>(() => {
    const r = sel.fullRect.value
    if (!r || !s.sheet.value || !cmd.canRanges.value || deps.editor.value) return null
    return r.bottom >= MAX_ROWS || r.right >= MAX_COLS ? null : r
  })
  const drag = shallowRef<FillDrag | null>(null)
  /** The handle that captured the pointer at the start of the drag, to release it on Esc. */
  let held: Element | null = null
  /** The handle's center: the selection's bottom-right corner in the scroller's content; it stays put while a drag draws its outline, since it holds the pointer. */
  const handle = computed(() => {
    const r = source.value
    if (!r) return null
    const b = v.boxOf(r.top, r.left, r.bottom, r.right)
    return { x: b.left + b.width, y: b.top + b.height, frozen: b.frozen, dragging: !!drag.value?.to }
  })
  /** A drag's outline: the selection and the cells it adds. */
  const frame = computed(() => {
    const d = drag.value
    if (!d?.to) return null
    const u = unionRect(d.from, d.to)
    return v.boxOf(u.top, u.left, u.bottom, u.right)
  })

  /** The cells a drag adds with the pointer at a point of the page (`box` is the viewport's): along the axis in which it moved farthest from the selection, up to the pointer's cell. */
  function targetAt(from: CellRect, box: DOMRect, clientX: number, clientY: number): CellRect | null {
    const p = v.cellAt(clientX - box.left, clientY - box.top)
    const rows = geo.rowOffsets.value
    const cols = geo.offsets.value
    const top = rows[from.top - 1] ?? 0
    const bottom = rows[from.bottom] ?? top
    const left = cols[from.left - 1] ?? 0
    const right = cols[from.right] ?? left
    const dy = p.y >= bottom ? p.y - bottom : p.y < top ? top - p.y : 0
    const dx = p.x >= right ? p.x - right : p.x < left ? left - p.x : 0
    if (dy === 0 && dx === 0) return null
    if (dy >= dx) {
      if (p.y >= bottom) {
        const end = Math.min(p.d, MAX_ROWS)
        return end > from.bottom ? { ...from, top: from.bottom + 1, bottom: end } : null
      }
      const start = Math.max(p.d, 1)
      return start < from.top ? { ...from, top: start, bottom: from.top - 1 } : null
    }
    if (p.x >= right) {
      const end = Math.min(p.col, MAX_COLS)
      return end > from.right ? { ...from, left: from.right + 1, right: end } : null
    }
    const start = Math.max(p.col, 1)
    return start < from.left ? { ...from, left: start, right: from.left - 1 } : null
  }
  /** Scrolls the grid while the pointer is at or beyond the viewport's edge (`box`), so a drag reaches cells off screen. */
  function edgeScroll(el: HTMLElement, box: DOMRect, clientX: number, clientY: number): void {
    if (!box.height) return
    const step = (past: number) => Math.min(EDGE_STEP_PX, Math.max(0, past))
    const dy = step(clientY - (box.bottom - EDGE_PX)) - step(box.top + v.headH.value + EDGE_PX - clientY)
    const dx = step(clientX - (box.right - EDGE_PX)) - step(box.left + v.leadW.value + EDGE_PX - clientX)
    if (dy) el.scrollTop += dy
    if (dx) el.scrollLeft += dx
    s.scrollTop.value = el.scrollTop
    s.scrollLeft.value = el.scrollLeft
  }

  function onDown(e: PointerEvent): void {
    const from = source.value
    if (!from || e.button !== 0) return
    e.preventDefault()
    e.stopPropagation()
    held = e.currentTarget as Element | null
    held?.setPointerCapture?.(e.pointerId)
    drag.value = { from, to: null, pointer: e.pointerId, touch: e.pointerType === 'touch' }
    document.addEventListener('keydown', onKey, true)
    s.focusGrid()
  }
  function onMove(e: PointerEvent): void {
    const d = drag.value
    const el = s.scroller.value
    if (!d || e.pointerId !== d.pointer || !el) return
    // One measurement per move: the viewport does not move as it scrolls, and a read after the scroll's write would trigger another layout.
    const box = el.getBoundingClientRect()
    edgeScroll(el, box, e.clientX, e.clientY)
    const to = targetAt(d.from, box, e.clientX, e.clientY)
    if (to?.top !== d.to?.top || to?.left !== d.to?.left || to?.bottom !== d.to?.bottom || to?.right !== d.to?.right) drag.value = { ...d, to }
  }
  function onUp(e: PointerEvent): void {
    const d = drag.value
    if (!d || e.pointerId !== d.pointer) return
    end()
    // Excel for Windows checks Ctrl at release, Excel for Mac Option (Control there is the right click).
    if (d.to) return void fill(d.from, d.to, e.ctrlKey || (e.altKey && onMac()) ? 'series' : undefined)
    // A tap on the handle's wide touch target meant the cell under it.
    if (d.touch) {
      const box = s.scroller.value?.getBoundingClientRect()
      if (!box) return
      const p = v.cellAt(e.clientX - box.left, e.clientY - box.top)
      sel.pickCell(s.rowOf(p.d), p.col)
    }
  }
  const onCancel = end
  /** Esc during a drag, before the grid's keys: Excel cancels the drag, so the release fills nothing; the pointer is released. */
  function onKey(e: KeyboardEvent): void {
    const d = drag.value
    if (e.key !== 'Escape' || !d) return
    e.preventDefault()
    e.stopPropagation()
    const el = held
    end()
    el?.releasePointerCapture?.(d.pointer)
  }
  /** The drag is over: no outline, no key listener, no element holding the pointer. */
  function end(): void {
    drag.value = null
    held = null
    document.removeEventListener('keydown', onKey, true)
  }
  if (getCurrentScope()) onScopeDispose(end)
  const read = useSheetReads(s)

  /** Fills the outline of `from` plus the added cells `to` from `from` as a single undo step; the selection becomes the outlined range, as Excel leaves it. */
  async function fill(from: CellRect, to: CellRect, type?: FillType): Promise<void> {
    const lock = ed.lockIn(to)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const sheet = s.sheetName.value
    const frame = unionRect(from, to)
    const before = { selected: s.selected.value, extent: s.extent.value }
    const out = await cmd.structure({ op: 'fillRange', sheet, range: rangeAddress(from), to: rangeAddress(frame), ...(type ? { type } : {}) })
    if (out?.ok && sheet === s.sheetName.value && s.selected.value === before.selected && s.extent.value === before.extent) sel.selectRect(frame)
  }

  /**
   * Excel's Fill Down and Fill Right: the selection's top row (left column) copied over its other rows (columns) through
   * the source's `structure` (`copyRange`: relative references shift as Excel's fill shifts them, and on a sheet filtered
   * in the workbook only the visible rows are filled, as Excel's Fill Down fills them); a selection one row tall (one
   * column wide) takes the row above it (the column to its left). One undo step; the selection stays. Null when nothing
   * was sent: nothing above row 1 or left of column A, no `structure`, a display filter, an open editor, a locked cell.
   */
  async function fillFrom(direction: 'down' | 'right'): Promise<EditOutcome | null> {
    const r = sel.fullRect.value
    if (!r || !s.sheet.value || !cmd.canRanges.value || deps.editor.value) return null
    const down = direction === 'down'
    const first = down ? r.top : r.left
    const single = first === (down ? r.bottom : r.right)
    if (single && first === 1) return null
    const line = single ? first - 1 : first
    const from = down ? { ...r, top: line, bottom: line } : { ...r, left: line, right: line }
    const to = single ? r : down ? { ...r, top: r.top + 1 } : { ...r, left: r.left + 1 }
    const lock = ed.lockIn(to)
    if (lock !== null) {
      deps.notify(s.messages.value.locked(lock))
      return null
    }
    const sheet = s.sheetName.value
    return cmd.structure({ op: 'copyRange', from: { sheet, range: rangeAddress(from) }, to: { sheet, range: rangeAddress(to) } })
  }
  /** Ctrl+D and Ctrl+R on an editable sheet (Cmd or Control on a Mac, as Excel for Mac takes either), with any layout; true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (!isShortcut(e) || e.shiftKey || e.altKey || !s.editable.value) return false
    const letter = shortcutLetter(e)
    if (letter !== 'd' && letter !== 'r') return false
    // the browser's bookmark and reload keys stay with the sheet even where it cannot fill, as Excel keeps them
    e.preventDefault()
    void fillFrom(letter === 'd' ? 'down' : 'right')
    return true
  }

  /** Excel's double click on the handle: down to the last data row of the column to the left of the selection, otherwise to its right; displaying another sheet meanwhile stops it. */
  async function autofill(): Promise<void> {
    const from = source.value
    if (!from) return
    const signal = read()
    for (const col of [from.left - 1, from.right + 1]) {
      if (col < 1 || col > MAX_COLS) continue
      const end = await dataEnd(col, from.bottom + 1, signal)
      if (signal.aborted) return
      if (end > from.bottom) return fill(from, { ...from, top: from.bottom + 1, bottom: end })
    }
  }
  /** The last row of the run of non-empty cells in a column from `row` down, as the source has it, read in windows of the cell limit. */
  async function dataEnd(col: number, row: number, signal: AbortSignal): Promise<number> {
    const sheet = s.sheetName.value
    const last = s.dataRows.value
    const step = s.source().maxCells ?? DEFAULT_MAX_CELLS
    let r = row
    while (r <= last) {
      let got
      try {
        got = await s.source().window({ sheet, row: r, rows: Math.min(step, last - r + 1), col, cols: 1 }, signal)
      } catch {
        return row - 1
      }
      if ('pending' in got) return row - 1
      if (!got.text.length) return r - 1
      for (let i = 0; i < got.text.length; i++) {
        if (isBlank(got.text[i]![0], got.values[i]?.[0]) && got.formulas[cellAddress(r + i, col)] === undefined) return r + i - 1
      }
      r += got.text.length
    }
    return last
  }

  return { source, drag, handle, frame, onDown, onMove, onUp, onCancel, autofill, fillFrom, key }
}

export type Fill = ReturnType<typeof useFill>
