// Excel's fill handle: the square at the selection's corner. Dragged down, up, right or left, its frame grows along the
// axis the pointer went furthest, and the release copies the selection over the cells added through the source's
// `structure` (`copyRange`, which tiles); a double click fills down as far as the data of the column beside goes.
// TODO(xlcalc): Excel's fill handle extends series (1, 2 → 3, 4; dates by day; Item1 → Item2) and tiles a part of the
// source at the end; `copyRange` copies whole tiles alone, so the frame grows by the source's size and a series is copied.
import { computed, getCurrentScope, onScopeDispose, shallowRef, type Ref } from 'vue'
import { cellAddress, MAX_COLS, MAX_ROWS, rangeAddress } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import type { Commands } from './commands'
import type { Editing, EditorState } from './editing'
import type { Geometry } from './geometry'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** Within this many pixels of the frame's edge a drag scrolls the grid, by at most `EDGE_STEP_PX` a move. */
const EDGE_PX = 24
const EDGE_STEP_PX = 40

export interface FillDeps {
  notify: (text: string) => void
  editor: Ref<EditorState | null>
}

/** A drag of the handle: the selection it copies, the cells added so far, the pointer. */
export interface FillDrag {
  from: CellRect
  to: CellRect | null
  pointer: number
  touch: boolean
}

const union = (a: CellRect, b: CellRect): CellRect => ({ top: Math.min(a.top, b.top), left: Math.min(a.left, b.left), bottom: Math.max(a.bottom, b.bottom), right: Math.max(a.right, b.right) })
/** Cells `n` rounded up to whole tiles of `size`, as `copyRange` tiles them. */
const tiles = (n: number, size: number) => Math.ceil(n / size) * size

export function useFill(s: SheetState, geo: Geometry, v: Viewport, sel: Selection, ed: Editing, cmd: Commands, deps: FillDeps) {
  /** The range the handle copies: a finite selection of an editable sheet with `structure`, not under the display filter or while a field is open. */
  const source = computed<CellRect | null>(() => {
    const r = sel.fullRect.value
    if (!r || !s.sheet.value || !s.editable.value || !cmd.can('structure') || s.visible.value || deps.editor.value) return null
    return r.bottom >= MAX_ROWS || r.right >= MAX_COLS ? null : r
  })
  const drag = shallowRef<FillDrag | null>(null)
  /** The handle that took the pointer at the drag's start, to let it go on Esc. */
  let held: Element | null = null
  /** The handle's centre: the selection's bottom-right corner in the scroller's content; it stays while a drag draws its frame, as it holds the pointer. */
  const handle = computed(() => {
    const r = source.value
    if (!r) return null
    const b = v.boxOf(r.top, r.left, r.bottom, r.right)
    return { x: b.left + b.width, y: b.top + b.height, frozen: b.frozen, dragging: !!drag.value?.to }
  })
  /** The frame of a drag: the selection and the cells it adds. */
  const frame = computed(() => {
    const d = drag.value
    if (!d?.to) return null
    const u = union(d.from, d.to)
    return v.boxOf(u.top, u.left, u.bottom, u.right)
  })

  /** The cells a drag adds with the pointer at a point of the page: along the axis it went furthest from the selection, in whole tiles. */
  function targetAt(from: CellRect, clientX: number, clientY: number): CellRect | null {
    const el = s.scroller.value
    if (!el) return null
    const box = el.getBoundingClientRect()
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
    const h = from.bottom - from.top + 1
    const w = from.right - from.left + 1
    if (dy >= dx) {
      if (p.y >= bottom) {
        const n = Math.min(tiles(p.d - from.bottom, h), Math.floor((MAX_ROWS - from.bottom) / h) * h)
        return n > 0 ? { ...from, top: from.bottom + 1, bottom: from.bottom + n } : null
      }
      const n = Math.min(tiles(from.top - p.d, h), Math.floor((from.top - 1) / h) * h)
      return n > 0 ? { ...from, top: from.top - n, bottom: from.top - 1 } : null
    }
    if (p.x >= right) {
      const n = Math.min(tiles(p.col - from.right, w), Math.floor((MAX_COLS - from.right) / w) * w)
      return n > 0 ? { ...from, left: from.right + 1, right: from.right + n } : null
    }
    const n = Math.min(tiles(from.left - p.col, w), Math.floor((from.left - 1) / w) * w)
    return n > 0 ? { ...from, left: from.left - n, right: from.left - 1 } : null
  }
  /** Scroll the grid while the pointer is at or past the frame's edge, so a drag reaches cells off screen. */
  function edgeScroll(clientX: number, clientY: number): void {
    const el = s.scroller.value
    if (!el) return
    const box = el.getBoundingClientRect()
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
    s.scroller.value?.focus({ preventScroll: true })
  }
  function onMove(e: PointerEvent): void {
    const d = drag.value
    if (!d || e.pointerId !== d.pointer) return
    edgeScroll(e.clientX, e.clientY)
    const to = targetAt(d.from, e.clientX, e.clientY)
    if (to?.top !== d.to?.top || to?.left !== d.to?.left || to?.bottom !== d.to?.bottom || to?.right !== d.to?.right) drag.value = { ...d, to }
  }
  function onUp(e: PointerEvent): void {
    const d = drag.value
    if (!d || e.pointerId !== d.pointer) return
    end()
    if (d.to) return void fill(d.from, d.to)
    // A finger's tap on the handle's wide target meant the cell under it.
    if (d.touch) {
      const box = s.scroller.value?.getBoundingClientRect()
      if (!box) return
      const p = v.cellAt(e.clientX - box.left, e.clientY - box.top)
      sel.pickCell(s.rowOf(p.d), p.col)
    }
  }
  const onCancel = end
  /** Esc while dragged, before the grid's keys: Excel drops the drag, so the release copies nothing; the pointer is let go. */
  function onKey(e: KeyboardEvent): void {
    const d = drag.value
    if (e.key !== 'Escape' || !d) return
    e.preventDefault()
    e.stopPropagation()
    const el = held
    end()
    el?.releasePointerCapture?.(d.pointer)
  }
  /** The drag is over: no frame, no key listener, no element holding the pointer. */
  function end(): void {
    drag.value = null
    held = null
    document.removeEventListener('keydown', onKey, true)
  }
  if (getCurrentScope()) onScopeDispose(end)

  /** Copy `from` over `to` in one step of undo; the selection becomes both, as Excel leaves it. */
  async function fill(from: CellRect, to: CellRect): Promise<void> {
    const lock = ed.lockIn(to)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const sheet = s.sheetName.value
    const before = { selected: s.selected.value, extent: s.extent.value }
    const out = await cmd.structure({ op: 'copyRange', from: { sheet, range: rangeAddress(from) }, to: { sheet, range: rangeAddress(to) } })
    if (out?.ok && sheet === s.sheetName.value && s.selected.value === before.selected && s.extent.value === before.extent) sel.selectRect(union(from, to))
  }

  /** Excel's double click on the handle: down to the last row of the data in the column left of the selection, else right of it, in whole tiles. */
  async function autofill(): Promise<void> {
    const from = source.value
    if (!from) return
    const h = from.bottom - from.top + 1
    for (const col of [from.left - 1, from.right + 1]) {
      if (col < 1 || col > MAX_COLS) continue
      const end = await dataEnd(col, from.bottom + 1)
      const n = Math.floor((end - from.bottom) / h) * h
      if (end > from.bottom) return n > 0 ? fill(from, { ...from, top: from.bottom + 1, bottom: from.bottom + n }) : undefined
    }
  }
  /** The last row of the run of filled cells in a column from `row` down, as the source has it, read in windows of the cell limit. */
  async function dataEnd(col: number, row: number): Promise<number> {
    const sheet = s.sheetName.value
    const last = s.dataRows.value
    const step = s.source().maxCells ?? 20_000
    let r = row
    while (r <= last) {
      let got
      try {
        got = await s.source().window({ sheet, row: r, rows: Math.min(step, last - r + 1), col, cols: 1 })
      } catch {
        return row - 1
      }
      if ('pending' in got) return row - 1
      if (!got.text.length) return r - 1
      for (let i = 0; i < got.text.length; i++) {
        const value = got.values[i]?.[0]
        const empty = (got.text[i]![0] ?? '') === '' && (value === null || value === undefined || value === '') && got.formulas[cellAddress(r + i, col)] === undefined
        if (empty) return r + i - 1
      }
      r += got.text.length
    }
    return last
  }

  return { source, drag, handle, frame, onDown, onMove, onUp, onCancel, autofill }
}

export type Fill = ReturnType<typeof useFill>
