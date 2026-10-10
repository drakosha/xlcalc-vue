// What the grid draws now: frozen panes with the lock, the body's rows and columns under the scroll, and spacers for the rest.
import { computed, onScopeDispose, ref, type ComputedRef, type Ref } from 'vue'
import type { SheetInfo } from '../source'
import type { Geometry } from './geometry'
import type { Memory } from './storage'

/** Height of the column letters row (fixed in CSS); frozen rows stick below it. */
export const HEAD_H = 24
/** Rows and columns drawn beyond the screen, so a short scroll needs no re-render. */
const RENDER_ROWS = 8
const RENDER_COLS = 2

export interface ViewportInput {
  sheet: ComputedRef<SheetInfo | null>
  rowCount: ComputedRef<number>
  colCount: ComputedRef<number>
  scrollTop: Ref<number>
  scrollLeft: Ref<number>
  viewH: Ref<number>
  viewW: Ref<number>
  memory: Memory
}

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v))

export function useViewport(i: ViewportInput, geo: Geometry) {
  const showNums = computed(() => i.sheet.value?.rowNumbers !== false)
  /** A phone-width frame: a narrower row numbers column. */
  const narrow = computed(() => i.viewW.value < 600)
  /** A finger points (a coarse pointer): targets get a finger's size. */
  const coarse = ref(false)
  if (typeof window !== 'undefined' && typeof window.matchMedia === 'function') {
    const query = window.matchMedia('(pointer: coarse)')
    const update = () => (coarse.value = query.matches)
    update()
    query.addEventListener?.('change', update)
    onScopeDispose(() => query.removeEventListener?.('change', update))
  }
  /** A phone's frame or a finger: buttons, fields and tabs at least 44 × 44 px, the grid itself as it is. */
  const touch = computed(() => narrow.value || coarse.value)
  const numW = computed(() => {
    if (!showNums.value) return 0
    if (!narrow.value) return 56
    return clamp(String(i.rowCount.value).length * 7 + 12, 28, 56)
  })

  // ── freeze: from the book, with a lock to drop it; it drops by itself when it eats over half the frame ──
  const bookFreeze = computed(() => {
    const f = i.sheet.value?.freeze
    return !!f && (f.rows > 0 || f.cols > 0)
  })
  const freezeChoice = ref<boolean | null>(null)
  const freezeOn = computed(() => {
    if (!bookFreeze.value) return false
    if (freezeChoice.value !== null) return freezeChoice.value
    const f = i.sheet.value!.freeze!
    const h = geo.rowOffsets.value[Math.min(f.rows, i.rowCount.value)] ?? 0
    const w = geo.offsets.value[Math.min(f.cols, i.colCount.value)] ?? 0
    return h <= i.viewH.value / 2 && w <= i.viewW.value / 2
  })
  const freezeKey = () => `freeze:${i.sheet.value?.name ?? ''}`
  const readFreeze = () => (freezeChoice.value = i.memory.recall<boolean>(freezeKey()))
  function toggleFreeze(): void {
    freezeChoice.value = !freezeOn.value
    i.memory.remember(freezeKey(), freezeChoice.value)
  }
  /** Frozen rows and columns; the whole sheet cannot be frozen, the body keeps one at least. */
  const frozenRows = computed(() => (freezeOn.value ? clamp(i.sheet.value?.freeze?.rows ?? 0, 0, Math.max(0, i.rowCount.value - 1)) : 0))
  const frozenCols = computed(() => (freezeOn.value ? clamp(i.sheet.value?.freeze?.cols ?? 0, 0, Math.max(0, i.colCount.value - 1)) : 0))
  const frozenH = computed(() => geo.rowOffsets.value[frozenRows.value] ?? 0)
  const frozenW = computed(() => geo.offsets.value[frozenCols.value] ?? 0)
  const frozenRowList = computed(() => {
    const out: number[] = []
    for (let r = 1; r <= frozenRows.value; r++) if (geo.heightOf(r) > 0) out.push(r)
    return out
  })
  const frozenColList = computed(() => {
    const out: number[] = []
    for (let c = 1; c <= frozenCols.value; c++) if (geo.widthOf(c) > 0) out.push(c)
    return out
  })
  /** Sticky place of a frozen row or column: below the letters, right of the numbers. */
  const stickyTop = (d: number) => `${HEAD_H + (geo.rowOffsets.value[d - 1] ?? 0)}px`
  const stickyLeft = (c: number) => `${numW.value + (geo.offsets.value[c - 1] ?? 0)}px`

  // ── the body: what is not frozen, under the scroll ──
  const bodyRow0 = computed(() => frozenRows.value + 1)
  const bodyCol0 = computed(() => frozenCols.value + 1)
  const firstRow = computed(() =>
    clamp(geo.rowAt(i.scrollTop.value + frozenH.value) - RENDER_ROWS, bodyRow0.value, Math.max(bodyRow0.value, i.rowCount.value)),
  )
  const lastRow = computed(() => clamp(geo.rowAt(i.scrollTop.value + i.viewH.value) + RENDER_ROWS, bodyRow0.value, i.rowCount.value))
  const firstCol = computed(() =>
    clamp(geo.colAt(i.scrollLeft.value + frozenW.value) - RENDER_COLS, bodyCol0.value, Math.max(bodyCol0.value, i.colCount.value)),
  )
  const lastCol = computed(() => clamp(geo.colAt(i.scrollLeft.value + i.viewW.value) + RENDER_COLS, bodyCol0.value, i.colCount.value))
  /** Shown body rows; hidden ones (height 0) are not drawn at all. */
  const shownRows = computed(() => {
    const out: number[] = []
    for (let d = firstRow.value; d <= lastRow.value; d++) if (geo.heightOf(d) > 0) out.push(d)
    return out
  })
  const shownCols = computed(() => {
    const out: number[] = []
    for (let c = firstCol.value; c <= lastCol.value; c++) if (geo.widthOf(c) > 0) out.push(c)
    return out
  })
  const topPad = computed(() => Math.max(0, (geo.rowOffsets.value[firstRow.value - 1] ?? 0) - frozenH.value))
  const bottomPad = computed(() => Math.max(0, (geo.rowOffsets.value[i.rowCount.value] ?? 0) - (geo.rowOffsets.value[lastRow.value] ?? 0)))
  const leftPad = computed(() => Math.max(0, (geo.offsets.value[firstCol.value - 1] ?? 0) - frozenW.value))
  const rightPad = computed(() => Math.max(0, (geo.offsets.value[i.colCount.value] ?? 0) - (geo.offsets.value[lastCol.value] ?? 0)))
  const tableWidth = computed(() => numW.value + (geo.offsets.value[i.colCount.value] ?? 0))

  return {
    showNums,
    narrow,
    touch,
    numW,
    bookFreeze,
    freezeOn,
    readFreeze,
    toggleFreeze,
    frozenRows,
    frozenCols,
    frozenH,
    frozenW,
    frozenRowList,
    frozenColList,
    stickyTop,
    stickyLeft,
    firstRow,
    lastRow,
    firstCol,
    lastCol,
    shownRows,
    shownCols,
    topPad,
    bottomPad,
    leftPad,
    rightPad,
    tableWidth,
  }
}

export type Viewport = ReturnType<typeof useViewport>
