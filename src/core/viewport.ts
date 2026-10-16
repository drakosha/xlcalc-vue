// What the grid draws now: frozen panes with the lock, the body's rows and columns under the scroll, spacers for the rest,
// and the room of Excel's outline bars beside the numbers and above the letters.
import { computed, ref, type ComputedRef, type Ref } from 'vue'
import { NARROW_PX, useCoarsePointer } from '../frame'
import { clamp } from '../rect'
import type { OutlineGroup, SheetInfo } from '../source'
import type { Geometry } from './geometry'
import type { Memory } from './storage'

/** Height of the column letters row (fixed in CSS); frozen rows stick below it. */
export const HEAD_H = 24
/** A slot of an outline bar for each level, in px: Excel's buttons are about 13 px; a finger gets wider slots. */
const OUTLINE_SLOT_PX = 16
const OUTLINE_SLOT_TOUCH_PX = 22
/** Rows and columns drawn beyond the screen, so a short scroll needs no re-render. */
const RENDER_ROWS = 8
const RENDER_COLS = 2
/** The drawn piece's edges move by steps of rows and columns: a scroll redraws the grid once a step, not once a row. */
const ROW_STEP = 4
const COL_STEP = 2

interface ViewportInput {
  sheet: ComputedRef<SheetInfo | null>
  /** The display filter in force: its rows are not the sheet's in order, so the rows' outline bar is not drawn. */
  visible: Readonly<Ref<object | null>>
  rowCount: ComputedRef<number>
  colCount: ComputedRef<number>
  scrollTop: Ref<number>
  scrollLeft: Ref<number>
  viewH: Ref<number>
  viewW: Ref<number>
  memory: Memory
}

/** The previous list when the new one holds the same lines: what reads it is not drawn again for a size changed. */
const same = (prev: number[] | undefined, next: number[]): number[] => (prev && prev.length === next.length && prev.every((x, i) => x === next[i]) ? prev : next)
/** The lines `from()..to()` with a size (hidden ones have none, and are not drawn at all), the same list while they are the same. */
const shownLines = (from: () => number, to: () => number, sizeOf: (line: number) => number) =>
  computed((prev?: number[]) => {
    const out: number[] = []
    for (let x = from(), last = to(); x <= last; x++) if (sizeOf(x) > 0) out.push(x)
    return same(prev, out)
  })
/**
 * The drawn lines around the screen's `first..last`: `spare` beyond each end, and up to `step - 1` more, as the first one
 * on screen stands from the start of its step; both edges move together, once the screen crosses a step.
 */
function drawnLines(first: number, last: number, spare: number, step: number): [number, number] {
  const lead = (first - 1) % step
  return [first - lead - spare, last + step - 1 - lead + spare]
}

export function useViewport(i: ViewportInput, geo: Geometry) {
  const showNums = computed(() => i.sheet.value?.rowNumbers !== false)
  /** A phone-width frame: a narrower row numbers column. */
  const narrow = computed(() => i.viewW.value < NARROW_PX)
  /** A finger points (a coarse pointer): targets get a finger's size. */
  const coarse = useCoarsePointer()
  /** A phone's frame or a finger: buttons, fields and tabs at least 44 × 44 px, the grid itself as it is. */
  const touch = computed(() => narrow.value || coarse.value)
  const numW = computed(() => {
    if (!showNums.value) return 0
    if (!narrow.value) return 56
    return clamp(String(i.rowCount.value).length * 7 + 12, 28, 56)
  })

  // ── Excel's outline bars: a slot for each level of groups and one more (the level buttons 1…N+1) ──
  const deepest = (groups: OutlineGroup[] | undefined) => (groups ?? []).reduce((n, g) => Math.max(n, g.level), 0)
  /** Slots of the rows' bar (left of the numbers) and of the columns' bar (above the letters); 0 without groups. */
  const outlineSlots = computed(() => {
    const o = i.sheet.value?.outline
    const rows = i.visible.value ? 0 : deepest(o?.rows)
    const cols = deepest(o?.columns)
    return { rows: rows ? rows + 1 : 0, cols: cols ? cols + 1 : 0 }
  })
  const slotPx = computed(() => (touch.value ? OUTLINE_SLOT_TOUCH_PX : OUTLINE_SLOT_PX))
  const barW = computed(() => outlineSlots.value.rows * slotPx.value)
  const barH = computed(() => outlineSlots.value.cols * slotPx.value)
  /** The header's height (letters and the columns' bar) and the width left of the cells (the rows' bar and numbers). */
  const headH = computed(() => HEAD_H + barH.value)
  /** List mode's totals row at the bottom, as high as the letters' row. */
  const footH = computed(() => (i.sheet.value?.totals ? HEAD_H : 0))
  const leadW = computed(() => barW.value + numW.value)

  // ── freeze: from the book, with a lock to drop it (a view: the book keeps its panes); it drops by itself when it eats over half the frame ──
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
  /** A freeze the person made with a command shows over the lock and the half-frame rule; an unfreeze forgets the lock's choice. */
  function showFreeze(sheet: string, on: boolean): void {
    if (i.sheet.value?.name !== sheet) return
    freezeChoice.value = on ? true : null
    i.memory.remember(freezeKey(), freezeChoice.value)
  }
  /** Excel's Freeze Panes from A1: the rows above and the columns left of the frame's middle, counted from the sheet's first. */
  function middlePanes(): { rows: number; cols: number } {
    const d = geo.rowAt(Math.max(0, (i.viewH.value - headH.value) / 2))
    const c = geo.colAt(Math.max(0, (i.viewW.value - leadW.value) / 2))
    return { rows: d - 1, cols: c - 1 }
  }
  /** Frozen rows and columns; the whole sheet cannot be frozen, the body keeps one at least. */
  const frozenRows = computed(() => (freezeOn.value ? clamp(i.sheet.value?.freeze?.rows ?? 0, 0, Math.max(0, i.rowCount.value - 1)) : 0))
  const frozenCols = computed(() => (freezeOn.value ? clamp(i.sheet.value?.freeze?.cols ?? 0, 0, Math.max(0, i.colCount.value - 1)) : 0))
  const frozenH = computed(() => geo.rowOffsets.value[frozenRows.value] ?? 0)
  const frozenW = computed(() => geo.offsets.value[frozenCols.value] ?? 0)
  const frozenRowList = shownLines(() => 1, () => frozenRows.value, geo.heightOf)
  const frozenColList = shownLines(() => 1, () => frozenCols.value, geo.widthOf)
  /** Sticky place of a frozen row or column: below the letters, right of the numbers. */
  const stickyTop = (d: number) => `${headH.value + (geo.rowOffsets.value[d - 1] ?? 0)}px`
  const stickyLeft = (c: number) => `${leadW.value + (geo.offsets.value[c - 1] ?? 0)}px`

  // ── the body: what is not frozen, under the scroll ──
  const bodyRow0 = computed(() => frozenRows.value + 1)
  const bodyCol0 = computed(() => frozenCols.value + 1)
  const drawnRows = computed(() => drawnLines(geo.rowAt(i.scrollTop.value + frozenH.value), geo.rowAt(i.scrollTop.value + i.viewH.value), RENDER_ROWS, ROW_STEP))
  const drawnCols = computed(() => drawnLines(geo.colAt(i.scrollLeft.value + frozenW.value), geo.colAt(i.scrollLeft.value + i.viewW.value), RENDER_COLS, COL_STEP))
  const firstRow = computed(() => clamp(drawnRows.value[0], bodyRow0.value, Math.max(bodyRow0.value, i.rowCount.value)))
  const lastRow = computed(() => clamp(drawnRows.value[1], bodyRow0.value, i.rowCount.value))
  const firstCol = computed(() => clamp(drawnCols.value[0], bodyCol0.value, Math.max(bodyCol0.value, i.colCount.value)))
  const lastCol = computed(() => clamp(drawnCols.value[1], bodyCol0.value, i.colCount.value))
  /** Shown body rows and columns. */
  const shownRows = shownLines(() => firstRow.value, () => lastRow.value, geo.heightOf)
  const shownCols = shownLines(() => firstCol.value, () => lastCol.value, geo.widthOf)
  const topPad = computed(() => Math.max(0, (geo.rowOffsets.value[firstRow.value - 1] ?? 0) - frozenH.value))
  const bottomPad = computed(() => Math.max(0, (geo.rowOffsets.value[i.rowCount.value] ?? 0) - (geo.rowOffsets.value[lastRow.value] ?? 0)))
  const leftPad = computed(() => Math.max(0, (geo.offsets.value[firstCol.value - 1] ?? 0) - frozenW.value))
  const rightPad = computed(() => Math.max(0, (geo.offsets.value[i.colCount.value] ?? 0) - (geo.offsets.value[lastCol.value] ?? 0)))
  const tableWidth = computed(() => leadW.value + (geo.offsets.value[i.colCount.value] ?? 0))

  /** Where display rows d1..d2 and columns c1..c2 are drawn in the scroller's content; one wholly in frozen rows or columns follows the scroll, as they stick. */
  function boxOf(d1: number, c1: number, d2: number, c2: number): { top: number; left: number; width: number; height: number; frozen: boolean } {
    const rows = geo.rowOffsets.value
    const cols = geo.offsets.value
    const top = (rows[d1 - 1] ?? 0) + (d2 <= frozenRows.value ? i.scrollTop.value : 0)
    const left = (cols[c1 - 1] ?? 0) + (c2 <= frozenCols.value ? i.scrollLeft.value : 0)
    return { top: headH.value + top, left: leadW.value + left, width: (cols[c2] ?? 0) - (cols[c1 - 1] ?? 0), height: (rows[d2] ?? 0) - (rows[d1 - 1] ?? 0), frozen: d2 <= frozenRows.value || c2 <= frozenCols.value }
  }
  /** What of the grid is in view, in the same coordinates: a popup goes over its cell where it has no room below. */
  const viewBox = computed(() => ({
    top: i.scrollTop.value + headH.value,
    left: i.scrollLeft.value + leadW.value,
    width: Math.max(0, i.viewW.value - leadW.value),
    height: Math.max(0, i.viewH.value - headH.value),
  }))
  /** The display row and column under a point of the frame (pixels from its top-left corner), with the point in the offsets' space (`x`, `y`). */
  function cellAt(px: number, py: number): { d: number; col: number; x: number; y: number } {
    const cy = Math.max(0, py - headH.value)
    const cx = Math.max(0, px - leadW.value)
    const y = cy < frozenH.value ? cy : cy + i.scrollTop.value
    const x = cx < frozenW.value ? cx : cx + i.scrollLeft.value
    return { d: geo.rowAt(y), col: geo.colAt(x), x, y }
  }

  return {
    showNums,
    narrow,
    touch,
    numW,
    outlineSlots,
    slotPx,
    barW,
    barH,
    headH,
    footH,
    leadW,
    bookFreeze,
    freezeOn,
    readFreeze,
    toggleFreeze,
    showFreeze,
    middlePanes,
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
    boxOf,
    viewBox,
    cellAt,
  }
}

export type Viewport = ReturnType<typeof useViewport>
