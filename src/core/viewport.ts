// What the grid draws now: frozen panes with the lock, the body's rows and columns under the scroll, spacers for the rest,
// and the space for Excel's outline bars next to the row numbers and above the column letters.
import { computed, ref, type ComputedRef, type Ref } from 'vue'
import { NARROW_PX, useCoarsePointer, useHairline, useShortViewport } from '../frame'
import { clamp } from '../rect'
import type { OutlineGroup, SheetInfo } from '../contract/source'
import type { Geometry } from './geometry'
import type { Memory } from './storage'

/** The height of the column letters row (fixed in CSS); frozen rows stick below it. */
export const HEAD_H = 24
/** A slot of an outline bar for each level, in px: Excel's buttons are about 13 px; touch input gets wider slots. */
const OUTLINE_SLOT_PX = 16
const OUTLINE_SLOT_TOUCH_PX = 22
/** Rows and columns drawn beyond the screen, so a short scroll needs no re-render. */
const RENDER_ROWS = 8
const RENDER_COLS = 2
/** The drawn area's edges move in steps of rows and columns: scrolling redraws the grid once per step, not once per row. */
const ROW_STEP = 4
const COL_STEP = 2

interface ViewportInput {
  sheet: ComputedRef<SheetInfo | null>
  /** The display filter in effect: its rows are not the sheet's rows in order, so the row outline bar is not drawn. */
  visible: Readonly<Ref<object | null>>
  rowCount: ComputedRef<number>
  colCount: ComputedRef<number>
  scrollTop: Ref<number>
  scrollLeft: Ref<number>
  viewH: Ref<number>
  viewW: Ref<number>
  sheetW: Ref<number>
  zoom: Ref<number>
  memory: Memory
}

/** The previous list when the new one holds the same rows or columns: what reads it is not redrawn for a changed size. */
const same = (prev: number[] | undefined, next: number[]): number[] => (prev && prev.length === next.length && prev.every((x, i) => x === next[i]) ? prev : next)
/** The rows or columns `from()..to()` with a size (hidden ones have none and are not drawn at all), the same list while they are the same. */
const shownLines = (from: () => number, to: () => number, sizeOf: (line: number) => number) =>
  computed((prev?: number[]) => {
    const out: number[] = []
    for (let x = from(), last = to(); x <= last; x++) if (sizeOf(x) > 0) out.push(x)
    return same(prev, out)
  })
/**
 * The drawn rows or columns around the screen's `first..last`: `spare` beyond each end, and up to `step - 1` more, since the
 * first one on screen counts from the start of its step; both edges move together, once the screen crosses a step.
 */
function drawnLines(first: number, last: number, spare: number, step: number): [number, number] {
  const lead = (first - 1) % step
  return [first - lead - spare, last + step - 1 - lead + spare]
}

/** What a scroller draws and reads: the rows and columns under its scroll and its size. */
export interface ScrollInput {
  rowCount: Readonly<Ref<number>>
  colCount: Readonly<Ref<number>>
  scrollTop: Readonly<Ref<number>>
  scrollLeft: Readonly<Ref<number>>
  viewH: Readonly<Ref<number>>
  viewW: Readonly<Ref<number>>
}
/** The parts of the grid that do not scroll: frozen rows and columns and the headings above and to the left of the cells. */
interface FixedParts {
  frozenRows: Readonly<Ref<number>>
  frozenCols: Readonly<Ref<number>>
  frozenH: Readonly<Ref<number>>
  frozenW: Readonly<Ref<number>>
  headH: Readonly<Ref<number>>
  leadW: Readonly<Ref<number>>
}

/** The body under a scroll: the rows and columns drawn, spacers for the rest, and the part in view; a split window's panes each have their own. */
export function scrolledArea(i: ScrollInput, geo: Geometry, f: FixedParts) {
  const { frozenRows, frozenCols, frozenH, frozenW, headH, leadW } = f
  const bodyRow0 = computed(() => frozenRows.value + 1)
  const bodyCol0 = computed(() => frozenCols.value + 1)
  const drawnRows = computed(() => drawnLines(geo.rowAt(i.scrollTop.value + frozenH.value), geo.rowAt(i.scrollTop.value + i.viewH.value), RENDER_ROWS, ROW_STEP))
  const drawnCols = computed(() => drawnLines(geo.colAt(i.scrollLeft.value + frozenW.value), geo.colAt(i.scrollLeft.value + i.viewW.value), RENDER_COLS, COL_STEP))
  const firstRow = computed(() => clamp(drawnRows.value[0], bodyRow0.value, Math.max(bodyRow0.value, i.rowCount.value)))
  const lastRow = computed(() => clamp(drawnRows.value[1], bodyRow0.value, i.rowCount.value))
  const firstCol = computed(() => clamp(drawnCols.value[0], bodyCol0.value, Math.max(bodyCol0.value, i.colCount.value)))
  const lastCol = computed(() => clamp(drawnCols.value[1], bodyCol0.value, i.colCount.value))
  /** Visible body rows and columns. */
  const shownRows = shownLines(() => firstRow.value, () => lastRow.value, geo.heightOf)
  const shownCols = shownLines(() => firstCol.value, () => lastCol.value, geo.widthOf)
  const topPad = computed(() => Math.max(0, (geo.rowOffsets.value[firstRow.value - 1] ?? 0) - frozenH.value))
  const bottomPad = computed(() => Math.max(0, (geo.rowOffsets.value[i.rowCount.value] ?? 0) - (geo.rowOffsets.value[lastRow.value] ?? 0)))
  const leftPad = computed(() => Math.max(0, (geo.offsets.value[firstCol.value - 1] ?? 0) - frozenW.value))
  const rightPad = computed(() => Math.max(0, (geo.offsets.value[i.colCount.value] ?? 0) - (geo.offsets.value[lastCol.value] ?? 0)))
  /** The part of the grid in view, in the scroller's content coordinates: a popup goes over its cell where it has no room below. */
  const viewBox = computed(() => ({
    top: i.scrollTop.value + headH.value,
    left: i.scrollLeft.value + leadW.value,
    width: Math.max(0, i.viewW.value - leadW.value),
    height: Math.max(0, i.viewH.value - headH.value),
  }))
  return { firstRow, lastRow, firstCol, lastCol, shownRows, shownCols, topPad, bottomPad, leftPad, rightPad, viewBox }
}

export function useViewport(i: ViewportInput, geo: Geometry) {
  /** Excel's View > Show > Headings: row numbers and column letters; a list's `rowNumbers` hides the numbers alone. */
  const showHeads = computed(() => i.sheet.value?.headings !== false)
  const showNums = computed(() => showHeads.value && i.sheet.value?.rowNumbers !== false)
  /**
   * A gridline's width in unzoomed pixels: 1, and zoomed out what draws as one screen pixel, since browsers draw a thinner
   * border one pixel wide and the rows would then outgrow the geometry. `lineOf` widens a border line the same way.
   */
  const line = computed(() => (i.zoom.value < 1 ? 1 / i.zoom.value : 1))
  const lineOf = (px: number): number => (px > 0 ? Math.max(px, line.value) : 0)
  const hairline = useHairline()
  /** A border's width as the browser draws it under the zoom, in unzoomed pixels: rounded down to whole hairlines, at least one. */
  const drawnOf = (px: number): number => {
    const step = hairline.value
    return px > 0 ? (Math.max(1, Math.floor((px * i.zoom.value) / step + 1e-6)) * step) / i.zoom.value : 0
  }
  /** The hairline's width in unzoomed pixels: the step borders are drawn in. */
  const hairlineOf = (): number => hairline.value / i.zoom.value
  /**
   * The stroke of a double border this wide, in hairlines, where every browser draws it the same; 0 where it is drawn solid
   * or browsers differ. Of n hairlines, fewer than three are solid; Chromium rounds a stroke to n / 3, WebKit rounds it up
   * (four hairlines are its 2 + 2, solid), so they differ when n is one more than a multiple of three.
   */
  const doubleStroke = (px: number): number => {
    const n = Math.round((drawnOf(px) * i.zoom.value) / hairline.value)
    return n >= 3 && n % 3 !== 1 ? Math.round(n / 3) : 0
  }
  /**
   * A narrow sheet, by its own width on screen whatever the zoom: a narrower row number column, panels below the grid.
   * Not by the grid's width, which a panel beside it narrows: the panel would then move below and back without end.
   */
  const narrow = computed(() => i.sheetW.value < NARROW_PX)
  /** A coarse pointer (touch): buttons, inputs and tabs at least 44 × 44 px, the grid itself unchanged. */
  const touch = useCoarsePointer()
  /** A finger on a narrow sheet: the formula and the tabs on their own lines, the filter's list on the full screen. A mouse in a narrow window stays the desktop. */
  const short = useShortViewport()
  const phone = computed(() => touch.value && (narrow.value || (short.value && i.sheetW.value < 1000)))
  const numW = computed(() => {
    if (!showNums.value) return 0
    if (!narrow.value) return 56
    return clamp(String(i.rowCount.value).length * 7 + 12, 28, 56)
  })

  // ── Excel's outline bars: a slot for each group level and one more (the level buttons 1…N+1) ──
  const deepest = (groups: OutlineGroup[] | undefined) => (groups ?? []).reduce((n, g) => Math.max(n, g.level), 0)
  /** Slots of the row bar (left of the row numbers) and of the column bar (above the column letters); 0 without groups. */
  const outlineSlots = computed(() => {
    const o = i.sheet.value?.outline
    const rows = i.visible.value ? 0 : deepest(o?.rows)
    const cols = deepest(o?.columns)
    return { rows: rows ? rows + 1 : 0, cols: cols ? cols + 1 : 0 }
  })
  const slotPx = computed(() => (touch.value ? OUTLINE_SLOT_TOUCH_PX : OUTLINE_SLOT_PX))
  const barW = computed(() => outlineSlots.value.rows * slotPx.value)
  const barH = computed(() => outlineSlots.value.cols * slotPx.value)
  /** The header's height (the letters and the column bar) and the width to the left of the cells (the row bar and the numbers). */
  /** The letters row goes with the headings, the level buttons with it, as Excel for Mac draws no level buttons without headings. */
  const headRow = computed(() => showHeads.value)
  const headH = computed(() => (headRow.value ? HEAD_H : 0) + barH.value)
  /** List mode's totals row at the bottom, as tall as the column letters row. */
  const footH = computed(() => (i.sheet.value?.totals ? HEAD_H : 0))
  const leadW = computed(() => barW.value + numW.value)

  // ── freeze: from the workbook, with a lock to hide it (a view setting: the workbook keeps its panes); it hides itself when it takes over half the viewport ──
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
  /** Panes the user froze with a command are displayed regardless of the lock and the half-viewport rule; unfreezing clears the lock's choice. */
  function showFreeze(sheet: string, on: boolean): void {
    if (i.sheet.value?.name !== sheet) return
    freezeChoice.value = on ? true : null
    i.memory.remember(freezeKey(), freezeChoice.value)
  }
  /** Excel's Freeze Panes with A1 active: the rows above and the columns to the left of the viewport's middle, counted from the sheet's first. */
  function middlePanes(): { rows: number; cols: number } {
    const d = geo.rowAt(Math.max(0, (i.viewH.value - headH.value) / 2))
    const c = geo.colAt(Math.max(0, (i.viewW.value - leadW.value) / 2))
    return { rows: d - 1, cols: c - 1 }
  }
  /** Frozen rows and columns; the entire sheet cannot be frozen, as the body keeps at least one. */
  const frozenRows = computed(() => (freezeOn.value ? clamp(i.sheet.value?.freeze?.rows ?? 0, 0, Math.max(0, i.rowCount.value - 1)) : 0))
  const frozenCols = computed(() => (freezeOn.value ? clamp(i.sheet.value?.freeze?.cols ?? 0, 0, Math.max(0, i.colCount.value - 1)) : 0))
  const frozenH = computed(() => geo.rowOffsets.value[frozenRows.value] ?? 0)
  const frozenW = computed(() => geo.offsets.value[frozenCols.value] ?? 0)
  const frozenRowList = shownLines(() => 1, () => frozenRows.value, geo.heightOf)
  const frozenColList = shownLines(() => 1, () => frozenCols.value, geo.widthOf)
  /** The sticky position of a frozen row or column: below the column letters, to the right of the row numbers. */
  const stickyTop = (d: number) => `${headH.value + (geo.rowOffsets.value[d - 1] ?? 0)}px`
  const stickyLeft = (c: number) => `${leadW.value + (geo.offsets.value[c - 1] ?? 0)}px`

  // ── the body: what is not frozen, under the scroll ──
  const area = scrolledArea(i, geo, { frozenRows, frozenCols, frozenH, frozenW, headH, leadW })
  const tableWidth = computed(() => leadW.value + (geo.offsets.value[i.colCount.value] ?? 0))

  /** Where display rows d1..d2 and columns c1..c2 are drawn in the scroller's content; one entirely in frozen rows or columns follows the scroll, since they stick. */
  function boxOf(d1: number, c1: number, d2: number, c2: number): { top: number; left: number; width: number; height: number; frozen: boolean } {
    const rows = geo.rowOffsets.value
    const cols = geo.offsets.value
    const top = (rows[d1 - 1] ?? 0) + (d2 <= frozenRows.value ? i.scrollTop.value : 0)
    const left = (cols[c1 - 1] ?? 0) + (c2 <= frozenCols.value ? i.scrollLeft.value : 0)
    return { top: headH.value + top, left: leadW.value + left, width: (cols[c2] ?? 0) - (cols[c1 - 1] ?? 0), height: (rows[d2] ?? 0) - (rows[d1 - 1] ?? 0), frozen: d2 <= frozenRows.value || c2 <= frozenCols.value }
  }
  /** The display row and column under a point of the viewport (pixels from its top-left corner), with the point in the offsets' space (`x`, `y`). */
  function cellAt(px: number, py: number): { d: number; col: number; x: number; y: number } {
    const cy = Math.max(0, py - headH.value)
    const cx = Math.max(0, px - leadW.value)
    const y = cy < frozenH.value ? cy : cy + i.scrollTop.value
    const x = cx < frozenW.value ? cx : cx + i.scrollLeft.value
    return { d: geo.rowAt(y), col: geo.colAt(x), x, y }
  }
  /** The same under a point of the page (a pointer's `clientX` and `clientY`), with `box` the scroller's rectangle: the page's pixels are the grid's zoomed ones. */
  const cellAtPoint = (clientX: number, clientY: number, box: { left: number; top: number }) => cellAt((clientX - box.left) / i.zoom.value, (clientY - box.top) / i.zoom.value)

  return {
    showHeads,
    showNums,
    headRow,
    line,
    lineOf,
    drawnOf,
    doubleStroke,
    hairlineOf,
    narrow,
    phone,
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
    ...area,
    tableWidth,
    boxOf,
    cellAt,
    cellAtPoint,
  }
}

export type Viewport = ReturnType<typeof useViewport>
