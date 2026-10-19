// Grid geometry: Excel units to pixels, prefix sums of widths and heights, and lookups by offset.
import { computed, triggerRef, type Ref, type ShallowRef } from 'vue'
import { columnWidthFromPoints, columnWidthToPoints, MAX_ROW_HEIGHT } from 'xlcalc/text'

/** A column's resizing limits, in pixels. */
export const MIN_COL_PX = 20
export const MAX_COL_PX = 900

/** The widest column xlsx stores. */
const MAX_COL_WIDTH = 255

/** A column width as xlsx stores it, in whole pixels in the library's geometry (Excel for Mac's, 53 points without a width), the one xlcalc uses to lay out `####`, overflow and wrapping. */
export function colPx(width: number | null | undefined): number {
  if (width !== null && width !== undefined && width <= 0) return 0
  return Math.round((columnWidthToPoints(width) * 4) / 3)
}

/** The widest column the library's geometry accepts, in points. */
const MAX_COL_PT = 1535

/** Pixels back to the stored width, in whole points as the library stores them. */
export const widthFromPx = (px: number): number => Math.min(MAX_COL_WIDTH, columnWidthFromPoints(Math.min(MAX_COL_PT, Math.max(0, (px * 3) / 4))))

/** Points to pixels at 96 dpi. */
export const rowPx = (pt: number): number => (pt <= 0 ? 0 : Math.round((pt * 4) / 3))

/** A row dragged by its edge: down to 2 px, up to Excel's tallest row. */
export const MIN_ROW_PX = 2
export const MAX_ROW_PX = rowPx(MAX_ROW_HEIGHT)

/** Pixels back to points, rounded to Excel's quarter point. */
export const ptFromPx = (px: number): number => Math.round(((px * 3) / 4) * 4) / 4

/** A font size in points to the grid's pixels: 11 pt is the grid's base 12.5 px, the rest in proportion. */
export const fontPx = (pt: number): number => Math.round(((pt * 12.5) / 11) * 10) / 10

/** Prefix sums: out[i] is the far edge of item i (out[0] = 0). */
export function prefixSums(count: number, sizeOf: (i: number) => number): Float64Array {
  const out = new Float64Array(count + 1)
  for (let i = 1; i <= count; i++) out[i] = out[i - 1]! + sizeOf(i)
  return out
}

const sameSet = (a: Set<number>, b: Set<number>): boolean => a.size === b.size && [...a].every((x) => b.has(x))

/** The item under an offset, by binary search over prefix sums; clamped to 1..count. */
export function indexAt(offsets: Float64Array, count: number, x: number): number {
  let lo = 1
  let hi = count
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (offsets[mid]! > x) hi = mid
    else lo = mid + 1
  }
  return Math.min(Math.max(1, count), Math.max(1, lo))
}

interface GeometryInput {
  rowCount: Readonly<Ref<number>>
  colCount: Readonly<Ref<number>>
  /** Column widths as the workbook stores them (null: the default). */
  bookWidths: ShallowRef<Map<number, number | null>>
  /** Widths set with the mouse, in pixels by column; they take precedence over the workbook's. */
  userWidths: ShallowRef<Record<number, number>>
  /** Row heights in pixels that differ from the default (0 hides a row): changed in place through `setHeights`, or replaced as a whole. */
  heights: ShallowRef<Map<number, number>>
  defaultRowPx: Ref<number>
  /** The display filter in effect: another one is another order of rows. */
  visible: Readonly<Ref<unknown>>
  /** The sheet row displayed at a display index (heights are the sheet rows') and back. */
  rowOf: (d: number) => number
  indexOf: (r: number) => number | undefined
}

/** Offsets of columns and rows, and the item under an offset; rows are display indexes (`d`). */
export function useGeometry(i: GeometryInput) {
  const widthOf = (c: number): number => i.userWidths.value[c] ?? colPx(i.bookWidths.value.get(c))
  const heightOf = (d: number): number => i.heights.value.get(i.rowOf(d)) ?? i.defaultRowPx.value
  const offsets = computed(() => prefixSums(i.colCount.value, widthOf))
  /** Display rows whose heights `setHeights` changed since the offsets were summed, and what they were summed over. */
  let changed: { from: number; to: number } | null = null
  let summed: { heights: Map<number, number>; def: number; count: number; visible: unknown } | null = null
  /** Sheet row heights in pixels, in place (undefined: the default); true when one changed. */
  function setHeights(rows: Iterable<readonly [number, number | undefined]>): boolean {
    const heights = i.heights.value
    let any = false
    for (const [r, px] of rows) {
      if (heights.get(r) === px) continue
      if (px === undefined) heights.delete(r)
      else heights.set(r, px)
      any = true
      const d = i.indexOf(r)
      if (d !== undefined) changed = changed ? { from: Math.min(changed.from, d), to: Math.max(changed.to, d) } : { from: d, to: d }
    }
    if (any) triggerRef(i.heights)
    return any
  }
  /** Row offsets: summed again from the first row `setHeights` changed and shifted after the last (a million rows sum slowly, a shift is a copy); summed in full after anything else. */
  const rowOffsets = computed((prev?: Float64Array) => {
    const heights = i.heights.value
    const def = i.defaultRowPx.value
    const count = i.rowCount.value
    const visible = i.visible.value
    const size = (d: number) => heights.get(i.rowOf(d)) ?? def
    const same = !!prev && summed?.heights === heights && summed.def === def && summed.count === count && summed.visible === visible
    const range = changed
    changed = null
    summed = { heights, def, count, visible }
    if (!same) return prefixSums(count, size)
    if (!range || range.from > count) return prev
    const to = Math.min(range.to, count)
    const out = new Float64Array(count + 1)
    out.set(prev.subarray(0, range.from))
    for (let d = range.from; d <= to; d++) out[d] = out[d - 1]! + size(d)
    const shift = out[to]! - prev[to]!
    for (let d = to + 1; d <= count; d++) out[d] = prev[d]! + shift
    return out
  })
  const colAt = (x: number): number => indexAt(offsets.value, i.colCount.value, x)
  const rowAt = (y: number): number => indexAt(rowOffsets.value, i.rowCount.value, y)
  /**
   * Hidden display rows and columns (no height, no width), the same set while they are the same: a cell finds its
   * neighbors by these, so a size changed elsewhere does not redraw it. Null for rows hidden by default (a workbook whose
   * default height is 0), which are found by their heights.
   */
  const hiddenRows = computed((prev?: Set<number> | null) => {
    if (i.defaultRowPx.value === 0) return null
    const out = new Set<number>()
    for (const [r, px] of i.heights.value) {
      const d = px === 0 ? i.indexOf(r) : undefined
      if (d !== undefined) out.add(d)
    }
    return prev && sameSet(prev, out) ? prev : out
  })
  const hiddenCols = computed((prev?: Set<number>) => {
    const out = new Set<number>()
    for (const [c, width] of i.bookWidths.value) if (colPx(width) === 0 && i.userWidths.value[c] === undefined) out.add(c)
    for (const [c, px] of Object.entries(i.userWidths.value)) if (px === 0) out.add(Number(c))
    return prev && sameSet(prev, out) ? prev : out
  })
  const rowHidden = (d: number): boolean => {
    const set = hiddenRows.value
    return set ? set.has(d) : heightOf(d) === 0
  }
  /** The next visible row in a direction, skipping hidden ones; at the edge, the same row. */
  function nextRow(d: number, dir: 1 | -1): number {
    let n = d + dir
    while (n >= 1 && n <= i.rowCount.value && rowHidden(n)) n += dir
    return n >= 1 && n <= i.rowCount.value ? n : d
  }
  function nextCol(c: number, dir: 1 | -1): number {
    let n = c + dir
    while (n >= 1 && n <= i.colCount.value && hiddenCols.value.has(n)) n += dir
    return n >= 1 && n <= i.colCount.value ? n : c
  }
  /** Whether a visible row or column lies further in a direction (the edge for a border lookup). */
  const hasRow = (d: number, dir: 1 | -1): boolean => nextRow(d, dir) !== d
  const hasCol = (c: number, dir: 1 | -1): boolean => nextCol(c, dir) !== c
  /** The height of rows d..d+rs-1 together (a merged range). */
  const spanHeight = (d: number, rs = 1): number => {
    const off = rowOffsets.value
    return (off[Math.min(d + rs - 1, i.rowCount.value)] ?? 0) - (off[d - 1] ?? 0)
  }
  const spanWidth = (c: number, cs = 1): number => {
    const off = offsets.value
    return (off[Math.min(c + cs - 1, i.colCount.value)] ?? 0) - (off[c - 1] ?? 0)
  }
  return { widthOf, heightOf, offsets, setHeights, rowOffsets, colAt, rowAt, nextRow, nextCol, hasRow, hasCol, spanHeight, spanWidth }
}

export type Geometry = ReturnType<typeof useGeometry>
