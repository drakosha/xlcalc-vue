// Grid geometry: Excel units to pixels, prefix sums of widths and heights, and lookups by offset.
import { computed, type Ref, type ShallowRef } from 'vue'

/** Excel's default row height in points. */
export const DEFAULT_ROW_PT = 15
/** Excel's maximum digit width of the default font, in pixels at 100 %. */
const DIGIT_PX = 7
/** Resizing limits of a column, in pixels. */
export const MIN_COL_PX = 20
export const MAX_COL_PX = 900

/** The width of a column without its own, as xlcalc counts it: Excel for Mac with Calibri 11. */
const DEFAULT_COL_WIDTH = 8.83203125

/** A column width as xlsx stores it (digits with padding) to pixels, Excel's formula; null is the library's default, 62 px. */
export function colPx(width: number | null | undefined): number {
  const w = width ?? DEFAULT_COL_WIDTH
  if (w <= 0) return 0
  return Math.trunc(((256 * w + Math.trunc(128 / DIGIT_PX)) / 256) * DIGIT_PX)
}

/** A column's pixels as Excel's Column Width box shows them: characters of the default font, to a hundredth. */
export const charsOfPx = (px: number): number => Math.max(0, Math.trunc(((px - 5) / DIGIT_PX) * 100 + 0.5) / 100)

/** Characters typed into the Column Width box to the stored width, as Excel records it. */
export const widthFromChars = (chars: number): number => Math.trunc(((chars * DIGIT_PX + 5) / DIGIT_PX) * 256) / 256

/** Pixels back to the stored width, as Excel records a width set with the mouse. */
export const widthFromPx = (px: number): number => widthFromChars(charsOfPx(px))

/** Points to pixels at 96 dpi. */
export const rowPx = (pt: number): number => (pt <= 0 ? 0 : Math.round((pt * 4) / 3))

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

export interface GeometryInput {
  rowCount: Readonly<Ref<number>>
  colCount: Readonly<Ref<number>>
  /** Book widths in pixels by column, as the windows brought them. */
  bookWidths: ShallowRef<Map<number, number>>
  /** Widths set with the mouse, pixels by column; they win over the book. */
  userWidths: Ref<Record<number, number>>
  /** Row heights in pixels that differ from the default (0 hides a row). */
  heights: ShallowRef<Map<number, number>>
  defaultRowPx: Ref<number>
  /** The sheet row shown at a display index (heights are the sheet rows'). */
  rowOf: (d: number) => number
}

/** Offsets of columns and rows, and the item under an offset; rows are display indexes (`d`). */
export function useGeometry(i: GeometryInput) {
  const widthOf = (c: number): number => i.userWidths.value[c] ?? i.bookWidths.value.get(c) ?? colPx(null)
  const heightOf = (d: number): number => i.heights.value.get(i.rowOf(d)) ?? i.defaultRowPx.value
  const offsets = computed(() => prefixSums(i.colCount.value, widthOf))
  const rowOffsets = computed(() => {
    const heights = i.heights.value
    const def = i.defaultRowPx.value
    return prefixSums(i.rowCount.value, (d) => heights.get(i.rowOf(d)) ?? def)
  })
  const colAt = (x: number): number => indexAt(offsets.value, i.colCount.value, x)
  const rowAt = (y: number): number => indexAt(rowOffsets.value, i.rowCount.value, y)
  /** The next shown row in a direction, skipping hidden ones; at the edge the same row. */
  function nextRow(d: number, dir: 1 | -1): number {
    let n = d + dir
    while (n >= 1 && n <= i.rowCount.value && heightOf(n) === 0) n += dir
    return n >= 1 && n <= i.rowCount.value ? n : d
  }
  function nextCol(c: number, dir: 1 | -1): number {
    let n = c + dir
    while (n >= 1 && n <= i.colCount.value && widthOf(n) === 0) n += dir
    return n >= 1 && n <= i.colCount.value ? n : c
  }
  /** Whether a shown row or column lies beyond in a direction (the edge of a border lookup). */
  const hasRow = (d: number, dir: 1 | -1): boolean => nextRow(d, dir) !== d
  const hasCol = (c: number, dir: 1 | -1): boolean => nextCol(c, dir) !== c
  /** Height of rows d..d+rs-1 together (a merged area). */
  const spanHeight = (d: number, rs = 1): number => {
    const off = rowOffsets.value
    return (off[Math.min(d + rs - 1, i.rowCount.value)] ?? 0) - (off[d - 1] ?? 0)
  }
  const spanWidth = (c: number, cs = 1): number => {
    const off = offsets.value
    return (off[Math.min(c + cs - 1, i.colCount.value)] ?? 0) - (off[c - 1] ?? 0)
  }
  return { widthOf, heightOf, offsets, rowOffsets, colAt, rowAt, nextRow, nextCol, hasRow, hasCol, spanHeight, spanWidth }
}

export type Geometry = ReturnType<typeof useGeometry>
