// Rectangles of cells: one cell's, the one two corners span, the union of two; and a number kept within bounds.
import type { CellPosition, CellRect } from 'xlcalc/contract'

/** The rectangle of one cell. */
export const cellRect = (row: number, col: number): CellRect => ({ top: row, left: col, bottom: row, right: col })

/** The rectangle two corners span, in any order. */
export const rectOf = (a: CellPosition, b: CellPosition): CellRect => ({ top: Math.min(a.row, b.row), left: Math.min(a.col, b.col), bottom: Math.max(a.row, b.row), right: Math.max(a.col, b.col) })

/** The smallest rectangle holding both. */
export const unionRect = (a: CellRect, b: CellRect): CellRect => ({ top: Math.min(a.top, b.top), left: Math.min(a.left, b.left), bottom: Math.max(a.bottom, b.bottom), right: Math.max(a.right, b.right) })

/** `v` kept within `lo`..`hi`. */
export const clamp = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v))
