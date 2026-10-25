// Cell ranges: a single cell's, the one two corners span, the union of two; and a number clamped to bounds.
import type { CellPosition, CellRect } from 'xlcalc/contract'

/** The range of a single cell. */
export const cellRect = (row: number, col: number): CellRect => ({ top: row, left: col, bottom: row, right: col })

/** The range two corners span, in any order. */
export const rectOf = (a: CellPosition, b: CellPosition): CellRect => ({ top: Math.min(a.row, b.row), left: Math.min(a.col, b.col), bottom: Math.max(a.row, b.row), right: Math.max(a.col, b.col) })

/** The smallest range containing both. */
export const unionRect = (a: CellRect, b: CellRect): CellRect => ({ top: Math.min(a.top, b.top), left: Math.min(a.left, b.left), bottom: Math.max(a.bottom, b.bottom), right: Math.max(a.right, b.right) })

/** `v` clamped to `lo`..`hi`. */
export const clamp = (v: number, lo: number, hi: number): number => Math.min(hi, Math.max(lo, v))

/** What remains of `a` outside `b`: up to four ranges, the bands above and below `b` and the parts left and right of it between them. */
export function rectMinus(a: CellRect, b: CellRect): CellRect[] {
  if (b.top > a.bottom || b.bottom < a.top || b.left > a.right || b.right < a.left) return [a]
  const out: CellRect[] = []
  const top = Math.max(a.top, b.top)
  const bottom = Math.min(a.bottom, b.bottom)
  if (a.top < b.top) out.push({ top: a.top, bottom: b.top - 1, left: a.left, right: a.right })
  if (a.bottom > b.bottom) out.push({ top: b.bottom + 1, bottom: a.bottom, left: a.left, right: a.right })
  if (a.left < b.left) out.push({ top, bottom, left: a.left, right: b.left - 1 })
  if (a.right > b.right) out.push({ top, bottom, left: b.right + 1, right: a.right })
  return out
}
