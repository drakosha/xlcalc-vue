// The range Excel's Data > Filter puts an AutoFilter over, as measured in Excel 16 for Mac (Range.AutoFilter): the selection,
// or a single cell's current region, trimmed to the cells with data in it and taken down to where the data around it end.
import type { CellRect } from 'xlcalc/contract'
import { currentRegion } from './tables'

/** The part of `r` within `area`, null when they do not meet. */
function within(r: CellRect, area: CellRect): CellRect | null {
  const out = { top: Math.max(r.top, area.top), left: Math.max(r.left, area.left), bottom: Math.min(r.bottom, area.bottom), right: Math.min(r.right, area.right) }
  return out.top <= out.bottom && out.left <= out.right ? out : null
}

/** The smallest range holding every filled cell of `r`, null when it has none. */
function dataBox(r: CellRect, filled: (row: number, col: number) => boolean): CellRect | null {
  let box: CellRect | null = null
  for (let row = r.top; row <= r.bottom; row++)
    for (let col = r.left; col <= r.right; col++) {
      if (!filled(row, col)) continue
      if (!box) box = { top: row, left: col, bottom: row, right: col }
      else box = { top: Math.min(box.top, row), left: Math.min(box.left, col), bottom: Math.max(box.bottom, row), right: Math.max(box.right, col) }
    }
  return box
}

/**
 * Excel's range for Data > Filter: a single cell's current region, otherwise the selection, trimmed to its cells with data,
 * then taken down to the bottom of the current region around that (a header row selected reaches the end of its data).
 * Null when there is no data there, where Excel refuses. `filled` answers within `area` (the sheet's data or the part read).
 */
export function autoFilterRange(selection: CellRect, area: CellRect, filled: (row: number, col: number) => boolean): CellRect | null {
  const single = selection.top === selection.bottom && selection.left === selection.right
  const from = single ? currentRegion({ row: selection.top, col: selection.left }, area, filled) : within(selection, area)
  const box = from && dataBox(from, filled)
  return box && { ...box, bottom: currentRegion(box, area, filled).bottom }
}
