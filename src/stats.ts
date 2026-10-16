// Totals of cells as Excel's status bar gives them, with one rule of what counts, for the sources and the grid.
import type { CellValue } from 'xlcalc/contract'
import type { RangeStats } from './source'

/** Totals of no cells. */
export const noStats = (): RangeStats => ({ cells: 0, numbers: 0, sum: 0, min: null, max: null })

/** A cell into the totals: one that shows nothing and holds no value does not count; a number goes into the sum and the ends. */
export function addCell(st: RangeStats, text: string, value: CellValue | undefined): void {
  if (text === '' && (value === null || value === undefined || value === '')) return
  st.cells++
  if (typeof value !== 'number') return
  st.numbers++
  st.sum += value
  st.min = st.min === null ? value : Math.min(st.min, value)
  st.max = st.max === null ? value : Math.max(st.max, value)
}
