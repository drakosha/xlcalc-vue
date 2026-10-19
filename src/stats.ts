// Cell totals as Excel's status bar shows them, by the counting rule ./local's session follows, and what a blank cell is.
import type { CellValue } from 'xlcalc/contract'
import type { RangeStats } from './contract/source'

/** The totals of no cells. */
export const noStats = (): RangeStats => ({ cells: 0, numbers: 0, sum: 0, min: null, max: null })

/** A cell that displays nothing and contains no value (formulas aside, which the caller handles where they count). */
export const isBlank = (text: string | undefined, value: CellValue | undefined): boolean => (text ?? '') === '' && (value === null || value === undefined || value === '')

/** Adds a cell to the totals: a blank one does not count; a number goes into the sum, min and max. */
export function addCell(st: RangeStats, text: string, value: CellValue | undefined): void {
  if (isBlank(text, value)) return
  st.cells++
  if (typeof value !== 'number') return
  st.numbers++
  st.sum += value
  st.min = st.min === null ? value : Math.min(st.min, value)
  st.max = st.max === null ? value : Math.max(st.max, value)
}
