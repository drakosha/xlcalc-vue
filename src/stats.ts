// Totals of cells as Excel's status bar gives them, by the rule of what counts that ./local's session keeps, and what a blank cell is.
import type { CellValue } from 'xlcalc/contract'
import type { RangeStats } from './contract/source'

/** Totals of no cells. */
export const noStats = (): RangeStats => ({ cells: 0, numbers: 0, sum: 0, min: null, max: null })

/** A cell that shows nothing and holds no value (a formula aside, which the caller weighs where it counts). */
export const isBlank = (text: string | undefined, value: CellValue | undefined): boolean => (text ?? '') === '' && (value === null || value === undefined || value === '')

/** A cell into the totals: a blank one does not count; a number goes into the sum and the ends. */
export function addCell(st: RangeStats, text: string, value: CellValue | undefined): void {
  if (isBlank(text, value)) return
  st.cells++
  if (typeof value !== 'number') return
  st.numbers++
  st.sum += value
  st.min = st.min === null ? value : Math.min(st.min, value)
  st.max = st.max === null ? value : Math.max(st.max, value)
}
