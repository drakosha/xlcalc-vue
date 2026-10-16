// References the grid shows of a formula: to cells, columns or rows of one sheet; a table's and one across sheets stay text.
import type { FormulaReference, GridReference } from 'xlcalc/contract'

/** Whether a reference `scanReferences` found is one the grid shows: not a table's, not over several sheets. */
export const isGridRef = (r: FormulaReference): r is GridReference => r.kind !== 'table' && r.sheetLast === undefined
