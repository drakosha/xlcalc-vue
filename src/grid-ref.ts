// References in a formula that the grid displays: to cells, columns or rows on a single sheet; table and 3D references stay text.
import type { FormulaReference, GridReference } from 'xlcalc/contract'

/** Whether a reference found by `scanReferences` is one the grid displays: not a table reference, not across several sheets. */
export const isGridRef = (r: FormulaReference): r is GridReference => r.kind !== 'table' && r.sheetLast === undefined
