// PivotTables in the local source: xlcalc's PivotTable engine (`xlcalc/pivot`), loaded as its own chunk for a workbook that
// has PivotTables, since without them nothing needs it; registered once, it serves every session of the page, a pool's threads too.
import { registerPivot, type BookInput } from 'xlcalc'

/** Whether the workbook holds PivotTables or their caches, which the engine refreshes and lays out. */
export const hasPivots = (book: BookInput): boolean => !!book.pivotCaches?.length || book.sheets.some((s) => !!s.pivotTables?.length)

let loading: Promise<void> | null = null
/** Loads xlcalc's PivotTable engine once for the page; a failed load is tried again by the next call. */
export function loadPivot(): Promise<void> {
  loading ??= import('xlcalc/pivot').then(
    (m) => registerPivot(m.pivotEngine),
    (e: unknown) => {
      loading = null
      throw e
    },
  )
  return loading
}
