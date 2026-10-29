// Excel's Workbook Statistics (Review > Workbook Statistics): the displayed sheet's end and counts and the workbook's, read
// from the source's `workbookStatistics` and shown with the rows Excel for Mac's window shows.
import { computed, ref } from 'vue'
import type { StatisticsCounts, WorkbookStatistics } from 'xlcalc/contract'
import { isAbort } from '../abort'
import type { SheetMessages } from '../contract/messages'
import { useSheetReads, type SheetState } from './state'

/** A row of the window: its label's text key and the value it shows. */
export type StatisticsRow = [label: keyof SheetMessages, value: string]

/** The counts in the order Excel lists them, with their labels. */
const COUNTS: [keyof StatisticsCounts, keyof SheetMessages][] = [
  ['cells', 'statisticsCells'],
  ['tables', 'statisticsTables'],
  ['pivotTables', 'statisticsPivotTables'],
  ['formulas', 'statisticsFormulas'],
  ['charts', 'statisticsCharts'],
  ['images', 'statisticsImages'],
  ['formControls', 'statisticsFormControls'],
  ['objects', 'statisticsObjects'],
  ['comments', 'statisticsComments'],
  ['notes', 'statisticsNotes'],
]
/** The rows Excel for Mac shows whatever their count. */
const ALWAYS = new Set<keyof StatisticsCounts>(['cells', 'tables', 'formulas'])

/**
 * The window's rows as Excel for Mac shows them (the library's 24 cases): the sheet's end where it has one, the cells,
 * tables and formulas always, any other count only above 0, a count the workbook does not hold never, and the workbook's
 * notes never.
 */
export function statisticsRows(st: WorkbookStatistics): { sheet: StatisticsRow[]; workbook: StatisticsRow[] } {
  const rows = (c: StatisticsCounts, skip: Set<keyof StatisticsCounts>): StatisticsRow[] =>
    COUNTS.flatMap(([k, label]): StatisticsRow[] => {
      const n = c[k]
      return n !== null && !skip.has(k) && (ALWAYS.has(k) || n > 0) ? [[label, String(n)]] : []
    })
  return {
    sheet: [...(st.sheet.end ? [['statisticsEnd', st.sheet.end] as StatisticsRow] : []), ...rows(st.sheet, new Set())],
    workbook: [['statisticsSheets', String(st.workbook.sheets)], ...rows(st.workbook, new Set(['notes']))],
  }
}

export function useStatistics(s: SheetState, notify: (text: string) => void) {
  const read = useSheetReads(s)
  /** The statistics shown, of the sheet displayed when they were asked for; null while the window is closed. */
  const shown = ref<WorkbookStatistics | null>(null)
  const available = computed(() => !!s.sheet.value && !!s.sourceMethod('workbookStatistics'))

  /** Asks the source and opens the window with its answer; a failure goes to the message line. */
  function open(): boolean {
    const get = s.sourceMethod('workbookStatistics')
    if (!get || !s.sheet.value) return false
    const outdated = s.outdated()
    get(s.sheetName.value, read()).then(
      (st) => {
        if (!outdated()) shown.value = st
      },
      (e: unknown) => {
        if (!outdated() && !isAbort(e)) notify(e instanceof Error ? e.message : String(e))
      },
    )
    return true
  }
  function close(): void {
    shown.value = null
    s.focusGrid()
  }
  return { available, shown, open, close }
}

export type Statistics = ReturnType<typeof useStatistics>
