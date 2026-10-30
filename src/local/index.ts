// The `@xlcalc/vue/local` entry point: a sheet source over an xlcalc workbook calculated in the browser, plus xlsx and CSV reading and writing.
// It needs `xlcalc` (a peer the host installs): the engine, `xlcalc/worker` for the pool, `xlcalc/xlsx` and `xlcalc/csv` for files.
export { localSource } from './source'
export { openXlsx, saveXlsx, groupIssues, type IssueGroup, type XlsxIssuesOptions } from './xlsx'
export { openCsv, saveCsv } from './csv'
export type { LocalSource, LocalSourceOptions } from './types'
