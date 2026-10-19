// The `@xlcalc/vue/local` entry point: a sheet source over an xlcalc workbook calculated in the browser, plus xlsx reading and writing.
// It needs `xlcalc` (a peer the host installs): the engine, `xlcalc/worker` for the pool and `xlcalc/xlsx` for files.
export { localSource } from './source'
export { openXlsx, saveXlsx, groupIssues, type IssueGroup, type XlsxIssuesOptions } from './xlsx'
export type { LocalSource, LocalSourceOptions } from './types'
