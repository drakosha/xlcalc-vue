// The `xlcalc-sheet/local` entry: a sheet source over an xlcalc book computed in the browser, and xlsx in and out.
// It needs `xlcalc` (a peer the host installs): the engine, `xlcalc/worker` for the pool and `xlcalc/xlsx` for files.
export { localSource } from './source'
export { openXlsx, saveXlsx, groupIssues, type IssueGroup, type XlsxIssuesOptions } from './xlsx'
export type { LocalSource, LocalSourceOptions } from './types'
