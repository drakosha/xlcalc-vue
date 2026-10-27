// The `@xlcalc/vue/remote` entry point: a sheet source over a workbook a server serves with `@xlcalc/vue/server`.
// It needs neither the engine nor `xlcalc`: the workbook is calculated on the server.
export { remoteSource } from './client'
export type { RemoteSource, RemoteSourceOptions, SourceAccess } from './types'
