// xlcalc's typed-entry module without the engine (`xlcalc/entry`, about 130 KB as built) as a separate chunk, loaded when first
// needed: the text a cell opens with when its source provides none, for the grid and the in-memory workbook.
export type Entry = typeof import('xlcalc/entry')

let loading: Promise<Entry> | null = null

/** The module, imported once; a failed import is retried on the next call. */
export function loadEntry(): Promise<Entry> {
  loading ??= import('xlcalc/entry').catch((e: unknown) => {
    loading = null
    throw e
  })
  return loading
}
