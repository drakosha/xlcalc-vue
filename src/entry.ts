// xlcalc's typed-entry module without the engine (`xlcalc/entry`, about 146 KB as built) as a separate chunk, loaded when first
// needed: the text a cell opens with when its source provides none (the grid, the in-memory workbook), and the Number Format list's samples.
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
