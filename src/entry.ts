// xlcalc's typed entry without the engine (`xlcalc/entry`, some 130 KB as built) as a chunk of its own, loaded on first need: the
// text a cell opens with when its source gives none, the grid's and the book in memory's.
export type Entry = typeof import('xlcalc/entry')

let loading: Promise<Entry> | null = null

/** The module, imported once; a failed import is tried again on the next call. */
export function loadEntry(): Promise<Entry> {
  loading ??= import('xlcalc/entry').catch((e: unknown) => {
    loading = null
    throw e
  })
  return loading
}
