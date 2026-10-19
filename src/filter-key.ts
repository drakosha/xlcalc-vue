// A display filter's stable key, shared by memorySource and ./local: free of the engine, as the main entry point requires.
import type { SheetFilter } from './contract/source'

/** A filter's stable key, for caching its row list. */
export const filterKey = (f: SheetFilter): string =>
  JSON.stringify([f.head, Object.entries(f.cols ?? {}).sort(([a], [b]) => Number(a) - Number(b)), f.sort ?? null])
