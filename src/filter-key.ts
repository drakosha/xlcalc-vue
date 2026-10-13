// A display filter's stable key, shared by memorySource and ./local: free of the engine, as the main entry needs.
import type { SheetFilter } from './source'

/** A stable key of a filter, for remembering its list of rows. */
export const filterKey = (f: SheetFilter): string =>
  JSON.stringify([f.head, Object.entries(f.cols ?? {}).sort(([a], [b]) => Number(a) - Number(b)), f.sort ?? null])
