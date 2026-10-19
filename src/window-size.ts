// A window's size: the sources' cell limit when none is given, and the rows a window of so many columns may contain within it.

/** The maximum number of cells in one window when the source specifies none (`SheetSource.maxCells`). */
export const DEFAULT_MAX_CELLS = 20_000

/** The rows of a window of `cols` columns: as many as requested, at most what `maxCells` cells allow. */
export const rowsFor = (maxCells: number, rows: number, cols: number): number => Math.max(0, Math.min(rows, Math.floor(maxCells / Math.max(1, cols))))
