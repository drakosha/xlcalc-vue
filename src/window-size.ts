// A window's size: the sources' cell limit when none is given, and the rows a window of so many columns may take within it.

/** The cell limit of one window when the source names none (`SheetSource.maxCells`). */
export const DEFAULT_MAX_CELLS = 20_000

/** The rows of a window of `cols` columns: as many as asked, at most what `maxCells` cells hold. */
export const rowsFor = (maxCells: number, rows: number, cols: number): number => Math.max(0, Math.min(rows, Math.floor(maxCells / Math.max(1, cols))))
