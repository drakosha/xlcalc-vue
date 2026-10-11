// Public types of the local source: its options and what it offers beyond `SheetSource`.
import type { AsyncSessionOptions, BookInput, Calc, CalcPoolOptions, ComputeResult, SessionHistory } from 'xlcalc'
import type { EditOutcome, SheetSource } from '../source'

export interface LocalSourceOptions {
  /**
   * Where the book computes: `pool`, Web Workers of `xlcalc/worker` (the default where `Worker` exists); `inline`, the
   * page's own thread in portions (the default elsewhere, e.g. Node); or a calculator of the host's, which the source
   * does not close.
   */
  calc?: 'pool' | 'inline' | Calc
  /** Options of the pool the source starts (`calc: 'pool'`), e.g. `spawn` for a bundler that moves the thread's script. */
  pool?: CalcPoolOptions
  /** Options of the book's session (a live clock, limits, deadlines). */
  session?: AsyncSessionOptions
  /** The cell limit of one window, 20,000 by default. */
  maxCells?: number
  /** A notice every window carries (shown above the grid), e.g. what the xlsx reader left out. */
  notice?: string
}

/** Methods of `SheetSource` that the local source always has. */
type Given = 'edit' | 'paste' | 'fill' | 'undo' | 'redo' | 'structure' | 'layout' | 'format' | 'subscribe' | 'explain' | 'stats' | 'find' | 'visibleRows' | 'columnValues'

/** A sheet source over an xlcalc book: every method of `SheetSource`, and the host's handles on the book. */
export interface LocalSource extends Omit<SheetSource, Given | 'maxCells'>, Required<Pick<SheetSource, Given>> {
  readonly maxCells: number
  /** The book's locale tag: formulas are shown and typed values read in it. */
  readonly locale: string
  /** Settles when the book is computed; rejects when it could not be opened (the methods reject with the same error). */
  readonly ready: Promise<void>
  /** The steps `undo` and `redo` can take now, for a toolbar's buttons. */
  readonly history: SessionHistory
  /** The book with every change so far, to save or keep. */
  book(): Promise<BookInput>
  /** The book's computed values, as `writeXlsx` takes them. */
  result(): Promise<ComputeResult>
  /** Frozen panes of a sheet in the book (null unfreezes them); subscribers get `reload`. */
  freeze(sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome>
  /** Close the session, and the pool when the source started it. */
  close(): Promise<void>
}
