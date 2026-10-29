// Public types of the local source: its options and what it offers beyond `SheetSource`.
import type { AsyncSessionOptions, BookInput, BookSnapshot, Calc, CalcPoolOptions, ExternalBook, SessionHistory } from 'xlcalc'
import type { EditOutcome, SheetSource } from '../contract/source'
import type { SpellChecker } from '../contract/workbook'

export interface LocalSourceOptions {
  /**
   * Where the workbook is calculated: `pool`, in Web Workers of `xlcalc/worker` (the default where `Worker` exists);
   * `inline`, on the page's own thread in slices (the default elsewhere, e.g. Node); or by the host's calculator, which
   * the source does not close.
   */
  calc?: 'pool' | 'inline' | Calc
  /** Options for the pool the source starts (`calc: 'pool'`), e.g. `spawn` for a bundler that relocates the thread script. */
  pool?: CalcPoolOptions
  /** Options for the workbook's session (a live clock, limits, deadlines). */
  session?: AsyncSessionOptions
  /** The maximum number of cells in one window, 20,000 by default. */
  maxCells?: number
  /** A notice included in every window (displayed above the grid), e.g. what the xlsx reader left out. */
  notice?: string
  /** The host's word checker for Excel's Spelling; without it the source has no `spelling`. */
  spelling?: SpellChecker
  /**
   * A level of `xlcalc/regex` (0.66.0) for match mode 3 of XLOOKUP and XMATCH, loaded as its own chunk and registered
   * before the workbook is calculated: `pcreSyntax` reads Excel's PCRE2 syntax, `pcreEngine` also ends a heavy pattern at
   * Excel's match limit. xlcalc keeps the level for every calculation of the page from then on.
   */
  regex?: 'pcreSyntax' | 'pcreEngine'
}

/** Methods of `SheetSource` that the local source always has. */
type Given = 'calculation' | 'calculate' | 'setCalculation' | 'edit' | 'paste' | 'fill' | 'pick' | 'end' | 'undo' | 'redo' | 'structure' | 'layout' | 'format' | 'replace' | 'outline' | 'freeze' | 'subscribe' | 'validation' | 'functionNames' | 'functionSignature' | 'explain' | 'stats' | 'find' | 'visibleRows' | 'columnValues' | 'clear' | 'link' | 'note' | 'names' | 'setNames' | 'dependents' | 'rules' | 'setRules' | 'table' | 'autoSum' | 'autoSumFunctions'

/** A sheet source over an xlcalc workbook: every method of `SheetSource`, plus the host's handles on the workbook. */
export interface LocalSource extends Omit<SheetSource, Given | 'maxCells'>, Required<Pick<SheetSource, Given>> {
  readonly maxCells: number
  /** The workbook's locale tag: formulas are displayed and typed values parsed in it. */
  readonly locale: string
  /** Resolves when the workbook is calculated; rejects if it could not be opened (the methods reject with the same error). */
  readonly ready: Promise<void>
  /** The number of steps `undo` and `redo` can take now, for a toolbar's buttons. */
  readonly history: SessionHistory
  /** The workbook with all changes so far, to save or store. */
  book(): Promise<BookInput>
  /** The workbook and its calculated values captured at one moment, as `writeXlsx` takes them. */
  snapshot(): Promise<BookSnapshot>
  /**
   * The links to other workbooks with the values a host read from them (xlcalc 0.62.0's `setExternalBooks`): a new link
   * after the others, more values of one, values read again. The workbook is calculated again in one step of undo, and
   * typed text finds a new link by its file name; `saveXlsx` with `over` writes them into the file.
   */
  setExternalBooks(books: ExternalBook[]): Promise<EditOutcome>
  /** Closes the session, and the pool when the source started it. */
  close(): Promise<void>
}
