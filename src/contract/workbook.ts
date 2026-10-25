/**
 * What a workbook holds beside its cells' values, as `SheetSource` reads and writes it for Excel's dialogs: hyperlinks,
 * notes, defined names, conditional formatting and data validation rules, tables, formula dependents and spelling.
 * Formulas in these types are written as the source displays formulas, in its locale and without `=`.
 */
import type { ConditionalFormat, DataValidation, StructureOp } from 'xlcalc/contract'

/** What `SheetSource.clear` clears: Excel's Clear Formats (styles, number formats and the conditional formats over the cells) or Clear All (also the contents, hyperlinks, notes and data validation rules). */
export type ClearKind = 'formats' | 'all'

/** A hyperlink as Excel's Insert Link writes it: an address outside the workbook (`href`), a place in it (`location`: `'Sheet 2'!B12`, a defined name), or both, and its ScreenTip. */
export interface SheetLink {
  href?: string
  location?: string
  tooltip?: string
}

/** A cell's note (Excel's New Note), with its author's name as Excel shows it above the text. */
export interface SheetNote {
  author?: string
  text: string
}

/** A defined name as Excel's Name Manager lists it: what it refers to (a formula, in the source's locale, without `=`) and the sheet it is visible on, the workbook without `scope`. */
export interface SheetName {
  name: string
  ref: string
  scope?: string
}

/** A cell whose formula refers to the range `dependents` was asked about, for Excel's Trace Dependents arrows. */
export interface SheetDependent {
  sheet: string
  cell: string
}

/** The direct dependents of a range: at most as many as the source returns in one response, `more` when it left some out. */
export interface SheetDependents {
  cells: SheetDependent[]
  more: boolean
}

/** A sheet's rules as Excel's Conditional Formatting Rules Manager and Data Validation edit them; their formulas in the source's locale. */
export interface SheetRules {
  /** In priority order, highest first. */
  conditionalFormats: ConditionalFormat[]
  validations: DataValidation[]
}

/** Excel's table commands: Format as Table and Insert Table (`addTable`), Rename, Delete Table and Convert to Range. */
export type TableOp = Extract<StructureOp, { op: 'addTable' | 'renameTable' | 'deleteTable' | 'convertTable' }>

/**
 * Where Excel's Spelling looks next: within `range` or the entire sheet, from the active cell (`after`, from position `from` of its
 * text, after Ignore Once on a word of it) on, by rows, skipping the words of Ignore All. As Excel's defaults, words in uppercase and
 * words with digits are not checked.
 */
export interface SpellingRequest {
  sheet: string
  range?: string
  after?: string
  from?: number
  ignore?: string[]
}

/** A misspelt word in a cell's text constant (formulas are not checked): its position in the cell's text and the checker's suggestions, best first. */
export interface SpellingHit {
  cell: string
  word: string
  at: number
  suggestions: string[]
}

/** A word checker the host supplies (a dictionary such as hunspell's, and the user's own words): `memorySource` and `@xlcalc/vue/local` check spelling with it. */
export interface SpellChecker {
  /** Whether the word is spelt correctly. */
  check(word: string): boolean
  /** Corrections for a misspelt word, best first. */
  suggest(word: string): string[]
  /** Excel's Add to Dictionary: the word is kept as the user's own. */
  add?(word: string): void
}
