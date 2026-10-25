/**
 * The sheet source contract: everything the component knows about its data. Rows and columns are 1-based, cells
 * are A1 addresses. The source displays formulas in its own locale and parses what the user typed; the component knows
 * nothing about the engine, the formula locale or HTTP. Optional methods enable features: without `edit` the sheet
 * is read-only, without `find` there is no search, and so on.
 */
import type { CellStyle, CellValue, ConditionalAppearance, FunctionSignature, NumberFormat, StructureOp, StylePatch } from 'xlcalc/contract'
import type { ClearKind, SheetDependents, SheetLink, SheetName, SheetNote, SheetRules, SpellingHit, SpellingRequest, TableOp } from './workbook'

/** The sheet's data source; the component accesses its data only through this interface. */
export interface SheetSource {
  /** The workbook's sheets in order; read again after a structure change and on a `reload` event. */
  sheets(signal?: AbortSignal): Promise<SheetInfo[]>
  /** A rectangular block of a sheet's cells (a window), at most `maxCells` cells. */
  window(req: WindowRequest, signal?: AbortSignal): Promise<SheetWindow | SheetPending>
  /** An edit as typed (`=SUM(A1:A3)`, `1.5`, `'007`); `null` clears. Without it the sheet is read-only. `opts.accept`: see `EditOptions`. */
  edit?(entries: CellEntry[], opts?: EditOptions): Promise<EditOutcome>
  /** Clipboard text pasted at a cell as Excel pastes it, as a single undo step; `select` is the filled range. Without it the fields go to `edit`. */
  paste?(sheet: string, cell: string, text: string): Promise<EditOutcome>
  /** Ctrl+Enter: the text typed into the active cell, entered into every cell of `range`. Without it only the active cell gets it. */
  fill?(sheet: string, range: string, text: string, active: string, opts?: EditOptions): Promise<EditOutcome>
  /** Undoes the workbook's last change as a whole (`ok` with no changes when there is none); `select` is where it was. Without it there is no undo. */
  undo?(): Promise<EditOutcome>
  /** Redoes the last undone change. */
  redo?(): Promise<EditOutcome>
  /** A structure operation; without it there is no structure menu. */
  structure?(op: SheetOp): Promise<EditOutcome>
  /** Column widths and row heights set with the mouse; without it a changed width lasts until the page reloads. */
  layout?(edit: LayoutChange): Promise<EditOutcome>
  /**
   * A range's plain text as Excel puts it on the clipboard: values formatted by their number formats and not clipped to
   * the column width (a date in a narrow column is the date, not `##`), hidden rows and columns omitted, rows ending with
   * CRLF, the last one included unless `trailingNewline` is false (Excel for Mac). A range too large to copy is rejected
   * with an error. Without it the sheet copies the displayed text of its windows.
   */
  copyText?(sheet: string, range: string, opts?: { trailingNewline?: boolean; signal?: AbortSignal }): Promise<string>
  /**
   * Formats ranges (`B2:D9`, entire columns `B:D`, rows `2:5`) as a single undo step, as Excel formats a selection:
   * style fields apply on top of each cell's displayed style (null removes a field), `style: null` removes the cells' own
   * style, and a number format or null sets the format; several patches carry what differs by position (outside borders).
   * Without it there is no toolbar.
   */
  format?(sheet: string, patches: FormatPatch[]): Promise<EditOutcome>
  /** Changes made elsewhere: another tab, an agent, the clock. Returns the unsubscribe function. */
  subscribe?(listener: (event: SourceEvent) => void): () => void
  /** A cell's data validation as Excel displays it on the cell: its dropdown list and its input message; null when no rule applies to the cell. */
  validation?(sheet: string, cell: string, signal?: AbortSignal): Promise<SheetValidation | null>
  /**
   * An item chosen from a list rule's dropdown: item number `item` of `validation`'s `list` (with its `text`, to detect a
   * changed list), written as Excel writes a chosen item (its value, not parsed as typed text; the cell's number format
   * kept), without checking the rule, as a single undo step. Without it the item goes to `edit` as typed, with `accept`.
   */
  pick?(sheet: string, cell: string, item: number, text: string): Promise<EditOutcome>
  /** Excel's Freeze Panes: the rows above and the columns to the left of a cell kept in view (`{ rows: 2, cols: 1 }`), null unfreezes; a single undo step. Without it only the sheet's lock button hides the workbook's frozen panes. */
  freeze?(sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome>
  /** The functions a formula can call, by their names as typed in the source's locale (`СУММ` in ru-RU), for Excel's Formula AutoComplete. */
  functionNames?(signal?: AbortSignal): Promise<string[]>
  /**
   * A function's arguments for Excel's argument tooltip while a formula is typed (`SUM(number1, [number2], ...)`, in ru-RU
   * `СУММ(число1;[число2];...)`): the name as typed, in English or in the source's locale, in any letter case; null for a
   * function the source does not know. Without it there is no tooltip.
   */
  functionSignature?(name: string, signal?: AbortSignal): Promise<SheetFunctionSignature | null>
  /** A cell's precedent tree as text. */
  explain?(sheet: string, cell: string): Promise<string>
  /**
   * Excel's Ctrl+arrow from a cell (`Range.End`): the cell it goes to (`B12`), which is the end of the block of non-empty
   * cells, the next non-empty cell, or the edge of the sheet, skipping hidden rows or columns. Without it the sheet goes to
   * the edge of the data known to it.
   */
  end?(sheet: string, cell: string, direction: 'up' | 'down' | 'left' | 'right', signal?: AbortSignal): Promise<string>
  /** Totals of a range; without it the component counts over the loaded windows. */
  stats?(sheet: string, range: string, filter?: SheetFilter, signal?: AbortSignal): Promise<RangeStats>
  /** Excel's Find on a sheet: matches in Excel's order, by rows or by columns (under a display filter, only its visible rows). Without it there is no search. */
  find?(req: FindRequest, signal?: AbortSignal): Promise<FindResult>
  /** Excel's Replace All: every cell `find` finds for `req`, searched in formulas whatever `lookIn` says, as Excel replaces, with the matched text replaced by `by`, as a single undo step. */
  replace?(req: FindRequest, by: string): Promise<ReplaceOutcome>
  /** Excel's outline commands: Group, Ungroup, a group's button and a level button. Without it the outline is drawn but cannot be changed. */
  outline?(change: OutlineChange): Promise<EditOutcome>
  /** The sheet's rows in display order under the filter and sort; `rows: null` when there are no conditions. */
  visibleRows?(sheet: string, filter: SheetFilter, signal?: AbortSignal): Promise<{ rows: number[] | null }>
  /** A column's values for the filter funnel, over the rows visible under the other columns' conditions. */
  columnValues?(sheet: string, col: number, filter: SheetFilter, signal?: AbortSignal): Promise<{ values: { text: string; count: number }[]; more: number }>
  /** Excel's Clear Formats and Clear All over ranges, as a single undo step; Delete clears contents through `edit`. Without it neither is offered. */
  clear?(sheet: string, ranges: string[], what: ClearKind): Promise<EditOutcome>
  /** Excel's Insert Link over a range (null: Remove Hyperlinks), the cells' contents kept; a single undo step. Without it links can be followed but not edited. */
  link?(sheet: string, range: string, link: SheetLink | null): Promise<EditOutcome>
  /** Excel's New Note and Edit Note on a cell (null: Delete Note); a single undo step. Without it notes are shown but not edited. */
  note?(sheet: string, cell: string, note: SheetNote | null): Promise<EditOutcome>
  /** The workbook's defined names, for the Name Box, Go To, Formula AutoComplete and the Name Manager. */
  names?(signal?: AbortSignal): Promise<SheetName[]>
  /** Replaces the workbook's defined names (Excel's Name Manager and Define Name), as a single undo step; a name Excel rejects is refused with its message. */
  setNames?(names: SheetName[]): Promise<EditOutcome>
  /** The cells whose formulas refer directly to a range, for Excel's Trace Dependents. Without it there are no arrows. */
  dependents?(sheet: string, range: string, signal?: AbortSignal): Promise<SheetDependents>
  /** A sheet's conditional formatting and data validation rules, for their dialogs. */
  rules?(sheet: string, signal?: AbortSignal): Promise<SheetRules>
  /** Replaces a sheet's conditional formatting rules, its data validation rules, or both, as a single undo step. Without it the rules are applied but not edited. */
  setRules?(sheet: string, rules: Partial<SheetRules>): Promise<EditOutcome>
  /** Excel's table commands. Without it there is no Format as Table. */
  table?(op: TableOp): Promise<EditOutcome>
  /** Excel's Spelling: the next misspelt word from `req.after` on, wrapping around once, or null when there is none. Without it there is no Spelling. */
  spelling?(req: SpellingRequest, signal?: AbortSignal): Promise<SpellingHit | null>
  /** Excel's Add to Dictionary in the Spelling dialog: the word is the user's own from now on. */
  learnWord?(word: string): Promise<void>
  /** The maximum number of cells in one window, 20,000 by default. */
  readonly maxCells?: number
  /** The number of steps `undo` and `redo` can take, as of the last completed call: a toolbar enables its buttons by them. */
  readonly history?: { readonly undo: number; readonly redo: number }
}

/** A function's signature as returned by a source, with what Excel's Insert Function displays under the syntax: a one-line description of the function, in the source's locale. */
export type SheetFunctionSignature = FunctionSignature & { description?: string }

/** A range's formatting in `SheetSource.format`. */
export interface FormatPatch {
  range: string
  style?: StylePatch | null
  numberFormat?: NumberFormat | null
}

/** A display filter and sort: the data do not change. */
export interface SheetFilter {
  /** Header rows at the top, which are never filtered. */
  head: number
  cols?: Record<number, { values?: string[]; text?: string }>
  sort?: { col: number; dir: 'asc' | 'desc' }
}

export interface RangeStats {
  cells: number
  numbers: number
  sum: number
  min: number | null
  max: number | null
}

export interface SheetInfo {
  name: string
  maxRow: number
  maxCol: number
  freeze?: { rows: number; cols: number }
  /** Not shown on the tabs; `very` (xlsx's veryHidden) is also not offered for Unhide, since Excel never offers it. */
  hidden?: boolean | 'very'
  /** Ranges that cannot be edited, with the reason for a tooltip (for example, ranges managed by a script). */
  locked?: { range: string; reason: string }[]
  /** List mode: column titles in the header instead of letters; `sortable` makes a header click sort. */
  columns?: { title: string; align?: 'left' | 'center' | 'right'; sortable?: boolean }[]
  /** Row numbers on the left, shown by default; a list has none. */
  rowNumbers?: boolean
  /** A totals row pinned to the bottom: the displayed text for each of the sheet's columns (null: empty). */
  totals?: (string | null)[]
  /** Excel's outline (Data > Group): the sheet's groups of rows and columns; a sheet without groups has none. */
  outline?: SheetOutline
  /** The workbook's default font (its Normal style), in which a cell without a font of its own is drawn; the component's `--xs-font` without it. */
  font?: { name: string; size: number }
  /** The sheet's saved zoom in percent (xlsx `zoomScale`), with which the sheet opens; 100 without it. */
  zoom?: number
}

/** A sheet's groups as Excel draws them in the bars next to the row numbers and above the column letters. */
export interface SheetOutline {
  /** Groups of rows by their first row, the outer group before the inner one when both start on the same row. */
  rows: OutlineGroup[]
  columns: OutlineGroup[]
  /** A group's summary row, which holds its button, is below it (Excel's default) or above it. */
  summaryBelow: boolean
  /** A group's summary column is to the right of it (Excel's default) or to the left of it. */
  summaryRight: boolean
}

/**
 * A group: the rows or columns `from`..`to` (columns by number) at outline level `level` or deeper, with no such row
 * or column adjacent to them; `level` is 1 to 7, so every row or column inside a level-2 group is also inside a level-1
 * group. `collapsed`: its rows or columns are hidden and its button shows +.
 */
export interface OutlineGroup {
  from: number
  to: number
  level: number
  collapsed: boolean
}

/**
 * Excel's outline commands on a sheet, each a single undo step: Group and Ungroup move entire rows or columns one level
 * deeper or shallower (`rows: '3:6'`, `columns: 'B:D'`); collapse and expand are a group's button (Excel's Hide Detail
 * and Show Detail), with the group identified by its rows or columns and level as `SheetOutline` returns it; `showLevel`
 * is a level button: groups at `level` and deeper are collapsed, shallower ones expanded, and one past the deepest level
 * expands all.
 */
export type OutlineChange =
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }
  | { op: 'collapse' | 'expand'; sheet: string; rows: string; level: number }
  | { op: 'collapse' | 'expand'; sheet: string; columns: string; level: number }
  | { op: 'showLevel'; sheet: string; axis: 'rows' | 'columns'; level: number }

/** What Excel's Find searches for and where, as set by the options of its dialog. */
export interface FindRequest {
  sheet: string
  /** The text to find; `*` matches any characters, `?` one character, and `~` before either of them or before itself matches that character, as in Excel. */
  query: string
  /** Only within this range, as Excel searches within a selection of more than one cell; the entire sheet without it. */
  range?: string
  /** Match case; without it the case is ignored. */
  matchCase?: boolean
  /** Match entire cell contents: the entire text matches the query, not a part of it. */
  wholeCell?: boolean
  /** Look in: the text a cell opens with for editing (formulas as typed in the workbook's locale, Excel's default) or its displayed text. */
  lookIn?: 'formulas' | 'values'
  /** Search by columns: matches column by column; row by row without it. */
  byColumns?: boolean
  /**
   * Find Next from this cell (`B12`), as Excel's Find continues from the active cell (`Range.Find`'s After): the matches in
   * Find's order starting from the first one after it, wrapping around at the end, with the cell itself last. Without it they
   * start from the range's first cell. A source that does not support it returns them from the first cell, and Find Next
   * moves through the matches it returned.
   */
  after?: string
  /** The display filter in effect: only its visible rows are searched. */
  filter?: SheetFilter
}

/** The matches of a search, at most as many as the source returns in one response; `total` counts them all. */
export interface FindResult {
  hits: { row: number; col: number }[]
  total: number
}

/**
 * The outcome of Replace All and how many cells it changed (Excel's "We made N replacements"); `refused`: the cells whose
 * new text Excel would not accept (a formula that no longer parses), left unchanged while the others are written, with the
 * reason.
 */
export type ReplaceOutcome = EditOutcome & { replaced: number; refused?: { sheet: string; cell: string; message: string }[] }

/** With a filter, `row` is a display index (a position among the visible rows). */
export interface WindowRequest {
  sheet: string
  row: number
  rows: number
  col: number
  cols: number
  filter?: SheetFilter
}

export interface SheetWindow {
  row: number
  col: number
  /** Displayed text, the window's rows × columns; an empty cell is ''. */
  text: string[][]
  /** Values of the same cells, used for selection totals and alignment. */
  values: CellValue[][]
  /** Formulas by address as they are displayed and edited (in the source's locale), without `=`. */
  formulas: Record<string, string>
  /** The text each cell opens with for editing, as Excel's formula bar displays it; a cell missing from it opens empty. Without it the component produces the text from the value and the formula. */
  editText?: Record<string, string>
  /** The color from the number format's section (`#RRGGBB`). */
  colors: Record<string, string>
  styleIds: Record<string, number>
  styles: CellStyle[]
  /** The number format code each cell is displayed with (its own, otherwise its row's or column's); a cell missing from it is General. For a toolbar. */
  numberFormats?: Record<string, string>
  /** The window's column widths (in the units of xlcalc's `ColumnInput.width`, null for the default), row heights in points. */
  widths: (number | null)[]
  rowHeights: Record<number, number>
  merges: string[]
  /** The default row height in points. */
  defaultRowHeight?: number
  /** The sheet's tables, from which their header and border are drawn. */
  tables?: { name: string; range: string; headerRow: number }[]
  /** The cells' conditional formatting as xlcalc calculates it: a style on top of the cell's own, a data bar, an icon, a hidden value. */
  appearance?: Record<string, ConditionalAppearance>
  /** Hyperlinks: an external address, a location in the workbook (`Sheet!B12`), a tooltip, a host-specific target (such as a file of the host). */
  links?: Record<string, { href?: string; location?: string; tooltip?: string; target?: string }>
  notes?: Record<string, { author?: string; text: string }>
  /** A notice above the grid (for example, "values calculated on the server"). */
  notice?: string
  /** Drill-down cells (list mode): underlined, and a click or Enter emits `cellDrill`. */
  drill?: Record<string, true>
}

/** The window is not ready yet (the host's workbook is recalculating): the component waits and requests it again. */
export interface SheetPending {
  pending: { elapsedMs: number }
}

export interface CellEntry {
  sheet: string
  cell: string
  text: string | null
}

/** Options of an edit. */
export interface EditOptions {
  /** The user kept a value the cell's rule rejects, with Yes on a warning or OK on an information alert: write it without checking the rule again. */
  accept?: true
  /** The text was edited starting from the cell's own text (F2, a double click, the formula bar), not typed over the cell: Excel's automatic percentage entry then adds no `%` (`6%` in a `0%` cell is 0.06). */
  inPlace?: true
}

/** A cell's data validation as the user sees it: Excel's dropdown of a list rule and the message displayed while the cell is selected. */
export interface SheetValidation {
  /** A list rule that shows its dropdown: the items as text, in order; empty when its formula returns none. */
  list?: string[]
  /** The input message (in Excel, a title of up to 32 characters and text of up to 255). */
  prompt?: { title?: string; text: string }
}

/** A data validation rule's alert for a value that violates it, as Excel displays it: stop (Retry, Cancel), warning (Continue? Yes, No, Cancel), information (OK, Cancel). */
export interface SheetValidationAlert {
  severity: 'stop' | 'warning' | 'information'
  /** The rule's title, otherwise the component's default ("Microsoft Excel"). */
  title?: string
  /** The rule's message, otherwise the source's description of the rule. */
  message: string
  /** The cell where the value was rejected. */
  sheet: string
  cell: string
}

export type SheetOp = Exclude<StructureOp, { op: 'addTable' | 'renameTable' | 'deleteTable' | 'convertTable' }>

/** Columns or rows changed: a width or height set with the mouse, rows or columns hidden or unhidden, columns or rows fitted to their contents (Excel's double click on an edge). */
export type LayoutChange =
  | { sheet: string; columns: string; width?: number; hidden?: boolean; fit?: true }
  | { sheet: string; rows: string; height?: number; hidden?: boolean; fit?: true }

export type EditOutcome =
  | {
      ok: true
      changes: CellUpdate[]
      /** Where the selection goes after `paste`, `fill`, `undo` or `redo`: the range they filled or restored; `edit` returns none (the active cell has already moved). */
      select?: { sheet: string; range: string }
    }
  | {
      ok: false
      /**
       * conflict: the workbook changed elsewhere; refused: the cell is locked or the operation is not allowed;
       * invalid: the formula does not parse (`at` is the position in the typed text); failed: anything else.
       */
      reason: 'conflict' | 'refused' | 'invalid' | 'failed'
      message: string
      /**
       * The source's code for the reason, for a host that reacts to it rather than to the text. `memorySource` and
       * `@xlcalc/vue/local` give xlcalc's codes: an outline command's `outline_deep`, `not_grouped` or `detail`, and in
       * `./local` the code of any structure operation xlcalc refuses.
       */
      code?: string
      at?: number
      /**
       * For `invalid`: the formula Excel suggests instead in its "We found a typo in your formula" dialog, in the source's
       * locale (unclosed parentheses closed, `=(1+2` → `=(1+2)`); the component offers to accept it and sends the edit again with it.
       */
      correction?: string
      /** A value that violates a data validation rule (`refused`): the component displays Excel's alert, and sends the edit again with `accept` after Yes on a warning or OK on an information alert. */
      alert?: SheetValidationAlert
    }

/** A cell's new state after an edit or a recalculation. */
export interface CellUpdate {
  sheet: string
  cell: string
  text: string
  value: CellValue
  color?: string
  /** Only for edited cells: the formula as displayed, or null when the cell no longer has one. */
  formula?: string | null
  /** Only for edited cells, from a source whose windows have `editText`: the text the cell now opens with for editing. */
  editText?: string
  /** Only for edited cells, from a source whose windows have `numberFormats`: the format code the cell is now displayed with (null: General), as when a typed date sets one. */
  numberFormat?: string | null
}

export type SourceEvent =
  | { kind: 'changed'; changes: CellUpdate[] }
  /** Many or unknown things changed: read the sheets and windows again. */
  | { kind: 'reload'; message?: string }
  /** Cells in these ranges may look different (conditional formats): windows that overlap them are re-read in place, keeping the sheets, sizes and display filter as they are. */
  | { kind: 'repaint'; ranges: { sheet: string; range: string }[] }
