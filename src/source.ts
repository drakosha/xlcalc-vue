/**
 * The sheet source contract: everything the component knows about its data. Rows and columns are 1-based, cells
 * are A1 addresses. The source shows formulas in its own locale and parses what a person typed; the component knows
 * nothing about the engine, the formula locale or HTTP. Optional methods turn features on: without `edit` the sheet
 * is read-only, without `find` there is no search, and so on.
 */
import type { CellStyle, CellValue, ConditionalAppearance, NumberFormat, StructureOp, StylePatch } from 'xlcalc/contract'

/** Where a sheet takes its data from; the component knows only this. */
export interface SheetSource {
  /** The book's sheets in order; read again after a structure edit and on a `reload` event. */
  sheets(signal?: AbortSignal): Promise<SheetInfo[]>
  /** A rectangle of a sheet, at most `maxCells` cells. */
  window(req: WindowRequest, signal?: AbortSignal): Promise<SheetWindow | SheetPending>
  /** An edit as typed (`=SUM(A1:A3)`, `1.5`, `'007`); `null` clears. Without it the sheet is read-only. `opts.accept`: see `EditOptions`. */
  edit?(entries: CellEntry[], opts?: EditOptions): Promise<EditOutcome>
  /** Clipboard text pasted at a cell as Excel pastes it, one step of undo; `select` is what it filled. Without it the fields go to `edit`. */
  paste?(sheet: string, cell: string, text: string): Promise<EditOutcome>
  /** Control+Return: the text typed into the active cell entered into every cell of `range`. Without it only the active cell takes it. */
  fill?(sheet: string, range: string, text: string, active: string, opts?: EditOptions): Promise<EditOutcome>
  /** Take back the book's last change, whole (`ok` with no changes when there is none); `select` is where it was. Without it there is no undo. */
  undo?(): Promise<EditOutcome>
  /** Do the last undone change again. */
  redo?(): Promise<EditOutcome>
  /** A structure operation; without it there is no structure menu. */
  structure?(op: SheetOp): Promise<EditOutcome>
  /** Column widths and row heights set with the mouse; without it a width lives until the page reloads. */
  layout?(edit: LayoutChange): Promise<EditOutcome>
  /**
   * Formatting of ranges (`B2:D9`, whole columns `B:D`, rows `2:5`) in one step of undo, as Excel formats a selection:
   * style fields go over the style each cell shows (null removes a field), `style: null` removes the cells' own style, a
   * number format or null; several patches carry what differs by place (outside borders). Without it there is no toolbar.
   */
  format?(sheet: string, patches: FormatPatch[]): Promise<EditOutcome>
  /** Changes made elsewhere: another tab, an agent, the clock. Returns the unsubscribe function. */
  subscribe?(listener: (event: SourceEvent) => void): () => void
  /** A cell's data validation as Excel shows it on the cell: the list of its dropdown and its input message; null when no rule is over the cell. */
  validation?(sheet: string, cell: string, signal?: AbortSignal): Promise<SheetValidation | null>
  /**
   * An item picked from a list rule's dropdown: the `item`th of `validation`'s `list` (`text`, to tell when the list changed),
   * written as Excel writes a picked item (its value, not read as typed text; the cell's format kept), the rule not asked,
   * one step of undo. Without it the item goes to `edit` as typed, with `accept`.
   */
  pick?(sheet: string, cell: string, item: number, text: string): Promise<EditOutcome>
  /** Excel's Freeze Panes: the rows above and the columns left of a cell kept in view (`{ rows: 2, cols: 1 }`), null unfreezes; one step of undo. Without it the sheet's lock alone hides a book's freeze. */
  freeze?(sheet: string, panes: { rows: number; cols: number } | null): Promise<EditOutcome>
  /** The functions a formula can call, by their names as typed in the source's locale (`СУММ` in ru-RU), for Excel's Formula AutoComplete. */
  functionNames?(signal?: AbortSignal): Promise<string[]>
  /** A cell's precedents tree as text. */
  explain?(sheet: string, cell: string): Promise<string>
  /** Totals of a rectangle; without it the component counts over the loaded windows. */
  stats?(sheet: string, range: string, filter?: SheetFilter, signal?: AbortSignal): Promise<RangeStats>
  /** Excel's Find on a sheet: hits in Excel's order, by rows or by columns (under a display filter only its visible rows). Without it there is no search. */
  find?(req: FindRequest, signal?: AbortSignal): Promise<FindResult>
  /** Excel's Replace All: every cell `find` finds for `req`, looked for in formulas whatever `lookIn` says, as Excel replaces, with the matched text replaced by `by`, in one step of undo. */
  replace?(req: FindRequest, by: string): Promise<ReplaceOutcome>
  /** Excel's outline commands: Group, Ungroup, a group's button and a level button. Without it the outline is drawn but does not change. */
  outline?(change: OutlineChange): Promise<EditOutcome>
  /** The sheet's rows in display order under the filter and sort; `rows: null` when there are no conditions. */
  visibleRows?(sheet: string, filter: SheetFilter, signal?: AbortSignal): Promise<{ rows: number[] | null; total: number }>
  /** A column's values for the filter funnel, over the rows visible under the other columns' conditions. */
  columnValues?(sheet: string, col: number, filter: SheetFilter, signal?: AbortSignal): Promise<{ values: { text: string; count: number }[]; more: number }>
  /** The cell limit of one window, 20,000 by default. */
  readonly maxCells?: number
  /** The steps `undo` and `redo` can take, as of the last finished call: a toolbar turns its buttons on by them. */
  readonly history?: { readonly undo: number; readonly redo: number }
}

/** A range's formatting in `SheetSource.format`. */
export interface FormatPatch {
  range: string
  style?: StylePatch | null
  numberFormat?: NumberFormat | null
}

/** A display filter and sort, as the the app has them: the data do not change. */
export interface SheetFilter {
  /** Header rows on top that are never filtered. */
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
  /** Not on the tabs; `very` (xlsx's veryHidden) is not offered to show again either, as Excel offers it nowhere. */
  hidden?: boolean | 'very'
  /** Rectangles that cannot be edited, with the reason for a tooltip (the app: the script's areas). */
  locked?: { range: string; reason: string }[]
  /** List mode: column titles in the header instead of letters; `sortable` lets a header click sort. */
  columns?: { title: string; align?: 'left' | 'center' | 'right'; sortable?: boolean }[]
  /** Row numbers on the left, on by default; a list has none. */
  rowNumbers?: boolean
  /** A totals row held at the bottom: the shown text by the sheet's columns (null: empty). */
  totals?: (string | null)[]
  /** Excel's outline (Data > Group): the sheet's groups of rows and columns; a sheet without groups has none. */
  outline?: SheetOutline
}

/** A sheet's groups as Excel draws them in the bars beside the row numbers and above the column letters. */
export interface SheetOutline {
  /** Groups of rows by their first row, the outer before the inner where two begin together. */
  rows: OutlineGroup[]
  columns: OutlineGroup[]
  /** A group's summary row, which holds its button, is below it (Excel's default) or above it. */
  summaryBelow: boolean
  /** A group's summary column is right of it (Excel's default) or left of it. */
  summaryRight: boolean
}

/**
 * A group: the lines `from`..`to` (rows, or columns by number) of outline level `level` or deeper, with no such line
 * beside them; `level` is 1 to 7, so every line inside a group of level 2 is inside one of level 1 too. `collapsed`: its
 * lines are hidden and its button shows +.
 */
export interface OutlineGroup {
  from: number
  to: number
  level: number
  collapsed: boolean
}

/**
 * Excel's outline commands on a sheet, each one step of undo: Group and Ungroup put whole lines one level deeper or
 * shallower (`rows: '3:6'`, `columns: 'B:D'`); collapse and expand are a group's button (Excel's Hide Detail and Show
 * Detail), the group named by its lines and level as `SheetOutline` gives it; `showLevel` is a level button: groups of
 * `level` and deeper collapsed, the shallower ones expanded, one past the deepest level expanding all.
 */
export type OutlineChange =
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }
  | { op: 'collapse' | 'expand'; sheet: string; rows: string; level: number }
  | { op: 'collapse' | 'expand'; sheet: string; columns: string; level: number }
  | { op: 'showLevel'; sheet: string; axis: 'rows' | 'columns'; level: number }

/** What Excel's Find looks for and where, as its dialog's options give it. */
export interface FindRequest {
  sheet: string
  /** The text to find; `*` stands for any characters, `?` for one, and `~` before either of them or itself for that character, as in Excel. */
  query: string
  /** Only within this range, as Excel finds within a selection of more than one cell; the whole sheet without it. */
  range?: string
  /** Match case; without it the case is ignored. */
  matchCase?: boolean
  /** Match entire cell contents: the whole text matches the query, not a part of it. */
  wholeCell?: boolean
  /** Look in: the text a cell opens with (formulas as typed in the book's locale, Excel's default) or the text it shows. */
  lookIn?: 'formulas' | 'values'
  /** Search by columns: hits column by column; row by row without it. */
  byColumns?: boolean
  /** The display filter in force: only its visible rows. */
  filter?: SheetFilter
}

/** Hits of a search, at most as many as the source gives in one answer; `total` counts them all. `text` is what matched: the formula or the shown text. */
export interface FindResult {
  hits: { row: number; col: number; text: string }[]
  total: number
}

/** The outcome of Replace All and how many cells it changed (Excel's "We made N replacements"). */
export type ReplaceOutcome = EditOutcome & { replaced: number }

/** With a filter `row` is a display index (a place among the visible rows), as in the the app. */
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
  /** Shown text, the window's rows × columns; an empty cell is ''. */
  text: string[][]
  /** Values of the same cells: selection totals and alignment use them. */
  values: CellValue[][]
  /** Formulas by address as they are shown and edited (the source's locale), without `=`. */
  formulas: Record<string, string>
  /** The text each cell opens with for editing, as Excel's formula bar shows it; a cell left out opens empty. Without it the component writes it from the value and the formula. */
  editText?: Record<string, string>
  /** The colour of the format's section (`#RRGGBB`). */
  colors: Record<string, string>
  styleIds: Record<string, number>
  styles: CellStyle[]
  /** The number format each cell shows by, as a code (its own, else its row's or column's); a cell left out is General. For a toolbar. */
  numberFormats?: Record<string, string>
  /** The window's column widths (units of xlcalc's `ColumnInput.width`, null for the default), row heights in points. */
  widths: (number | null)[]
  rowHeights: Record<number, number>
  merges: string[]
  /** The default row height in points. */
  defaultRowHeight?: number
  /** The sheet's tables: their header and border are drawn from these. */
  tables?: { name: string; range: string; headerRow: number }[]
  /** Conditional formatting data bars: the share of the cell's width (0..1) and the bar's colour. */
  bars?: Record<string, { ratio: number; color: string }>
  /** What the sheet's conditional formats give cells, as xlcalc computes it: a style over the cell's, a data bar, an icon, the value hidden. Drawn in place of `bars` where given. */
  appearance?: Record<string, ConditionalAppearance>
  /** Hyperlinks: an external address, a place in the book (`Sheet!B12`), a tooltip, a host's target (a the app file). */
  links?: Record<string, { href?: string; location?: string; tooltip?: string; target?: string }>
  notes?: Record<string, { author?: string; text: string }>
  /** A notice above the grid (the the app's "values computed by the the app"). */
  notice?: string
  /** Cells with a drill-down (list mode): underlined, a click or Enter emits `cellDrill`. */
  drill?: Record<string, true>
}

/** The window is not ready yet (the the app's book is recomputing): the component waits and asks again. */
export interface SheetPending {
  pending: { elapsedMs: number }
}

export interface CellEntry {
  sheet: string
  cell: string
  text: string | null
}

/** How an edit goes on. */
export interface EditOptions {
  /** The person kept a value the cell's rule does not take, by a warning's Yes or an information's OK: write it without asking the rule again. */
  accept?: true
  /** The text was edited from the cell's own (F2, a double click, the formula bar), not typed over the cell: Excel's percent editor adds no `%` then (`6%` in a `0%` cell is 0.06). */
  inPlace?: true
}

/** A cell's data validation as the person meets it: Excel's dropdown of a list rule and the message shown while the cell is selected. */
export interface SheetValidation {
  /** A list rule that shows its dropdown: the items as their text, in order; empty when its formula gives none. */
  list?: string[]
  /** The input message (Excel's title up to 32 characters, text up to 255). */
  prompt?: { title?: string; text: string }
}

/** The alert of a rule of data validation for a value it does not take, as Excel shows it: stop (Retry, Cancel), warning (Continue? Yes, No, Cancel), information (OK, Cancel). */
export interface SheetValidationAlert {
  severity: 'stop' | 'warning' | 'information'
  /** The rule's title, else the component's default ("Microsoft Excel"). */
  title?: string
  /** The rule's message, else the source's text of the rule. */
  message: string
  /** The cell the value was refused in. */
  sheet: string
  cell: string
}

export type SheetOp = Exclude<StructureOp, { op: 'addTable' | 'renameTable' | 'deleteTable' | 'convertTable' }>

/** Columns or rows changed: a width or height set with the mouse, lines hidden or shown, columns fitted to what they show (Excel's double click on an edge). */
export type LayoutChange =
  | { sheet: string; columns: string; width?: number; hidden?: boolean; fit?: true }
  | { sheet: string; rows: string; height?: number; hidden?: boolean }

export type EditOutcome =
  | {
      ok: true
      changes: CellUpdate[]
      /** Where the selection goes after `paste`, `fill`, `undo` or `redo`: what they filled or took back; `edit` gives none (the cursor has moved on). */
      select?: { sheet: string; range: string }
    }
  | {
      ok: false
      /**
       * conflict: the book changed elsewhere; refused: the cell is locked or the operation is not allowed;
       * invalid: the formula does not parse (`at` is the position in the typed text); busy: try again later.
       */
      reason: 'conflict' | 'refused' | 'invalid' | 'busy' | 'failed'
      message: string
      cells?: { sheet: string; cell: string }[]
      at?: number
      /** A value a rule of data validation does not take (`refused`): the component shows Excel's alert, and sends the edit again with `accept` after a warning's Yes or an information's OK. */
      alert?: SheetValidationAlert
    }

/** A cell's new state after an edit or a recomputation. */
export interface CellUpdate {
  sheet: string
  cell: string
  text: string
  value: CellValue
  color?: string
  /** Only for edited cells: the formula as shown, or null when the cell no longer has one. */
  formula?: string | null
  /** Only for edited cells, from a source whose windows have `editText`: the text the cell opens with now. */
  editText?: string
  /** Only for edited cells, from a source whose windows have `numberFormats`: the format code it shows by now (null: General), as a typed date sets one. */
  numberFormat?: string | null
}

export type SourceEvent =
  | { kind: 'changed'; changes: CellUpdate[] }
  /** Much or something unknown changed: read the sheets and windows again. */
  | { kind: 'reload'; message?: string }
  /** Cells of these ranges may look otherwise (conditional formats): windows that meet them are read again in place, the sheets, sizes and display filter as they are. */
  | { kind: 'repaint'; ranges: { sheet: string; range: string }[] }
