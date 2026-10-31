/**
 * The sheet source contract: everything the component knows about its data. Rows and columns are 1-based, cells
 * are A1 addresses. The source displays formulas in its own locale and parses what the user typed; the component knows
 * nothing about the engine, the formula locale or HTTP. Optional methods enable features: without `edit` the sheet
 * is read-only, without `find` there is no search, and so on.
 */
import type { CellStyle, CellValue, ConditionalAppearance, FunctionSignature, NumberFormat, PivotCache, PivotRef, PivotTable, SelectionLabels, SheetProtection, Slicer, SlicerCache, SpecialCells, StructureOp, SplitPanes, StylePatch, Timeline, TimelineCache, WorkbookStatistics } from 'xlcalc/contract'
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
  /**
   * A structure operation; without it there is no structure menu. xlcalc's operations on PivotTables are among them: the
   * sheet sends Excel's Refresh (Alt+F5) of the PivotTable under the active cell as `refreshPivots` naming it.
   */
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
  /** Excel's Tab Color: a sheet's tab color (`#RRGGBB`), null for No Color; a single undo step. Without it the tab menu has no Tab Color. */
  tabColor?(sheet: string, color: string | null): Promise<EditOutcome>
  /** Excel's View > Show on a sheet: its gridlines and its headings, kept with the workbook; not an undo step, as in Excel. Without it the sheet keeps them in the browser. */
  show?(sheet: string, show: SheetShow): Promise<EditOutcome>
  /** Excel's View > Split on a sheet: its split window (`SheetInfo.split`), null to remove it; kept with the workbook, not an undo step, as in Excel. Without it the sheet offers no Split. */
  split?(sheet: string, split: SheetSplit | null): Promise<EditOutcome>
  /**
   * Excel's Protect Sheet with what users may still do (`{}` is the dialog's default protection), or null for Unprotect
   * Sheet; a single undo step, after which the sheets are read again (`SheetInfo.protection`). There is no password: the
   * workbook holds none. Without it the sheet tab's menu offers neither command; a source that has it refuses on a protected
   * sheet what Excel refuses there (`EditOutcome` code `protected`).
   */
  protect?(sheet: string, protection: SheetProtection | null): Promise<EditOutcome>
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
  /**
   * Excel's Data > Filter: puts the sheet's AutoFilter (`SheetInfo.autoFilter`) over `range`, its header row first (`A1:D20`),
   * or removes it with null, showing the rows its filter hid; a single undo step. Without it a sheet keeps the AutoFilter it
   * has, which cannot be turned on or off, and a sheet without one has the display filter's funnels on its column letters.
   */
  setAutoFilter?(sheet: string, range: string | null): Promise<EditOutcome>
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
  /**
   * Excel's AutoSum (Alt+=) at a cell: the range it offers and its `SUM`, or with `fn` (one of `autoSumFunctions`) the same
   * range under that function, as Excel's AutoSum list enters it; null when the source cannot guess one. Without it there is no AutoSum.
   */
  autoSum?(sheet: string, cell: string, signal?: AbortSignal, fn?: AutoSumFunction): Promise<SheetAutoSum | null>
  /** The functions of Excel's AutoSum list that `autoSum` builds besides Sum; without it the list offers Sum and More Functions… only. */
  readonly autoSumFunctions?: readonly AutoSumFunction[]
  /**
   * Excel's Go To Special over a range: the rectangles Excel selects (`['B2:D2', 'B4']`), none when Excel finds no cells. A
   * range of one cell stands for the sheet's used range, as in Excel. `kind` is xlcalc's `SpecialCells` (blank cells,
   * constants and formulas by the kinds of their values, visible cells) or `lastCell`, the used range's last cell whatever
   * the range. Without it there is no Go To Special.
   */
  specialCells?(sheet: string, range: string, kind: SheetSpecialCells, signal?: AbortSignal): Promise<string[]>
  /**
   * Excel's Create from Selection over a range: the workbook names its labels make (`labels`: the sides that hold them, as the
   * dialog's boxes; at least one), each referring absolutely to its cells with the sheet, in the source's locale, in the order
   * Excel makes them; none for labels alone or labels without text. The sheet asks about names the workbook already has and
   * writes them with `setNames`. Without it there is no Create from Selection.
   */
  namesFromSelection?(sheet: string, range: string, labels: SelectionLabels, signal?: AbortSignal): Promise<SheetName[]>
  /**
   * Excel's Workbook Statistics (Review > Workbook Statistics) with `sheet` active: the sheet's end and counts and the
   * workbook's (xlcalc's `WorkbookStatistics`), a count the workbook does not hold being null. Without it there is no
   * Workbook Statistics.
   */
  workbookStatistics?(sheet: string, signal?: AbortSignal): Promise<WorkbookStatistics>
  /** The maximum number of cells in one window, 20,000 by default. */
  readonly maxCells?: number
  /** The number of steps `undo` and `redo` can take, as of the last completed call: a toolbar enables its buttons by them. */
  readonly history?: { readonly undo: number; readonly redo: number }
  /**
   * The workbook's calculation option (Excel's Formulas > Calculation Options) and whether cells wait for a calculation,
   * read when the sheets are read; a `calculation` event tells a change. Without it the workbook calculates automatically.
   */
  calculation?(signal?: AbortSignal): Promise<SheetCalculation>
  /** Sets the calculation option, which the workbook keeps; not a step of undo. Automatic calculates what waits. Without it the option cannot be changed. */
  setCalculation?(mode: CalculationMode): Promise<EditOutcome>
  /**
   * Excel's Calculate Now (F9), or with `sheet` Calculate Sheet (Shift+F9): the cells that wait for a calculation show
   * their values, returned as `changes` and sent as a `changed` event; not a step of undo. Without it there is no F9.
   */
  calculate?(sheet?: string): Promise<EditOutcome>
  /**
   * Excel's Data > Refresh All (Ctrl+Alt+F5): the workbook's calls outside itself asked again, of the functions `names`
   * (upper-case) or of every one (xlcalc's `refreshConnections`), and the cells they change returned as `changes` and sent as
   * a `changed` event; answers still on their way arrive as `changed` events. Without `names` its PivotTables are refreshed
   * from their sources too, after the calls (xlcalc's `refreshPivots`), and the sheets are read again. Asking the calls again
   * is not a step of undo, as in Excel; refreshing the PivotTables is one, as xlcalc's. Without it there is no Refresh All.
   */
  refreshConnections?(names?: string[]): Promise<EditOutcome>
  /**
   * The PivotTable `name` on `sheet` as xlcalc holds it, with its pivot cache's source and fields, for the PivotTable Fields
   * pane and the dialogs over a PivotTable; null when the sheet has no such PivotTable. Read again after a `reload` or
   * `sheets` event. Without it there is no PivotTable Fields pane.
   */
  pivotTable?(sheet: string, name: string, signal?: AbortSignal): Promise<SheetPivotTable | null>
  /**
   * What the cells of the PivotTable `name` on `sheet` show, as it was last laid out: its fields' items as labels, the
   * header cells with Excel's filter buttons, and what each of `cells` is (an item, a total, a value, a header), for the
   * cell menu, the filter buttons and the dialogs over a PivotTable; null when the sheet has no such PivotTable. Without
   * it a PivotTable's cells have no buttons and its cell menu no commands but Refresh and Field List.
   */
  pivotCells?(sheet: string, name: string, cells?: string[], signal?: AbortSignal): Promise<SheetPivotCells | null>
}

/**
 * Excel's calculation options: `automatic`; `autoNoTable`, automatic except for data tables (What-If Analysis); `manual`,
 * where an edit leaves the formulas that depend on it with their values until the user calculates.
 */
export type CalculationMode = 'automatic' | 'autoNoTable' | 'manual'

/** The workbook's calculation: its option and, in Manual, whether cells wait for a calculation (Excel's status bar then shows Calculate). */
export interface SheetCalculation {
  mode: CalculationMode
  pending: boolean
}

/**
 * Excel's AutoSum at a cell: the range it offers to sum (the run of numbers above the cell, else to its left, up to a
 * subtotal; the subtotals alone when the nearest is one), null without numbers, and the formula it enters, as the source
 * displays formulas, without `=` (`SUM(B2:B9)`, `СУММ(B2:B9)` in ru-RU, `SUM()` without a range).
 */
export interface SheetAutoSum {
  range: string | null
  formula: string
}

/** What `SheetSource.specialCells` finds: xlcalc's `SpecialCells`, or the last cell of the sheet's used range (Excel's Last cell). */
export type SheetSpecialCells = SpecialCells | { type: 'lastCell' }

/** The functions of Excel's AutoSum list: Sum, Average, Count Numbers (`COUNT`), Max, Min. */
export type AutoSumFunction = 'sum' | 'average' | 'count' | 'max' | 'min'

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
  /** The last row the conditions and the sort apply to (an AutoFilter's last row); the rows below it follow in their order. Every row below the header without it. */
  last?: number
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
  /** Excel's Tab Color (`#RRGGBB`): another tab is filled with it, the active one underlined in it. */
  tabColor?: string
  /** False hides the gridlines (Excel's View > Show > Gridlines, xlsx `showGridLines`); shown by default. */
  gridlines?: boolean
  /** False hides the row numbers and column letters (Excel's View > Show > Headings, xlsx `showRowColHeaders`); shown by default. */
  headings?: boolean
  /**
   * Excel's AutoFilter outside a table: the range it acts on, header row first (`A1:D20`); the filter buttons are on that
   * row, and the display filter applies to the rows below it. Data typed under it may grow the range (xlcalc's
   * `getAutoFilterRange`), and a source that grows it sends `sheets`.
   */
  autoFilter?: { range: string }
  /**
   * Excel's Protect Sheet: the sheet is protected, with what users may still do (`{}`: selecting cells only, Excel's
   * default). Its cells whose style is locked (`CellStyle.protection`, locked unless a style says otherwise) take no edit,
   * and a hidden cell's formula is not displayed; the component disables the commands Excel disables. A source enforces it.
   */
  protection?: SheetProtection
  /** Excel's split window (View > Split) where the sheet has no frozen panes, which take its place. */
  split?: SheetSplit
  /**
   * The sheet's PivotTables as last laid out, from which the sheet finds the PivotTable under a cell for its commands
   * (Refresh); a sheet without them has none. A change that lays one out again elsewhere sends `reload` or `sheets`.
   */
  pivots?: SheetPivot[]
  /**
   * The sheet's slicers and timelines over PivotTables, drawn over the cells at their places with the choice they hold; a
   * sheet without them has none. An operation on them, a refresh or an undo sends `reload` or `sheets`, and they are read again.
   */
  slicers?: SheetSlicer[]
  timelines?: SheetTimeline[]
}

/**
 * A PivotTable on a sheet by its areas, each a range as Excel writes it: `range` the table without its Filters area (Excel's
 * `TableRange1`); `filters` its Filters area above it, as wide as its filters, without the blank row between (`PageRange`);
 * `headers` the table's rows above its values: the value fields' and column fields' headers and the column labels;
 * `rowLabels` the row labels left of the values, down to the grand total; `values` the values with their totals
 * (`DataBodyRange`). An area the PivotTable does not show is absent (`values` without a value field).
 */
export interface SheetPivot {
  /** Unique among the sheet's PivotTables: an operation on the PivotTable names it by the sheet and this name (xlcalc's `PivotRef`). */
  name: string
  range: string
  filters?: string
  headers?: string
  rowLabels?: string
  values?: string
  /** The `id` of its pivot cache: the PivotTables a slicer or a timeline can filter together share it (Report Connections). */
  cache?: number
  /** The fields of its cache that Insert Slicer offers, in the cache's order, a field of dates marked for Insert Timeline. */
  fields?: SheetPivotField[]
}

/** A field of a PivotTable's cache by its name in the source, as a slicer or a timeline names it; `dates` when it holds dates. */
export interface SheetPivotField {
  name: string
  dates?: true
}

/**
 * A slicer on a sheet: xlcalc's `Slicer` (its name, its cache's name, `caption`, `showCaption`, `place`, `columns`,
 * `rowHeight` in EMU, `startItem`, `locked`, `style`) with what its cache holds: the source field, the PivotTables it
 * filters, their pivot cache, the sorting and display of the items, and the items in the slicer's order. A slicer
 * without `place` is not drawn.
 */
export interface SheetSlicer extends Slicer, Pick<SlicerCache, 'sort' | 'customListSort' | 'noData' | 'showMissing'> {
  field: string
  pivots: PivotRef[]
  /** The pivot cache's `id`; absent for a table's slicer, whose items the sheet does not change. */
  pivotCache?: number
  items: SheetSlicerItem[]
}

/**
 * An item of a slicer, in the slicer's order (with data first unless its cache says otherwise): `item` its index in the
 * cache field's items, which `editSlicer` names it by; `label` as the slicer shows it; `selected` false when not chosen;
 * `noData` when no record of it passes the other filters and the slicer marks such items.
 */
export interface SheetSlicerItem {
  item: number
  label: string
  selected?: false
  noData?: true
}

/**
 * A timeline on a sheet: xlcalc's `Timeline` (its name, its cache's name, `caption`, `place`, `level`, `selectionLevel`,
 * what it shows, `scrollPosition`, `style`) with what its cache holds: the date field, the PivotTables it filters, their
 * pivot cache, the dates chosen (`selection`, none for all periods) and the dates it spans (`bounds`, whole years).
 */
export interface SheetTimeline extends Timeline, Pick<TimelineCache, 'selection' | 'bounds'> {
  field: string
  pivots: PivotRef[]
  pivotCache: number
}

/**
 * A PivotTable's definition (`SheetSource.pivotTable`): xlcalc's `PivotTable` without its saved lines, its fields on the
 * axes and its value fields by the index of a cache field, and its pivot cache without the rows it saved: the source, and
 * the fields with their items and the kinds of value each holds (`PivotCacheField.facts`), in the cache's order.
 */
export interface SheetPivotTable {
  table: Omit<PivotTable, 'rowLines' | 'columnLines'>
  cache: Omit<PivotCache, 'records'>
}

/**
 * What a PivotTable's cells show (`SheetSource.pivotCells`): `labels`, per cache field on an axis, its items' labels in the
 * order of the PivotTable field's `items` (a caption, or the value as the PivotTable shows it), null for a field on no
 * axis; `buttons`, the cells with Excel's filter button (Row Labels and Column Labels in compact form, each field's header
 * otherwise, the item of a field of the Filters area), none with Display field captions and filter drop downs off;
 * `cells`, what each cell asked for is, null outside the PivotTable.
 */
export interface SheetPivotCells {
  labels: (string[] | null)[]
  buttons: SheetPivotButton[]
  cells?: (SheetPivotCell | null)[]
}

/** A filter button of a PivotTable: its cell, and the field (by its index in the cache) on the axis it filters first. */
export interface SheetPivotButton {
  cell: string
  axis: PivotAxis
  field: number
}

/** The areas of a PivotTable's fields: the rows, the columns and the Filters area (`pages`). */
export type PivotAxis = 'rows' | 'columns' | 'pages'

/**
 * A cell of a PivotTable (`SheetPivotCells.cells`): `kind` is an item's label (`item`), a subtotal's label or a grand
 * total's (`subtotal`, `grand`), a value with its totals (`value`), a header (`header`: Row Labels, Column Labels or a
 * field's name), a value field's name (`valueHeader`), a field of the Filters area's name or item (`pageField`, `pageItem`),
 * or none of them (`blank`). `field` is the cache field of the label, header or line, `item` its index in the PivotTable
 * field's `items`, `value` the value field's index in `values`.
 */
export interface SheetPivotCell {
  kind: 'item' | 'subtotal' | 'grand' | 'value' | 'header' | 'valueHeader' | 'pageField' | 'pageItem' | 'blank'
  axis?: PivotAxis
  field?: number
  item?: number
  value?: number
}

/**
 * Excel's split window, xlcalc's `SplitPanes`: the vertical line `x` points from the window's left edge and the horizontal
 * one `y` points from its top (at least one), as the file holds them, headings counted in, at the view's zoom; `topLeft`,
 * the top-left cell of the pane after both lines; `activePane`, the pane that holds the active cell (the first without it).
 */
export type SheetSplit = SplitPanes

/** What Excel's View > Show sets on a sheet: whether its gridlines and its headings show. */
export interface SheetShow {
  gridlines?: boolean
  headings?: boolean
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
  /** Notes; `shown` is drawn all the time, in its box next to its cell. */
  notes?: Record<string, { author?: string; text: string; shown?: boolean }>
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
       * `./local` the code of any structure operation xlcalc refuses. `protected`: a protected sheet does not allow it, with
       * Excel's alert as the message; `pivot`: a PivotTable's cells do not take it, or an operation on a PivotTable failed,
       * with Excel's alert or the reason as the message. The component displays both as Excel does, in an alert.
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
  /** The calculation option changed, or cells began or stopped waiting for a calculation. */
  | { kind: 'calculation'; calculation: SheetCalculation }
  /** The sheets' information changed without their cells (an AutoFilter's range grown by data typed under it): read the sheets again, keeping the windows and the display filter's rows. */
  | { kind: 'sheets' }
