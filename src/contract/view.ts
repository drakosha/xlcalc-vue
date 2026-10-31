// The sheet's public types: SheetView's props, state, events and methods, the toolbar's source, list mode's sort.
import type { Component } from 'vue'
import type { CellStyle, CellValue, IconSetName, SheetProtection } from 'xlcalc/contract'
import type { ClearRulesFrom, ColorScalePreset, DataBarColor, QuickRule } from './conditional'
import type { SheetMessages } from './messages'
import type { AutoSumFunction, CalculationMode, EditOutcome, FormatPatch, OutlineChange, SheetOp, SheetSource, SheetWindow } from './source'
import type { ClearKind } from './workbook'

/** Properties of `SheetView`. */
export interface SheetViewProps {
  source: SheetSource
  /** Interface texts; a partial set overrides the English defaults. */
  messages?: Partial<SheetMessages>
  /** Key in the browser's storage for the frozen panes display choice, Show Formulas and column widths; without it they last as long as the component. */
  storageKey?: string
  /** A location to show: `Sheet!B12:D15`, `B7`, `Sheet`; each new object is a new jump. */
  focus?: { ref: string } | null
  /** The host page's address for a location; without it there is no "Link to the cell" button. */
  cellUrl?: (ref: string) => string
  readOnly?: boolean
  /** Shows the formula bar above the grid; on by default. */
  formulaBar?: boolean
  /** The workbook's locale for texts the component produces itself: a cell's edit text without `editText`, and selection totals. */
  locale?: string
  /** Hides the tabs of a workbook with a single sheet, for hosts that want it; Excel always shows them, which is the default. */
  hideLoneTab?: boolean
  /** Double-clicking a column edge restores the workbook's width, for hosts that want it; by default it fits the columns to their contents, as Excel does. */
  edgeRestoresWidth?: boolean
  /** List mode: the order the host sorted the list in, shown as an arrow on the column's title; null for none. */
  sort?: ListSort | null
  /** List mode: the first column opens its row (`rowOpen`), and so does Enter on a cell that is not a drill-down. */
  openRows?: boolean
  /** The host opens links itself through `openLink`: for a host whose listener the component cannot see, such as a custom element's. */
  hostLinks?: boolean
  /**
   * Excel's zoom by sheet name, in percent (10 to 400), as the host kept it from the `zoom` event; a sheet without one opens
   * at the workbook's saved zoom (`SheetInfo.zoom`), otherwise 100%. A new object applies its changed entries, to the active sheet at once.
   */
  zoom?: Record<string, number>
  /** Excel's Watch Window list, as the host kept it from the `watches` event; a new array replaces the list. */
  watches?: SheetWatch[]
  /** The author's name for a new note, which Excel writes in bold on its first line (Excel takes the user name from its Options); without it a new note has no author. */
  noteAuthor?: string
  /**
   * Draws Excel's PivotTable Fields pane to the right of the grid (below it on a phone) while it is shown
   * (`SheetViewState.fieldList`); on by default, false for a host that places `PivotFieldsPane` itself.
   */
  pivotFields?: boolean
}

/** A cell in Excel's Watch Window: its sheet and address (`B12`). */
export interface SheetWatch {
  sheet: string
  cell: string
}

/** Why the workbook cannot be displayed: the source's error and its text. */
export interface SheetFailure {
  error: unknown
  text: string
}

/** What `state` reports: the sheet, the selection, and the active cell's formula and displayed value. */
export interface SheetViewState {
  sheet: string
  ref: string
  formula?: string
  value?: string
  /** The active cell's style, for a toolbar; absent while its window is not loaded or when it has none. */
  style?: CellStyle
  /** The active cell's number format code, from a source whose windows include them; absent for General. */
  numberFormat?: string
  /** The active cell's value from the loaded window (absent for an empty cell), for a toolbar's Increase and Decrease Decimal. */
  cellValue?: CellValue
  /** The width of the active cell's column as the workbook stores it (xlcalc's `ColumnInput.width`), absent for the default: Increase Decimal on General depends on it. */
  columnWidth?: number
  /** The merged range containing the active cell (`A1:B2`), from the loaded windows: Excel's Merge & Center shows as pressed on it. */
  merged?: string
  /** The merge commands are available (`SheetViewApi.merge`): the source has `structure`, the sheet is editable, and no display filter is on. */
  canMerge?: true
  /** AutoSum is available (`SheetViewApi.autoSum`): the source has `autoSum` and the sheet is editable. */
  canAutoSum?: true
  /** Format Painter is available (`SheetViewApi.formatPainter`): the source has `structure`, the sheet is editable, and no display filter is on. */
  canFormatPainter?: true
  /** Format Painter holds a brush: `once` paints the next range, `keep` until it is turned off; a ribbon shows its button pressed. */
  painting?: FormatPainterMode
  /** The display filter has a condition or a sort: a ribbon shows its Filter button pressed and enables Clear. */
  filtered?: true
  /** The sheet's AutoFilter range (`SheetInfo.autoFilter`): a ribbon shows Data > Filter pressed. */
  autoFilter?: string
  /** Data > Filter is available (`toggleAutoFilter`): the source has `setAutoFilter` and the sheet is editable. */
  canAutoFilter?: true
  /** Excel's Show Formulas is on for the displayed sheet (Ctrl+` or `SheetViewApi.showFormulas`): a ribbon shows its button pressed. */
  showFormulas?: true
  /** The sheet's zoom in percent; absent at 100%. */
  zoom?: number
  /** The displayed sheet hides its gridlines or its headings (View > Show): a ribbon shows their boxes unchecked. */
  gridlines?: false
  headings?: false
  /** Excel's Show All Notes is on (`SheetViewApi.showNotes`): a ribbon shows its button pressed. */
  allNotes?: true
  /** The sheet's window is split (Excel's View > Split, `SheetInfo.split`), so a ribbon shows its button pressed. */
  split?: true
  /** The workbook's default font (`SheetInfo.font`), which a cell without a font of its own displays: a toolbar's Font and Font Size boxes show it for such a cell. */
  font?: { name: string; size: number }
  /** The workbook's calculation option, from a source with `calculation`: a ribbon shows the checked item of Calculation Options. */
  calculation?: CalculationMode
  /** In Manual, cells wait for a calculation: the status bar shows Calculate. */
  calculationPending?: true
  /** Refresh All is available (`SheetViewApi.refreshAll`): the source has `refreshConnections` and the sheet is editable. */
  canRefreshAll?: true
  /**
   * The PivotTable under the active cell (`SheetInfo.pivots`, its Filters area too), by its name: a ribbon shows Excel's
   * PivotTable Analyze and Design tabs while it is there.
   */
  pivot?: string
  /**
   * Refresh is available (`SheetViewApi.refreshPivot`): the active cell is in a PivotTable, the source has `structure`, and
   * the sheet is editable and not protected, as Excel refreshes no PivotTable on a protected sheet whatever it allows.
   */
  canRefreshPivot?: true
  /**
   * The PivotTable Fields pane is shown: the active cell is in a PivotTable, the source has `pivotTable`, and Excel's Field
   * List is on (`SheetViewApi.fieldList`), as Excel shows the pane only while the active cell is in a PivotTable.
   */
  fieldList?: true
  /** Insert PivotTable is available (`SheetViewApi.insertPivotTable`): the source has `structure` and the sheet is editable. */
  canInsertPivot?: true
  /** The sheet is protected (`SheetInfo.protection`), with what users may still do: a toolbar disables formatting without `formatCells`, as Excel does. */
  protection?: SheetProtection
  /** A selection of several ranges (Go To Special's): each of them as Excel writes it, the active one (`ref`) among them; absent for one range. A toolbar formats every one. */
  areas?: string[]
}

/** The `zoom` event: the user changed a sheet's zoom (Ctrl+wheel, a pinch, the status bar, the Zoom dialog), in percent. */
export interface SheetZoomEvent {
  sheet: string
  zoom: number
}

/** A hyperlink the component does not open itself, with the cell it belongs to. */
export type SheetLinkEvent = NonNullable<SheetWindow['links']>[string] & { sheet: string; cell: string }

/**
 * What `SheetView` exposes to its host (through a template `ref`): changes made as the sheet's own keyboard commands make
 * them, with the outcome's selection set without a flash or a Back history entry. Each method resolves to the source's
 * outcome, or to null when nothing was sent (the source lacks the method, the sheet is read-only, or an undo is pending).
 */
export interface SheetViewApi {
  undo(): Promise<EditOutcome | null>
  redo(): Promise<EditOutcome | null>
  /** Formats the active sheet in one call to the source's `format`. */
  format(patches: SheetFormatPatch[]): Promise<EditOutcome | null>
  /** A structure operation through the source; the sheet displayed afterwards is the one Excel displays. */
  structure(op: SheetOp): Promise<EditOutcome | null>
  /** An outline command through the source (Group, Ungroup, a group's button, a level button); the sheet reloads afterwards. */
  outline(change: OutlineChange): Promise<EditOutcome | null>
  /** Opens Excel's Format Cells dialog (Ctrl+1) for the selection, with the active cell's formatting; false when the sheet cannot be formatted. */
  formatCells(): boolean
  /**
   * Excel's Freeze Panes through the source's `freeze`, as a single undo step, after which the sheet reloads: `panes` (the
   * default) keeps the rows above and the columns to the left of the active cell in view (with A1 active, the middle of the
   * viewport, where Excel splits it), `unfreeze` removes the workbook's panes, `topRow` and `firstColumn` freeze one row or
   * column. Frozen panes are displayed whatever the lock button's earlier state; null when nothing was sent (no `freeze`,
   * a read-only sheet, nothing to freeze or unfreeze).
   */
  freezePanes(mode?: FreezeMode): Promise<EditOutcome | null>
  /** Excel's Reapply (Ctrl+Alt+L): applies the display filter's conditions and sort again to the current data; false when no filter is on. */
  reapplyFilter(): Promise<boolean>
  /** Excel's Trace Precedents as a panel next to the grid: the source's tree of the active cell's precedents, following the active cell; false if the source has no `explain`. */
  explain(): boolean
  /**
   * Excel's Merge & Center (`mergeCenter`), Merge Across, Merge Cells and Unmerge Cells over the selection through the
   * source's `structure`, as a single undo step, with Excel's warning first when more than one cell would lose its value. null
   * when nothing was sent: no `structure`, a read-only sheet, a display filter, a single cell outside a merged range, or Cancel
   * in the warning.
   */
  merge(command: MergeCommand): Promise<EditOutcome | null>
  /** Opens Excel's Go To dialog (F5, Ctrl+G) with the locations its jumps have left; false while no sheet is displayed. */
  goTo(): boolean
  /**
   * Opens Excel's Go To Special dialog (Home > Find & Select, Go To's Special…) through the source's `specialCells`: OK selects
   * the cells it finds, several ranges at once, keeping the active cell when one of them holds it. false without
   * `specialCells`, while no sheet is displayed or under a display filter.
   */
  goToSpecial(): boolean
  /**
   * Opens Excel's Symbol dialog (Insert > Symbol): Insert puts the character at the caret of an open edit, otherwise starts
   * editing the active cell with it, as typing does; false on a read-only sheet or without an active cell.
   */
  insertSymbol(): boolean
  /**
   * Excel's Fill Down (Ctrl+D): the selection's top row copied over its other rows through the source's `structure`
   * (`copyRange`, relative references shifted), or the row above it into a selection one row tall, as a single undo step.
   * null when nothing was sent: no `structure`, a read-only sheet, a display filter, row 1 alone, or a locked cell.
   */
  fillDown(): Promise<EditOutcome | null>
  /** Excel's Fill Right (Ctrl+R): `fillDown` along the columns, from the selection's left column or the column to its left. */
  fillRight(): Promise<EditOutcome | null>
  /** Excel's Show Formulas (Ctrl+`) on the displayed sheet, toggled without `on`: a view setting kept as the frozen panes choice is, the workbook unchanged; whether it is on afterwards. */
  showFormulas(on?: boolean): boolean
  /**
   * Excel's View > Show > Gridlines on the displayed sheet, toggled without `on`; whether they show afterwards. Kept through
   * the source's `show` (not an undo step, as in Excel), and in the browser for a read-only sheet or a source without it.
   */
  showGridlines(on?: boolean): boolean
  /** Excel's View > Show > Headings (row numbers and column letters) on the displayed sheet, as `showGridlines`. */
  showHeadings(on?: boolean): boolean
  /** Excel's Show All Notes: every note drawn in its box all the time, toggled without `on`; a view setting the source does not get, as in Excel; whether it is on afterwards. */
  showNotes(on?: boolean): boolean
  /**
   * Excel's View > Split, toggled without `on`: splits the window above and to the left of the active cell (at the middle
   * of the viewport with the top-left cell in view active, as Excel does) or removes the split, through the source's
   * `split`, not an undo step; whether the sheet is split afterwards. false without `split`, on a read-only sheet, with
   * frozen panes or under a display filter. The sheet draws the split's panes, which scroll apart, and its bars.
   */
  split(on?: boolean): boolean
  /** Excel's Define Name: the New Name dialog for the selection; false when names cannot be defined (no `names` or `setNames`, a read-only sheet) or a cell is being edited. */
  defineName(): boolean
  /** Excel's Name Manager (Ctrl+F3); false when the source has no `names` or a cell is being edited. Without `setNames` or on a read-only sheet it only lists the names. */
  nameManager(): boolean
  /**
   * Excel's Create from Selection (Ctrl+Shift+F3, Cmd+Shift+F3 on a Mac): the Create Names dialog for the selection, with the
   * boxes Excel checks; OK makes the names through the source's `namesFromSelection` and `setNames` as a single undo step,
   * asking first about a name the workbook defines otherwise. false without `namesFromSelection`, `names` or `setNames`, on a
   * read-only sheet or while a cell is being edited.
   */
  createNames(): boolean
  /** Opens Excel's Workbook Statistics for the displayed sheet once the source's `workbookStatistics` answers; false without it or while no sheet is displayed. */
  workbookStatistics(): boolean
  /**
   * Excel's zoom of the active sheet: a percentage (rounded and kept within 10 to 400), or `selection` for Zoom to Selection,
   * which fits the selection into the viewport and scrolls it to the top-left corner. Returns the zoom set; it is not
   * reported by the `zoom` event, which reports the user's changes.
   */
  setZoom(zoom: number | 'selection'): number
  /** Opens Excel's Watch Window over the sheet (`true`, the default), or closes it (`false`); returns whether it is open. */
  watchWindow(open?: boolean): boolean
  /**
   * Excel's Trace Dependents through the source's `dependents`: arrows from the active cell to the cells whose formulas refer
   * to it, and with each further call the next level; false when nothing was added (no `dependents`, no further level).
   */
  traceDependents(): Promise<boolean>
  /** Excel's Remove Arrows: every tracer arrow; false when there were none. */
  removeArrows(): boolean
  /** Excel's Insert Table and Format as Table (Ctrl+T, Ctrl+L): the Create Table dialog for the selection or the data around the active cell; false without the source's `table`, or on a read-only sheet. */
  insertTable(): Promise<boolean>
  /** Excel's Convert to Range on the table at the active cell, after Excel's question; null when nothing was sent (no table there, no `table`, No). */
  convertToRange(): Promise<EditOutcome | null>
  /** Delete Table: the table at the active cell with its cells; null when nothing was sent. */
  deleteTable(): Promise<EditOutcome | null>
  /** Excel's Spelling (F7) through the source's `spelling`: the dialog from the active cell, within a selection of several cells; false without `spelling`. */
  checkSpelling(): boolean
  /**
   * Excel's AutoSum (Alt+=) through the source's `autoSum`: the active cell's edit opens with `=SUM(range)`, the range
   * selected, so Enter enters it and typing or the arrows replace it; with `fn`, the function of Excel's AutoSum list
   * (`=AVERAGE(range)` and so on). false without `autoSum`, with a function the source does not list in `autoSumFunctions`,
   * on a read-only sheet, while a cell is being edited or when the source guesses nothing.
   */
  autoSum(fn?: AutoSumFunction): Promise<boolean>
  /** Opens Excel's Insert Function dialog (the formula bar's fx, AutoSum's More Functions…); false on a read-only sheet or without the source's `functionNames`. */
  insertFunction(): boolean
  /**
   * Opens Excel's Paste Special dialog (Ctrl+Alt+V, Ctrl+Cmd+V on a Mac) for the copied range: its Paste choices, Operation,
   * Skip Blanks and Transpose, through the source's `structure` (`copyRange` with `paste`, `transpose`, `operation` and
   * `skipBlanks`, xlcalc 0.65). false without a copy
   * in copy mode (a cut has no Paste Special), without `structure`, on a read-only sheet or under a display filter.
   */
  pasteSpecial(): boolean
  /**
   * Excel's Format Painter: `once` (the default) takes the selection's formats and paints them onto the next range selected
   * with the mouse or a touch, `keep` paints until Esc or `off`, as a double click on Excel's button; each painting is a single
   * undo step through `structure` (`copyRange` with `paste: 'formats'`). Returns whether painting is on afterwards.
   */
  formatPainter(mode?: FormatPainterMode | 'off'): boolean
  /**
   * Data > Filter for a ribbon: opens the display filter's menu (the funnel on the column letter: its sort, values and text
   * condition) for the active cell's column; false without the source's `visibleRows` and `columnValues`. The data do not change.
   */
  filterColumn(): boolean
  /** Data > Clear: every condition and the sort of the display filter at once; false when none is on. */
  clearFilter(): boolean
  /**
   * Sort A to Z (`asc`) or Z to A (`desc`) by the active cell's column, as the display filter sorts: the rows below the header
   * are displayed in that order and the data do not change (Excel's sort rewrites them); false without the display filter.
   */
  sortColumn(dir: 'asc' | 'desc'): boolean
  /**
   * Excel's Data > Filter (Ctrl+Shift+L) through the source's `setAutoFilter`, as a single undo step: removes the sheet's
   * AutoFilter with its conditions, or puts one over the selection, or over the active cell's current region when a single
   * cell is selected, trimmed to its data as Excel trims it; the sheet reloads afterwards. null when nothing was sent: no
   * `setAutoFilter`, a read-only sheet, or no data there (the sheet displays Excel's message).
   */
  toggleAutoFilter(): Promise<EditOutcome | null>
  /**
   * Excel's Home > Clear > Clear Formats (`formats`) or Clear All (`all`) over the selection through the source's `clear`, as a
   * single undo step (under a display filter, over its visible rows), after which the windows over it are read again. null when
   * nothing was sent: no `clear`, a read-only sheet, or a locked cell in the selection.
   */
  clear(what: ClearKind): Promise<EditOutcome | null>
  /**
   * Excel's Insert Link (Ctrl+K) over the selection: opens its dialog, as Edit Hyperlink when the active cell has a link; OK
   * writes the link through the source's `link`. false when links cannot be edited: no `link`, a read-only sheet, a locked cell.
   */
  insertLink(): boolean
  /**
   * Opens Excel's Data Validation dialog (Data > Data Validation) for the selection over the source's `rules` and `setRules`,
   * after Excel's question when the selection holds rules of more than one kind or cells without the rule; false when the
   * sheet is read-only or the source lacks either method.
   */
  dataValidation(): boolean
  // Excel's Conditional Formatting menu: each needs the source's `rules` and `setRules` and an editable sheet (false or null otherwise)
  /** Opens a quick rule's dialog (Highlight Cells Rules, Top/Bottom Rules) for the selection; OK adds the rule first in priority. */
  quickRule(kind: QuickRule): boolean
  /** Adds Excel's data bar of that color over the selection, with a gradient fill (the default) or a solid one. */
  dataBars(color: DataBarColor, fill?: 'gradient' | 'solid'): Promise<EditOutcome | null>
  /** Adds the color scale of Excel's gallery over the selection. */
  colorScale(preset: ColorScalePreset): Promise<EditOutcome | null>
  /** Adds the icon set over the selection, with Excel's default thresholds: equal shares in percent. */
  iconSet(icons: IconSetName): Promise<EditOutcome | null>
  /** Excel's Clear Rules from Selected Cells (a rule keeps the rest of its cells) or from Entire Sheet; null when there is nothing to clear. */
  clearRules(from: ClearRulesFrom): Promise<EditOutcome | null>
  /** Opens Excel's New Formatting Rule dialog for the selection. */
  newRule(): boolean
  /** Opens Excel's Conditional Formatting Rules Manager. */
  manageRules(): boolean
  /**
   * Excel's Calculate Now (F9, `workbook`, the default) or Calculate Sheet (Shift+F9, `sheet`) through the source's
   * `calculate`; null when nothing was sent (no `calculate`, a read-only sheet).
   */
  calculate(scope?: 'workbook' | 'sheet'): Promise<EditOutcome | null>
  /** Excel's Calculation Options through the source's `setCalculation`; null when nothing was sent (no `setCalculation`, a read-only sheet). */
  setCalculation(mode: CalculationMode): Promise<EditOutcome | null>
  /** Excel's Data > Refresh All (Ctrl+Alt+F5) through the source's `refreshConnections`; null when nothing was sent (no `refreshConnections`, a read-only sheet). */
  refreshAll(): Promise<EditOutcome | null>
  /**
   * Excel's PivotTable Analyze > Refresh (Alt+F5) of the PivotTable under the active cell through the source's `structure`
   * (`refreshPivots` naming it): its cache read again from its source, with every PivotTable over it, as a single undo step;
   * the sheet reloads afterwards. null when nothing was sent (no PivotTable there, no `structure`, a read-only or protected sheet).
   */
  refreshPivot(): Promise<EditOutcome | null>
  /**
   * Excel's Insert > PivotTable: opens its dialog with the table at the active cell, or the selection, or the active cell's
   * current region; OK creates the PivotTable on a new sheet at A3 or at the location given, through the source's
   * `structure` (`addPivotTable`) as a single undo step, and selects its top-left cell, where the PivotTable Fields pane
   * shows. false without `structure` or on a read-only sheet.
   */
  insertPivotTable(): Promise<boolean>
  /**
   * Excel's PivotTable Analyze > Field List, toggled without `show`: whether the PivotTable Fields pane shows while the
   * active cell is in a PivotTable; returns whether it is on afterwards. On by default, as in Excel.
   */
  fieldList(show?: boolean): boolean
  /**
   * Asks for a dialog over a field of a PivotTable, as the PivotTable Fields pane does: Field Settings, Value Field Settings,
   * or the field's sort and filter menu. The sheet sends it as its `pivotDialog` event; false while the sheet is read-only.
   */
  pivotDialog(request: PivotDialogRequest): boolean
  /**
   * Opens Excel's Protect Sheet dialog for the displayed sheet: what users may still do, without a password; OK protects it
   * through the source's `protect`, as a single undo step. false without `protect`, on a read-only sheet, or on a protected one.
   */
  protectSheet(): boolean
  /** Excel's Unprotect Sheet on the displayed sheet through the source's `protect`; null when nothing was sent (no `protect`, a read-only sheet, an unprotected one). */
  unprotectSheet(): Promise<EditOutcome | null>
}

/**
 * A dialog over a field of a PivotTable that the PivotTable Fields pane asks for: `fieldSettings` (Field Settings… of a field
 * on the Filters, Columns or Rows area), `valueFieldSettings` (Value Field Settings… of a value field, `value` its index in
 * the PivotTable's `values` as the source holds it), `fieldFilter` (the field's sort and filter menu, the arrow beside it in the field list). `field`
 * is the index of the cache field; `at` is where the pane's control is, in the viewport's coordinates, for a menu.
 */
export interface PivotDialogRequest {
  dialog: 'fieldSettings' | 'valueFieldSettings' | 'fieldFilter'
  sheet: string
  pivot: string
  field: number
  value?: number
  at?: { left: number; top: number; width: number; height: number }
}

/** Properties of `PivotFieldsPane`, the PivotTable Fields pane a host places itself (with `SheetView`'s `pivotFields` off). */
export interface PivotFieldsPaneProps {
  /** The source the pane reads the PivotTable from (`pivotTable`) and whose events it follows. */
  source: SheetSource
  /** The sheet whose changes the pane makes (`structure`), whose Field List it closes and whose dialogs it asks for. */
  sheet: SheetViewApi | null
  /** The sheet's `state` event: the pane shows the PivotTable of `pivot` while `fieldList` is on. */
  state: SheetViewState | null
  messages?: Partial<SheetMessages>
}

/** A patch for `SheetView.format`: applied to its own range, or to the selection without one. */
export type SheetFormatPatch = Omit<FormatPatch, 'range'> & { range?: string }

/** Format Painter's brush: `once` paints one range, `keep` paints until it is turned off. */
export type FormatPainterMode = 'once' | 'keep'

/** Excel's Freeze Panes commands: the panes of the active cell, Unfreeze Panes, Freeze Top Row, Freeze First Column. */
export type FreezeMode = 'panes' | 'unfreeze' | 'topRow' | 'firstColumn'

/** Excel's merge commands: Merge & Center on the button, and Merge Across, Merge Cells and Unmerge Cells in its dropdown. */
export type MergeCommand = 'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells'

/** The order a list is sorted in, by a column of the sheet (1-based): the host sorts, and the sheet shows the arrow. */
export interface ListSort {
  col: number
  dir: 'asc' | 'desc'
}

/** The source the toolbar formats through: a `SheetSource`, whose `history` tells whether Undo and Redo have a step. */
export type ToolbarSource = SheetSource

/**
 * A group of `SheetToolbar`'s buttons, as `items` names it: Undo and Redo, Format Painter, Bold, Italic and Underline (`fontStyle`), the
 * Font and Font Size boxes with Increase and Decrease Font Size (`font`), the Number Format box with its buttons (`number`),
 * font and fill color, borders, Top, Middle and Bottom Align, horizontal alignment with indent and Wrap Text (`align`),
 * Merge & Center, Format Cells, AutoSum, Insert > PivotTable, and Data > Refresh All.
 */
export type ToolbarGroup = 'undo' | 'formatPainter' | 'fontStyle' | 'font' | 'number' | 'colors' | 'borders' | 'verticalAlign' | 'align' | 'merge' | 'formatCells' | 'autoSum' | 'pivotTable' | 'refreshAll'

/** A host's button on `SheetToolbar`, drawn as the toolbar's own and, where it does not fit, moved into its menu. */
export interface ToolbarButton {
  /** Unique among the host's buttons; the button's hook is `data-tb-host="<id>"`. */
  id: string
  /** The tooltip and accessible name, and the button's text without an icon. */
  label: string
  /** A Vue component (it gets the toolbar's icon class) or SVG markup, drawn at the toolbar's icon size. */
  icon?: Component | string
  /** The click's action; with `menu` too, the button is split, as AutoSum's: this part runs, its arrow opens the list. */
  run?: () => void
  /** A dropdown list the toolbar draws as its own lists (the narrow menu too); without `run` a click opens it, as Excel's Conditional Formatting. */
  menu?: readonly ToolbarMenuItem[]
  disabled?: boolean
  /** Shown as pressed (`aria-pressed`), as a toggle. */
  pressed?: boolean
}

/** An item of a host's button's list; its hook is `data-tb-host-item="<id>"`. */
export interface ToolbarMenuItem {
  id: string
  label: string
  run: () => void
  disabled?: boolean
}

/** An item of `SheetToolbar`'s `items`: one of its groups, or a host's button. */
export type ToolbarItem = ToolbarGroup | ToolbarButton
