// The sheet's public types: SheetView's props, state, events and methods, the toolbar's source, list mode's sort.
import type { CellStyle } from 'xlcalc/contract'
import type { SheetMessages } from './messages'
import type { EditOutcome, FormatPatch, OutlineChange, SheetOp, SheetSource, SheetWindow } from './source'

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
   * at 100%. A new object applies its changed entries, to the active sheet at once.
   */
  zoom?: Record<string, number>
  /** Excel's Watch Window list, as the host kept it from the `watches` event; a new array replaces the list. */
  watches?: SheetWatch[]
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
  /** The merged range containing the active cell (`A1:B2`), from the loaded windows: Excel's Merge & Center shows as pressed on it. */
  merged?: string
  /** The merge commands are available (`SheetViewApi.merge`): the source has `structure`, the sheet is editable, and no display filter is on. */
  canMerge?: true
  /** Excel's Show Formulas is on for the displayed sheet (Ctrl+` or `SheetViewApi.showFormulas`): a ribbon shows its button pressed. */
  showFormulas?: true
  /** The sheet's zoom in percent; absent at 100%. */
  zoom?: number
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
  /** Excel's Define Name: the New Name dialog for the selection; false when names cannot be defined (no `names` or `setNames`, a read-only sheet) or a cell is being edited. */
  defineName(): boolean
  /** Excel's Name Manager (Ctrl+F3); false when the source has no `names` or a cell is being edited. Without `setNames` or on a read-only sheet it only lists the names. */
  nameManager(): boolean
  /**
   * Excel's zoom of the active sheet: a percentage (rounded and kept within 10 to 400), or `selection` for Zoom to Selection,
   * which fits the selection into the viewport and scrolls it to the top-left corner. Returns the zoom set; it is not
   * reported by the `zoom` event, which reports the user's changes.
   */
  setZoom(zoom: number | 'selection'): number
  /** Opens Excel's Watch Window over the sheet (`true`, the default), or closes it (`false`); returns whether it is open. */
  watchWindow(open?: boolean): boolean
}

/** A patch for `SheetView.format`: applied to its own range, or to the selection without one. */
export type SheetFormatPatch = Omit<FormatPatch, 'range'> & { range?: string }

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
