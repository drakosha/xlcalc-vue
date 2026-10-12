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
  /** An edit as typed (`=SUM(A1:A3)`, `1.5`, `'007`); `null` clears. Without it the sheet is read-only. */
  edit?(entries: CellEntry[]): Promise<EditOutcome>
  /** Clipboard text pasted at a cell as Excel pastes it, one step of undo; `select` is what it filled. Without it the fields go to `edit`. */
  paste?(sheet: string, cell: string, text: string): Promise<EditOutcome>
  /** Control+Return: the text typed into the active cell entered into every cell of `range`. Without it only the active cell takes it. */
  fill?(sheet: string, range: string, text: string, active: string): Promise<EditOutcome>
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
  /** A cell's precedents tree as text. */
  explain?(sheet: string, cell: string): Promise<string>
  /** Totals of a rectangle; without it the component counts over the loaded windows. */
  stats?(sheet: string, range: string, filter?: SheetFilter, signal?: AbortSignal): Promise<RangeStats>
  /** Case-insensitive substring search, hits in row order (with a filter only the visible rows). */
  find?(sheet: string, query: string, filter?: SheetFilter, signal?: AbortSignal): Promise<{ hits: { row: number; col: number; text: string }[]; total: number }>
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
}

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
