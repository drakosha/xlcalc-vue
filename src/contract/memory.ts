// memorySource's public types: the in-memory workbook, its sheets, cells and validation rules, and the source's extras.
import type { CellStyle, CellValue, ConditionalAppearance } from 'xlcalc/contract'
import type { CellUpdate, EditOutcome, SheetInfo, SheetSource, SheetValidationAlert, SheetWindow, SourceEvent } from './source'

/** A cell of the in-memory workbook; a bare value is a cell with that value. */
export interface MemoryCell {
  value?: CellValue
  /** Displayed text; by default the value as plain text. */
  text?: string
  /** The formula as displayed, without `=`; with no engine its value is whatever `value` specifies. */
  formula?: string
  /** A style from the workbook's `styles`, by index. */
  style?: number
  /** The color from the number format's section. */
  color?: string
  /** The number format code the cell is displayed with, as xlsx stores it (`0.0%`, `dd.mm.yyyy`); General without it. The text the cell opens with for editing follows it. */
  numberFormat?: string
  link?: { href?: string; location?: string; tooltip?: string; target?: string }
  note?: { author?: string; text: string }
  /** List mode: the cell is a drill-down (`SheetWindow.drill`). */
  drill?: true
  /** The cell's conditional formatting, as xlcalc calculates it (`ConditionalAppearance`); the in-memory workbook keeps it as given. */
  appearance?: ConditionalAppearance
}

/**
 * Excel's outline of an in-memory sheet: outline levels (1 to 7) of rows and columns by row, column or span, keyed as in
 * `widths` and `heights` (`'3:9': 1`, `'C:E': 2`), the summary rows and columns marked collapsed (`'10'`, `'F'`; a group
 * inside a collapsed one stays collapsed when the outer one is expanded), and where the summaries are (below and to the
 * right by default). A group's rows or columns are hidden by a size of 0.
 */
export interface MemoryOutline {
  rows?: Record<string, number>
  columns?: Record<string, number>
  collapsed?: string[]
  summaryBelow?: boolean
  summaryRight?: boolean
}

export interface MemorySheet extends Omit<SheetInfo, 'maxRow' | 'maxCol' | 'outline'> {
  /** Cells by A1 address. */
  cells: Record<string, MemoryCell | CellValue>
  /** Column widths by letters (`B`) or a span (`B:D`), in xlsx units; 0 hides. */
  widths?: Record<string, number>
  /** Row heights by row (`5`) or a span (`5:9`), in points; 0 hides. */
  heights?: Record<string, number>
  defaultRowHeight?: number
  merges?: string[]
  tables?: { name: string; range: string; headerRow: number }[]
  notice?: string
  outline?: MemoryOutline
  /** Data validation rules: a cell's dropdown and input message, and the alert for a typed value that violates a rule. */
  validations?: MemoryValidation[]
}

export interface MemoryBook {
  sheets: MemorySheet[]
  styles?: CellStyle[]
  /** The functions for Formula AutoComplete, by their names as typed in the workbook's formulas, in the order its list displays them; without them the source has no `functionNames`. */
  functions?: string[]
}

export interface MemorySourceOptions {
  /** Milliseconds every response takes, to see the component waiting; 0 by default. */
  delay?: number
  maxCells?: number
}

export interface MemorySource extends SheetSource {
  /** Responds to the next edits with these failed outcomes instead of applying them (for tests and the example). */
  failNext(...outcomes: Extract<EditOutcome, { ok: false }>[]): void
  /** Sends an event to the subscribers, as another tab would. */
  emit(event: SourceEvent): void
  /** Changes cells as another tab would: the changes are applied and sent to the subscribers as `changed`. */
  setCells(sheet: string, entries: Record<string, string | null>): CellUpdate[]
  /** Simulates a workbook recalculating for `ms`: windows return `pending` until then. */
  recompute(ms: number): void
}

/** A data validation rule of an in-memory sheet, as Excel's dialog creates one: a list or whole numbers within bounds, an input message, the alert. */
export interface MemoryValidation {
  /** The cells: A1 ranges separated by spaces (`B2:B20 D2`, `C:C`). */
  range: string
  /** A list: the value must be one of the items, case-sensitively, ignoring leading and trailing spaces; the cell shows them in a dropdown. */
  list?: string[]
  /** Whole numbers from `min` to `max`; omit either bound for none. */
  whole?: { min?: number; max?: number }
  /** The input message displayed while a cell is selected. */
  prompt?: { title?: string; text: string }
  /** The alert for a value that violates the rule: its style (stop by default), title and message (Excel's own text without one). */
  alert?: { severity?: SheetValidationAlert['severity']; title?: string; message?: string }
}
