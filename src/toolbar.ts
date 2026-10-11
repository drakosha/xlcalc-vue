// The formatting toolbar's rules: Excel's buttons as patches of `SheetSource.format` over the selection, toggled by what
// the active cell shows, and the number format it shows by among those offered.
import type { CellStyle, NumberFormat, StylePatch } from 'xlcalc/contract'
import { parseRange, rangeAddress } from 'xlcalc/a1'
import type { FormatPatch, SheetSource } from './source'

/** The kinds of Excel's Number Format list the toolbar offers. */
export type NumberKind = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'time' | 'text'

/**
 * A number format the toolbar offers: what it sets (a code, or a built-in format's number) and the code a cell shows by
 * with it in the book's locale (`SheetViewState.numberFormat`), by which the toolbar knows it on the active cell.
 */
export interface ToolbarNumberFormat {
  kind: NumberKind
  format: NumberFormat
  code: string
}

/** A source the toolbar formats through: `SheetSource`, whose `history` says whether undo and redo have a step. */
export type ToolbarSource = SheetSource

export type FontToggle = 'bold' | 'italic' | 'underline'
export type HorizontalAlign = 'left' | 'center' | 'right'
export type BorderKind = 'all' | 'outside' | 'bottom' | 'none'

/** Excel's palette: the Office theme's colours, then the standard ones. */
export const PALETTE = [
  '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47',
  '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0',
]
/** The colours Excel's split buttons apply before one is picked. */
export const FIRST_FONT_COLOUR = '#FF0000'
export const FIRST_FILL_COLOUR = '#FFFF00'

const THIN = { style: 'thin' } as const

/** Whether the active cell shows bold, italic or an underline, as Excel presses its buttons. */
export function fontOn(style: CellStyle | undefined, what: FontToggle): boolean {
  return !!style?.font?.[what]
}

/** Bold, italic or underline over the selection: off when the active cell shows it, else on, as Excel toggles them. */
export function fontToggle(style: CellStyle | undefined, what: FontToggle): StylePatch {
  const on = fontOn(style, what)
  return { font: { [what]: on ? null : what === 'underline' ? 'single' : true } }
}

/** A horizontal alignment over the selection; pressed again on a cell aligned so, back to General. */
export function alignToggle(style: CellStyle | undefined, h: HorizontalAlign): StylePatch {
  return { alignment: { horizontal: style?.alignment?.horizontal === h ? null : h } }
}

export function wrapToggle(style: CellStyle | undefined): StylePatch {
  return { alignment: { wrap: style?.alignment?.wrap ? null : true } }
}

/**
 * Borders over a selection as Excel draws them: every cell's four sides, the outline (its four edges, one patch each, in
 * one call and so one step of undo), the bottom edge, or none. Nothing for a ref that is no range.
 */
export function borderPatches(ref: string, kind: BorderKind): FormatPatch[] {
  const r = parseRange(ref)
  if (!r) return []
  const { top, left, bottom, right } = r
  const all = rangeAddress({ top, left, bottom, right })
  const rowAt = (row: number) => rangeAddress({ top: row, left, bottom: row, right })
  const colAt = (col: number) => rangeAddress({ top, left: col, bottom, right: col })
  switch (kind) {
    case 'all':
      return [{ range: all, style: { border: { top: THIN, right: THIN, bottom: THIN, left: THIN } } }]
    case 'outside':
      return [
        { range: rowAt(top), style: { border: { top: THIN } } },
        { range: rowAt(bottom), style: { border: { bottom: THIN } } },
        { range: colAt(left), style: { border: { left: THIN } } },
        { range: colAt(right), style: { border: { right: THIN } } },
      ]
    case 'bottom':
      return [{ range: rowAt(bottom), style: { border: { bottom: THIN } } }]
    case 'none':
      return [{ range: all, style: { border: null } }]
  }
}

/**
 * The offered kind the active cell's format code is: General without a code, else the one offered with the same code;
 * null for any other format (Excel's Custom). Which kind a code of its own is would take the library's rule.
 */
export function numberKindOf(code: string | undefined, offered: readonly ToolbarNumberFormat[]): NumberKind | null {
  if (code === undefined || code === 'General') return 'general'
  return offered.find((f) => f.code === code)?.kind ?? null
}
