// The formatting toolbar's rules: Excel's buttons as `SheetSource.format` patches over the selection, toggled according to
// the active cell's formatting, and the borders that the toolbar and the Format Cells dialog draw.
import type { BorderLine, CellRect, CellStyle, NumberFormat, StylePatch } from 'xlcalc/contract'
import { parseRange, rangeAddress } from 'xlcalc/a1'
import { stepDecimals } from 'xlcalc/text'
import type { FormatPatch, SheetSource } from './contract/source'
import type { SheetViewState, ToolbarGroup } from './contract/view'

/** The toolbar's groups in their default order, as `SheetToolbar`'s `items` takes them: a host spreads it to add its buttons. */
export const TOOLBAR_GROUPS: readonly ToolbarGroup[] = ['undo', 'fontStyle', 'font', 'number', 'colors', 'borders', 'verticalAlign', 'align', 'merge', 'formatCells', 'autoSum']

export type FontToggle = 'bold' | 'italic' | 'underline'
export type HorizontalAlign = 'left' | 'center' | 'right'
export type VerticalAlign = 'top' | 'center' | 'bottom'
export type BorderKind = 'all' | 'outside' | 'bottom' | 'none'

/** Excel's font sizes in its list; any size from 1 to 409 may be typed. */
export const FONT_SIZES: readonly number[] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
/** Fonts offered in the lists of the font inputs, Excel's common ones (a browser does not tell which are installed); any name may be typed. */
export const FONT_NAMES: readonly string[] = ['Aptos', 'Aptos Narrow', 'Arial', 'Calibri', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana']
/** Excel's font size limits, in points. */
const MIN_SIZE = 1
const MAX_SIZE = 409
/** The size the grid draws a cell without its own size at (11 pt, Excel's default), from which the size buttons step. */
const GRID_SIZE = 11
/** Excel's indent applies with the left, right and distributed alignments; its levels go up to 250. */
export const INDENTED: ReadonlySet<string> = new Set(['left', 'right', 'distributed'])
const MAX_INDENT = 250
/** Excel's Percent Style and its Ctrl+Shift+%: the built-in format 9, `0%`. */
export const PERCENT_STYLE = 9
/** Excel's Accounting Number Format and Comma Style: the built-in formats 44 and 43, displayed in the workbook's locale (its currency for 44). */
export const ACCOUNTING_STYLE = 44
export const COMMA_STYLE = 43

/**
 * Excel's Increase Decimal and Decrease Decimal from the active cell, as xlcalc's `stepDecimals` gives them (its format,
 * value and column width; Excel puts the result on the whole selection); null when the step changes nothing.
 */
export function decimalStep(state: Pick<SheetViewState, 'numberFormat' | 'cellValue' | 'columnWidth'>, up: boolean, locale: string | undefined): NumberFormat | null {
  const got = stepDecimals(state.numberFormat, state.cellValue ?? null, up ? 1 : -1, { ...(locale ? { locale } : {}), ...(state.columnWidth !== undefined ? { width: state.columnWidth } : {}) })
  return got === undefined || got === state.numberFormat ? null : got
}

/** The workbook's default font, which a cell without its own font name or size displays. */
export interface BookFont {
  name?: string
  size?: number
}
/**
 * The font the Font and Font Size boxes show for the active cell: its own, otherwise the workbook's (`book`, the sheet
 * state's `font`); without either the boxes are empty.
 */
export function shownFont(style: CellStyle | undefined, book: BookFont = {}): { name: string; size: number | undefined } {
  return { name: style?.font?.name ?? book.name ?? '', size: style?.font?.size ?? book.size }
}

/** A typed font size, with a decimal point or comma: undefined for an empty input, NaN for anything that is not a number. */
export function parseFontSize(text: string): number | undefined {
  const t = text.trim().replace(',', '.')
  return t ? Number(t) : undefined
}
/** A typed font size as Excel takes it: rounded to half points, null outside 1 to 409 or for anything that is not a number. */
export function fontSizeOf(text: string): number | null {
  const size = parseFontSize(text)
  return size !== undefined && size >= MIN_SIZE && size <= MAX_SIZE ? Math.round(size * 2) / 2 : null
}

/**
 * Excel's Increase and Decrease Font Size from a size: the next size of its list, beyond 72 in steps of 10 up to 409, and
 * below 8 a point at a time down to 1, as Office's Grow Font and Shrink Font step; the same size at either limit.
 */
export function stepFontSize(size: number, up: boolean): number {
  const top = FONT_SIZES[FONT_SIZES.length - 1]!
  const bottom = FONT_SIZES[0]!
  if (up) return FONT_SIZES.find((s) => s > size) ?? Math.min(MAX_SIZE, Math.floor(size / 10) * 10 + 10)
  if (size > top) return Math.max(top, Math.ceil(size / 10) * 10 - 10)
  return [...FONT_SIZES].reverse().find((s) => s < size) ?? Math.max(MIN_SIZE, Math.min(bottom, Math.ceil(size)) - 1)
}
/** The size buttons over the selection: one step from the active cell's size; nothing at the limit. */
export function fontSizeStep(style: CellStyle | undefined, up: boolean, book: BookFont = {}): StylePatch | null {
  const size = shownFont(style, book).size ?? GRID_SIZE
  const next = stepFontSize(size, up)
  return next === size ? null : { font: { size: next } }
}

/** Whether the active cell is aligned so vertically, as Excel shows Top, Middle and Bottom Align pressed: bottom without an alignment of its own. */
export function verticalOn(style: CellStyle | undefined, v: VerticalAlign): boolean {
  return (style?.alignment?.vertical ?? 'bottom') === v
}
/** A vertical alignment over the selection; clicked again on a cell already aligned so, it returns to Bottom, Excel's default, which a cell does not store. */
export function verticalToggle(style: CellStyle | undefined, v: VerticalAlign): StylePatch {
  return { alignment: { vertical: v === 'bottom' || verticalOn(style, v) ? null : v } }
}

/**
 * Excel's Increase and Decrease Indent over the selection, a level from the active cell's: Increase keeps the left, right
 * and distributed alignments and turns any other into Left, which has an indent; Decrease takes a level off without changing
 * the alignment and does nothing at 0 or with an alignment that has no indent. Nothing at either limit.
 */
export function indentStep(style: CellStyle | undefined, up: boolean): StylePatch | null {
  const h = style?.alignment?.horizontal ?? 'general'
  const level = INDENTED.has(h) ? (style?.alignment?.indent ?? 0) : 0
  if (up) return level >= MAX_INDENT ? null : { alignment: INDENTED.has(h) ? { indent: level + 1 } : { horizontal: 'left', indent: level + 1 } }
  return level > 0 ? { alignment: { indent: level - 1 || null } } : null
}

/** Excel's palette: the Office theme colors, then the standard ones. */
export const PALETTE = [
  '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47',
  '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0',
]
/** The colors Excel's split buttons apply before one is picked. */
export const FIRST_FONT_COLOUR = '#FF0000'
export const FIRST_FILL_COLOUR = '#FFFF00'

const THIN = { style: 'thin' } as const

/** The edges of a selection that Excel's Border tab draws on: its outline and the lines inside it, horizontal and vertical. */
export type BorderEdge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV'
export const BORDER_EDGES: readonly BorderEdge[] = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']
/** Lines on edges: a line, or null to remove the edge's line. */
export type EdgeLines = Partial<Record<BorderEdge, BorderLine | null>>

/** A color as the palette and the pickers compare it: `#RRGGBB` in uppercase, null for none. */
export const upperColour = (c: string | null | undefined): string | null => c?.toUpperCase() ?? null

/** Whether two lines are the same: their style and color. */
export const sameLine = (a: BorderLine | null | undefined, b: BorderLine | null | undefined): boolean =>
  !a || !b ? a === b : a.style === b.style && upperColour(a.color) === upperColour(b.color)

/** The edges a range has: the inside lines only across more than one row or more than one column. */
export function edgesOf(r: CellRect): BorderEdge[] {
  return BORDER_EDGES.filter((e) => (e === 'insideH' ? r.bottom > r.top : e === 'insideV' ? r.right > r.left : true))
}

/**
 * Lines on the edges of a selection as Excel draws them: an outline edge on its row or column of cells, an inside line on
 * the sides of both cells it lies between, each as a patch in one call and so a single undo step; one patch over the entire
 * range when every edge it has gets the same line, or none. Nothing for a reference that is not a range.
 */
export function edgePatches(ref: string, lines: EdgeLines): FormatPatch[] {
  const r = parseRange(ref)
  if (!r) return []
  const { top, left, bottom, right } = r
  const range = (t: number, l: number, b: number, rt: number) => rangeAddress({ top: t, left: l, bottom: b, right: rt })
  const edges = edgesOf(r)
  const first = lines[edges[0]!]
  if (first !== undefined && edges.every((e) => lines[e] !== undefined && sameLine(lines[e], first))) {
    const style = first ? { border: { top: first, right: first, bottom: first, left: first } } : { border: null }
    return [{ range: range(top, left, bottom, right), style }]
  }
  const out: FormatPatch[] = []
  const put = (rng: string, side: 'top' | 'bottom' | 'left' | 'right', line: BorderLine | null) => out.push({ range: rng, style: { border: { [side]: line } } })
  for (const e of edges) {
    const line = lines[e]
    if (line === undefined) continue
    if (e === 'top') put(range(top, left, top, right), 'top', line)
    else if (e === 'bottom') put(range(bottom, left, bottom, right), 'bottom', line)
    else if (e === 'left') put(range(top, left, bottom, left), 'left', line)
    else if (e === 'right') put(range(top, right, bottom, right), 'right', line)
    else if (e === 'insideH') {
      put(range(top, left, bottom - 1, right), 'bottom', line)
      put(range(top + 1, left, bottom, right), 'top', line)
    } else {
      put(range(top, left, bottom, right - 1), 'right', line)
      put(range(top, left + 1, bottom, right), 'left', line)
    }
  }
  return out
}

/** Whether the active cell is bold, italic or underlined, as Excel shows its buttons pressed. */
export function fontOn(style: CellStyle | undefined, what: FontToggle): boolean {
  return !!style?.font?.[what]
}

/** Bold, italic or underline over the selection: off when the active cell has it, otherwise on, as Excel toggles them. */
export function fontToggle(style: CellStyle | undefined, what: FontToggle): StylePatch {
  const on = fontOn(style, what)
  return { font: { [what]: on ? null : what === 'underline' ? 'single' : true } }
}

/** A horizontal alignment over the selection; clicked again on a cell already aligned so, it returns to General. */
export function alignToggle(style: CellStyle | undefined, h: HorizontalAlign): StylePatch {
  return { alignment: { horizontal: style?.alignment?.horizontal === h ? null : h } }
}

export function wrapToggle(style: CellStyle | undefined): StylePatch {
  return { alignment: { wrap: style?.alignment?.wrap ? null : true } }
}

/**
 * Borders over a selection as the toolbar's menu draws them: every cell's four sides, the outline (its four edges, one
 * patch each, in one call and so a single undo step), the bottom edge, or none. Nothing for a reference that is not a range.
 */
export function borderPatches(ref: string, kind: BorderKind): FormatPatch[] {
  const every = (line: BorderLine | null): EdgeLines => Object.fromEntries(BORDER_EDGES.map((e) => [e, line]))
  switch (kind) {
    case 'all':
      return edgePatches(ref, every(THIN))
    case 'outside':
      return edgePatches(ref, { top: THIN, bottom: THIN, left: THIN, right: THIN })
    case 'bottom':
      return edgePatches(ref, { bottom: THIN })
    case 'none':
      return edgePatches(ref, every(null))
  }
}
