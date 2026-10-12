// The formatting toolbar's rules: Excel's buttons as patches of `SheetSource.format` over the selection, toggled by what
// the active cell shows, and the number format it shows by among those offered.
import type { BorderLine, CellRect, CellStyle, NumberFormat, StylePatch } from 'xlcalc/contract'
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

/** The edges of a selection Excel's Border tab draws on: its outline and the lines inside it, across and down. */
export type BorderEdge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV'
export const BORDER_EDGES: readonly BorderEdge[] = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']
/** Lines on edges: a line, or null to take the edge's line off. */
export type EdgeLines = Partial<Record<BorderEdge, BorderLine | null>>

/** Whether two lines are the same: their kind and colour. */
export const sameLine = (a: BorderLine | null | undefined, b: BorderLine | null | undefined): boolean =>
  !a || !b ? a === b : a.style === b.style && (a.color?.toUpperCase() ?? null) === (b.color?.toUpperCase() ?? null)

/** The edges a range has: the inside lines only across more than one row or down more than one column. */
export function edgesOf(r: CellRect): BorderEdge[] {
  return BORDER_EDGES.filter((e) => (e === 'insideH' ? r.bottom > r.top : e === 'insideV' ? r.right > r.left : true))
}

/**
 * Lines on the edges of a selection as Excel draws them: an outline edge on its row or column of cells, an inside line on
 * both cells' sides it lies between, each a patch in one call and so one step of undo; one patch over the whole range
 * when every edge it has takes the same line, or none. Nothing for a ref that is no range.
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
 * Borders over a selection as the toolbar's menu draws them: every cell's four sides, the outline (its four edges, one
 * patch each, in one call and so one step of undo), the bottom edge, or none. Nothing for a ref that is no range.
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

/**
 * The offered kind the active cell's format code is: General without a code, else the one offered with the same code;
 * null for any other format (Excel's Custom). Which kind a code of its own is would take the library's rule.
 */
export function numberKindOf(code: string | undefined, offered: readonly ToolbarNumberFormat[]): NumberKind | null {
  if (code === undefined || code === 'General') return 'general'
  return offered.find((f) => f.code === code)?.kind ?? null
}
