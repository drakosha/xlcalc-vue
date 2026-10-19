// The formatting toolbar's rules: Excel's buttons as `SheetSource.format` patches over the selection, toggled according to
// the active cell's formatting, and the borders that the toolbar and the Format Cells dialog draw.
import type { BorderLine, CellRect, CellStyle, StylePatch } from 'xlcalc/contract'
import { parseRange, rangeAddress } from 'xlcalc/a1'
import type { FormatPatch, SheetSource } from './contract/source'

export type FontToggle = 'bold' | 'italic' | 'underline'
export type HorizontalAlign = 'left' | 'center' | 'right'
export type BorderKind = 'all' | 'outside' | 'bottom' | 'none'

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
