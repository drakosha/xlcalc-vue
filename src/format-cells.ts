// Excel's Format Cells dialog without its look: the fields as the active cell's look opens them, what each tab's buttons
// do to them, and the patches of `SheetView.format` that OK sends, only for what the person changed.
import type { AlignmentStyle, BorderLine, CellStyle, FontStyle, FormatChoice, NumberFormat } from 'xlcalc/contract'
import { parseRange } from 'xlcalc/a1'
import type { SheetFormatPatch } from './core/commands'
import type { SheetViewState } from './core/sheet'
import { BORDER_EDGES, edgePatches, edgesOf, sameLine, upperColour, type BorderEdge, type EdgeLines } from './toolbar'

export type FormatTab = 'number' | 'alignment' | 'font' | 'border' | 'fill'
export const FORMAT_TABS: readonly FormatTab[] = ['number', 'alignment', 'font', 'border', 'fill']

type Horizontal = NonNullable<AlignmentStyle['horizontal']>
type Vertical = NonNullable<AlignmentStyle['vertical']>
export const HORIZONTALS: readonly Horizontal[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
export const VERTICALS: readonly Vertical[] = ['top', 'center', 'bottom', 'justify', 'distributed']
/** Excel's indent goes with the left, right and distributed alignments. */
export const INDENTED: ReadonlySet<string> = new Set<Horizontal>(['left', 'right', 'distributed'])
/** The kinds of line Excel's Border tab offers, in the order of its list. */
export const LINE_STYLES: readonly BorderLine['style'][] = ['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double']
type FontStyleKind = 'regular' | 'italic' | 'bold' | 'boldItalic'
type Underline = 'none' | NonNullable<FontStyle['underline']>
/** The font styles and underlines in the order of Excel's lists. */
export const FONT_STYLES: readonly FontStyleKind[] = ['regular', 'italic', 'bold', 'boldItalic']
export const UNDERLINES: readonly Underline[] = ['none', 'single', 'double']
/** Excel's font sizes in its list; any size from 1 to 409 may be typed. */
export const FONT_SIZES: readonly number[] = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
/** Fonts offered in the font field's list; any name may be typed. */
export const FONT_NAMES: readonly string[] = ['Aptos', 'Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Verdana']

/** The dialog's fields; text fields hold what is typed until OK reads them. */
export interface FormatDraft {
  /** The code the cells show by: an item's of Excel's list, or one typed (the box names its category). */
  numberCode: string
  /** What OK sets: an item's format (a built-in one by number), or the code typed; null leaves the format as it is. */
  numberFormat: NumberFormat | null
  horizontal: Horizontal
  vertical: Vertical
  wrap: boolean
  indent: string
  fontName: string
  fontSize: string
  fontStyle: FontStyleKind
  underline: Underline
  strike: boolean
  /** null: Automatic. */
  fontColour: string | null
  /** null: No Color. */
  fill: string | null
  /** The line the Border tab's buttons draw with. */
  line: BorderLine
  /** Edges as the Border tab set them; an edge left out stays as it is. */
  edges: EdgeLines
}

/** The dialog as it opened: the fields, the range it formats and what its edges were (known for one cell, unknown for a range). */
export interface FormatStart {
  draft: FormatDraft
  ref: string
  /** One cell: the inside edges are off and the outline is the cell's own sides. */
  single: boolean
  edges: EdgeLines
}

const fontStyleOf = (f: FontStyle | undefined): FontStyleKind => (f?.bold ? (f.italic ? 'boldItalic' : 'bold') : f?.italic ? 'italic' : 'regular')
const isBold = (k: FontStyleKind) => k === 'bold' || k === 'boldItalic'
const isItalic = (k: FontStyleKind) => k === 'italic' || k === 'boldItalic'
/** The indent typed: a whole number of characters, NaN for what is none. */
const parseIndent = (text: string): number => (/^\d+$/.test(text.trim()) ? Number(text.trim()) : NaN)
/** The font size typed, with a point or a comma: undefined for an empty field (the book's size), NaN for what is no number. */
function parseSize(text: string): number | undefined {
  const t = text.trim().replace(',', '.')
  return t ? Number(t) : undefined
}

/** The fields as the active cell shows them: its style, its number format's code, Excel's thin black line to draw with. */
export function startOf(state: SheetViewState): FormatStart {
  const s: CellStyle = state.style ?? {}
  const r = parseRange(state.ref)
  const single = !r || (r.top === r.bottom && r.left === r.right)
  const own = s.border ?? {}
  const draft: FormatDraft = {
    numberCode: state.numberFormat ?? 'General',
    numberFormat: null,
    horizontal: s.alignment?.horizontal ?? 'general',
    vertical: s.alignment?.vertical ?? 'bottom',
    wrap: !!s.alignment?.wrap,
    indent: String(s.alignment?.indent ?? 0),
    fontName: s.font?.name ?? '',
    fontSize: s.font?.size === undefined ? '' : String(s.font.size),
    fontStyle: fontStyleOf(s.font),
    underline: s.font?.underline ?? 'none',
    strike: !!s.font?.strike,
    fontColour: upperColour(s.font?.color),
    fill: upperColour(s.fill?.color),
    line: { style: 'thin' },
    edges: {},
  }
  // a cell's own sides are its outline; a range's edges are not known from its active cell, as Excel's grey buttons
  const edges: EdgeLines = single ? { top: own.top ?? null, bottom: own.bottom ?? null, left: own.left ?? null, right: own.right ?? null } : {}
  return { draft, ref: state.ref, single, edges }
}

/** The edges the selection has: one cell has no inside lines. */
export const edgesIn = (start: FormatStart): BorderEdge[] => {
  const r = parseRange(start.ref)
  return r ? edgesOf(r) : []
}
/** What an edge shows now: as set in the dialog, else as it was; undefined where that is not known. */
export const shownEdge = (start: FormatStart, draft: FormatDraft, e: BorderEdge): BorderLine | null | undefined =>
  draft.edges[e] !== undefined ? draft.edges[e] : start.edges[e]

/** An edge's button, as Excel's: the line on it, or off when it already has that line. */
export function pressEdge(start: FormatStart, draft: FormatDraft, e: BorderEdge): void {
  draft.edges = { ...draft.edges, [e]: sameLine(shownEdge(start, draft, e), draft.line) ? null : { ...draft.line } }
}
/** Excel's presets: None takes every line off, Outline and Inside draw the line round the selection or inside it. */
export function pressPreset(start: FormatStart, draft: FormatDraft, preset: 'none' | 'outline' | 'inside'): void {
  const next: EdgeLines = { ...draft.edges }
  for (const e of edgesIn(start)) {
    const inside = e === 'insideH' || e === 'insideV'
    if (preset === 'none') next[e] = null
    else if ((preset === 'inside') === inside) next[e] = { ...draft.line }
  }
  draft.edges = next
}

/** A field OK cannot read, by its name: the code, the indent, the size. */
interface DraftErrors {
  numberCode?: true
  indent?: true
  fontSize?: true
}
/** Excel's limits: a font of 1 to 409 points, an indent of 0 to 250, a code of the cell's own not empty. */
export function draftErrors(draft: FormatDraft): DraftErrors {
  const out: DraftErrors = {}
  if (!draft.numberCode.trim()) out.numberCode = true
  if (!(parseIndent(draft.indent) <= 250)) out.indent = true
  const size = parseSize(draft.fontSize)
  if (size !== undefined && !(size >= 1 && size <= 409)) out.fontSize = true
  return out
}

/** An item of Excel's list: its format, and its code in the field to be changed from there, as Excel's Custom starts. */
export function pickNumber(draft: FormatDraft, item: FormatChoice): void {
  draft.numberCode = item.code
  draft.numberFormat = item.format
}
/** A code typed into the field: OK sets it as written. */
export function typeNumber(draft: FormatDraft, code: string): void {
  draft.numberCode = code
  draft.numberFormat = code.trim()
}

/** The font size typed, to Excel's half points; undefined for an empty field (the book's size). */
function sizeOf(text: string): number | undefined {
  const size = parseSize(text)
  return size === undefined ? undefined : Math.round(size * 2) / 2
}

/** An alignment chosen in the list; one without an indent takes it back to 0, as Excel's field greys out. */
export function setHorizontal(draft: FormatDraft, h: string): void {
  const next = HORIZONTALS.find((x) => x === h)
  if (!next) return
  draft.horizontal = next
  if (!INDENTED.has(next)) draft.indent = '0'
}

/**
 * OK's patches: one over the selection with what changed of the font, the alignment, the fill and the number format, and
 * the Border tab's edges on their ranges, all for one call and so one step of undo; none when nothing changed. A field
 * back at its default takes the cell's own off (null), as Excel stores no default.
 */
export function patchesOf(start: FormatStart, draft: FormatDraft): SheetFormatPatch[] {
  const was = start.draft
  const font: Record<string, unknown> = {}
  const name = draft.fontName.trim()
  if (name !== was.fontName) font.name = name || null
  const size = sizeOf(draft.fontSize)
  if (size !== sizeOf(was.fontSize)) font.size = size ?? null
  if (isBold(draft.fontStyle) !== isBold(was.fontStyle)) font.bold = isBold(draft.fontStyle) || null
  if (isItalic(draft.fontStyle) !== isItalic(was.fontStyle)) font.italic = isItalic(draft.fontStyle) || null
  if (draft.underline !== was.underline) font.underline = draft.underline === 'none' ? null : draft.underline
  if (draft.strike !== was.strike) font.strike = draft.strike || null
  if (draft.fontColour !== was.fontColour) font.color = draft.fontColour
  const alignment: Record<string, unknown> = {}
  if (draft.horizontal !== was.horizontal) alignment.horizontal = draft.horizontal === 'general' ? null : draft.horizontal
  if (draft.vertical !== was.vertical) alignment.vertical = draft.vertical === 'bottom' ? null : draft.vertical
  if (draft.wrap !== was.wrap) alignment.wrap = draft.wrap || null
  const indent = parseIndent(draft.indent)
  if (indent !== parseIndent(was.indent)) alignment.indent = indent || null
  const style: Record<string, unknown> = {}
  if (Object.keys(font).length) style.font = font
  if (Object.keys(alignment).length) style.alignment = alignment
  if (draft.fill !== was.fill) style.fill = draft.fill ? { color: draft.fill } : null
  const out: SheetFormatPatch[] = []
  const own: SheetFormatPatch = {}
  if (Object.keys(style).length) own.style = style
  if (draft.numberFormat !== null && draft.numberCode.trim() !== was.numberCode) own.numberFormat = draft.numberFormat
  if (own.style || own.numberFormat !== undefined) out.push(own)
  // only edges the person set to something else than they were; an unknown edge set at all
  const lines: EdgeLines = {}
  for (const e of BORDER_EDGES) {
    const now = draft.edges[e]
    if (now !== undefined && (start.edges[e] === undefined || !sameLine(now, start.edges[e]))) lines[e] = now
  }
  if (Object.keys(lines).length) out.push(...edgePatches(start.ref, lines))
  return out
}
