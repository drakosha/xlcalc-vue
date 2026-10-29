// Excel's text orientation: a cell's rotation from its style, and the box a rotated line of text takes, for the grid to
// align it in its cell as level text is aligned.
import type { CellStyle } from 'xlcalc/contract'
import { fontPx } from './geometry'

/** A line's height for its font size, as the grid's line-height. */
const LINE = 1.3
/** A character's average width for its font size, where the browser cannot measure (no canvas). */
const CHAR = 0.55

/** The text's orientation: degrees from −90 to 90 (up for positive ones), `vertical` for stacked letters; null for level text. */
export function rotationOf(style: CellStyle | undefined): number | 'vertical' | null {
  const r = style?.alignment?.rotation
  if (r === 'vertical') return r
  return typeof r === 'number' && r !== 0 ? Math.max(-90, Math.min(90, r)) : null
}

/** The font a cell draws its text in, for measuring. */
export interface TextFont {
  family: string
  px: number
  bold: boolean
  italic: boolean
}

/** A cell's font: its own name and size, else the workbook's default font, else the interface font at the grid's base size. */
export function textFont(style: CellStyle | undefined, book: { name: string; size: number } | null): TextFont {
  const f = style?.font
  const name = f?.name ?? book?.name
  const size = f?.size ?? book?.size
  return { family: name ? `"${name}", system-ui, sans-serif` : 'system-ui, sans-serif', px: fontPx(size ?? 11), bold: !!f?.bold, italic: !!f?.italic }
}

let context: OffscreenCanvasRenderingContext2D | null | undefined
/** The width of a line of text in pixels: measured on a canvas where there is one, otherwise estimated. */
function lineWidth(text: string, font: TextFont): number {
  if (context === undefined) context = typeof OffscreenCanvas === 'function' ? new OffscreenCanvas(1, 1).getContext('2d') : null
  if (!context) return text.length * font.px * CHAR
  context.font = `${font.italic ? 'italic ' : ''}${font.bold ? '700 ' : ''}${font.px}px ${font.family}`
  return context.measureText(text).width
}

/** A rotated line's bounding box in pixels, and its angle in degrees (counterclockwise for positive ones). */
export interface RotatedBox {
  width: number
  height: number
  angle: number
}

/** The box a line of text takes rotated by `angle` degrees: the grid aligns it in its cell and turns the line inside it. */
export function rotatedBox(text: string, font: TextFont, angle: number): RotatedBox {
  const long = lineWidth(text, font)
  const high = font.px * LINE
  const rad = (Math.abs(angle) * Math.PI) / 180
  const sin = Math.sin(rad)
  const cos = Math.cos(rad)
  // a sine or cosine of an exact right angle is not exactly 0 or 1 in floating point
  const near = (x: number) => Math.round(x * 1e6) / 1e6
  return { width: Math.ceil(near(long * cos + high * sin)), height: Math.ceil(near(long * sin + high * cos)), angle }
}
