// A book's cell style (CellStyle) as CSS: font, fill, borders and alignment, Excel's rules where CSS has no word for them.
import type { BorderLine, CellStyle, CellValue } from 'xlcalc/contract'
import { fontPx } from './geometry'

export type Horizontal = 'left' | 'center' | 'right' | 'fill' | 'justify' | 'distributed' | 'centerContinuous'
export type Vertical = 'top' | 'center' | 'bottom'

/** Pixels of one indent level (Excel's step at the default font). */
export const INDENT_PX = 9

/** Excel's General alignment by the value: numbers right, text left, booleans and errors centred. */
export function generalAlign(value: CellValue | undefined): 'left' | 'center' | 'right' {
  if (typeof value === 'number') return 'right'
  if (typeof value === 'boolean' || (value !== null && typeof value === 'object')) return 'center'
  return 'left'
}

/** The horizontal alignment a cell shows. */
export function horizontalOf(style: CellStyle | undefined, value: CellValue | undefined): Horizontal {
  const h = style?.alignment?.horizontal
  return !h || h === 'general' ? generalAlign(value) : h
}

/** The vertical alignment a cell shows; Excel's default is the bottom. */
export function verticalOf(style: CellStyle | undefined): Vertical {
  const v = style?.alignment?.vertical
  if (v === 'top' || v === 'justify') return 'top'
  if (v === 'center' || v === 'distributed') return 'center'
  return 'bottom'
}

/** Whether the text wraps: by the style, and always for justify and distributed, as Excel does. */
export function wrapsOf(style: CellStyle | undefined): boolean {
  const a = style?.alignment
  return !!a && (a.wrap === true || a.horizontal === 'justify' || a.horizontal === 'distributed' || a.vertical === 'justify' || a.vertical === 'distributed')
}

/** Line width in pixels of an xlsx border kind. */
export function borderWidth(line: BorderLine | undefined): number {
  if (!line) return 0
  switch (line.style) {
    case 'thick':
    case 'double':
      return 3
    case 'medium':
    case 'mediumDashed':
    case 'mediumDashDot':
    case 'mediumDashDotDot':
    case 'slantDashDot':
      return 2
    default:
      return 1
  }
}

/** An xlsx border line as a CSS border; CSS has no dash-dot, so those are dashed. Without a colour: Excel's automatic. */
export function borderCss(line: BorderLine | undefined): string | undefined {
  if (!line) return undefined
  const color = line.color ?? 'var(--xs-border-auto)'
  const width = borderWidth(line)
  switch (line.style) {
    case 'double':
      return `3px double ${color}`
    case 'dotted':
    case 'hair':
      return `1px dotted ${color}`
    case 'thin':
    case 'medium':
    case 'thick':
      return `${width}px solid ${color}`
    default:
      return `${width}px dashed ${color}`
  }
}

/** Font, colour and fill of a cell; `section` is the colour of the number format's section, which wins over the font's. */
export function fontCss(style: CellStyle | undefined, section: string | undefined): Record<string, string> {
  const css: Record<string, string> = {}
  const f = style?.font
  if (f?.name) css.fontFamily = `"${f.name.replace(/["\\]/g, '')}", var(--xs-font)`
  if (f?.size) css.fontSize = `${fontPx(f.size)}px`
  if (f?.bold) css.fontWeight = '700'
  if (f?.italic) css.fontStyle = 'italic'
  const lines = [f?.underline ? 'underline' : '', f?.strike ? 'line-through' : ''].filter(Boolean)
  if (lines.length) css.textDecorationLine = lines.join(' ')
  if (f?.underline === 'double') css.textDecorationStyle = 'double'
  const color = section ?? f?.color
  if (color) css.color = color
  if (style?.fill?.color) css.backgroundColor = style.fill.color
  return css
}

/** Content box alignment: CSS for the box that holds the text, by horizontal and vertical alignment and indent. */
export function alignCss(h: Horizontal, indent: number): Record<string, string> {
  const css: Record<string, string> = {}
  const pad = indent > 0 ? `${indent * INDENT_PX}px` : ''
  switch (h) {
    case 'right':
      css.textAlign = 'right'
      if (pad) css.paddingRight = pad
      break
    case 'center':
    case 'centerContinuous':
      css.textAlign = 'center'
      break
    case 'justify':
      css.textAlign = 'justify'
      if (pad) css.paddingLeft = pad
      break
    case 'distributed':
      css.textAlign = 'justify'
      css.textAlignLast = 'justify'
      if (pad) css.paddingLeft = css.paddingRight = pad
      break
    default:
      css.textAlign = 'left'
      if (pad) css.paddingLeft = pad
  }
  return css
}

/** Excel's Fill alignment repeats the text across the cell; enough copies to cover its width, the rest is clipped. */
export function fillText(text: string, widthPx: number): string {
  if (!text) return ''
  const copies = Math.min(500, Math.ceil(widthPx / Math.max(4, text.length * 5)) + 1)
  return text.repeat(copies)
}

/** Whether a value's text may overflow into empty neighbours: only text does, as in Excel. */
export const overflows = (value: CellValue | undefined): boolean => typeof value === 'string'
