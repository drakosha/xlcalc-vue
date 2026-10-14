// The toolbar's icons as inline SVGs in the look of icons.ts; the colour buttons draw their colour as a bar under the glyph.
import { h, type FunctionalComponent } from 'vue'

const ATTRS = { class: 'xs-icon', viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' }

const icon = (paths: string[], extra: Record<string, unknown> = {}): FunctionalComponent => {
  const c: FunctionalComponent = () => h('svg', { ...ATTRS, ...extra }, paths.map((d) => h('path', { d })))
  return c
}

/** A glyph over a bar of `colour` (none: an outline), as Excel's Font Color and Fill Color buttons show the colour they apply. */
const colourIcon = (paths: string[]): FunctionalComponent<{ colour?: string | null }> => {
  const c: FunctionalComponent<{ colour?: string | null }> = (props) =>
    h('svg', ATTRS, [
      ...paths.map((d) => h('path', { d })),
      h('rect', { x: 3, y: 19, width: 18, height: 3.5, rx: 0.5, fill: props.colour ?? 'none', stroke: props.colour ? 'none' : 'currentColor', 'stroke-width': 1 }),
    ])
  c.props = ['colour']
  return c
}

export const IconUndo = icon(['M9 14 4 9l5-5', 'M4 9h10.5a5.5 5.5 0 0 1 0 11H11'])
export const IconRedo = icon(['m15 14 5-5-5-5', 'M20 9H9.5a5.5 5.5 0 0 0 0 11H13'])
export const IconBold = icon(['M7 4h6a3.5 3.5 0 0 1 0 7H7z', 'M7 11h7a3.5 3.5 0 0 1 0 7H7z'], { 'stroke-width': 2.6 })
export const IconItalic = icon(['M19 4h-9', 'M14 20H5', 'M15 4 9 20'])
export const IconUnderline = icon(['M6 4v6a6 6 0 0 0 12 0V4', 'M4 21h16'])
export const IconFontColour = colourIcon(['m6 16 6-13 6 13', 'M8.3 11h7.4'])
export const IconFillColour = colourIcon(['m17 9-6-6-7 7 6 6z', 'M4.5 10.5h12', 'M20 11.5s1.5 1.8 1.5 2.8a1.5 1.5 0 0 1-3 0c0-1 1.5-2.8 1.5-2.8z'])
export const IconBorderAll = icon(['M3 3h18v18H3z', 'M3 12h18', 'M12 3v18'])
export const IconBorderOutside = icon(['M3 3h18v18H3z', 'M12 8v.01', 'M12 12v.01', 'M12 16v.01', 'M8 12h.01', 'M16 12h.01'])
export const IconBorderBottom = icon(['M3 21h18', 'M3 3v.01', 'M3 7.5v.01', 'M3 12v.01', 'M3 16.5v.01', 'M21 3v.01', 'M21 7.5v.01', 'M21 12v.01', 'M21 16.5v.01', 'M7.5 3h.01', 'M12 3h.01', 'M16.5 3h.01'])
export const IconBorderNone = icon(['M3 3h18v18H3z'], { 'stroke-dasharray': '2 3' })
export const IconAlignLeft = icon(['M3 6h18', 'M3 12h12', 'M3 18h15'])
export const IconAlignCenter = icon(['M3 6h18', 'M7 12h10', 'M5 18h14'])
export const IconAlignRight = icon(['M3 6h18', 'M9 12h12', 'M6 18h15'])
export const IconWrap = icon(['M3 6h18', 'M3 12h15a3 3 0 1 1 0 6h-4', 'm16 16-2 2 2 2', 'M3 18h7'])
export const IconMore = icon(['M5 12h.01', 'M12 12h.01', 'M19 12h.01'], { 'stroke-width': 3 })
export const IconFormatCells = icon(['M3 4h18v16H3z', 'M3 10h18', 'M9 10v10', 'M13 14h5', 'M13 17h3'])
/** Merge & Center: a cell's box over the cells it takes in, with the arrow both ways. */
export const IconMerge = icon(['M3 5h18v14H3z', 'M7 12h10', 'm9 9.5-2.5 2.5L9 14.5', 'm15 9.5 2.5 2.5-2.5 2.5'])
