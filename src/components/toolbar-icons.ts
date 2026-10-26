// The toolbar's icons as inline SVGs in the style of icons.ts; the color buttons draw their color as a bar under the glyph.
import { h, type FunctionalComponent } from 'vue'
import { ICON_ATTRS, icon } from './icons'

/** A glyph over a bar of `colour` (none: an outline), as Excel's Font Color and Fill Color buttons show the color they apply. */
const colourIcon = (paths: string[]): FunctionalComponent<{ colour?: string | null }> => {
  const c: FunctionalComponent<{ colour?: string | null }> = (props) =>
    h('svg', ICON_ATTRS, [
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
/** Merge & Center: a cell's box over the cells it covers, with a two-way arrow. */
export const IconMerge = icon(['M3 5h18v14H3z', 'M7 12h10', 'm9 9.5-2.5 2.5L9 14.5', 'm15 9.5 2.5 2.5-2.5 2.5'])
/** Increase and Decrease Font Size: a letter A with a caret up or down beside it, as Excel's buttons draw them. */
export const IconFontUp = icon(['m3 20 5.5-14 5.5 14', 'M5.2 15h6.6', 'm16 9 2.5-3 2.5 3'])
export const IconFontDown = icon(['m3 20 5.5-14 5.5 14', 'M5.2 15h6.6', 'm16 6 2.5 3 2.5-3'])
/** Top, Middle and Bottom Align: an arrow to a line at the top, to the middle from both sides, to a line at the bottom. */
export const IconAlignTop = icon(['M4 4h16', 'M12 20V9', 'm8 13 4-4 4 4'])
export const IconAlignMiddle = icon(['M4 12h16', 'M12 2v6', 'm9 5 3 3 3-3', 'M12 22v-6', 'm9 19 3-3 3 3'])
export const IconAlignBottom = icon(['M4 20h16', 'M12 4v11', 'm8 11 4 4 4-4'])
/** Decrease and Increase Indent: lines of text with an arrow to the left or to the right before the middle ones. */
export const IconIndentDecrease = icon(['M3 5h18', 'M11 10h10', 'M11 14h10', 'M3 19h18', 'm7 9-3 3 3 3'])
export const IconIndentIncrease = icon(['M3 5h18', 'M11 10h10', 'M11 14h10', 'M3 19h18', 'm4 9 3 3-3 3'])
/** Percent Style: the percent sign. */
// Accounting: a banknote; Comma Style: a thousands comma; Increase and Decrease Decimal: a decimal point with an arrow toward more or fewer zeros.
export const IconAccounting = icon(['M2 6h20v12H2z', 'M12 9.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 1 0 0-5', 'M5.5 9v.01', 'M18.5 15v.01'])
export const IconComma = icon(['M7 15.5v1.5a2 2 0 0 1-2 2', 'M12 15.5v1.5a2 2 0 0 1-2 2', 'M17 15.5v1.5a2 2 0 0 1-2 2'], { 'stroke-width': 2.4 })
export const IconDecimalUp = icon(['M3 17v.01', 'M7.5 8a2.5 4.5 0 1 0 0 9 2.5 4.5 0 1 0 0-9', 'M14 10.5h7', 'm18 7.5 3 3-3 3'], { 'stroke-width': 1.8 })
export const IconDecimalDown = icon(['M3 17v.01', 'M7.5 8a2.5 4.5 0 1 0 0 9 2.5 4.5 0 1 0 0-9', 'M21 10.5h-7', 'm17 7.5-3 3 3 3'], { 'stroke-width': 1.8 })
export const IconPercent = icon(['M19 5 5 19', 'M6.5 4a2.5 2.5 0 1 0 0 5 2.5 2.5 0 1 0 0-5', 'M17.5 15a2.5 2.5 0 1 0 0 5 2.5 2.5 0 1 0 0-5'])
export const IconAutoSum = icon(['M18 5V4H6l7 8-7 8h12v-1'])
