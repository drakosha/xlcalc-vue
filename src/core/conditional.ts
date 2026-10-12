// What conditional formats give a cell, drawn as Excel draws it: a rule's style over the cell's, a data bar as background layers.
import type { BarLook, CellStyle, ConditionalStyle } from 'xlcalc/contract'

/** A background layer of a cell: its image, size and position. */
export interface Layer {
  img: string
  size: string
  pos: string
}

/** A rule's style over the style the cell shows, as Excel lays a differential format: each field the rule sets replaces the cell's. */
export function lookStyle(style: CellStyle | undefined, rule: ConditionalStyle | undefined): CellStyle | undefined {
  if (!rule) return style
  const out: CellStyle = { ...style }
  if (rule.font) out.font = { ...style?.font, ...rule.font }
  if (rule.fill) out.fill = rule.fill
  if (rule.border) out.border = { ...style?.border, ...rule.border }
  return out
}

/** A share of the bar's room as a CSS percentage, rounded so the text stays short. */
const pct = (f: number) => `${Math.round(Math.min(1, Math.max(0, f)) * 100_000) / 1000}%`
/** Where a gradient fill fades to at the bar's far end: Excel's is close to white (measured in Excel for Mac 16). */
const pale = (color: string) => `color-mix(in srgb, ${color} 18%, white)`

/**
 * A data bar as background layers, top first: the axis, the fill, the border under it. Excel draws a bar 1 px inside the
 * cell's sides, 1 px below its top and 3 px above its bottom line (1 and 2 points at 15 points high); a gradient is the
 * bar's colour at its axis (or start) fading toward white, the border a 1 px line round the bar, the axis a dashed 1 px
 * line across the whole cell.
 */
export function barLayers(bar: BarLook): Layer[] {
  const out: Layer[] = []
  const dir = bar.rightToLeft ? 'to left' : 'to right'
  if (bar.axis !== undefined) {
    // The 1 px image's left edge at 1px + f × (width − 2px); a percentage places it by (width − 1px).
    const f = bar.rightToLeft ? 1 - bar.axis : bar.axis
    out.push({ img: `repeating-linear-gradient(to bottom, ${bar.axisColor ?? '#000000'} 0px 3px, transparent 3px 4px)`, size: '1px 100%', pos: `calc(${pct(f)} + ${Math.round((1 - f) * 100) / 100}px) 0px` })
  }
  if (!(bar.end > bar.start)) return out
  const s = pct(bar.start)
  const e = pct(bar.end)
  // The dark side is at the axis: the end of a negative bar, the start of every other.
  const atEnd = bar.axis !== undefined && Math.abs(bar.end - bar.axis) < Math.abs(bar.start - bar.axis)
  const [from, to] = !bar.gradient ? [bar.color, bar.color] : atEnd ? [pale(bar.color), bar.color] : [bar.color, pale(bar.color)]
  if (!bar.border) {
    out.push({ img: `linear-gradient(${dir}, transparent ${s}, ${from} ${s}, ${to} ${e}, transparent ${e})`, size: 'calc(100% - 2px) calc(100% - 4px)', pos: 'left 1px top 1px' })
    return out
  }
  out.push(
    { img: `linear-gradient(${dir}, transparent calc(${s} + 1px), ${from} calc(${s} + 1px), ${to} calc(${e} - 1px), transparent calc(${e} - 1px))`, size: 'calc(100% - 2px) calc(100% - 6px)', pos: 'left 1px top 2px' },
    { img: `linear-gradient(${dir}, transparent ${s}, ${bar.border} ${s}, ${bar.border} ${e}, transparent ${e})`, size: 'calc(100% - 2px) calc(100% - 4px)', pos: 'left 1px top 1px' },
  )
  return out
}
