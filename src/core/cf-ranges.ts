// The cells of conditional formatting rules: the rectangle a rule's formulas are written for, Clear Rules from Selected Cells
// cutting the selection out of the rules' ranges, and the "Applies to" box's text as Excel's Rules Manager shows and takes it.
import { colLetters, MAX_COLS, MAX_ROWS, parseRange, rangeAddress, sameName } from 'xlcalc/a1'
import type { CellRect, ConditionalFormat } from 'xlcalc/contract'

/** The rectangles of a rule's ranges; a range that does not parse is left out. */
export const rectsOf = (ranges: readonly string[]): CellRect[] =>
  ranges.flatMap((r) => {
    const rect = parseRange(r)
    return rect ? [{ top: rect.top, left: rect.left, bottom: rect.bottom, right: rect.right }] : []
  })

/** The top-left cell of the rectangle enclosing the ranges, which a rule's formulas are written for; null without ranges. */
export function topLeftOf(ranges: readonly string[]): { row: number; col: number } | null {
  const rects = rectsOf(ranges)
  if (!rects.length) return null
  return { row: Math.min(...rects.map((r) => r.top)), col: Math.min(...rects.map((r) => r.left)) }
}

const meets = (a: CellRect, b: CellRect) => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right

/** Whether a rule applies to a cell of the rectangle, as the Rules Manager's Current Selection lists it. */
export const ruleMeets = (rule: ConditionalFormat, rect: CellRect): boolean => rectsOf(rule.ranges).some((r) => meets(r, rect))

/** What is left of a rectangle with `cut` taken out: the rows above and below it across its width, then the parts to its left and right. */
export function subtract(r: CellRect, cut: CellRect): CellRect[] {
  if (!meets(r, cut)) return [r]
  const out: CellRect[] = []
  const top = Math.max(r.top, cut.top)
  const bottom = Math.min(r.bottom, cut.bottom)
  if (r.top < cut.top) out.push({ ...r, bottom: cut.top - 1 })
  if (cut.left > r.left) out.push({ top, bottom, left: r.left, right: cut.left - 1 })
  if (cut.right < r.right) out.push({ top, bottom, left: cut.right + 1, right: r.right })
  if (r.bottom > cut.bottom) out.push({ ...r, top: cut.bottom + 1 })
  return out
}

/**
 * Clear Rules from Selected Cells: each rule loses the cells of `cut` and keeps the rest of its ranges, and a rule with no
 * cells left is removed. A rule whose first cell moved has its formulas rewritten for the new one (`shift`), so each cell
 * left keeps the condition it had. Null when no rule applies to the selection.
 */
export function clearRulesIn(
  rules: readonly ConditionalFormat[],
  cut: CellRect,
  shift: (rule: ConditionalFormat, rows: number, cols: number) => ConditionalFormat,
): ConditionalFormat[] | null {
  if (!rules.some((r) => ruleMeets(r, cut))) return null
  const out: ConditionalFormat[] = []
  for (const rule of rules) {
    if (!ruleMeets(rule, cut)) {
      out.push(rule)
      continue
    }
    const ranges = rectsOf(rule.ranges).flatMap((r) => subtract(r, cut)).map(rangeAddress)
    if (!ranges.length) continue
    const was = topLeftOf(rule.ranges)!
    const now = topLeftOf(ranges)!
    out.push(shift({ ...rule, ranges }, now.row - was.row, now.col - was.col))
  }
  return out
}

/** A rectangle as Excel's "Applies to" box writes it, with `$`: `$A$1:$B$5`, `$F:$F`, `$3:$5`, `$C$2`. */
function absolute(r: CellRect): string {
  if (r.top === 1 && r.bottom >= MAX_ROWS) return `$${colLetters(r.left)}:$${colLetters(r.right)}`
  if (r.left === 1 && r.right >= MAX_COLS) return `$${r.top}:$${r.bottom}`
  const a = `$${colLetters(r.left)}$${r.top}`
  return r.top === r.bottom && r.left === r.right ? a : `${a}:$${colLetters(r.right)}$${r.bottom}`
}

/** The "Applies to" box's text: `=` and the ranges, separated by the locale's list separator (`,` in en-US, `;` in ru-RU). */
export const appliesText = (ranges: readonly string[], separator: string): string => `=${rectsOf(ranges).map(absolute).join(separator)}`

/** The ranges typed into "Applies to", with or without `=` and `$`, on this sheet or naming it; null for anything else. */
export function parseApplies(text: string, separator: string, sheet: string): string[] | null {
  const body = text.trim().replace(/^=/, '')
  if (!body) return null
  const out: string[] = []
  for (const raw of body.split(separator)) {
    let part = raw.trim()
    const bang = part.lastIndexOf('!')
    if (bang >= 0) {
      const name = part.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'")
      if (!sameName(name, sheet)) return null
      part = part.slice(bang + 1)
    }
    const r = parseRange(part)
    if (!r) return null
    out.push(rangeAddress(r))
  }
  return out
}
