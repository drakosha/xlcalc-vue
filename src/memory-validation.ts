// Data validation of the book in memory: a sheet's simple rules (a list, whole numbers within bounds), what a cell shows of
// them and the alert for a typed value one does not take, without an engine.
import { parseCell, parseRange, rectContains } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { cellRect } from './rect'
import type { CellEntry, EditOutcome, SheetValidation, SheetValidationAlert } from './contract/source'
import type { MemoryValidation } from './contract/memory'

/** Excel's text of an alert whose rule has none (Excel for Windows en-US). */
const EXCEL_MESSAGE = "This value doesn't match the data validation restrictions defined for this cell."
const STRICTNESS = { information: 0, warning: 1, stop: 2 } as const

/** A number as the book in memory reads a typed one: digits, a minus and a point, nothing else. */
export const PLAIN_NUMBER = /^-?\d+(\.\d+)?$/

/** A sheet's rules with their rectangles, read once. */
export class MemoryRules {
  private readonly rules: { rule: MemoryValidation; rects: CellRect[] }[]
  constructor(rules: MemoryValidation[] = []) {
    this.rules = rules.map((rule) => ({ rule, rects: rule.range.split(/\s+/).flatMap((r) => parseRange(r) ?? []) }))
  }

  /** The rule over a cell (the first, as no two rules of a sheet meet in Excel). */
  at(cell: string): MemoryValidation | null {
    const at = parseCell(cell)
    if (!at) return null
    const one = cellRect(at.row, at.col)
    return this.rules.find((x) => x.rects.some((r) => rectContains(r, one)))?.rule ?? null
  }

  /** What the cell shows of its rule: the list's items and the input message; null without a rule. */
  show(cell: string): SheetValidation | null {
    const rule = this.at(cell)
    if (!rule) return null
    return { ...(rule.list ? { list: [...rule.list] } : {}), ...(rule.prompt ? { prompt: { ...rule.prompt } } : {}) }
  }
}

/** The alert of the strictest rule typed entries break, the first of that style, as a refusal; null when all are taken. */
export function memoryRefusal(entries: CellEntry[], rulesOf: (sheet: string) => MemoryRules | undefined): Extract<EditOutcome, { ok: false }> | null {
  let worst: SheetValidationAlert | null = null
  for (const e of entries) {
    const rule = rulesOf(e.sheet)?.at(e.cell)
    if (!rule || takes(rule, e.text)) continue
    const severity = rule.alert?.severity ?? 'stop'
    if (worst && STRICTNESS[severity] <= STRICTNESS[worst.severity]) continue
    const title = rule.alert?.title
    worst = { severity, ...(title ? { title } : {}), message: rule.alert?.message ?? EXCEL_MESSAGE, sheet: e.sheet, cell: e.cell.toUpperCase() }
  }
  return worst && { ok: false, reason: 'refused', message: worst.message, alert: worst }
}

/** Whether a rule takes a typed text: a blank and a formula (no engine computes it) pass, as Excel's Ignore blank. */
function takes(rule: MemoryValidation, text: string | null): boolean {
  if (text === null || text === '' || (text.startsWith('=') && text.length > 1)) return true
  if (rule.list && !rule.list.some((item) => item.trim() === text.trim())) return false
  if (rule.whole) {
    const n = PLAIN_NUMBER.test(text.trim()) ? Number(text) : NaN
    if (!Number.isInteger(n) || n < (rule.whole.min ?? -Infinity) || n > (rule.whole.max ?? Infinity)) return false
  }
  return true
}
