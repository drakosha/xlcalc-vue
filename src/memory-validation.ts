// Data validation of the book in memory: a sheet's simple rules (a list, whole numbers within bounds), what a cell shows of
// them and the alert for a typed value one does not take, without an engine.
import { parseCell, parseRange, rectContains } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import { cellRect } from './rect'
import type { CellEntry, EditOutcome, SheetValidation, SheetValidationAlert } from './source'

/** A rule of data validation of a memory sheet, as Excel's dialog makes one: a list or whole numbers within bounds, an input message, the alert. */
export interface MemoryValidation {
  /** The cells: A1 ranges separated by spaces (`B2:B20 D2`, `C:C`). */
  range: string
  /** A list: the value is one of the items, in their case, spaces at either end aside; the cell shows them in a dropdown. */
  list?: string[]
  /** Whole numbers from `min` to `max`, either bound left out for none. */
  whole?: { min?: number; max?: number }
  /** The input message shown while a cell is selected. */
  prompt?: { title?: string; text: string }
  /** The alert for a value the rule does not take: stop by default, the title and message (Excel's own text without one). */
  alert?: { severity?: SheetValidationAlert['severity']; title?: string; message?: string }
}

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
