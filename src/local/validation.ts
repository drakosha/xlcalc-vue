// Data validation of the local source: what Excel displays on a cell with a rule (its dropdown items, its input message)
// from the session's rule, and a value that violates a rule as a rejection with the alert Excel shows for it.
import type { CellValidation, CellValidationRefusal } from 'xlcalc'
import type { EditOutcome, SheetValidation } from '../contract/source'

/** A cell's rule as the component displays it: a list's items when it shows its dropdown (a formula's error yields none), and the input message unless it is turned off; null without a rule. */
export function validationOf(got: CellValidation | null): SheetValidation | null {
  if (!got) return null
  const { rule } = got
  const out: SheetValidation = {}
  if (rule.type === 'list' && rule.dropdown !== false) out.list = Array.isArray(got.list) ? got.list.map((i) => i.text) : []
  if (rule.showPrompt !== false && (rule.prompt || rule.promptTitle)) out.prompt = { ...(rule.promptTitle ? { title: rule.promptTitle } : {}), text: rule.prompt ?? '' }
  return out
}

/** A value rejected by its rule, with Excel's alert: the rule's style, its title (the component's default without one) and its message (the library's text without one). */
export function alertRefusalOf(r: CellValidationRefusal): Extract<EditOutcome, { ok: false }> {
  const title = r.rule.errorTitle
  return {
    ok: false,
    reason: 'refused',
    message: r.message,
    alert: { severity: r.severity, ...(title ? { title } : {}), message: r.message, sheet: r.sheet, cell: r.cell },
  }
}
