// Data validation of the local source: what Excel shows on a cell with a rule (its dropdown's items, its input message)
// from the session's rule, and a value a rule does not take as a refusal with the alert Excel shows for it.
import type { CellValidation, CellValidationRefusal } from 'xlcalc'
import type { EditOutcome, SheetValidation } from '../contract/source'

/** A cell's rule as the component shows it: a list's items when it shows its dropdown (a formula's error gives none), the input message unless it is off; null without a rule. */
export function validationOf(got: CellValidation | null): SheetValidation | null {
  if (!got) return null
  const { rule } = got
  const out: SheetValidation = {}
  if (rule.type === 'list' && rule.dropdown !== false) out.list = Array.isArray(got.list) ? got.list.map((i) => i.text) : []
  if (rule.showPrompt !== false && (rule.prompt || rule.promptTitle)) out.prompt = { ...(rule.promptTitle ? { title: rule.promptTitle } : {}), text: rule.prompt ?? '' }
  return out
}

/** A value refused by its rule, with Excel's alert: the rule's style, its title (the component's default without one) and its message, the library's text without one. */
export function alertRefusalOf(r: CellValidationRefusal): Extract<EditOutcome, { ok: false }> {
  const title = r.rule.errorTitle
  return {
    ok: false,
    reason: 'refused',
    message: r.message,
    alert: { severity: r.severity, ...(title ? { title } : {}), message: r.message, sheet: r.sheet, cell: r.cell },
  }
}
