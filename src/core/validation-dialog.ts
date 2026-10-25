// Excel's Data Validation dialog (Data > Data Validation) over the source's `rules` and `setRules`: the sheet's rules read
// when it opens, Excel's questions for a selection with rules of several kinds or cells without the rule, and OK writing the
// sheet's whole list of rules in one call, after which the windows over the changed cells are read again.
import { shallowRef, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { CellPosition, CellRect, DataValidation, ScanOptions } from 'xlcalc/contract'
import { loadEntry } from '../entry'
import { errorText } from '../failure'
import { draftOf, planOf, rulesAfter, type RuleSettings, type RuleText, type ValidationDraft } from '../validation-rules'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import type { AlertAnswer, SheetAlert } from './validation'
import type { Windows } from './windows'

export interface ValidationDialogDeps {
  /** How the workbook's formula text is parsed; null while its locale loads. */
  scan: () => ScanOptions | null
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
  notify: (text: string) => void
  /** The active cell's rule changed: its dropdown and input message are requested again. */
  changed: () => void
}

/** The open dialog: the selection and active cell it was opened for, the sheet's rules then, the rule shown, and its first inputs. */
export interface ValidationDialogState {
  sheet: string
  sel: CellRect
  at: CellPosition
  rules: DataValidation[]
  /** The rule whose settings the dialog opened with; null for Excel's defaults. */
  shown: DataValidation | null
  draft: ValidationDraft
  text: RuleText
}

/** The parse options once the workbook's locale has loaded. */
function scanReady(scan: () => ScanOptions | null): Promise<ScanOptions> {
  const now = scan()
  if (now) return Promise.resolve(now)
  return new Promise((resolve) => {
    const stop = watch(scan, (v) => {
      if (!v) return
      stop()
      resolve(v)
    })
  })
}

export function useValidationDialog(s: SheetState, sel: Selection, w: Windows, deps: ValidationDialogDeps) {
  const read = useSheetReads(s)
  const state = shallowRef<ValidationDialogState | null>(null)
  /** The dialog can be opened: the sheet is editable and the source reads and writes rules. */
  const available = () => s.editable.value && !!s.sourceMethod('rules') && !!s.sourceMethod('setRules')

  async function start(sheet: string, r: CellRect, at: CellPosition): Promise<void> {
    const signal = read()
    const outdated = s.outdated()
    let rules: DataValidation[]
    let text: RuleText
    try {
      const [got, entry, scan] = await Promise.all([s.sourceMethod('rules')!(sheet, signal), loadEntry(), scanReady(deps.scan)])
      rules = got.validations
      text = { scan, entry }
    } catch (e) {
      if (!signal.aborted && !outdated()) deps.notify(errorText(e) || s.messages.value.loadFailed)
      return
    }
    if (signal.aborted || outdated()) return
    const plan = planOf(rules, r, at, text.scan)
    let shown: DataValidation | null = null
    if (plan.kind === 'one') shown = plan.rule
    else if (plan.kind !== 'none') {
      const m = s.messages.value
      const many = plan.kind === 'many'
      const answer = await deps.ask({ severity: many ? 'rulesMany' : 'rulesExtend', message: many ? m.dvMany : m.dvExtend, sheet, cell: cellAddress(at.row, at.col) })
      if (answer === 'cancel' || signal.aborted || outdated()) return
      // Yes extends the rule's settings to the whole selection; OK on "erase" and No start from Excel's defaults
      if (plan.kind === 'extend' && answer === 'accept') shown = plan.rule
    }
    state.value = { sheet, sel: r, at, rules, shown, draft: draftOf(shown, at, text), text }
  }

  /** Opens the dialog for the selection, after Excel's question where it asks one; false when the sheet cannot edit rules. */
  function open(): boolean {
    const r = sel.fullRect.value
    const at = s.selected.value
    if (!available() || !s.sheet.value || !r || !at || state.value) return false
    void start(s.sheetName.value, r, at)
    return true
  }

  function close(): void {
    state.value = null
    s.focusGrid()
  }

  /**
   * OK: the new rule (or none) over the selection, and with `applyAll` over the other cells with the shown rule's settings,
   * written as the sheet's whole list in one `setRules`; the dialog closes. The source's refusal, or null once written.
   */
  async function apply(rule: RuleSettings | null, applyAll: boolean): Promise<string | null> {
    const d = state.value
    const call = s.sourceMethod('setRules')
    if (!d || !call) return null
    const next = rulesAfter(d.rules, d.sel, d.at, rule, applyAll ? d.shown : null, d.text.scan)
    const outdated = s.outdated()
    let out
    try {
      out = await call(d.sheet, { validations: next.rules })
    } catch (e) {
      out = { ok: false as const, message: errorText(e) }
    }
    if (outdated() || state.value !== d) return null
    if (!out.ok) return out.message || s.messages.value.editFailed
    s.hooks.changing()
    w.applyChanges(out.changes)
    // the outcome carries values, not rules: the windows over the changed cells are read again
    w.repaint(next.ranges.map((range) => ({ sheet: d.sheet, range })))
    deps.changed()
    close()
    return null
  }

  return { state, available, open, close, apply }
}

export type ValidationDialog = ReturnType<typeof useValidationDialog>
