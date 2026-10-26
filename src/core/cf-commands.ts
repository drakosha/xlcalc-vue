// Excel's Conditional Formatting menu over the source's `rules` and `setRules`: the quick rules' dialogs, the galleries, Clear
// Rules, New Rule and the Rules Manager. Each change replaces the sheet's list in one call (a single undo step), and the
// windows over the rules it touched are read again, since the outcome carries values only.
import { shallowRef } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { ConditionalFormat, ConditionalStyle, IconSetName, Locale, ScanOptions } from 'xlcalc/contract'
import type { ClearRulesFrom, ColorScalePreset, DataBarColor, QuickRule } from '../contract/conditional'
import type { EditOutcome } from '../contract/source'
import { loadEntry } from '../entry'
import { errorText } from '../failure'
import { shiftFormula, shiftRule } from './cf-formula'
import { colorScaleRule, dataBarRule, iconSetRule, quickRuleOf, type QuickInput } from './cf-presets'
import { clearRulesIn, topLeftOf } from './cf-ranges'
import type { Editing } from './editing'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** Where a dialog's rule goes: the sheet, its ranges, and the active cell that typed formulas are relative to, as in Excel. */
export interface RuleTarget {
  sheet: string
  ranges: string[]
  active: { row: number; col: number }
}

/** The New or Edit Formatting Rule dialog: its target, the rule edited (null for a new one), and what OK does with the rule. */
export interface RuleDialog extends RuleTarget {
  base: ConditionalFormat | null
  done: (rule: ConditionalFormat) => void
}

/** A rule's Format Cells: the format it opens with, what OK does with the new one, and what Cancel does. */
export interface StyleDialog {
  style: ConditionalStyle
  done: (style: ConditionalStyle) => void
  cancel?: () => void
}

/** The Rules Manager: the sheet, the selection it lists rules for, and the sheet's rules once read. */
export interface ManagerState extends RuleTarget {
  rules: ConditionalFormat[] | null
}

export interface ConditionalDeps {
  notify: (text: string) => void
  /** How the workbook's formula text is parsed (its locale); null while the locale loads. */
  scan: () => ScanOptions | null
}

export function useConditional(s: SheetState, sel: Selection, w: Windows, ed: Editing, deps: ConditionalDeps) {
  const quick = shallowRef<(RuleTarget & { kind: QuickRule }) | null>(null)
  const rule = shallowRef<RuleDialog | null>(null)
  const style = shallowRef<StyleDialog | null>(null)
  const manager = shallowRef<ManagerState | null>(null)

  const scan = (): ScanOptions => deps.scan() ?? {}
  /** The workbook's locale object, en-US without one. */
  const locale = (): Locale | undefined => {
    const l = scan().locale
    return typeof l === 'object' ? l : undefined
  }

  /** The source's rule methods, while the sheet is editable and has a selection; otherwise null. */
  function methods() {
    const read = s.sourceMethod('rules')
    const write = s.sourceMethod('setRules')
    return s.editable.value && read && write && s.sheet.value && sel.fullRect.value && s.selected.value ? { read, write } : null
  }
  const available = () => methods() !== null

  /** The selection as a rule's target. */
  function target(): RuleTarget {
    const a = s.selected.value!
    return { sheet: s.sheetName.value, ranges: [sel.selectionRef.value], active: { row: a.row, col: a.col } }
  }

  /** A formula typed relative to the active cell, written for the first cell of `ranges`, as Excel writes a dialog's formula. */
  function stored(t: RuleTarget, ranges: string[] = t.ranges): (formula: string) => string {
    const tl = topLeftOf(ranges) ?? t.active
    return (f) => shiftFormula(f, tl.row - t.active.row, tl.col - t.active.col, scan())
  }
  /** A rule's formula as a dialog shows it: relative to the active cell. */
  function shown(t: RuleTarget, ranges: string[]): (formula: string) => string {
    const tl = topLeftOf(ranges) ?? t.active
    return (f) => shiftFormula(f, t.active.row - tl.row, t.active.col - tl.col, scan())
  }
  const activeAddress = (t: RuleTarget) => cellAddress(t.active.row, t.active.col)

  /** The sheet's rules, or null when the source failed (its message is displayed). */
  async function read(sheet: string): Promise<ConditionalFormat[] | null> {
    const m = methods()
    if (!m) return null
    try {
      return (await m.read(sheet)).conditionalFormats
    } catch (e) {
      deps.notify(errorText(e) || s.messages.value.editFailed)
      return null
    }
  }

  /** Replaces the sheet's rules in one call; after a success the windows over the old and the new rules are read again. */
  async function write(sheet: string, next: ConditionalFormat[], before: readonly ConditionalFormat[]): Promise<EditOutcome | null> {
    const m = methods()
    if (!m) return null
    const outdated = s.outdated()
    // plain data: a dialog's reactive proxies cannot be cloned into a source's worker thread
    const plain = JSON.parse(JSON.stringify(next)) as ConditionalFormat[]
    const out = await ed.run(() => m.write(sheet, { conditionalFormats: plain }), [], null, null, false)
    if (out.ok && !outdated()) w.repaint([...before, ...next].flatMap((r) => r.ranges.map((range) => ({ sheet, range }))))
    return out
  }

  /** A new rule first in priority, as Excel adds one. */
  async function add(sheet: string, made: ConditionalFormat): Promise<EditOutcome | null> {
    const rules = await read(sheet)
    return rules ? write(sheet, [made, ...rules], rules) : null
  }
  /** A gallery rule over the selection. */
  const addOver = (make: (ranges: string[]) => ConditionalFormat): Promise<EditOutcome | null> => {
    if (!available()) return Promise.resolve(null)
    const t = target()
    return add(t.sheet, make(t.ranges))
  }

  // ── the quick rules ──
  function quickRule(kind: QuickRule): boolean {
    if (!available()) return false
    void loadEntry().catch(() => undefined)
    quick.value = { kind, ...target() }
    return true
  }
  /** OK in a quick dialog: the rule over its target; formulas and a text reference typed for the active cell are written for the first cell. */
  function quickOk(q: QuickInput): void {
    const t = quick.value
    if (!t) return
    quick.value = null
    s.focusGrid()
    const write = stored(t)
    void add(t.sheet, quickRuleOf(t.kind, t.ranges, { ...q, ...(q.formulas ? { formulas: q.formulas.map(write) } : {}), ...(q.textFormula ? { text: write(q.text ?? ''), textFormula: write(q.textFormula) } : {}) }))
  }
  function quickClose(): void {
    quick.value = null
    s.focusGrid()
  }

  // ── Clear Rules ──
  async function clearRules(from: ClearRulesFrom): Promise<EditOutcome | null> {
    if (!available()) return null
    const sheet = s.sheetName.value
    const rect = sel.fullRect.value!
    const rules = await read(sheet)
    if (!rules?.length) return null
    const next = from === 'sheet' ? [] : clearRulesIn(rules, rect, (r, rows, cols) => shiftRule(r, rows, cols, scan()))
    return next ? write(sheet, next, rules) : null
  }

  // ── New Formatting Rule, a rule's Format Cells ──
  function newRule(): boolean {
    if (!available()) return false
    void loadEntry().catch(() => undefined)
    const t = target()
    rule.value = { ...t, base: null, done: (made) => void add(t.sheet, made) }
    return true
  }
  /** Opens the rule dialog from the Rules Manager. */
  function openRule(d: RuleDialog): void {
    void loadEntry().catch(() => undefined)
    rule.value = d
  }
  function ruleClose(): void {
    rule.value = null
    if (!manager.value) s.focusGrid()
  }
  const openStyle = (d: StyleDialog) => (style.value = d)
  /** The rule's Format Cells closes: OK passes the format on, Cancel lets its opener know. */
  function styleClose(made: ConditionalStyle | null): void {
    const d = style.value
    style.value = null
    if (made) d?.done(made)
    else d?.cancel?.()
  }

  // ── the Rules Manager ──
  function manageRules(): boolean {
    if (!available()) return false
    void loadEntry().catch(() => undefined)
    const t = target()
    manager.value = { ...t, rules: null }
    void reread()
    return true
  }
  /** Reads the sheet's rules into the open manager. */
  async function reread(): Promise<void> {
    const at = manager.value
    if (!at) return
    const rules = await read(at.sheet)
    if (manager.value !== at) return
    if (!rules) manager.value = null
    else manager.value = { ...at, rules }
  }
  /** Apply and OK: the manager's list written as the sheet's rules; Apply reads them back and keeps the manager open. */
  async function managerWrite(next: ConditionalFormat[], close: boolean): Promise<void> {
    const at = manager.value
    if (!at?.rules) return
    if (close) managerClose()
    const out = await write(at.sheet, next, at.rules)
    if (!close && out?.ok && manager.value === at) await reread()
  }
  /** Every dialog closes: a new workbook replaced the one they were for. */
  function closeAll(): void {
    quick.value = null
    manager.value = null
    rule.value = null
    style.value = null
  }
  function managerClose(): void {
    closeAll()
    s.focusGrid()
  }

  /** What the host's ribbon calls, as `SheetViewApi` lists them. */
  const api = {
    quickRule,
    dataBars: (color: DataBarColor, fill: 'gradient' | 'solid' = 'gradient') => addOver((r) => dataBarRule(r, color, fill)),
    colorScale: (preset: ColorScalePreset) => addOver((r) => colorScaleRule(r, preset)),
    iconSet: (icons: IconSetName) => addOver((r) => iconSetRule(r, icons)),
    clearRules,
    newRule,
    manageRules,
  }

  return {
    api,
    quick,
    rule,
    style,
    manager,
    locale,
    stored,
    shown,
    activeAddress,
    quickOk,
    quickClose,
    openRule,
    ruleClose,
    openStyle,
    styleClose,
    managerWrite,
    managerClose,
    closeAll,
  }
}

export type Conditional = ReturnType<typeof useConditional>
