// Excel's Data Validation dialog without its markup: what the selection's rules make Excel ask before it opens, the boxes
// as a rule fills them for the active cell, the rule OK makes from them with Excel's checks, and the sheet's rules after it.
import { parseRange, rangeAddress, rectsMeet } from 'xlcalc/a1'
import type { CellPosition, CellRect, DataValidation, Locale, ScanOptions, ValidationOperator, ValidationType } from 'xlcalc/contract'
import type { Entry } from './entry'
import { shiftFormula } from './formula-shift'
import { rectMinus } from './rect'

/** The types in the order of Excel's Allow box, and the comparisons of its Data box. */
export const VALIDATION_TYPES: readonly ValidationType[] = ['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom']
export const VALIDATION_OPERATORS: readonly ValidationOperator[] = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
/** Excel's limits on a rule's texts: the input message's title and text, the error alert's title and text. */
export const VALIDATION_LIMITS = { promptTitle: 32, prompt: 255, errorTitle: 32, error: 225 } as const

/** The boxes of the Settings tab, as Excel names them by the type and the comparison. */
export type ValidationBox = 'min' | 'max' | 'value' | 'length' | 'start' | 'end' | 'date' | 'startTime' | 'endTime' | 'time' | 'source' | 'formula'
export type ValidationTab = 'settings' | 'input' | 'error'
export const VALIDATION_TABS: readonly ValidationTab[] = ['settings', 'input', 'error']

/** The dialog's inputs; the boxes hold what is typed until OK reads them. */
export interface ValidationDraft {
  type: ValidationType
  operator: ValidationOperator
  /** The first box (Minimum, Value, Source, Formula) and the second (Maximum). */
  first: string
  second: string
  allowBlank: boolean
  dropdown: boolean
  /** Apply these changes to all other cells with the same settings. */
  applyAll: boolean
  showPrompt: boolean
  promptTitle: string
  prompt: string
  showError: boolean
  errorStyle: 'stop' | 'warning' | 'information'
  errorTitle: string
  error: string
}

/** How formula text reads and writes in the workbook's locale: its references (`xlcalc/text`), and typed values (`xlcalc/entry`). */
export interface RuleText {
  scan: ScanOptions
  entry: Entry
}

/** What Excel does before the dialog opens: no rule in the selection, one covering it, one covering part of it ("extend?"), or several ("erase?"). */
export type ValidationPlan = { kind: 'none' } | { kind: 'one' | 'extend'; rule: DataValidation } | { kind: 'many' }

/** Why OK is refused: an empty box, a box whose value is not one, or a maximum below the minimum. */
export type DraftError = { kind: 'empty' | 'invalid'; box: ValidationBox } | { kind: 'order'; min: ValidationBox; max: ValidationBox }

/** A rule's settings without its cells, with formulas written for the active cell. */
export type RuleSettings = Omit<DataValidation, 'range'>

const COMPARED: ReadonlySet<ValidationType> = new Set(['whole', 'decimal', 'date', 'time', 'textLength'])
const BOUNDED: ReadonlySet<ValidationOperator> = new Set(['between', 'notBetween'])
/** The number formats a date and a time bound are shown in: Excel's short date and its time with seconds. */
const DATE_FORMAT = 14
const TIME_FORMAT = 21

export const isCompared = (type: ValidationType): boolean => COMPARED.has(type)

/** The boxes Excel's Settings tab shows for a type and a comparison, first box first. */
export function boxesOf(type: ValidationType, operator: ValidationOperator): ValidationBox[] {
  if (type === 'list') return ['source']
  if (type === 'custom') return ['formula']
  if (!COMPARED.has(type)) return []
  const [lo, hi, one]: [ValidationBox, ValidationBox, ValidationBox] =
    type === 'date' ? ['start', 'end', 'date'] : type === 'time' ? ['startTime', 'endTime', 'time'] : ['min', 'max', type === 'textLength' ? 'length' : 'value']
  if (BOUNDED.has(operator)) return [lo, hi]
  if (operator === 'equal' || operator === 'notEqual') return [one]
  return [operator.startsWith('greater') ? lo : hi]
}

export const emptyDraft = (): ValidationDraft => ({
  type: 'any',
  operator: 'between',
  first: '',
  second: '',
  allowBlank: true,
  dropdown: true,
  applyAll: false,
  showPrompt: true,
  promptTitle: '',
  prompt: '',
  showError: true,
  errorStyle: 'stop',
  errorTitle: '',
  error: '',
})

/** A rule's ranges (`B2:B9 D2`) as rectangles; a part that is not one is skipped. */
export const rectsOf = (range: string): CellRect[] => range.split(/\s+/).flatMap((p) => (p && parseRange(p)) || [])

/** The top-left cell of the rectangle enclosing the ranges, which a rule's formulas are written for. */
function cornerOf(rects: CellRect[]): CellPosition {
  return { row: Math.min(...rects.map((r) => r.top)), col: Math.min(...rects.map((r) => r.left)) }
}

/** The settings with their formulas moved `rows` down and `cols` right, as a copy moves them. */
function shifted<T extends RuleSettings>(rule: T, rows: number, cols: number, scan: ScanOptions): T {
  const out = { ...rule }
  if (rule.formula1 !== undefined) out.formula1 = shiftFormula(rule.formula1, rows, cols, scan)
  if (rule.formula2 !== undefined) out.formula2 = shiftFormula(rule.formula2, rows, cols, scan)
  return out
}

/** A rule's settings with its formulas for cell `at`. */
function settingsAt(rule: DataValidation, at: CellPosition, scan: ScanOptions): RuleSettings {
  const { range, ...rest } = rule
  const corner = cornerOf(rectsOf(range))
  return shifted(rest, at.row - corner.row, at.col - corner.col, scan)
}

/** What makes two rules the same settings for Excel: every field with its default filled in, and the formulas written for A1. */
function settingsKey(rule: DataValidation, scan: ScanOptions): string {
  const s = settingsAt(rule, { row: 1, col: 1 }, scan)
  const compared = COMPARED.has(s.type)
  const operator = compared ? (s.operator ?? 'between') : null
  return JSON.stringify([
    s.type,
    operator,
    s.type === 'any' ? null : (s.formula1 ?? ''),
    operator && BOUNDED.has(operator) ? (s.formula2 ?? '') : null,
    s.allowBlank ?? true,
    s.type === 'list' ? (s.dropdown ?? true) : null,
    s.showPrompt ?? true,
    s.promptTitle ?? '',
    s.prompt ?? '',
    s.showError ?? true,
    s.errorStyle ?? 'stop',
    s.errorTitle ?? '',
    s.error ?? '',
  ])
}

/** `rects` with every part covered by `cuts` removed. */
const without = (rects: CellRect[], cuts: CellRect[]): CellRect[] => cuts.reduce((left, cut) => left.flatMap((r) => rectMinus(r, cut)), rects)

/**
 * What Excel finds in the selection when Data Validation opens: no rule; one rule's settings over every cell; the same
 * settings over some cells only, when Excel asks whether to extend them; rules of more than one kind, when it asks whether to
 * erase them. The rule shown is the one containing the active cell where it is one of them.
 */
export function planOf(rules: DataValidation[], sel: CellRect, at: CellPosition, scan: ScanOptions): ValidationPlan {
  const hits = rules.filter((r) => rectsOf(r.range).some((x) => rectsMeet(x, sel)))
  if (!hits.length) return { kind: 'none' }
  if (new Set(hits.map((r) => settingsKey(r, scan))).size > 1) return { kind: 'many' }
  const active: CellRect = { top: at.row, bottom: at.row, left: at.col, right: at.col }
  const rule = hits.find((r) => rectsOf(r.range).some((x) => rectsMeet(x, active))) ?? hits[0]!
  return { kind: without([sel], hits.flatMap((r) => rectsOf(r.range))).length ? 'extend' : 'one', rule }
}

/** The workbook's locale object; none for en-US. */
const localeOf = (t: RuleText): Locale | undefined => (typeof t.scan.locale === 'object' ? t.scan.locale : undefined)
/** The locale's separator of a literal list's items in Excel's Source box, as of function arguments: `,` in en-US, `;` in ru-RU. */
const listSeparator = (t: RuleText): string => localeOf(t)?.listSeparator ?? ','

/** A formula that is a number as typed in the locale (`10`, `-1,5` in ru-RU), which Excel's boxes show without `=`. */
function numberIn(formula: string, t: RuleText): number | null {
  const dec = localeOf(t)?.decimalSeparator ?? '.'
  const esc = dec.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return new RegExp(`^-?\\d+(?:${esc}\\d+)?(?:E[+-]?\\d+)?$`, 'i').test(formula) ? Number(formula.replace(dec, '.')) : null
}

/** A literal list's items (`"a,b"`, as a file stores it in every locale); null for a formula. */
function listItems(formula: string): string[] | null {
  return /^"(?:[^"]|"")*"$/.test(formula) ? formula.slice(1, -1).replace(/""/g, '"').split(',') : null
}

/** A bound as Excel's box shows it: a number as is, a date or a time in the locale's short format, anything else as a formula. */
function boundText(formula: string | undefined, type: ValidationType, t: RuleText): string {
  if (formula === undefined) return ''
  const n = numberIn(formula, t)
  if (n === null) return `=${formula}`
  if (type !== 'date' && type !== 'time') return formula
  return t.entry.entryText(n, { locale: t.scan.locale, numberFormat: type === 'date' ? DATE_FORMAT : TIME_FORMAT })
}

/** The dialog's inputs for a rule (Excel's defaults without one), its formulas shown for the active cell `at`, with `=`. */
export function draftOf(rule: DataValidation | null, at: CellPosition, t: RuleText): ValidationDraft {
  const d = emptyDraft()
  if (!rule) return d
  const s = settingsAt(rule, at, t.scan)
  d.type = s.type
  d.operator = s.operator ?? 'between'
  if (s.type === 'list') {
    const items = s.formula1 === undefined ? null : listItems(s.formula1)
    d.first = items ? items.join(listSeparator(t)) : s.formula1 === undefined ? '' : `=${s.formula1}`
  } else if (s.type === 'custom') d.first = s.formula1 === undefined ? '' : `=${s.formula1}`
  else if (COMPARED.has(s.type)) {
    d.first = boundText(s.formula1, s.type, t)
    if (BOUNDED.has(d.operator)) d.second = boundText(s.formula2, s.type, t)
  }
  d.allowBlank = s.allowBlank ?? true
  d.dropdown = s.dropdown ?? true
  d.showPrompt = s.showPrompt ?? true
  d.promptTitle = s.promptTitle ?? ''
  d.prompt = s.prompt ?? ''
  d.showError = s.showError ?? true
  d.errorStyle = s.errorStyle ?? 'stop'
  d.errorTitle = s.errorTitle ?? ''
  d.error = s.error ?? ''
  return d
}

/** A number as a formula in the locale, as the formula bar writes it (`1,5` in ru-RU). */
const numberFormula = (n: number, t: RuleText): string => t.entry.entryText(n, { locale: t.scan.locale })

/** A typed constant as a formula: a number, a boolean in the locale's words, or text in quotes; null for one that is none. */
function constantFormula(text: string, t: RuleText): string | null {
  const got = t.entry.parseEntry(text, { locale: t.scan.locale })
  if (!got.ok || !got.input) return null
  const v = got.input.value
  if (typeof v === 'number') return numberFormula(v, t)
  if (typeof v === 'boolean') return (localeOf(t)?.booleans ?? { TRUE: 'TRUE', FALSE: 'FALSE' })[v ? 'TRUE' : 'FALSE']
  if (typeof v === 'string') return `"${v.replace(/"/g, '""')}"`
  return null
}

/** A bound box read as Excel reads it: `=` starts a formula, anything else is a value typed as into a cell, which must be a number (a date, a time). */
function boundFormula(text: string, box: ValidationBox, t: RuleText): { formula: string; n: number | null } | DraftError {
  if (!text.trim()) return { kind: 'empty', box }
  if (text.startsWith('=')) return text.length > 1 ? { formula: text.slice(1), n: null } : { kind: 'invalid', box }
  const got = t.entry.parseEntry(text, { locale: t.scan.locale })
  if (!got.ok || typeof got.input?.value !== 'number') return { kind: 'invalid', box }
  return { formula: numberFormula(got.input.value, t), n: got.input.value }
}

const isError = (x: object): x is DraftError => 'kind' in x

/**
 * The rule OK makes from the dialog's inputs, its formulas for the active cell, or null for none (Any value without a
 * message: the cells lose their rule); or why Excel refuses it. Only what differs from a new rule's defaults is written.
 */
export function ruleOf(d: ValidationDraft, t: RuleText): { ok: true; rule: RuleSettings | null } | { ok: false; error: DraftError } {
  const rule: RuleSettings = { type: d.type }
  if (COMPARED.has(d.type)) {
    const boxes = boxesOf(d.type, d.operator)
    const first = boundFormula(d.first, boxes[0]!, t)
    if (isError(first)) return { ok: false, error: first }
    if (d.operator !== 'between') rule.operator = d.operator
    rule.formula1 = first.formula
    if (boxes[1]) {
      const second = boundFormula(d.second, boxes[1], t)
      if (isError(second)) return { ok: false, error: second }
      if (first.n !== null && second.n !== null && second.n < first.n) return { ok: false, error: { kind: 'order', min: boxes[0]!, max: boxes[1] } }
      rule.formula2 = second.formula
    }
  } else if (d.type === 'list') {
    if (!d.first.trim()) return { ok: false, error: { kind: 'empty', box: 'source' } }
    if (d.first.startsWith('=')) rule.formula1 = d.first.slice(1)
    // a literal list typed with the locale's list separator, stored with commas as a file stores it in every locale
    else rule.formula1 = `"${d.first.split(listSeparator(t)).join(',').replace(/"/g, '""')}"`
    if (!d.dropdown) rule.dropdown = false
  } else if (d.type === 'custom') {
    if (!d.first.trim()) return { ok: false, error: { kind: 'empty', box: 'formula' } }
    const formula = d.first.startsWith('=') ? d.first.slice(1) : constantFormula(d.first, t)
    if (!formula) return { ok: false, error: { kind: 'invalid', box: 'formula' } }
    rule.formula1 = formula
  } else if (!d.promptTitle && !d.prompt && !d.errorTitle && !d.error) return { ok: true, rule: null }
  if (d.type !== 'any' && !d.allowBlank) rule.allowBlank = false
  if (!d.showPrompt) rule.showPrompt = false
  if (d.promptTitle) rule.promptTitle = d.promptTitle.slice(0, VALIDATION_LIMITS.promptTitle)
  if (d.prompt) rule.prompt = d.prompt.slice(0, VALIDATION_LIMITS.prompt)
  if (!d.showError) rule.showError = false
  if (d.errorStyle !== 'stop') rule.errorStyle = d.errorStyle
  if (d.errorTitle) rule.errorTitle = d.errorTitle.slice(0, VALIDATION_LIMITS.errorTitle)
  if (d.error) rule.error = d.error.slice(0, VALIDATION_LIMITS.error)
  return { ok: true, rule }
}

/**
 * The sheet's rules after OK: the new rule (`made`, its formulas for the active cell `at`, or null for none) over the
 * selection, and with `sameAs` (Apply these changes to all other cells with the same settings) over every cell with that
 * rule's settings too; a rule partly covered keeps its other cells, with its formulas written for its new top-left cell.
 * Also the ranges whose rules changed.
 */
export function rulesAfter(
  rules: DataValidation[],
  sel: CellRect,
  at: CellPosition,
  made: RuleSettings | null,
  sameAs: DataValidation | null,
  scan: ScanOptions,
): { rules: DataValidation[]; ranges: string[] } {
  const key = sameAs ? settingsKey(sameAs, scan) : null
  const absorbed = new Set(key === null ? [] : rules.filter((r) => settingsKey(r, scan) === key))
  const target = [sel, ...without([...absorbed].flatMap((r) => rectsOf(r.range)), [sel])]
  const placed = (settings: RuleSettings, rects: CellRect[], from: CellPosition): DataValidation => {
    const corner = cornerOf(rects)
    return { range: rects.map(rangeAddress).join(' '), ...shifted(settings, corner.row - from.row, corner.col - from.col, scan) }
  }
  const out: DataValidation[] = []
  for (const r of rules) {
    if (absorbed.has(r)) continue
    const rects = rectsOf(r.range)
    if (!rects.some((x) => target.some((y) => rectsMeet(x, y)))) {
      out.push(r)
      continue
    }
    const left = without(rects, target)
    if (!left.length) continue
    const { range: _, ...settings } = r
    out.push(placed(settings, left, cornerOf(rects)))
  }
  if (made) out.push(placed(made, target, at))
  return { rules: out, ranges: target.map(rangeAddress) }
}
