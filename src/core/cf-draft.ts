// Excel's New Formatting Rule and Edit Formatting Rule dialogs without their markup: the rule types and their inputs as a
// rule fills them, and the rule OK writes, keeping what the dialog does not show (Stop If True, negative bars, custom icons).
import type { ConditionalFormat, ConditionalStyle, ConditionalValue, IconSetName, Locale, TimePeriodFormat, TextFormat, CellIsFormat } from 'xlcalc/contract'
import type { Entry } from '../entry'
import { boxFormula, hasRelative, scanOf, textRuleFormula } from './cf-formula'
import { DATA_BAR_COLORS, iconCount, iconSetRule } from './cf-presets'

export type RuleKind = 'values' | 'contain' | 'rank' | 'average' | 'unique' | 'formula'
/** The rule types in the order of Excel's list. */
export const RULE_KINDS: readonly RuleKind[] = ['values', 'contain', 'rank', 'average', 'unique', 'formula']
export type ValuesStyle = 'scale2' | 'scale3' | 'dataBar' | 'iconSet'
export const VALUES_STYLES: readonly ValuesStyle[] = ['scale2', 'scale3', 'dataBar', 'iconSet']
export type ContainKind = 'cellValue' | 'text' | 'dates' | 'blanks' | 'noBlanks' | 'errors' | 'noErrors'
export const CONTAIN_KINDS: readonly ContainKind[] = ['cellValue', 'text', 'dates', 'blanks', 'noBlanks', 'errors', 'noErrors']
export type Operator = CellIsFormat['operator']
export const OPERATORS: readonly Operator[] = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
export type TextOp = TextFormat['type']
export const TEXT_OPS: readonly TextOp[] = ['containsText', 'notContainsText', 'beginsWith', 'endsWith']
export type Period = TimePeriodFormat['period']
/** The periods in the order of Excel's list. */
export const PERIODS: readonly Period[] = ['yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek', 'lastMonth', 'thisMonth', 'nextMonth']
export type AverageKind = 'above' | 'below' | 'equalAbove' | 'equalBelow' | 'std1Above' | 'std2Above' | 'std3Above' | 'std1Below' | 'std2Below' | 'std3Below'
export const AVERAGE_KINDS: readonly AverageKind[] = ['above', 'below', 'equalAbove', 'equalBelow', 'std1Above', 'std2Above', 'std3Above', 'std1Below', 'std2Below', 'std3Below']
export type ValueType = ConditionalValue['type']
/** The threshold types each box offers: a scale's lowest and highest point, its midpoint, a data bar's bounds, an icon's threshold. */
export const LOW_TYPES: readonly ValueType[] = ['min', 'num', 'percent', 'formula', 'percentile']
export const MID_TYPES: readonly ValueType[] = ['num', 'percent', 'formula', 'percentile']
export const HIGH_TYPES: readonly ValueType[] = ['max', 'num', 'percent', 'formula', 'percentile']
export const BAR_LOW_TYPES: readonly ValueType[] = ['autoMin', 'min', 'num', 'percent', 'formula', 'percentile']
export const BAR_HIGH_TYPES: readonly ValueType[] = ['autoMax', 'max', 'num', 'percent', 'formula', 'percentile']
export const ICON_TYPES: readonly ValueType[] = ['num', 'percent', 'formula', 'percentile']
/** Excel's icon sets in the order of its Icon Style list. */
export const ICON_SETS: readonly IconSetName[] = [
  '3Arrows', '3ArrowsGray', '3Triangles', '4Arrows', '4ArrowsGray', '5Arrows', '5ArrowsGray',
  '3TrafficLights1', '3TrafficLights2', '3Signs', '4TrafficLights', '4RedToBlack',
  '3Symbols', '3Symbols2', '3Flags', '3Stars', '4Rating', '5Quarters', '5Rating', '5Boxes',
]

export interface PointDraft {
  type: ValueType
  /** The box's text: a number, or a formula after `=`. */
  value: string
  color: string
}
export interface IconPointDraft {
  gte: boolean
  type: ValueType
  value: string
}

/** The dialog's inputs for every rule type, so switching types keeps what was typed. */
export interface RuleDraft {
  kind: RuleKind
  values: ValuesStyle
  scale2: PointDraft[]
  scale3: PointDraft[]
  barMin: PointDraft
  barMax: PointDraft
  barColor: string
  barGradient: boolean
  barBorder: boolean
  barBorderColor: string
  barOnly: boolean
  barDirection: 'context' | 'leftToRight' | 'rightToLeft'
  iconSet: IconSetName
  iconReverse: boolean
  iconOnly: boolean
  /** The thresholds of the icons after the first, from the lowest. */
  iconPoints: IconPointDraft[]
  contain: ContainKind
  operator: Operator
  value1: string
  value2: string
  textOp: TextOp
  text: string
  period: Period
  top: boolean
  rank: string
  percent: boolean
  average: AverageKind
  unique: boolean
  formula: string
  style: ConditionalStyle
}

const point = (type: ValueType, color: string, value = ''): PointDraft => ({ type, value, color })
/** A threshold's box: a number as is, any other formula after `=`. */
const boxOf = (v: ConditionalValue): string => (v.value === undefined ? '' : /^-?[\d.,]+$/.test(v.value) ? v.value : `=${v.value}`)
const pointOf = (v: ConditionalValue, color: string): PointDraft => ({ type: v.type, value: boxOf(v), color })
const iconPointsOf = (set: IconSetName, points?: { value: ConditionalValue; gte?: boolean }[]) =>
  (points ?? iconSetRule([], set).points).slice(1).map((p) => ({ gte: p.gte !== false, type: p.value.type, value: boxOf(p.value) }))

/** A new rule's inputs, as Excel's dialog opens: a two-color scale from orange to yellow, the rest at Excel's defaults. */
export function newDraft(): RuleDraft {
  return {
    kind: 'values',
    values: 'scale2',
    scale2: [point('min', '#FF7128'), point('max', '#FFEF9C')],
    scale3: [point('min', '#F8696B'), point('percentile', '#FFEB84', '50'), point('max', '#63BE7B')],
    barMin: point('autoMin', ''),
    barMax: point('autoMax', ''),
    barColor: DATA_BAR_COLORS.blue,
    barGradient: true,
    barBorder: true,
    barBorderColor: DATA_BAR_COLORS.blue,
    barOnly: false,
    barDirection: 'context',
    iconSet: '3TrafficLights1',
    iconReverse: false,
    iconOnly: false,
    iconPoints: iconPointsOf('3TrafficLights1'),
    contain: 'cellValue',
    operator: 'between',
    value1: '',
    value2: '',
    textOp: 'containsText',
    text: '',
    period: 'yesterday',
    top: true,
    rank: '10',
    percent: false,
    average: 'above',
    unique: false,
    formula: '',
    style: {},
  }
}

/** A new set's thresholds, as Excel resets them when another icon style is picked. */
export function pickIconSet(d: RuleDraft, set: IconSetName): void {
  if (iconCount(set) !== iconCount(d.iconSet)) d.iconPoints = iconPointsOf(set)
  d.iconSet = set
}

const AVERAGE_OF: Record<AverageKind, { below?: true; equal?: true; stdDev?: number }> = {
  above: {},
  below: { below: true },
  equalAbove: { equal: true },
  equalBelow: { below: true, equal: true },
  std1Above: { stdDev: 1 },
  std2Above: { stdDev: 2 },
  std3Above: { stdDev: 3 },
  std1Below: { stdDev: 1, below: true },
  std2Below: { stdDev: 2, below: true },
  std3Below: { stdDev: 3, below: true },
}
const CONTAIN_TYPES: Partial<Record<ContainKind, ConditionalFormat['type']>> = { blanks: 'containsBlanks', noBlanks: 'notContainsBlanks', errors: 'containsErrors', noErrors: 'notContainsErrors' }

/** A text rule's text as its box shows it: after `=` when the rule's formula does not search for the text itself, which is then a reference. */
function textBox(rule: TextFormat): string {
  const quoted = `"${rule.text.replace(/"/g, '""')}"`
  return rule.formula !== undefined && !rule.formula.includes(quoted) ? `=${rule.text}` : rule.text
}

/** The inputs as an existing rule fills them; `shown` gives a formula as the dialog shows it (relative to the active cell). */
export function draftOf(rule: ConditionalFormat, shown: (formula: string) => string): RuleDraft {
  const d = newDraft()
  // plain data: the rule may be a reactive list's
  d.style = 'style' in rule && rule.style ? (JSON.parse(JSON.stringify(rule.style)) as ConditionalStyle) : {}
  switch (rule.type) {
    case 'colorScale': {
      d.kind = 'values'
      const pts = rule.points.map((p) => pointOf(p.value, p.color))
      if (pts.length === 3) {
        d.values = 'scale3'
        d.scale3 = pts
      } else {
        d.values = 'scale2'
        d.scale2 = pts.slice(0, 2)
      }
      break
    }
    case 'dataBar':
      Object.assign(d, {
        kind: 'values',
        values: 'dataBar',
        barMin: pointOf(rule.min, ''),
        barMax: pointOf(rule.max, ''),
        barColor: rule.color,
        barGradient: rule.gradient !== false,
        barBorder: rule.border !== undefined,
        barBorderColor: rule.border ?? rule.color,
        barOnly: rule.showValue === false,
        barDirection: rule.direction ?? 'context',
      })
      break
    case 'iconSet':
      Object.assign(d, { kind: 'values', values: 'iconSet', iconSet: rule.icons, iconReverse: !!rule.reverse, iconOnly: rule.showValue === false, iconPoints: iconPointsOf(rule.icons, rule.points) })
      break
    case 'cellIs':
      Object.assign(d, { kind: 'contain', contain: 'cellValue', operator: rule.operator, value1: `=${shown(rule.formulas[0] ?? '')}`, value2: rule.formulas[1] === undefined ? '' : `=${shown(rule.formulas[1])}` })
      break
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const box = textBox(rule)
      Object.assign(d, { kind: 'contain', contain: 'text', textOp: rule.type, text: box.startsWith('=') ? `=${shown(box.slice(1))}` : box })
      break
    }
    case 'timePeriod':
      Object.assign(d, { kind: 'contain', contain: 'dates', period: rule.period })
      break
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      Object.assign(d, { kind: 'contain', contain: (Object.keys(CONTAIN_TYPES) as ContainKind[]).find((k) => CONTAIN_TYPES[k] === rule.type)! })
      break
    case 'top10':
      Object.assign(d, { kind: 'rank', top: !rule.bottom, rank: String(rule.rank), percent: !!rule.percent })
      break
    case 'aboveAverage': {
      const avg = AVERAGE_KINDS.find((k) => {
        const a = AVERAGE_OF[k]
        return !!a.below === !!rule.below && !!a.equal === (!!rule.equal && !rule.stdDev) && (a.stdDev ?? 0) === (rule.stdDev ?? 0)
      })
      Object.assign(d, { kind: 'average', average: avg ?? 'above' })
      break
    }
    case 'duplicateValues':
    case 'uniqueValues':
      Object.assign(d, { kind: 'unique', unique: rule.type === 'uniqueValues' })
      break
    case 'expression':
      Object.assign(d, { kind: 'formula', formula: `=${shown(rule.formula)}` })
      break
  }
  return d
}

/** Why OK cannot write the rule: a box without a value, a formula that does not parse, a relative threshold, a number out of range. */
export type DraftError = { key: 'value' } | { key: 'formula'; message?: string } | { key: 'relative' } | { key: 'percent' } | { key: 'rank'; max: number }
export type DraftResult = { ok: true; rule: ConditionalFormat } | { ok: false; error: DraftError }

/** What OK needs besides the inputs: the rule's ranges, the locale, `xlcalc/entry`, and how a typed formula is written for the rule's first cell. */
export interface DraftContext {
  ranges: string[]
  locale: Locale | undefined
  entry: Entry
  /** A formula typed relative to the active cell, rewritten for the rule's first cell. */
  stored: (formula: string) => string
  /** The active cell's address, which a text rule's formula is typed for. */
  active: string
}

class Refusal {
  constructor(readonly error: DraftError) {}
}
const refuse = (error: DraftError): never => {
  throw new Refusal(error)
}

/** A threshold's box as the rule stores it: no value for the lowest and the highest; a number from 0 to 100 for a percent; an absolute formula. */
function valueOf(p: { type: ValueType; value: string }, c: DraftContext): ConditionalValue {
  if (p.type === 'min' || p.type === 'max' || p.type === 'autoMin' || p.type === 'autoMax') return { type: p.type }
  const value = p.value.trim().replace(/^=/, '')
  if (!value) refuse({ key: 'value' })
  if (hasRelative(value, scanOf(c.locale))) refuse({ key: 'relative' })
  const n = Number(value.replace(c.locale?.decimalSeparator ?? '.', '.'))
  if ((p.type === 'percent' || p.type === 'percentile') && Number.isFinite(n) && (n < 0 || n > 100)) refuse({ key: 'percent' })
  return { type: p.type, value }
}

/** A value box as the rule stores it, written for the rule's first cell. */
function storedBox(text: string, c: DraftContext): string {
  const got = boxFormula(text, c.entry, c.locale)
  if (!got.ok) return refuse(got.message ? { key: 'formula', message: got.message } : { key: 'value' })
  return text.trim().startsWith('=') ? c.stored(got.formula) : got.formula
}

/** The rule's own fields from the inputs. */
function build(d: RuleDraft, c: DraftContext): ConditionalFormat {
  const ranges = c.ranges
  const style = d.style
  if (d.kind === 'values') {
    if (d.values === 'scale2' || d.values === 'scale3') {
      const pts = d.values === 'scale2' ? d.scale2 : d.scale3
      return { type: 'colorScale', ranges, points: pts.map((p) => ({ value: valueOf(p, c), color: p.color })) }
    }
    if (d.values === 'dataBar') {
      return {
        type: 'dataBar',
        ranges,
        min: valueOf(d.barMin, c),
        max: valueOf(d.barMax, c),
        color: d.barColor,
        ...(d.barOnly ? { showValue: false } : {}),
        ...(d.barGradient ? {} : { gradient: false }),
        ...(d.barBorder ? { border: d.barBorderColor } : {}),
        ...(d.barDirection === 'context' ? {} : { direction: d.barDirection }),
      }
    }
    const first = { value: { type: 'percent' as const, value: '0' } }
    return {
      type: 'iconSet',
      ranges,
      icons: d.iconSet,
      points: [first, ...d.iconPoints.map((p) => ({ value: valueOf(p, c), ...(p.gte ? {} : { gte: false }) }))],
      ...(d.iconReverse ? { reverse: true } : {}),
      ...(d.iconOnly ? { showValue: false } : {}),
    }
  }
  if (d.kind === 'contain') {
    if (d.contain === 'cellValue') {
      const two = d.operator === 'between' || d.operator === 'notBetween'
      return { type: 'cellIs', ranges, operator: d.operator, formulas: two ? [storedBox(d.value1, c), storedBox(d.value2, c)] : [storedBox(d.value1, c)], style }
    }
    if (d.contain === 'text') {
      const t = d.text.trim()
      if (!t) refuse({ key: 'value' })
      if (!t.startsWith('=')) return { type: d.textOp, ranges, text: d.text, style }
      const ref = boxFormula(t, c.entry, c.locale)
      if (!ref.ok) return refuse(ref.message ? { key: 'formula', message: ref.message } : { key: 'value' })
      return { type: d.textOp, ranges, text: c.stored(ref.formula), formula: c.stored(textRuleFormula(d.textOp, ref.formula, c.active, c.locale)), style }
    }
    if (d.contain === 'dates') return { type: 'timePeriod', ranges, period: d.period, style }
    return { type: CONTAIN_TYPES[d.contain] as 'containsBlanks', ranges, style }
  }
  if (d.kind === 'rank') {
    const max = d.percent ? 100 : 1000
    const rank = /^\d+$/.test(d.rank.trim()) ? Number(d.rank.trim()) : NaN
    if (!(rank >= 1 && rank <= max)) refuse({ key: 'rank', max })
    return { type: 'top10', ranges, rank, ...(d.percent ? { percent: true } : {}), ...(d.top ? {} : { bottom: true }), style }
  }
  if (d.kind === 'average') return { type: 'aboveAverage', ranges, ...AVERAGE_OF[d.average], style }
  if (d.kind === 'unique') return { type: d.unique ? 'uniqueValues' : 'duplicateValues', ranges, style }
  const f = d.formula.trim()
  if (!f || f === '=') refuse({ key: 'value' })
  return { type: 'expression', ranges, formula: storedBox(f.startsWith('=') ? f : `=${f}`, c), style }
}

/** Fields of a data bar that the dialog does not show, kept when an existing data bar is edited (an icon set's custom icons are kept in `ruleOf`). */
const KEPT: Partial<Record<ConditionalFormat['type'], string[]>> = {
  dataBar: ['minLength', 'maxLength', 'negativeColor', 'negativeBorder', 'axis', 'axisColor', 'excel2007'],
}

/** The rule OK writes from the inputs, or why it cannot; an edited rule keeps its Stop If True and the fields the dialog does not show. */
export function ruleOf(d: RuleDraft, base: ConditionalFormat | null, c: DraftContext): DraftResult {
  let rule: ConditionalFormat
  try {
    rule = build(d, c)
  } catch (e) {
    if (e instanceof Refusal) return { ok: false, error: e.error }
    throw e
  }
  if (!base) return { ok: true, rule: rule.type === 'dataBar' ? { ...rule, negativeColor: '#FF0000', axisColor: '#000000', ...(rule.border ? { negativeBorder: '#FF0000' } : {}) } : rule }
  const kept: Record<string, unknown> = {}
  if (base.type === rule.type) for (const k of KEPT[rule.type] ?? []) if (k in base) kept[k] = (base as unknown as Record<string, unknown>)[k]
  // custom icons of an icon set stay with their thresholds while the set is the same
  if (base.type === 'iconSet' && rule.type === 'iconSet' && base.icons === rule.icons) rule.points = rule.points.map((p, i) => (base.points[i]?.icon !== undefined ? { ...p, icon: base.points[i]!.icon } : p))
  return { ok: true, rule: { ...rule, ...kept, ...(base.stopIfTrue ? { stopIfTrue: true } : {}) } as ConditionalFormat }
}

/** A rule for the dialog's preview, from the inputs as they are: no thresholds checked, no formulas read. */
export function previewOf(d: RuleDraft): ConditionalFormat {
  if (d.kind !== 'values') return { type: 'expression', ranges: [], formula: '', style: d.style }
  if (d.values === 'dataBar')
    return { type: 'dataBar', ranges: [], min: { type: 'min' }, max: { type: 'max' }, color: d.barColor, ...(d.barGradient ? {} : { gradient: false }), ...(d.barBorder ? { border: d.barBorderColor } : {}) }
  if (d.values === 'iconSet') return { ...iconSetRule([], d.iconSet), ...(d.iconReverse ? { reverse: true } : {}) }
  const pts = d.values === 'scale2' ? d.scale2 : d.scale3
  return { type: 'colorScale', ranges: [], points: pts.map((p) => ({ value: { type: 'min' }, color: p.color })) }
}
