// A rule as the Rules Manager's Rule column describes it: the condition in words, its formulas as the active cell sees them.
import type { ConditionalFormat } from 'xlcalc/contract'
import type { SheetMessages } from '../contract/messages'

/** A number or a text constant reads as itself; any other formula after `=`, as Excel's boxes show it. */
const operand = (f: string, shown: (f: string) => string) => (/^-?[\d.,]+$|^".*"$/.test(f) ? f : `=${shown(f)}`)

/** The Rule column's text for a rule; `shown` gives a formula relative to the active cell. */
export function describeRule(rule: ConditionalFormat, m: SheetMessages, shown: (formula: string) => string): string {
  switch (rule.type) {
    case 'cellIs':
      return m.cfDescCell(rule.operator, operand(rule.formulas[0] ?? '', shown), rule.formulas[1] === undefined ? undefined : operand(rule.formulas[1], shown))
    case 'expression':
      return m.cfDescFormula(`=${shown(rule.formula)}`)
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return m.cfDescText(rule.type, rule.text)
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return m.cfDescBlanks(rule.type)
    case 'timePeriod':
      return m.cfPeriod(rule.period)
    case 'top10':
      return m.cfDescTop(rule.rank, !!rule.percent, !!rule.bottom)
    case 'aboveAverage':
      return m.cfDescAverage(!!rule.below, !!rule.equal, rule.stdDev)
    case 'duplicateValues':
    case 'uniqueValues':
      return m.cfDescDuplicate(rule.type === 'uniqueValues')
    case 'colorScale':
      return m.cfDescScale
    case 'dataBar':
      return m.cfDescBar
    case 'iconSet':
      return m.cfDescIcons
  }
}
