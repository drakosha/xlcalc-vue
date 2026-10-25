// Conditional formatting rules without the dialogs: Clear Rules from Selected Cells cutting ranges, formulas moved for a rule's
// first cell, the Applies to box, the value boxes, Excel's presets, and the rule dialog's inputs read from a rule and written back.
import { describe, expect, it } from 'vitest'
import type { ConditionalFormat } from 'xlcalc/contract'
import * as entry from 'xlcalc/entry'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { draftOf, newDraft, pickIconSet, ruleOf, type DraftContext } from '../src/core/cf-draft'
import { describeRule } from '../src/core/cf-describe'
import { boxFormula, hasRelative, shiftFormula, shiftRule, textRuleFormula } from '../src/core/cf-formula'
import { colorScaleRule, dataBarRule, iconSetRule, PRESET_STYLES, quickRuleOf } from '../src/core/cf-presets'
import { appliesText, clearRulesIn, parseApplies, subtract, topLeftOf } from '../src/core/cf-ranges'
import { enMessages, ruMessages } from '../src'

const rect = (top: number, left: number, bottom: number, right: number) => ({ top, left, bottom, right })
const shift = (r: ConditionalFormat, rows: number, cols: number) => shiftRule(r, rows, cols, {})
const fill = { style: { fill: { color: '#FFC7CE' } } }

describe('Clear Rules from Selected Cells', () => {
  it('cuts the selection out of a rectangle: the rows above and below across it, then its left and right parts', () => {
    expect(subtract(rect(2, 2, 4, 4), rect(3, 3, 3, 3))).toEqual([rect(2, 2, 2, 4), rect(3, 2, 3, 2), rect(3, 4, 3, 4), rect(4, 2, 4, 4)])
    expect(subtract(rect(1, 1, 5, 1), rect(7, 1, 9, 1))).toEqual([rect(1, 1, 5, 1)])
  })

  it('keeps the rest of a rule, removes one with nothing left, and leaves the others alone', () => {
    const rules: ConditionalFormat[] = [
      { type: 'cellIs', ranges: ['A1:A10'], operator: 'greaterThan', formulas: ['5'], ...fill },
      { type: 'duplicateValues', ranges: ['A3:A4'], ...fill },
      { type: 'colorScale', ranges: ['C1:C9'], points: colorScaleRule([], 'whiteRed').points },
      { type: 'top10', ranges: ['F:F', 'H2'], rank: 3, ...fill },
    ]
    const out = clearRulesIn(rules, rect(3, 1, 5, 6), shift)!
    expect(out.map((r) => r.ranges)).toEqual([['A1:A2', 'A6:A10'], ['C1:C2', 'C6:C9'], ['F1:F2', 'F6:F1048576', 'H2']])
    expect(clearRulesIn(rules, rect(20, 10, 30, 12), shift)).toBeNull()
  })

  it('writes the formulas of a rule whose first cell moved for the new one, so each cell keeps its condition', () => {
    const rules: ConditionalFormat[] = [{ type: 'expression', ranges: ['A1:B10'], formula: '$C1>A$1+B2', ...fill }]
    const out = clearRulesIn(rules, rect(1, 1, 2, 2), shift)!
    expect(out[0]).toMatchObject({ ranges: ['A3:B10'], formula: '$C3>A$1+B4' })
    expect(topLeftOf(['J1:J3', 'H3:H6'])).toEqual({ row: 1, col: 8 })
  })
})

describe('formulas for the rule’s first cell', () => {
  it('moves relative references only, round the sheet’s edge as a rule computes them, in any locale', () => {
    expect(shiftFormula('A1+$B$2+C$3+$D4+"A1"', 2, 1, {})).toBe('B3+$B$2+D$3+$D6+"A1"')
    expect(shiftFormula('A1>0', -1, -1, {})).toBe('XFD1048576>0')
    expect(shiftFormula("SUM('My Sheet'!A1:B2,C:C,3:3)", 1, 1, {})).toBe("SUM('My Sheet'!B2:C3,D:D,4:4)")
    expect(shiftFormula('СУММ(A1;$B1)>1,5', 1, 0, { locale: ruRU })).toBe('СУММ(A2;$B2)>1,5')
  })

  it('tells a relative reference from an absolute one', () => {
    expect(hasRelative('$A$1*2', {})).toBe(false)
    expect(hasRelative('$A1', {})).toBe(true)
    expect(hasRelative('$A:$A', {})).toBe(false)
    expect(hasRelative('10', {})).toBe(false)
  })

  it('writes a text rule’s formula for a reference in the locale', () => {
    expect(textRuleFormula('containsText', '$A$1', 'B2', undefined)).toBe('NOT(ISERROR(SEARCH($A$1,B2)))')
    expect(textRuleFormula('beginsWith', '$A$1', 'B2', ruRU)).toBe('ЛЕВСИМВ(B2;ДЛСТР($A$1))=$A$1')
  })
})

describe('a value box', () => {
  it('stores a number, a date and a boolean as such, text in quotes, a formula after =', () => {
    const box = (t: string, l = undefined as typeof ruRU | undefined) => boxFormula(t, entry, l)
    expect(box('100')).toEqual({ ok: true, formula: '100' })
    expect(box('1/1/2020')).toEqual({ ok: true, formula: '43831' })
    expect(box('1,5', ruRU)).toEqual({ ok: true, formula: '1,5' })
    expect(box('ИСТИНА', ruRU)).toEqual({ ok: true, formula: 'ИСТИНА' })
    expect(box('say "hi"')).toEqual({ ok: true, formula: '"say ""hi"""' })
    expect(box('=$A$1*2')).toEqual({ ok: true, formula: '$A$1*2' })
    expect(box(' ')).toEqual({ ok: false })
    expect(box('=1+*2')).toMatchObject({ ok: false, message: expect.any(String) })
  })
})

describe('the Applies to box', () => {
  it('shows the ranges with $ and the locale’s separator, and takes them back with or without', () => {
    expect(appliesText(['A1:B5', 'F:F', '3:5', 'C2'], ',')).toBe('=$A$1:$B$5,$F:$F,$3:$5,$C$2')
    expect(parseApplies('=$A$1:$B$5;F:F', ';', 'S')).toEqual(['A1:B5', 'F:F'])
    expect(parseApplies("='S'!$A$1", ',', 's')).toEqual(['A1'])
    expect(parseApplies('=Other!A1', ',', 'S')).toBeNull()
    expect(parseApplies('=A1:', ',', 'S')).toBeNull()
  })
})

describe('Excel’s presets', () => {
  it('a gradient data bar has a border of its color and red negative bars; a solid one has no border', () => {
    expect(dataBarRule(['A1:A5'], 'blue')).toEqual({ type: 'dataBar', ranges: ['A1:A5'], min: { type: 'autoMin' }, max: { type: 'autoMax' }, color: '#638EC6', border: '#638EC6', negativeColor: '#FF0000', negativeBorder: '#FF0000', axisColor: '#000000' })
    expect(dataBarRule(['A1:A5'], 'orange', 'solid')).toMatchObject({ color: '#FFB628', gradient: false })
    expect(dataBarRule(['A1:A5'], 'orange', 'solid').border).toBeUndefined()
  })

  it('a three-color scale has its midpoint at the 50th percentile; the highest numbers get the first color of its name', () => {
    expect(colorScaleRule(['A1:A5'], 'greenYellowRed').points).toEqual([
      { value: { type: 'min' }, color: '#F8696B' },
      { value: { type: 'percentile', value: '50' }, color: '#FFEB84' },
      { value: { type: 'max' }, color: '#63BE7B' },
    ])
    expect(colorScaleRule(['A1'], 'whiteGreen').points.map((p) => p.color)).toEqual(['#63BE7B', '#FCFCFF'])
  })

  it('an icon set gets equal shares in percent', () => {
    expect(iconSetRule([], '3Arrows').points.map((p) => p.value.value)).toEqual(['0', '33', '67'])
    expect(iconSetRule([], '4Rating').points.map((p) => p.value.value)).toEqual(['0', '25', '50', '75'])
    expect(iconSetRule([], '5Boxes').points.map((p) => p.value.value)).toEqual(['0', '20', '40', '60', '80'])
  })

  it('a quick rule is the rule Excel writes', () => {
    const style = PRESET_STYLES.greenDarkGreen
    expect(quickRuleOf('equalTo', ['A1'], { formulas: ['"x"'], style })).toEqual({ type: 'cellIs', ranges: ['A1'], operator: 'equal', formulas: ['"x"'], style })
    expect(quickRuleOf('bottom10Percent', ['A1'], { rank: 5, style })).toEqual({ type: 'top10', ranges: ['A1'], rank: 5, percent: true, bottom: true, style })
    expect(quickRuleOf('belowAverage', ['A1'], { style })).toMatchObject({ type: 'aboveAverage', below: true })
    expect(quickRuleOf('duplicateValues', ['A1'], { unique: true, style }).type).toBe('uniqueValues')
  })
})

describe('the rule dialog’s inputs', () => {
  const ctx = (over: Partial<DraftContext> = {}): DraftContext => ({ ranges: ['A1:A9'], locale: undefined, entry, stored: (f) => f, active: 'A1', ...over })
  const roundTrip = (rule: ConditionalFormat) => ruleOf(draftOf(rule, (f) => f), rule, ctx())

  it('reads every kind of rule back unchanged', () => {
    const rules: ConditionalFormat[] = [
      { type: 'cellIs', ranges: ['A1:A9'], operator: 'between', formulas: ['1', '$B$1'], ...fill, stopIfTrue: true },
      { type: 'expression', ranges: ['A1:A9'], formula: '$B1>2', ...fill },
      { type: 'beginsWith', ranges: ['A1:A9'], text: 'ab', ...fill },
      { type: 'timePeriod', ranges: ['A1:A9'], period: 'lastMonth', ...fill },
      { type: 'notContainsErrors', ranges: ['A1:A9'], ...fill },
      { type: 'top10', ranges: ['A1:A9'], rank: 15, percent: true, ...fill },
      { type: 'aboveAverage', ranges: ['A1:A9'], below: true, stdDev: 2, ...fill },
      { type: 'aboveAverage', ranges: ['A1:A9'], below: true, equal: true, ...fill },
      { type: 'uniqueValues', ranges: ['A1:A9'], ...fill },
      colorScaleRule(['A1:A9'], 'blueWhiteRed'),
      { ...dataBarRule(['A1:A9'], 'green'), showValue: false, axis: 'middle' },
      { ...iconSetRule(['A1:A9'], '4Arrows'), reverse: true, points: iconSetRule([], '4Arrows').points.map((p, i) => (i === 2 ? { ...p, gte: false, icon: null } : p)) },
    ]
    for (const r of rules) expect(roundTrip(r)).toEqual({ ok: true, rule: r })
  })

  it('opens a new rule as Excel does and writes a formula typed for the active cell for the first cell', () => {
    const d = newDraft()
    expect(d.kind).toBe('values')
    d.kind = 'formula'
    d.formula = '=$B5>0'
    expect(ruleOf(d, null, ctx({ stored: (f) => shiftFormula(f, -4, 0, {}) }))).toEqual({ ok: true, rule: { type: 'expression', ranges: ['A1:A9'], formula: '$B1>0', style: {} } })
  })

  it('writes a text rule whose text is a reference with Excel’s formula', () => {
    const d = newDraft()
    Object.assign(d, { kind: 'contain', contain: 'text', textOp: 'containsText', text: '=$Z$1' })
    expect(ruleOf(d, null, ctx())).toMatchObject({ ok: true, rule: { type: 'containsText', text: '$Z$1', formula: 'NOT(ISERROR(SEARCH($Z$1,A1)))' } })
  })

  it('refuses an empty box, a relative threshold, a percent beyond 100 and a rank out of range', () => {
    const d = newDraft()
    Object.assign(d, { kind: 'contain', contain: 'cellValue', operator: 'greaterThan', value1: '' })
    expect(ruleOf(d, null, ctx())).toEqual({ ok: false, error: { key: 'value' } })
    const s = newDraft()
    s.scale2[0] = { type: 'num', value: '=A1', color: '#FF0000' }
    expect(ruleOf(s, null, ctx())).toEqual({ ok: false, error: { key: 'relative' } })
    s.scale2[0] = { type: 'percent', value: '120', color: '#FF0000' }
    expect(ruleOf(s, null, ctx())).toEqual({ ok: false, error: { key: 'percent' } })
    const r = newDraft()
    Object.assign(r, { kind: 'rank', rank: '101', percent: true })
    expect(ruleOf(r, null, ctx())).toEqual({ ok: false, error: { key: 'rank', max: 100 } })
  })

  it('resets the thresholds when an icon style with another number of icons is picked', () => {
    const d = newDraft()
    pickIconSet(d, '5Quarters')
    expect(d.iconPoints.map((p) => p.value)).toEqual(['20', '40', '60', '80'])
  })
})

describe('the Rules Manager’s descriptions', () => {
  it('describe a rule as Excel’s Rule column, formulas after =', () => {
    const shown = (f: string) => f
    expect(describeRule({ type: 'cellIs', ranges: [], operator: 'greaterThanOrEqual', formulas: ['5'] }, enMessages, shown)).toBe('Cell Value >= 5')
    expect(describeRule({ type: 'cellIs', ranges: [], operator: 'between', formulas: ['1', '$B$1'] }, enMessages, shown)).toBe('Cell Value between 1 and =$B$1')
    expect(describeRule({ type: 'expression', ranges: [], formula: '$A1>5' }, enMessages, shown)).toBe('Formula: =$A1>5')
    expect(describeRule({ type: 'top10', ranges: [], rank: 10, percent: true, bottom: true }, enMessages, shown)).toBe('Bottom 10%')
    expect(describeRule(colorScaleRule([], 'redWhite'), ruMessages, shown)).toBe('Градиентная цветовая шкала')
  })
})
