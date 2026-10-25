// @vitest-environment node
// Excel's Data Validation dialog without its markup: the boxes a rule fills for the active cell in en-US and ru-RU, the rule OK
// makes from them with Excel's checks, Excel's questions for the selection, and the sheet's rules after OK.
import { describe, expect, it } from 'vitest'
import type { DataValidation } from 'xlcalc/contract'
import * as entry from 'xlcalc/entry'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { shiftFormula } from '../src/formula-shift'
import { boxesOf, draftOf, emptyDraft, planOf, ruleOf, rulesAfter, type RuleText, type ValidationDraft } from '../src/validation-rules'

const en: RuleText = { scan: {}, entry }
const ru: RuleText = { scan: { locale: ruRU }, entry }
const at = (row: number, col: number) => ({ row, col })
const rect = (top: number, left: number, bottom: number, right: number) => ({ top, left, bottom, right })
const draft = (patch: Partial<ValidationDraft>): ValidationDraft => ({ ...emptyDraft(), ...patch })

describe('shiftFormula', () => {
  it('moves relative rows and columns, keeps absolute parts, names and text, in either locale', () => {
    expect(shiftFormula('B2>A2', 3, 0, {})).toBe('B5>A5')
    expect(shiftFormula('$A1+B$2+$C$3+Rate+"A1"', 2, 1, {})).toBe('$A3+C$2+$C$3+Rate+"A1"')
    expect(shiftFormula('СЧЁТЕСЛИ($A$1:$C$20;A1)=1', 1, 2, { locale: ruRU })).toBe('СЧЁТЕСЛИ($A$1:$C$20;C2)=1')
    expect(shiftFormula("SUM(C:C)+'Other sheet'!B2+3:4+E2#", 1, 1, {})).toBe("SUM(D:D)+'Other sheet'!C3+4:5+F3#")
  })
  it('wraps past the edge of the sheet, as relative references of rules do (unverified)', () => {
    expect(shiftFormula('A1', -1, -1, {})).toBe('XFD1048576')
  })
})

describe('the boxes of the Settings tab', () => {
  it('follow the type and the comparison, as Excel names them', () => {
    expect(boxesOf('whole', 'between')).toEqual(['min', 'max'])
    expect(boxesOf('decimal', 'greaterThan')).toEqual(['min'])
    expect(boxesOf('decimal', 'lessThanOrEqual')).toEqual(['max'])
    expect(boxesOf('textLength', 'equal')).toEqual(['length'])
    expect(boxesOf('date', 'notBetween')).toEqual(['start', 'end'])
    expect(boxesOf('time', 'notEqual')).toEqual(['time'])
    expect(boxesOf('list', 'between')).toEqual(['source'])
    expect(boxesOf('custom', 'between')).toEqual(['formula'])
    expect(boxesOf('any', 'between')).toEqual([])
  })
})

describe('draftOf: the boxes a rule fills', () => {
  it("shows a literal list with the locale's list separator, a list of cells as a formula", () => {
    const list: DataValidation = { range: 'A1:A9', type: 'list', formula1: '"Open,On hold,Closed"' }
    expect(draftOf(list, at(1, 1), en).first).toBe('Open,On hold,Closed')
    expect(draftOf(list, at(1, 1), ru).first).toBe('Open;On hold;Closed')
    expect(draftOf({ range: 'A1:A9', type: 'list', formula1: '"say ""hi"",x"' }, at(1, 1), en).first).toBe('say "hi",x')
    expect(draftOf({ range: 'A1:A9', type: 'list', formula1: '$D$2:$D$9' }, at(3, 1), en).first).toBe('=$D$2:$D$9')
  })

  it('shows numbers as typed, dates and times in the locale, other bounds as formulas for the active cell', () => {
    const whole: DataValidation = { range: 'B2:B9', type: 'whole', formula1: '1', formula2: 'C2*2' }
    expect(draftOf(whole, at(5, 2), en)).toMatchObject({ type: 'whole', operator: 'between', first: '1', second: '=C5*2' })
    expect(draftOf({ range: 'B2', type: 'decimal', operator: 'greaterThan', formula1: '1,5' }, at(2, 2), ru)).toMatchObject({ operator: 'greaterThan', first: '1,5', second: '' })
    const date: DataValidation = { range: 'C1', type: 'date', formula1: '45658', formula2: '45688' }
    expect(draftOf(date, at(1, 3), en)).toMatchObject({ first: '1/1/2025', second: '1/31/2025' })
    expect(draftOf(date, at(1, 3), ru)).toMatchObject({ first: '01.01.2025', second: '31.01.2025' })
    expect(draftOf({ range: 'C1', type: 'time', operator: 'lessThan', formula1: '0.375' }, at(1, 3), en).first).toBe('9:00:00 AM')
    expect(draftOf({ range: 'B2:B9', type: 'custom', formula1: 'B2>A2' }, at(5, 2), en).first).toBe('=B5>A5')
  })

  it("takes the rule's checks and messages, Excel's defaults without a rule", () => {
    const rule: DataValidation = { range: 'A1', type: 'list', formula1: '"a"', allowBlank: false, dropdown: false, promptTitle: 'T', prompt: 'P', showPrompt: false, errorStyle: 'warning', errorTitle: 'E', error: 'M', showError: false }
    expect(draftOf(rule, at(1, 1), en)).toEqual({
      type: 'list',
      operator: 'between',
      first: 'a',
      second: '',
      allowBlank: false,
      dropdown: false,
      applyAll: false,
      showPrompt: false,
      promptTitle: 'T',
      prompt: 'P',
      showError: false,
      errorStyle: 'warning',
      errorTitle: 'E',
      error: 'M',
    })
    expect(draftOf(null, at(1, 1), en)).toEqual(emptyDraft())
  })
})

describe('ruleOf: the rule OK makes', () => {
  it("stores a literal list with commas whatever the locale's separator, a formula without its =", () => {
    expect(ruleOf(draft({ type: 'list', first: 'а;б;в' }), ru)).toEqual({ ok: true, rule: { type: 'list', formula1: '"а,б,в"' } })
    expect(ruleOf(draft({ type: 'list', first: 'a,b' }), en)).toEqual({ ok: true, rule: { type: 'list', formula1: '"a,b"' } })
    expect(ruleOf(draft({ type: 'list', first: 'say "hi"' }), en)).toEqual({ ok: true, rule: { type: 'list', formula1: '"say ""hi"""' } })
    expect(ruleOf(draft({ type: 'list', first: '=$D$2:$D$9', dropdown: false }), en)).toEqual({ ok: true, rule: { type: 'list', formula1: '$D$2:$D$9', dropdown: false } })
  })

  it('reads a bound as typed into a cell: a number, a date or a time in the locale, or a formula after =', () => {
    expect(ruleOf(draft({ type: 'whole', first: '1', second: '=$E$1' }), en)).toEqual({ ok: true, rule: { type: 'whole', formula1: '1', formula2: '$E$1' } })
    expect(ruleOf(draft({ type: 'decimal', operator: 'greaterThan', first: '1,5' }), ru)).toEqual({ ok: true, rule: { type: 'decimal', operator: 'greaterThan', formula1: '1,5' } })
    expect(ruleOf(draft({ type: 'date', first: '1/1/2025', second: '1/31/2025' }), en)).toEqual({ ok: true, rule: { type: 'date', formula1: '45658', formula2: '45688' } })
    expect(ruleOf(draft({ type: 'date', operator: 'equal', first: '01.01.2025' }), ru)).toEqual({ ok: true, rule: { type: 'date', operator: 'equal', formula1: '45658' } })
    expect(ruleOf(draft({ type: 'time', operator: 'lessThan', first: '9:00' }), en)).toEqual({ ok: true, rule: { type: 'time', operator: 'lessThan', formula1: '0.375' } })
  })

  it('refuses an empty box, a bound that is not a number, and a maximum below the minimum, naming the boxes', () => {
    expect(ruleOf(draft({ type: 'list', first: ' ' }), en)).toEqual({ ok: false, error: { kind: 'empty', box: 'source' } })
    expect(ruleOf(draft({ type: 'custom' }), en)).toEqual({ ok: false, error: { kind: 'empty', box: 'formula' } })
    expect(ruleOf(draft({ type: 'whole', first: '1' }), en)).toEqual({ ok: false, error: { kind: 'empty', box: 'max' } })
    expect(ruleOf(draft({ type: 'textLength', operator: 'equal', first: 'abc' }), en)).toEqual({ ok: false, error: { kind: 'invalid', box: 'length' } })
    expect(ruleOf(draft({ type: 'whole', first: '10', second: '1' }), en)).toEqual({ ok: false, error: { kind: 'order', min: 'min', max: 'max' } })
    expect(ruleOf(draft({ type: 'date', first: '2/1/2025', second: '1/1/2025' }), en)).toEqual({ ok: false, error: { kind: 'order', min: 'start', max: 'end' } })
    // a formula bound is checked by the workbook, not here
    expect(ruleOf(draft({ type: 'whole', first: '10', second: '=$A$1' }), en).ok).toBe(true)
  })

  it('takes a custom formula after =, and a constant typed without it', () => {
    expect(ruleOf(draft({ type: 'custom', first: '=B2>A2' }), en)).toEqual({ ok: true, rule: { type: 'custom', formula1: 'B2>A2' } })
    expect(ruleOf(draft({ type: 'custom', first: 'истина' }), ru)).toEqual({ ok: true, rule: { type: 'custom', formula1: 'ИСТИНА' } })
    expect(ruleOf(draft({ type: 'custom', first: 'B2>A2' }), en)).toEqual({ ok: true, rule: { type: 'custom', formula1: '"B2>A2"' } })
  })

  it('writes the messages and checks that differ from a new rule, truncated to Excel’s limits; Any value without a message is no rule', () => {
    const long = 'x'.repeat(300)
    const made = ruleOf(draft({ type: 'any', promptTitle: long, prompt: long, showPrompt: false, errorStyle: 'information', errorTitle: 'E', error: long, showError: false, allowBlank: false }), en)
    expect(made).toEqual({
      ok: true,
      rule: { type: 'any', showPrompt: false, promptTitle: 'x'.repeat(32), prompt: 'x'.repeat(255), showError: false, errorStyle: 'information', errorTitle: 'E', error: 'x'.repeat(225) },
    })
    expect(ruleOf(draft({ type: 'whole', first: '1', second: '2', allowBlank: false }), en)).toEqual({ ok: true, rule: { type: 'whole', formula1: '1', formula2: '2', allowBlank: false } })
    expect(ruleOf(emptyDraft(), en)).toEqual({ ok: true, rule: null })
  })
})

describe("planOf: Excel's question before the dialog opens", () => {
  const rules: DataValidation[] = [
    { range: 'A1:A5', type: 'list', formula1: '"a,b"' },
    { range: 'B1:B5', type: 'whole', formula1: '1', formula2: '9' },
    { range: 'C1:C2 C4', type: 'list', formula1: '"a,b"' },
  ]
  it('finds no rule, one over every cell, one over some cells, or several kinds', () => {
    expect(planOf(rules, rect(1, 5, 3, 6), at(1, 5), {})).toEqual({ kind: 'none' })
    expect(planOf(rules, rect(2, 1, 3, 1), at(2, 1), {})).toEqual({ kind: 'one', rule: rules[0] })
    expect(planOf(rules, rect(4, 1, 7, 1), at(4, 1), {})).toEqual({ kind: 'extend', rule: rules[0] })
    expect(planOf(rules, rect(1, 1, 1, 2), at(1, 1), {})).toEqual({ kind: 'many' })
  })
  it('counts rules with the same settings as one kind, formulas compared as for their own cells', () => {
    expect(planOf(rules, rect(1, 1, 2, 1), at(1, 1), {}).kind).toBe('one')
    expect(planOf([rules[0]!, rules[2]!], rect(1, 1, 5, 3), at(1, 1), {})).toEqual({ kind: 'extend', rule: rules[0] })
    const custom: DataValidation[] = [
      { range: 'A1:A2', type: 'custom', formula1: 'A1>0' },
      { range: 'B5:B6', type: 'custom', formula1: 'B5>0' },
    ]
    expect(planOf(custom, rect(1, 1, 6, 2), at(1, 1), {}).kind).toBe('extend')
    expect(planOf([custom[0]!, { range: 'B5', type: 'custom', formula1: 'A5>0' }], rect(1, 1, 6, 2), at(1, 1), {}).kind).toBe('many')
  })
})

describe('rulesAfter: the sheet’s rules after OK', () => {
  it('splits a rule the selection covers part of, its formulas written for its new top-left cell', () => {
    const rules: DataValidation[] = [{ range: 'B2:B9', type: 'custom', formula1: 'B2>A2' }, { range: 'D1', type: 'list', formula1: '"x"' }]
    const out = rulesAfter(rules, rect(2, 2, 4, 2), at(2, 2), { type: 'whole', formula1: '1', formula2: '9' }, null, {})
    expect(out.rules).toEqual([
      { range: 'B5:B9', type: 'custom', formula1: 'B5>A5' },
      { range: 'D1', type: 'list', formula1: '"x"' },
      { range: 'B2:B4', type: 'whole', formula1: '1', formula2: '9' },
    ])
    expect(out.ranges).toEqual(['B2:B4'])
  })

  it('cuts a hole in a rule, and none removes the rule from the selection', () => {
    const out = rulesAfter([{ range: 'A1:C3', type: 'list', formula1: '"a"' }], rect(2, 2, 2, 2), at(2, 2), null, null, {})
    expect(out.rules).toEqual([{ range: 'A1:C1 A3:C3 A2 C2', type: 'list', formula1: '"a"' }])
  })

  it('writes the formulas typed for the active cell for the top-left cell of the selection', () => {
    const out = rulesAfter([], rect(2, 2, 6, 2), at(4, 2), { type: 'custom', formula1: 'B4>$A$1' }, null, {})
    expect(out.rules).toEqual([{ range: 'B2:B6', type: 'custom', formula1: 'B2>$A$1' }])
  })

  it('with Apply these changes to all other cells, takes in every rule with the same settings', () => {
    const rules: DataValidation[] = [
      { range: 'A1:A3', type: 'custom', formula1: 'A1>0' },
      { range: 'C5', type: 'custom', formula1: 'C5>0' },
      { range: 'E1', type: 'custom', formula1: 'E2>0' },
    ]
    const out = rulesAfter(rules, rect(2, 1, 2, 1), at(2, 1), { type: 'custom', formula1: 'A2<0' }, rules[0]!, {})
    expect(out.rules).toEqual([rules[2], { range: 'A2 A1 A3 C5', type: 'custom', formula1: 'A1<0' }])
    expect(out.ranges).toEqual(['A2', 'A1', 'A3', 'C5'])
  })
})
