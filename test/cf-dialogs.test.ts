// Excel's Conditional Formatting menu on the sheet over localSource: the quick rules' dialogs with their presets and Custom
// Format, the galleries, Clear Rules, New Rule with formulas typed for the active cell, and the Rules Manager.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput, ConditionalFormat } from 'xlcalc/contract'
import { memorySource, ruMessages, type SheetViewApi } from '../src'
import { goTo, key, mountSheet, settle, until } from './helpers'
import { closeAll, source } from './local-fixture'

afterEach(async () => {
  document.body.innerHTML = ''
  await closeAll()
})

const RED = { font: { color: '#9C0006' }, fill: { color: '#FFC7CE' } }
const book = (rules: ConditionalFormat[] = [], locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 }, A5: { value: 5 }, A6: { value: 6 }, B1: { value: 'ab' }, B2: { value: 'cd' } },
      conditionalFormats: rules,
    },
  ],
})

async function open(rules: ConditionalFormat[] = [], locale?: string) {
  const s = source(book(rules, locale))
  const w = await mountSheet(s, locale ? { locale, messages: ruMessages } : {})
  // the workbook is calculated and its first window drawn before a command
  await until(() => w.find('[data-cell="1:1"]').exists() && w.find('[data-cell="1:1"]').text() === '1')
  const api = w.vm as unknown as SheetViewApi
  const rulesOf = async () => (await s.rules('S')).conditionalFormats
  return { s, w, api, rulesOf }
}
const bg = (w: VueWrapper, r: number, c: number) => getComputedStyle(w.find(`[data-cell="${r}:${c}"]`).element).backgroundColor
const field = (w: VueWrapper, sel: string) => w.find(sel).element as HTMLInputElement

describe('the Conditional Formatting menu', () => {
  it('needs rules, setRules and an editable sheet: otherwise each method returns false or null', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    const api = w.vm as unknown as SheetViewApi
    expect(api.quickRule('greaterThan')).toBe(false)
    expect(api.manageRules()).toBe(false)
    expect(await api.dataBars('blue')).toBeNull()
    const ro = await mountSheet(source(book()), { readOnly: true })
    const roApi = ro.vm as unknown as SheetViewApi
    expect(roApi.newRule()).toBe(false)
    expect(await roApi.clearRules('sheet')).toBeNull()
  })

  it('Greater Than writes a cell value rule over the selection with Light Red Fill with Dark Red Text, first in priority, and the cells show it', async () => {
    const { w, api, rulesOf } = await open([{ type: 'duplicateValues', ranges: ['B1:B2'], style: { fill: { color: '#00FF00' } } }])
    await goTo(w, 'A1:A6')
    expect(api.quickRule('greaterThan')).toBe(true)
    await settle()
    expect(w.find('[data-cf-quick="greaterThan"]').text()).toContain('Format cells that are GREATER THAN:')
    await w.find('[data-cf-value="0"]').setValue('3')
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => bg(w, 4, 1) === '#FFC7CE')
    const rules = await rulesOf()
    expect(rules[0]).toEqual({ type: 'cellIs', ranges: ['A1:A6'], operator: 'greaterThan', formulas: ['3'], style: RED })
    expect(rules[1]).toMatchObject({ type: 'duplicateValues' })
    expect(bg(w, 3, 1)).not.toBe('#FFC7CE')
    expect(w.find('[data-cf-quick]').exists()).toBe(false)
  })

  it('reads the windows over the rules again after setRules, without an event from the source', async () => {
    const s = source(book())
    const { subscribe: _, ...quiet } = s
    const w = await mountSheet(quiet)
    await goTo(w, 'A1:A6')
    expect(await (w.vm as unknown as SheetViewApi).colorScale('whiteRed')).toMatchObject({ ok: true })
    await until(() => bg(w, 1, 1) === '#F8696B' && bg(w, 6, 1) === '#FCFCFF')
  })

  it('Between takes text in quotes and a reference after =; an empty box is refused', async () => {
    const { w, api, rulesOf } = await open()
    await goTo(w, 'B1:B2')
    api.quickRule('between')
    await settle()
    await w.find('[data-cf-quick]').trigger('submit')
    await settle()
    expect(w.find('.xs-cf-error').text()).toBe('Enter a value or a formula')
    await w.find('[data-cf-value="0"]').setValue('a')
    await w.find('[data-cf-value="1"]').setValue('=$A$1')
    await w.find('[data-cf-preset]').setValue('greenDarkGreen')
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => !w.find('[data-cf-quick]').exists())
    await settle()
    expect((await rulesOf())[0]).toMatchObject({ operator: 'between', formulas: ['"a"', '$A$1'], style: { font: { color: '#006100' }, fill: { color: '#C6EFCE' } } })
  })

  it('Custom Format opens the rule’s Format Cells: Font, Border and Fill, and the rule gets what was set there', async () => {
    const { w, api, rulesOf } = await open()
    await goTo(w, 'A1:A3')
    api.quickRule('top10')
    await settle()
    expect(field(w, '[data-cf-rank]').value).toBe('10')
    await w.find('[data-cf-rank]').setValue('2')
    await w.find('[data-cf-preset]').setValue('custom')
    await settle()
    expect(w.findAll('[data-cf-style] [data-fmt-tab]').map((t) => t.attributes('data-fmt-tab'))).toEqual(['font', 'border', 'fill'])
    await w.find('[data-cf-style] [data-fmt-font-style]').setValue('bold')
    await w.find('[data-cf-style] [data-fmt-tab="border"]').trigger('click')
    await w.find('[data-cf-style] [data-fmt-edge="bottom"]').trigger('click')
    await w.find('[data-cf-style] [data-fmt-tab="fill"]').trigger('click')
    await w.find('[data-cf-style] [data-fmt-colour-of="fill"] [data-fmt-colour-code]').setValue('#00ff00')
    await w.find('[data-cf-style]').trigger('submit')
    await settle()
    expect(w.find('[data-cf-style]').exists()).toBe(false)
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => !w.find('[data-cf-quick]').exists())
    await settle()
    expect((await rulesOf())[0]).toEqual({ type: 'top10', ranges: ['A1:A3'], rank: 2, style: { font: { bold: true }, fill: { color: '#00FF00' }, border: { bottom: { style: 'thin' } } } })
  })

  it('a quick dialog for a text, a date, duplicates and an average writes the rule Excel writes', async () => {
    const { s, w, api, rulesOf } = await open()
    await goTo(w, 'B1:B2')
    api.quickRule('textContains')
    await settle()
    await w.find('[data-cf-text]').setValue('b')
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => !w.find('[data-cf-quick]').exists())
    api.quickRule('dateOccurring')
    await settle()
    await w.find('[data-cf-period]').setValue('nextMonth')
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => !w.find('[data-cf-quick]').exists())
    api.quickRule('duplicateValues')
    await settle()
    await w.find('[data-cf-unique]').setValue('true')
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => !w.find('[data-cf-quick]').exists())
    api.quickRule('belowAverage')
    await settle()
    expect(w.find('[data-cf-quick]').text()).toContain('for the selected range with')
    await w.find('[data-cf-quick]').trigger('submit')
    await until(() => s.history?.undo === 4)
    const rules = await rulesOf()
    expect(rules.map((r) => r.type)).toEqual(['aboveAverage', 'uniqueValues', 'timePeriod', 'containsText'])
    expect(rules[0]).toMatchObject({ below: true })
    expect(rules[2]).toMatchObject({ period: 'nextMonth' })
    expect(rules[3]).toMatchObject({ text: 'b' })
  })

  it('the galleries add a data bar, a color scale and an icon set over the selection', async () => {
    const { w, api, rulesOf } = await open()
    await goTo(w, 'A1:A6')
    expect(await api.dataBars('green', 'solid')).toMatchObject({ ok: true })
    expect(await api.colorScale('redWhiteBlue')).toMatchObject({ ok: true })
    expect(await api.iconSet('3Arrows')).toMatchObject({ ok: true })
    const rules = await rulesOf()
    expect(rules.map((r) => r.type)).toEqual(['iconSet', 'colorScale', 'dataBar'])
    expect(rules[2]).toMatchObject({ ranges: ['A1:A6'], color: '#63C384', gradient: false })
    await until(() => w.find('[data-cell="6:1"] [data-icon="3Arrows:2"]').exists())
  })
})

describe('Clear Rules', () => {
  const two: ConditionalFormat[] = [
    { type: 'expression', ranges: ['A1:A6'], formula: 'A1>2', style: RED },
    { type: 'uniqueValues', ranges: ['B1:B2'], style: RED },
  ]

  it('from Selected Cells: a rule keeps the rest of its cells, its formula written for its new first cell', async () => {
    const { w, api, rulesOf } = await open(two)
    await goTo(w, 'A1:B3')
    await until(() => bg(w, 3, 1) === '#FFC7CE')
    expect(await api.clearRules('selection')).toMatchObject({ ok: true })
    expect(await rulesOf()).toEqual([{ type: 'expression', ranges: ['A4:A6'], formula: 'A4>2', style: RED }])
    await until(() => bg(w, 3, 1) !== '#FFC7CE' && bg(w, 4, 1) === '#FFC7CE')
    await goTo(w, 'D1')
    expect(await api.clearRules('selection')).toBeNull()
  })

  it('from Entire Sheet: every rule goes; with none there is nothing to send', async () => {
    const { w, api, rulesOf } = await open(two)
    await goTo(w, 'D1')
    expect(await api.clearRules('sheet')).toMatchObject({ ok: true })
    expect(await rulesOf()).toEqual([])
    expect(await api.clearRules('sheet')).toBeNull()
    expect(w.find('.xs-message').exists()).toBe(false)
  })
})

describe('New Formatting Rule', () => {
  it('takes a formula for the active cell and writes it for the selection’s first cell, as Excel does', async () => {
    const { w, api, rulesOf } = await open()
    await goTo(w, 'A3')
    await key(w, 'ArrowUp', { shiftKey: true })
    await settle()
    expect(api.newRule()).toBe(true)
    await settle()
    expect(w.find('[data-cf-kind="values"]').attributes('aria-selected')).toBe('true')
    await w.find('[data-cf-kind="formula"]').trigger('click')
    await w.find('[data-cf-formula]').setValue('=$A3>2')
    await w.find('[data-cf-format]').trigger('click')
    await settle()
    await w.find('[data-cf-style] [data-fmt-font-style]').setValue('italic')
    await w.find('[data-cf-style]').trigger('submit')
    await settle()
    expect(w.find('[data-cf-rule] [data-cf-preview="style"]').exists()).toBe(true)
    await w.find('[data-cf-rule]').trigger('submit')
    await until(() => !w.find('[data-cf-rule]').exists())
    await settle()
    expect((await rulesOf())[0]).toEqual({ type: 'expression', ranges: ['A2:A3'], formula: '$A2>2', style: { font: { italic: true } } })
  })

  it('a relative reference in a color scale’s threshold is refused with Excel’s message', async () => {
    const { w, api } = await open()
    await goTo(w, 'A1:A6')
    api.newRule()
    await settle()
    await w.find('[data-cf-type="min"]').setValue('num')
    await w.find('[data-cf-point-value="min"]').setValue('=A1')
    await w.find('[data-cf-rule]').trigger('submit')
    await settle()
    expect(w.find('[data-cf-rule] .xs-cf-error').text()).toContain('relative references')
  })
})

describe('the Rules Manager', () => {
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['A1:A6'], operator: 'greaterThan', formulas: ['4'], style: RED },
    { type: 'colorScale', ranges: ['A1:A6'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'max' }, color: '#63BE7B' }] },
    { type: 'uniqueValues', ranges: ['B1:B2'], style: { font: { bold: true } } },
  ]
  const rows = (w: VueWrapper) => w.findAll('[data-cf-manager] [data-cf-row]')
  const descs = (w: VueWrapper) => w.findAll('[data-cf-manager] [data-cf-desc]').map((d) => d.text())

  it('lists the rules of the selection or of the sheet, each with its description, format, Applies to and Stop If True; Duplicate Rule copies one', async () => {
    const { w, api } = await open(rules)
    await goTo(w, 'A2')
    expect(api.manageRules()).toBe(true)
    await until(() => rows(w).length > 0)
    expect(descs(w)).toEqual(['Cell Value > 4', 'Graded Color Scale'])
    await w.find('[data-cf-show]').setValue('sheet')
    expect(descs(w)).toEqual(['Cell Value > 4', 'Graded Color Scale', 'Unique Values'])
    expect(w.findAll('[data-cf-applies]').map((i) => (i.element as HTMLInputElement).value)).toEqual(['=$A$1:$A$6', '=$A$1:$A$6', '=$B$1:$B$2'])
    expect(w.findAll('[data-cf-stop]').map((i) => (i.element as HTMLInputElement).disabled)).toEqual([false, true, false])
    expect(w.findAll('[data-cf-preview]').map((p) => p.attributes('data-cf-preview'))).toEqual(['style', 'scale', 'style'])
    await rows(w)[0]!.trigger('click')
    await w.find('[data-cf-duplicate]').trigger('click')
    expect(descs(w)).toEqual(['Cell Value > 4', 'Cell Value > 4', 'Graded Color Scale', 'Unique Values'])
    expect(rows(w)[1]!.attributes('aria-selected')).toBe('true')
  })

  it('moves, edits, deletes and sets Stop If True; Close discards, Apply and OK write the list in one call', async () => {
    const { s, w, api, rulesOf } = await open(rules)
    await goTo(w, 'A2')
    api.manageRules()
    await until(() => rows(w).length > 0)
    await w.find('[data-cf-show]').setValue('sheet')
    await rows(w)[2]!.trigger('click')
    await w.find('[data-cf-up]').trigger('click')
    expect(descs(w)).toEqual(['Cell Value > 4', 'Unique Values', 'Graded Color Scale'])
    await w.find('[data-cf-close]').trigger('click')
    expect(w.find('[data-cf-manager]').exists()).toBe(false)
    expect((await rulesOf()).map((r) => r.type)).toEqual(['cellIs', 'colorScale', 'uniqueValues'])

    api.manageRules()
    await until(() => rows(w).length > 0)
    await w.find('[data-cf-show]').setValue('sheet')
    const before = s.history?.undo ?? 0
    await rows(w)[0]!.trigger('click')
    await w.find('[data-cf-down]').trigger('click')
    await w.findAll('[data-cf-stop]')[1]!.setValue(true)
    await w.findAll('[data-cf-applies]')[2]!.setValue('=B1:B9')
    await w.find('[data-cf-apply]').trigger('click')
    await until(() => (s.history?.undo ?? 0) === before + 1)
    await settle()
    expect(await rulesOf()).toEqual([rules[1], { ...rules[0], stopIfTrue: true }, { ...rules[2], ranges: ['B1:B9'] }])
    expect(w.find('[data-cf-manager]').exists()).toBe(true)

    // Edit Rule: the cell value rule's value; Delete Rule: the unique values rule
    await until(() => rows(w).length === 3)
    await rows(w)[1]!.trigger('click')
    await w.find('[data-cf-edit]').trigger('click')
    await settle()
    expect(field(w, '[data-cf-rule] [data-cf-value="0"]').value).toBe('=4')
    expect(w.find('[data-cf-rule] [data-cf-kind="contain"]').attributes('aria-selected')).toBe('true')
    await w.find('[data-cf-rule] [data-cf-value="0"]').setValue('=5')
    await w.find('[data-cf-rule]').trigger('submit')
    await until(() => !w.find('[data-cf-rule]').exists())
    expect(descs(w)[1]).toBe('Cell Value > 5')
    await rows(w)[2]!.trigger('click')
    await w.find('[data-cf-delete]').trigger('click')
    expect(rows(w)).toHaveLength(2)
    await w.find('[data-cf-ok]').trigger('click')
    await until(() => (s.history?.undo ?? 0) === before + 2)
    expect(w.find('[data-cf-manager]').exists()).toBe(false)
    expect(await rulesOf()).toEqual([rules[1], { ...rules[0], formulas: ['5'], stopIfTrue: true }])
  })

  it('keeps OK waiting while Applies to holds something that is not a range on the sheet', async () => {
    const { w, api } = await open(rules)
    api.manageRules()
    await until(() => rows(w).length > 0)
    await w.findAll('[data-cf-applies]')[0]!.setValue('=Other!A1')
    expect(w.find('[data-cf-manager] .xs-cf-error').exists()).toBe(true)
    expect((w.find('[data-cf-ok]').element as HTMLButtonElement).disabled).toBe(true)
  })

  it('in ru-RU shows Applies to with ; and Russian descriptions, and Edit Rule shows formulas in the locale', async () => {
    const { s, w, api, rulesOf } = await open([{ type: 'cellIs', ranges: ['A1:A2', 'A5'], operator: 'between', formulas: ['1.5', 'SUM($A$1,1)'], style: RED }], 'ru-RU')
    await goTo(w, 'A1')
    api.manageRules()
    await until(() => rows(w).length > 0)
    expect(field(w, '[data-cf-applies]').value).toBe('=$A$1:$A$2;$A$5')
    expect(descs(w)).toEqual(['Значение ячейки между 1,5 и =СУММ($A$1;1)'])
    await rows(w)[0]!.trigger('dblclick')
    await settle()
    expect(field(w, '[data-cf-rule] [data-cf-value="1"]').value).toBe('=СУММ($A$1;1)')
    await w.find('[data-cf-rule] [data-cf-value="0"]').setValue('2,5')
    await w.find('[data-cf-rule]').trigger('submit')
    await until(() => !w.find('[data-cf-rule]').exists())
    await w.find('[data-cf-ok]').trigger('click')
    await until(() => s.history?.undo === 1)
    expect((await rulesOf())[0]).toMatchObject({ formulas: ['2,5', 'СУММ($A$1;1)'] })
  })
})
