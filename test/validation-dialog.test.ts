// Excel's Data Validation dialog over the local source's `rules` and `setRules`: opened through `SheetViewApi.dataValidation`
// for the selection, with the active cell's rule (formulas with =, a literal list with the locale's separator), Excel's
// questions for several kinds of rules or cells without the rule, its checks before OK, and OK writing the sheet's rules.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput, DataValidation } from 'xlcalc'
import { enMessages, memorySource, ruMessages, type SheetViewApi } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, mountSheet, settle, until } from './helpers'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

const RULES: DataValidation[] = [
  { range: 'A2:A9', type: 'list', formula1: '"Open,On hold,Closed"', promptTitle: 'Status', prompt: 'Pick one' },
  { range: 'B2:B9', type: 'whole', formula1: '0', formula2: '100', errorStyle: 'warning', error: 'From 0 to 100' },
  { range: 'C2:C9', type: 'list', formula1: '$E$2:$E$4' },
]
const book = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'S', cells: { E2: { value: 'x' }, E3: { value: 'y' }, E4: { value: 'z' } }, validations: RULES }],
})

async function mountLocal(locale?: string, props: Record<string, unknown> = {}): Promise<{ w: VueWrapper; source: LocalSource; api: SheetViewApi }> {
  const source = localSource(book(locale), { calc: 'inline' })
  sources.push(source)
  await source.ready
  const w = await mountSheet(source, { locale: source.locale, ...(locale === 'ru-RU' ? { messages: ruMessages } : {}), ...props })
  await settle(12)
  return { w, source, api: w.vm as unknown as SheetViewApi }
}
const dialog = () => document.querySelector<HTMLElement>('[data-dv]')
const q = <T extends HTMLElement>(sel: string) => document.querySelector<T>(sel)!
const opened = () => until(() => !!dialog() || !!document.querySelector('[data-xs-alert]'))
async function open(w: VueWrapper, api: SheetViewApi, place: string): Promise<void> {
  await goTo(w, place)
  expect(api.dataValidation()).toBe(true)
  await opened()
  await settle()
}
async function set(sel: string, value: string | boolean): Promise<void> {
  const el = q<HTMLInputElement>(sel)
  if (typeof value === 'boolean') el.checked = value
  else el.value = value
  el.dispatchEvent(new Event(el.tagName === 'SELECT' || typeof value === 'boolean' ? 'change' : 'input'))
  await settle(1)
}
async function ok(): Promise<void> {
  q('[data-dv-ok]').click()
  await settle(12)
  await until(() => !dialog() || !!document.querySelector('[data-dv-message]'))
}
const rules = async (source: LocalSource) => (await source.rules!('S')).validations

describe('opening Data Validation', () => {
  it('is not offered by a source without rules, nor on a read-only sheet', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: {} }] }))
    expect((w.vm as unknown as SheetViewApi).dataValidation()).toBe(false)
    const { api } = await mountLocal(undefined, { readOnly: true })
    expect(api.dataValidation()).toBe(false)
  })

  it("shows the active cell's rule: a literal list with commas in en-US, its messages and alert", async () => {
    const { w, api } = await mountLocal()
    await open(w, api, 'A3')
    expect(dialog()?.getAttribute('aria-label')).toBe('Data Validation')
    expect(q<HTMLSelectElement>('[data-dv-type]').value).toBe('list')
    expect(q<HTMLInputElement>('[data-dv-box="source"]').value).toBe('Open,On hold,Closed')
    expect(q<HTMLInputElement>('[data-dv-dropdown]').checked).toBe(true)
    expect(q<HTMLInputElement>('[data-dv-prompt-title]').value).toBe('Status')
    expect(q<HTMLTextAreaElement>('[data-dv-prompt]').value).toBe('Pick one')
    expect([...document.querySelectorAll('[data-dv-tab]')].map((t) => t.textContent?.trim())).toEqual(['Settings', 'Input Message', 'Error Alert'])
  })

  it("in ru-RU shows the list with semicolons, Excel's Russian names, and a list of cells with =", async () => {
    const { w, api } = await mountLocal('ru-RU')
    await open(w, api, 'A3')
    expect(dialog()?.getAttribute('aria-label')).toBe('Проверка вводимых значений')
    expect(q<HTMLInputElement>('[data-dv-box="source"]').value).toBe('Open;On hold;Closed')
    expect(q<HTMLSelectElement>('[data-dv-type]').selectedOptions[0]?.textContent?.trim()).toBe('Список')
    q('[data-dv-cancel]').click()
    await settle()
    await open(w, api, 'C5')
    expect(q<HTMLInputElement>('[data-dv-box="source"]').value).toBe('=$E$2:$E$4')
  })

  it("asks Excel's question for rules of more than one kind: Cancel opens nothing, OK opens Excel's defaults", async () => {
    const { w, api } = await mountLocal()
    await open(w, api, 'A2:B2')
    expect(document.querySelector('[data-xs-alert-message]')?.textContent).toBe(enMessages.dvMany)
    q('[data-xs-alert-button="cancel"]').click()
    await settle(12)
    expect(dialog()).toBeNull()
    expect(api.dataValidation()).toBe(true)
    await opened()
    q('[data-xs-alert-button="ok"]').click()
    await until(() => !!dialog())
    expect(q<HTMLSelectElement>('[data-dv-type]').value).toBe('any')
  })
})

describe('OK', () => {
  it('writes a new rule over the selection, splitting the rule the selection covers part of', async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'B4:B5')
    expect(q<HTMLInputElement>('[data-dv-box="min"]').value).toBe('0')
    await set('[data-dv-box="max"]', '10')
    await ok()
    expect(dialog()).toBeNull()
    expect(await rules(source)).toEqual([
      RULES[0],
      { range: 'B2:B3 B6:B9', type: 'whole', formula1: '0', formula2: '100', errorStyle: 'warning', error: 'From 0 to 100' },
      RULES[2],
      { range: 'B4:B5', type: 'whole', formula1: '0', formula2: '10', errorStyle: 'warning', error: 'From 0 to 100' },
    ])
  })

  it('in ru-RU stores a literal list typed with semicolons as a file stores it, and the sheet shows the new list', async () => {
    const { w, source, api } = await mountLocal('ru-RU')
    await open(w, api, 'D2')
    await set('[data-dv-type]', 'list')
    await set('[data-dv-box="source"]', 'да;нет')
    await ok()
    expect((await rules(source)).at(-1)).toEqual({ range: 'D2', type: 'list', formula1: '"да,нет"' })
    expect(await source.validation!('S', 'D2')).toEqual({ list: ['да', 'нет'] })
    // the active cell's rule is requested again: its dropdown arrow appears
    await until(() => !!document.querySelector('[data-xs-dropdown-arrow]'))
  })

  it("shows Excel's message and writes nothing for a maximum below the minimum or an empty source", async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'B3')
    await set('[data-dv-box="min"]', '50')
    await set('[data-dv-box="max"]', '5')
    await ok()
    expect(q('[data-dv-message]').textContent).toBe('The Maximum must be greater than or equal to the Minimum.')
    await set('[data-dv-type]', 'list')
    await set('[data-dv-box="source"]', '')
    await ok()
    expect(q('[data-dv-message]').textContent).toBe('You must enter a value in the Source box.')
    expect(await rules(source)).toEqual(RULES)
  })

  it("shows the source's refusal of a formula that does not parse, and keeps the dialog open", async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'D2')
    await set('[data-dv-type]', 'custom')
    await set('[data-dv-box="formula"]', '=SUM(')
    await ok()
    expect(dialog()).not.toBeNull()
    expect(q('[data-dv-message]').textContent).not.toBe('')
    expect(await rules(source)).toEqual(RULES)
  })

  it('Clear All then OK removes the rule from the selection only', async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'A2:A3')
    q('[data-dv-clear]').click()
    await settle()
    expect(q<HTMLSelectElement>('[data-dv-type]').value).toBe('any')
    expect(q<HTMLInputElement>('[data-dv-prompt-title]').value).toBe('')
    await ok()
    expect((await rules(source))[0]).toEqual({ ...RULES[0], range: 'A4:A9' })
    expect(await source.validation!('S', 'A2')).toBeNull()
  })

  it('writes the input message and the error alert, and the active cell shows the new message', async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'D2')
    await set('[data-dv-type]', 'textLength')
    await set('[data-dv-operator]', 'lessThanOrEqual')
    await set('[data-dv-box="max"]', '5')
    await set('[data-dv-prompt-title]', 'Code')
    await set('[data-dv-prompt]', 'Five letters at most')
    await set('[data-dv-style]', 'information')
    await set('[data-dv-error]', 'Too long')
    await ok()
    expect((await rules(source)).at(-1)).toEqual({ range: 'D2', type: 'textLength', operator: 'lessThanOrEqual', formula1: '5', promptTitle: 'Code', prompt: 'Five letters at most', errorStyle: 'information', error: 'Too long' })
    await until(() => document.querySelector('[data-xs-prompt]')?.textContent?.includes('Five letters at most') ?? false)
  })
})

describe('a selection with cells without the rule', () => {
  it("asks whether to extend it: Yes opens with the rule's settings, and OK gives them to the whole selection", async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'C8:C11')
    expect(document.querySelector('[data-xs-alert-message]')?.textContent).toBe(enMessages.dvExtend)
    expect([...document.querySelectorAll('[data-xs-alert-button]')].map((b) => b.textContent?.trim())).toEqual(['Yes', 'No', 'Cancel'])
    q('[data-xs-alert-button="yes"]').click()
    await until(() => !!dialog())
    expect(q<HTMLInputElement>('[data-dv-box="source"]').value).toBe('=$E$2:$E$4')
    await ok()
    expect(await rules(source)).toEqual([RULES[0], RULES[1], { ...RULES[2], range: 'C2:C7' }, { range: 'C8:C11', type: 'list', formula1: '$E$2:$E$4' }])
  })

  it("No opens with Excel's defaults", async () => {
    const { w, api } = await mountLocal()
    await open(w, api, 'C8:C11')
    q('[data-xs-alert-button="no"]').click()
    await until(() => !!dialog())
    expect(q<HTMLSelectElement>('[data-dv-type]').value).toBe('any')
    expect(q<HTMLInputElement>('[data-dv-all]').disabled).toBe(true)
  })

  it('Apply these changes to all other cells with the same settings changes every cell of the rule', async () => {
    const { w, source, api } = await mountLocal()
    await open(w, api, 'B3')
    await set('[data-dv-all]', true)
    await set('[data-dv-box="max"]', '50')
    await ok()
    expect((await rules(source)).filter((r) => r.type === 'whole')).toEqual([{ range: 'B3 B2 B4:B9', type: 'whole', formula1: '0', formula2: '50', errorStyle: 'warning', error: 'From 0 to 100' }])
  })
})
