// Insert Function, Formula AutoComplete and the argument tooltip over localSource, in en-US and ru-RU: the categories xlcalc
// 0.68.0 added (User Defined, Cube, Web) in the order of Excel's dialog, the functions of 0.68.0 to 0.70.0 with their syntax and
// descriptions from `xlcalc/function-text`, and ru-RU's forecast functions under Excel for Windows' names.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { ruMessages, type SheetViewApi } from '../src'
import { editorKey, goTo, key, mountSheet, settle, type, until } from './helpers'
import { closeAll, source } from './local-fixture'

afterEach(async () => {
  document.body.innerHTML = ''
  await closeAll()
})

const sheet = (locale: string) => source({ locale, sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] })
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const options = (w: VueWrapper) => w.find('[data-xs-if-category]').findAll('option').map((o) => o.text())
const items = (w: VueWrapper) => w.findAll('[data-xs-if-item]').map((i) => i.text())
const syntax = (w: VueWrapper) => w.find('[data-xs-if-syntax]').text()
const about = (w: VueWrapper) => w.find('[data-xs-if-about]').text()

/** Opens Insert Function over a source in `locale` and waits for every function's signature, which name the categories. */
async function insertFunction(locale: string, messages?: typeof ruMessages): Promise<VueWrapper> {
  const s = sheet(locale)
  const w = await mountSheet(s, { locale, ...(messages ? { messages } : {}) })
  await s.ready
  await goTo(w, 'B2')
  await until(() => api(w).insertFunction())
  await settle()
  await until(() => w.find('[data-xs-if-category]').exists() && options(w).length > 2, 20_000)
  return w
}
/** Shows a category's functions and selects one by name. */
async function pick(w: VueWrapper, category: string, name: string): Promise<void> {
  await w.find('[data-xs-if-category]').setValue(category)
  await settle()
  const i = items(w).indexOf(name)
  expect(i).toBeGreaterThanOrEqual(0)
  await w.findAll('[data-xs-if-item]')[i]!.trigger('click')
}

describe('Insert Function over localSource', () => {
  it("in en-US lists every category in the order of Excel's dialog, User Defined, Cube and Web among them", async () => {
    const w = await insertFunction('en-US')
    expect(options(w)).toEqual(['Most Recently Used', 'All', 'Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text', 'Logical', 'Information', 'User Defined', 'Engineering', 'Cube', 'Compatibility', 'Web'])
    await pick(w, 'web', 'FILTERXML')
    expect(items(w)).toEqual(['ENCODEURL', 'FILTERXML', 'WEBSERVICE'])
    expect(syntax(w)).toBe('FILTERXML(xml, xpath)')
    expect(about(w)).not.toBe('')
    await pick(w, 'user', 'EUROCONVERT')
    expect(items(w)).toEqual(['CALL', 'EUROCONVERT', 'REGISTER.ID'])
    await pick(w, 'cube', 'CUBEVALUE')
    expect(syntax(w)).toBe('CUBEVALUE(connection, [member_expression1], [member_expression2], ...)')
    await pick(w, 'logical', 'LAMBDA')
    expect(items(w)).toEqual(expect.arrayContaining(['BYCOL', 'BYROW', 'LAMBDA', 'LET', 'MAKEARRAY', 'MAP', 'REDUCE', 'SCAN']))
    await pick(w, 'lookup', 'GROUPBY')
    expect(items(w)).toEqual(expect.arrayContaining(['FORMULATEXT', 'GROUPBY', 'INDIRECT', 'OFFSET', 'PIVOTBY', 'TRIMRANGE']))
    expect(about(w)).not.toBe('')
    await pick(w, 'information', 'SHEETS')
    expect(items(w)).toEqual(expect.arrayContaining(['CELL', 'INFO', 'ISFORMULA', 'ISOMITTED', 'SHEET', 'SHEETS']))
    await pick(w, 'text', 'REGEXEXTRACT')
    expect(items(w)).toEqual(expect.arrayContaining(['REGEXEXTRACT', 'REGEXREPLACE', 'REGEXTEST']))
    expect(syntax(w)).toBe('REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])')
    await pick(w, 'statistical', 'FORECAST.ETS')
    expect(items(w)).toEqual(expect.arrayContaining(['FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT', 'FORECAST.LINEAR']))
  }, 30_000)

  it('in ru-RU names the categories in Russian, and the functions with their Russian names, syntax and descriptions', async () => {
    const w = await insertFunction('ru-RU', ruMessages)
    expect(options(w)).toEqual(['10 недавно использовавшихся', 'Полный алфавитный перечень', 'Финансовые', 'Дата и время', 'Математические', 'Статистические', 'Ссылки и массивы', 'Работа с базой данных', 'Текстовые', 'Логические', 'Проверка свойств и значений', 'Определенные пользователем', 'Инженерные', 'Аналитические', 'Совместимость', 'Интернет'])
    await pick(w, 'web', 'ФИЛЬТР.XML')
    expect(items(w)).toEqual(['ВЕБСЛУЖБА', 'КОДИР.URL', 'ФИЛЬТР.XML'])
    expect(syntax(w)).toBe('ФИЛЬТР.XML(xml;xpath)')
    expect(about(w)).not.toBe('')
    await pick(w, 'lookup', 'ДВССЫЛ')
    expect(syntax(w)).toBe('ДВССЫЛ(ссылка_на_ячейку;[a1])')
    await pick(w, 'statistical', 'ПРЕДСКАЗ.ЕTS.ДОВИНТЕРВАЛ')
    expect(items(w)).toEqual(expect.arrayContaining(['ПРЕДСКАЗ.ETS', 'ПРЕДСКАЗ.ЕTS.ДОВИНТЕРВАЛ', 'ПРЕДСКАЗ.ETS.СЕЗОННОСТЬ', 'ПРЕДСКАЗ.ETS.СТАТ', 'ПРЕДСКАЗ.ЛИНЕЙН']))
    expect(items(w)).not.toContain('FORECAST.ETS')
    expect(syntax(w)).toMatch(/^ПРЕДСКАЗ\.ЕTS\.ДОВИНТЕРВАЛ\(целевая_дата;значения;временная_шкала;/)
    expect(about(w)).not.toBe('')
  }, 30_000)
})

describe('Formula AutoComplete and the argument tooltip over localSource in ru-RU', () => {
  const shown = () => [...document.querySelectorAll('.xs-fl-item')].map((li) => li.textContent!.trim())
  const tip = () => document.querySelector('[data-xs-tip]')?.textContent?.trim() ?? null

  it('the forecast functions by their characters: the Latin E of ПРЕДСКАЗ.ETS finds three, the Cyrillic Е of ПРЕДСКАЗ.ЕTS.ДОВИНТЕРВАЛ the fourth', async () => {
    const s = sheet('ru-RU')
    const w = await mountSheet(s, { locale: 'ru-RU' })
    await key(w, '=')
    await type(w, '=ПРЕДСКАЗ.E')
    await until(() => shown().length > 0)
    expect(shown()).toEqual(['ПРЕДСКАЗ.ETS', 'ПРЕДСКАЗ.ETS.СЕЗОННОСТЬ', 'ПРЕДСКАЗ.ETS.СТАТ'])
    await type(w, '=ПРЕДСКАЗ.Е')
    await settle()
    expect(shown()).toEqual(['ПРЕДСКАЗ.ЕTS.ДОВИНТЕРВАЛ'])
    await editorKey(w, 'Tab')
    await until(() => tip() !== null)
    expect(tip()).toMatch(/^ПРЕДСКАЗ\.ЕTS\.ДОВИНТЕРВАЛ\(целевая_дата;/)
  })

  it("an English name typed in a ru-RU workbook gets the Russian tooltip, as for LAMBDA's family", async () => {
    const s = sheet('ru-RU')
    const w = await mountSheet(s, { locale: 'ru-RU' })
    await key(w, '=')
    await type(w, '=FORECAST.ETS.CONFINT(')
    await until(() => tip() !== null)
    expect(tip()).toMatch(/^ПРЕДСКАЗ\.ЕTS\.ДОВИНТЕРВАЛ\(/)
    await type(w, '=GROUPBY(')
    await until(() => tip()?.startsWith('ГРУПППО(') ?? false)
    await type(w, '=REDUCE(0;A1:A3;')
    await until(() => tip()?.startsWith('REDUCE(') ?? false)
    expect(document.querySelector('[data-xs-tip] .xs-ft-on')?.textContent?.trim()).toBe('лямбда')
  })
})
