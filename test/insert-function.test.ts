// Excel's buttons left of the formula bar's field: Cancel and Enter while an edit is open, and Insert Function with its dialog over the source's names.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { FunctionCategory } from 'xlcalc/contract'
import { memorySource, ruMessages, type MemoryBook } from '../src'
import { active, barField, cellText, editorKey, editorOpen, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 1, A2: 2, C1: 'text' } }], functions: ['БДСУММ', 'СУММ', 'СРЗНАЧ', 'СУММЕСЛИ', 'МАКС'] })
const about: Record<string, { category: FunctionCategory; categoryName: string; description: string }> = {
  БДСУММ: { category: 'database', categoryName: 'Работа с базой данных', description: 'Сумма по базе.' },
  СУММ: { category: 'math', categoryName: 'Математические', description: 'Сумма аргументов.' },
  СУММЕСЛИ: { category: 'math', categoryName: 'Математические', description: 'Сумма по условию.' },
  СРЗНАЧ: { category: 'statistical', categoryName: 'Статистические', description: 'Среднее аргументов.' },
  МАКС: { category: 'statistical', categoryName: 'Статистические', description: 'Наибольшее значение.' },
}
/** A source with the book's functions and their signatures. */
function signed() {
  const source = memorySource(book())
  const asked: string[] = []
  source.functionSignature = async (name: string) => {
    asked.push(name)
    return { name: name.toUpperCase(), args: [{ name: 'число1', optional: false }], repeat: 1, text: `${name.toUpperCase()}(число1;[число2];...)`, ...about[name.toUpperCase()]! }
  }
  return { source, asked }
}
const button = (w: VueWrapper, what: 'cancel' | 'enter' | 'fx') => w.find(`[data-xs-bar-${what}]`)
async function press(w: VueWrapper, what: 'cancel' | 'enter' | 'fx'): Promise<void> {
  await button(w, what).trigger('mousedown')
  await button(w, what).trigger('click')
  await settle()
}
const dialog = (w: VueWrapper) => w.find('[data-xs-insert-function]')
const items = (w: VueWrapper) => w.findAll('[data-xs-if-item]').map((i) => i.text())
const editingBar = (w: VueWrapper) => w.find('.xs-bar').classes().includes('xs-bar-editing')
async function editingKeyEscape(w: VueWrapper): Promise<void> {
  await w.find('.xs-fx-input').trigger('keydown', { key: 'Escape' })
  await settle()
}

describe('Cancel and Enter by the formula bar', () => {
  it('are off outside an edit and on while one is open, labelled from the messages', async () => {
    const w = await mountSheet(memorySource(book()), { messages: ruMessages })
    expect(button(w, 'cancel').attributes('disabled')).toBeDefined()
    expect(button(w, 'enter').attributes('disabled')).toBeDefined()
    expect(button(w, 'cancel').attributes('aria-label')).toBe('Отмена')
    expect(button(w, 'enter').attributes('aria-label')).toBe('Ввод')
    expect(button(w, 'fx').attributes('aria-label')).toBe('Вставить функцию')
    await goTo(w, 'C1')
    await key(w, 'F2')
    await settle()
    expect(button(w, 'cancel').attributes('disabled')).toBeUndefined()
    expect(button(w, 'enter').attributes('disabled')).toBeUndefined()
  })

  it('Cancel drops the edit, the cell keeps its text', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'C1')
    await key(w, 'F2')
    await settle()
    await type(w, 'other')
    await press(w, 'cancel')
    expect(editorOpen(w)).toBe(false)
    expect(cellText(w, 1, 3)).toBe('text')
  })

  it('Enter writes the cell and the active cell stays', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'C1')
    await key(w, 'x')
    await type(w, 'new')
    await press(w, 'enter')
    expect(editorOpen(w)).toBe(false)
    expect(cellText(w, 1, 3)).toBe('new')
    expect(active(w)).toBe('1:3')
  })
})

describe('Insert Function', () => {
  it('is off without functionNames and on a read-only sheet', async () => {
    const plain = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    expect(button(plain, 'fx').attributes('disabled')).toBeDefined()
    const read = await mountSheet(memorySource(book()), { readOnly: true })
    expect(button(read, 'fx').attributes('disabled')).toBeDefined()
  })

  it('opens a dialog of the functions; the search narrows it by any part of the name, in any case; the syntax of the one selected shows', async () => {
    const { source } = signed()
    const w = await mountSheet(source, { messages: ruMessages })
    await goTo(w, 'B2')
    await press(w, 'fx')
    expect(dialog(w).exists()).toBe(true)
    expect(dialog(w).text()).toContain('Вставка функции')
    expect(w.find('[data-xs-if-search]').attributes('placeholder')).toBe('Поиск функции')
    expect(dialog(w).text()).toContain('Выберите функцию')
    expect(document.activeElement).toBe(w.find('[data-xs-if-search]').element)
    expect(items(w)).toEqual(['БДСУММ', 'СУММ', 'СРЗНАЧ', 'СУММЕСЛИ', 'МАКС'])
    // The name itself first, then those it begins, then the rest that hold it.
    await w.find('[data-xs-if-search]').setValue('сумм')
    await settle()
    expect(items(w)).toEqual(['СУММ', 'СУММЕСЛИ', 'БДСУММ'])
    expect(w.find('[data-xs-if-syntax]').text()).toBe('СУММ(число1;[число2];...)')
    await w.find('[data-xs-if-search]').trigger('keydown', { key: 'ArrowDown' })
    await settle()
    expect(w.find('[data-xs-if-syntax]').text()).toBe('СУММЕСЛИ(число1;[число2];...)')
  })

  it('OK opens the edit of the active cell in the formula bar with =NAME( and the caret after the bracket', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await press(w, 'fx')
    await w.find('[data-xs-if-search]').setValue('сумм')
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(editingBar(w)).toBe(true)
    expect(barField(w).value).toBe('=СУММ(')
    expect(document.activeElement).toBe(barField(w))
    expect([barField(w).selectionStart, barField(w).selectionEnd]).toEqual([6, 6])
  })

  it('with an edit open puts NAME( at the draft’s caret, the edit kept open meanwhile; an empty draft gets = first', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await key(w, '=')
    await type(w, '=1+2')
    const field = w.find('.xs-editor').element as HTMLTextAreaElement
    field.setSelectionRange(3, 3)
    await press(w, 'fx')
    expect(editorOpen(w)).toBe(true)
    await w.find('[data-xs-if-search]').setValue('макс')
    await w.find('[data-xs-if-search]').trigger('keydown', { key: 'Enter' })
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(field.value).toBe('=1+МАКС(2')
    expect([field.selectionStart, field.selectionEnd]).toEqual([8, 8])
    await editorKey(w, 'Escape')
    await goTo(w, 'B3')
    await key(w, 'F2')
    await settle()
    await press(w, 'fx')
    await w.find('[data-xs-if-category]').setValue('all')
    await w.findAll('[data-xs-if-item]')[2]!.trigger('dblclick')
    await settle()
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('=СРЗНАЧ(')
  })

  it('Esc and Cancel close it and leave the cell and the edit as they were', async () => {
    const { source } = signed()
    const w = await mountSheet(source)
    await goTo(w, 'C1')
    await press(w, 'fx')
    await w.find('[data-xs-if-search]').trigger('keydown', { key: 'Escape' })
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(editingBar(w)).toBe(false)
    expect(editorOpen(w)).toBe(false)
    await key(w, 'F2')
    await settle()
    await press(w, 'fx')
    await w.find('[data-xs-dialog-cancel]').trigger('click')
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(editorOpen(w)).toBe(true)
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('text')
  })

  it("a category narrows the list, in Excel's order by the source's names; the description of the one selected shows; OK puts it first among the recently used", async () => {
    const { source } = signed()
    const w = await mountSheet(source, { messages: ruMessages })
    await goTo(w, 'B2')
    await press(w, 'fx')
    const category = w.find('[data-xs-if-category]')
    expect(category.findAll('option').map((o) => o.text())).toEqual(['10 недавно использовавшихся', 'Полный алфавитный перечень', 'Математические', 'Статистические', 'Работа с базой данных'])
    // Nothing used yet: the whole list first.
    expect((category.element as HTMLSelectElement).value).toBe('all')
    await category.setValue('statistical')
    await settle()
    expect(items(w)).toEqual(['СРЗНАЧ', 'МАКС'])
    expect(w.find('[data-xs-if-about]').text()).toBe('Среднее аргументов.')
    await w.findAll('[data-xs-if-item]')[1]!.trigger('click')
    expect(w.find('[data-xs-if-about]').text()).toBe('Наибольшее значение.')
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await settle()
    expect(barField(w).value).toBe('=МАКС(')
    await editingKeyEscape(w)
    await press(w, 'fx')
    expect((w.find('[data-xs-if-category]').element as HTMLSelectElement).value).toBe('recent')
    expect(items(w)).toEqual(['МАКС'])
  })
})
