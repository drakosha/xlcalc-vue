// Excel's Symbol dialog: the grid by Unicode subset, the Character code box, the recently used symbols, the Special Characters tab, and Insert into the edit or a new one.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, ruMessages, type MemoryBook, type SheetViewApi } from '../src'
import { codeHex, DEFAULT_RECENT_SYMBOLS, parseCode, subsetOf, symbolCodes, SYMBOL_SPECIALS } from '../src/symbols'
import { cellText, editorKey, editorOpen, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'abc', B1: 'x' } }] })
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const dialog = (w: VueWrapper) => w.find('[data-xs-symbol]')
const draft = (w: VueWrapper) => (w.find('.xs-editor').element as HTMLTextAreaElement).value
const codeBox = (w: VueWrapper) => w.find('[data-xs-sym-code-box]').element as HTMLInputElement
const selected = (w: VueWrapper) => w.find('.xs-sym-grid .xs-sym-on').attributes('data-xs-sym-code')
async function open(w: VueWrapper): Promise<void> {
  expect(api(w).insertSymbol()).toBe(true)
  await settle()
}
async function insert(w: VueWrapper): Promise<void> {
  await w.find('[data-xs-sym-insert]').trigger('click')
  await settle()
}
async function typeCode(w: VueWrapper, text: string): Promise<void> {
  await w.find('[data-xs-sym-code-box]').setValue(text)
  await settle()
}
async function close(w: VueWrapper): Promise<void> {
  await w.find('[data-xs-dialog-cancel]').trigger('click')
  await settle()
}

describe('the Symbol data', () => {
  it('lists assigned characters only, in code order, by subset', () => {
    const codes = symbolCodes()
    expect(codes[0]).toBe(0x20)
    expect(codes).not.toContain(0x7f)
    expect(codes).not.toContain(0x9f)
    expect(codes).toContain(0x2014)
    expect([...codes].sort((a, b) => a - b)).toEqual(codes)
    expect(subsetOf(0x416)).toBe('cyrillic')
    expect(subsetOf(0x20ac)).toBe('currency')
    expect(subsetOf(0x1f600)).toBeNull()
  })
  it('reads the Character code box as hex Unicode, with or without U+', () => {
    expect(parseCode('00a9')).toBe(0xa9)
    expect(parseCode('U+2014')).toBe(0x2014)
    expect(parseCode('1F600')).toBe(0x1f600)
    expect(parseCode('0007')).toBeNull()
    expect(parseCode('zz')).toBeNull()
    expect(parseCode('110000')).toBeNull()
    expect(codeHex(0xa9)).toBe('00A9')
  })
})

describe('Symbol', () => {
  it('opens on an editable sheet only, in the texts of the messages', async () => {
    const read = await mountSheet(memorySource(book()), { readOnly: true })
    expect(api(read).insertSymbol()).toBe(false)
    const w = await mountSheet(memorySource(book()), { messages: ruMessages })
    await open(w)
    expect(dialog(w).find('.xs-modal-title').text()).toBe('Символ')
    expect(w.findAll('[data-xs-sym-tab]').map((t) => t.text())).toEqual(['Символы', 'Специальные знаки'])
    expect(document.activeElement).toBe(w.find('[data-xs-sym-grid]').element)
    expect(selected(w)).toBe('0020')
    expect(w.find('[data-xs-dialog-cancel]').text()).toBe('Отмена')
  })

  it('outside editing, Insert starts editing the active cell with the character as typing does; the dialog stays and the next one follows at the caret', async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    await typeCode(w, '00A9')
    expect(selected(w)).toBe('00A9')
    expect((w.find('[data-xs-sym-subset]').element as HTMLSelectElement).value).toBe('latin1')
    await insert(w)
    expect(editorOpen(w)).toBe(true)
    expect(draft(w)).toBe('©')
    expect(dialog(w).exists()).toBe(true)
    expect(w.find('[data-xs-dialog-cancel]').text()).toBe('Close')
    await typeCode(w, '2122')
    await insert(w)
    expect(draft(w)).toBe('©™')
    await close(w)
    expect(dialog(w).exists()).toBe(false)
    expect(editorOpen(w)).toBe(true)
    expect(document.activeElement).toBe(w.find('.xs-editor').element)
    await editorKey(w, 'Enter')
    await settle()
    expect(cellText(w, 1, 1)).toBe('©™')
  })

  it('in an open edit, Insert puts the character at the caret in place of the selection, and the edit stays open', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B1')
    await key(w, 'F2')
    await settle()
    await type(w, 'a-b')
    const el = w.find('.xs-editor').element as HTMLTextAreaElement
    el.setSelectionRange(1, 2)
    await open(w)
    expect(editorOpen(w)).toBe(true)
    await w.find('[data-xs-sym-tab="special"]').trigger('click')
    await w.find('[data-xs-sym-special="enDash"]').trigger('click')
    await insert(w)
    await w.find('[data-xs-sym-special="emDash"]').trigger('dblclick')
    await settle()
    expect(draft(w)).toBe('a–—b')
    await close(w)
    expect(el.selectionStart).toBe(3)
    await editorKey(w, 'Enter')
    await settle()
    expect(cellText(w, 1, 2)).toBe('a–—b')
  })

  it('the arrows move through the grid, a row is 16 characters, and Enter inserts', async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    const grid = w.find('[data-xs-sym-grid]')
    await grid.trigger('keydown', { key: 'ArrowRight' })
    expect(selected(w)).toBe('0021')
    await grid.trigger('keydown', { key: 'ArrowDown' })
    expect(selected(w)).toBe('0031')
    expect(codeBox(w).value).toBe('0031')
    await grid.trigger('keydown', { key: 'End' })
    expect(selected(w)).toBe('003F')
    await grid.trigger('keydown', { key: 'Home', ctrlKey: true })
    expect(selected(w)).toBe('0020')
    await grid.trigger('keydown', { key: 'ArrowLeft' })
    expect(selected(w)).toBe('0020')
    await grid.trigger('keydown', { key: 'ArrowRight' })
    await grid.trigger('keydown', { key: 'Enter' })
    await settle()
    expect(draft(w)).toBe('!')
  })

  it('the Subset box goes to the subset’s first character', async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    await w.find('[data-xs-sym-subset]').setValue('cyrillic')
    await settle()
    expect(selected(w)).toBe('0400')
    expect(codeBox(w).value).toBe('0400')
  })

  it("lists Office's recently used symbols at first, then those inserted, newest first, kept under storageKey", async () => {
    const w = await mountSheet(memorySource(book()), { storageKey: 'sym' })
    await open(w)
    const recent = () => w.findAll('[data-xs-sym-recent]').map((b) => b.attributes('data-xs-sym-recent'))
    expect(recent()).toEqual(DEFAULT_RECENT_SYMBOLS.map(codeHex))
    await typeCode(w, '03A9')
    await insert(w)
    await w.find('[data-xs-sym-recent="00A9"]').trigger('click')
    expect(selected(w)).toBe('00A9')
    await insert(w)
    expect(recent().slice(0, 3)).toEqual(['00A9', '03A9', '20AC'])
    expect(recent()).toHaveLength(16)
    await close(w)
    w.unmount()
    const again = await mountSheet(memorySource(book()), { storageKey: 'sym' })
    await open(again)
    expect(again.findAll('[data-xs-sym-recent]')[0]!.attributes('data-xs-sym-recent')).toBe('00A9')
    expect(again.find('.xs-sym-grid .xs-sym-on').attributes('data-xs-sym-code')).toBe('00A9')
  })

  it("the Special Characters tab lists Excel's characters by name; Esc closes the dialog back to the grid", async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    await w.find('[data-xs-sym-tab="special"]').trigger('click')
    expect(w.findAll('[data-xs-sym-special]')).toHaveLength(SYMBOL_SPECIALS.length)
    expect(w.find('[data-xs-sym-special="copyright"]').text()).toBe('©Copyright')
    const list = w.find('[data-xs-sym-specials]')
    await list.trigger('keydown', { key: 'ArrowDown' })
    expect(w.find('[data-xs-sym-special="enDash"]').attributes('aria-selected')).toBe('true')
    await dialog(w).trigger('keydown', { key: 'Escape' })
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(editorOpen(w)).toBe(false)
    expect(document.activeElement).toBe(w.find('.xs-scroller').element)
  })
})
