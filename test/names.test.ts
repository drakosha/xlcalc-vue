// Excel's defined names in the sheet: the Name Box goes to a name, shows it for its range, lists the names that refer to
// ranges and defines a new one; Define Name and the Name Manager replace the names through `setNames`; Formula
// AutoComplete offers names among the functions; the names are read once per source and again after a change.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, ruMessages, type MemoryBook, type SheetViewApi } from '../src'
import type { EditOutcome, SheetSource } from '../src/contract/source'
import type { SheetName } from '../src/contract/workbook'
import { absoluteRange, rangeRef } from '../src/core/place'
import { address, cellText, editorKey, goTo, key, mountSheet, settle, type, until } from './helpers'
import { book as localBook, closeAll, source as localSource } from './local-fixture'

afterEach(async () => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
  await closeAll()
})

const NAMES: SheetName[] = [
  { name: 'Sales', ref: 'S!$B$2:$C$3' },
  { name: 'Rate', ref: '0.2' },
  { name: 'Far', ref: "'Other sheet'!$D$4" },
  { name: 'Own', ref: "'Other sheet'!$A$1:$A$2", scope: 'Other sheet' },
  { name: 'cols', ref: 'S!$B:$C' },
]
const book = (names: SheetName[] = NAMES, functions?: string[]): MemoryBook => ({
  sheets: [
    { name: 'S', cells: { A1: 1, B2: 2 } },
    { name: 'Other sheet', cells: { A1: 3 } },
  ],
  names,
  ...(functions ? { functions } : {}),
})
/** An in-memory workbook whose `names` and `setNames` calls are counted. */
function spied(names: SheetName[] = NAMES, functions?: string[]) {
  const source = memorySource(book(names, functions))
  const read = vi.spyOn(source, 'names')
  const save = vi.spyOn(source, 'setNames')
  return { source, read, save }
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const sheetName = () => document.querySelector('.xs-tab-on')?.textContent?.trim()
const message = (w: VueWrapper) => w.find('.xs-message').text()
const $ = <T extends Element = HTMLElement>(sel: string) => document.querySelector<T>(sel)
const input = (sel: string) => $<HTMLInputElement>(sel)!
async function click(sel: string): Promise<void> {
  $(sel)!.click()
  await settle()
}
async function setInput(sel: string, value: string): Promise<void> {
  const el = input(sel)
  el.value = value
  el.dispatchEvent(new Event('input'))
  el.dispatchEvent(new Event('change'))
  await settle()
}

describe('the selection as Define Name writes it', () => {
  it('is absolute, with its sheet quoted where it needs quotes', () => {
    expect(absoluteRange({ top: 2, left: 2, bottom: 9, right: 4 })).toBe('$B$2:$D$9')
    expect(absoluteRange({ top: 3, left: 1, bottom: 3, right: 1 })).toBe('$A$3')
    expect(rangeRef('Other sheet', { top: 1, left: 2, bottom: 1048576, right: 2 })).toBe("'Other sheet'!$B:$B")
  })
})

describe('the Name Box', () => {
  it('goes to a name’s range, on another sheet too, and to a sheet’s own name typed with its sheet', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'sales')
    expect(address(w)).toBe('Sales')
    expect(w.find('.xs-cell.xs-active').attributes('data-cell')).toBe('2:2')
    await goTo(w, 'Far')
    expect(sheetName()).toBe('Other sheet')
    expect(address(w)).toBe('Far')
    await goTo(w, 'Own')
    expect(address(w)).toBe('Own')
    await goTo(w, 'S!A1')
    await goTo(w, "'Other sheet'!Own")
    expect(sheetName()).toBe('Other sheet')
    expect(address(w)).toBe('Own')
  })

  it('shows the name whose range is exactly the selection, otherwise the address', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B2:C3')
    expect(address(w)).toBe('Sales')
    await goTo(w, 'B2:C4')
    expect(address(w)).toBe('B2:C4')
    await goTo(w, 'B:C')
    expect(address(w)).toBe('cols')
  })

  it('says a name that does not refer to a range is not a reference, and a sheet’s own name is not visible from another sheet', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'Rate')
    expect(message(w)).toBe('Reference isn’t valid.')
    expect(address(w)).toBe('A1')
  })

  it('its arrow lists the names that refer to ranges and are visible on the sheet, alphabetically; a click goes to one', async () => {
    const w = await mountSheet(memorySource(book()))
    await click('[data-xs-namebox-arrow]')
    const shown = () => [...document.querySelectorAll<HTMLElement>('[data-xs-namebox-name]')].map((li) => li.dataset.xsNameboxName)
    expect(shown()).toEqual(['cols', 'Far', 'Sales'])
    await click('[data-xs-namebox-name="Far"]')
    expect(sheetName()).toBe('Other sheet')
    expect($('[data-xs-namebox-list]')).toBeNull()
    await click('[data-xs-namebox-arrow]')
    expect(shown()).toEqual(['cols', 'Far', 'Own', 'Sales'])
    expect($('[data-xs-namebox-name="Far"]')!.getAttribute('aria-selected')).toBe('true')
  })

  it('Alt+Down opens the list, Down and Up move, Enter goes, Esc closes', async () => {
    const w = await mountSheet(memorySource(book()))
    const box = w.find('.xs-address')
    await box.trigger('focus')
    await box.trigger('keydown', { key: 'ArrowDown', altKey: true })
    await box.trigger('keydown', { key: 'ArrowDown' })
    await box.trigger('keydown', { key: 'ArrowDown' })
    expect($('.xs-namebox-list .xs-popup-on')!.textContent!.trim()).toBe('Far')
    await box.trigger('keydown', { key: 'Escape' })
    expect($('[data-xs-namebox-list]')).toBeNull()
    await box.trigger('keydown', { key: 'ArrowDown', altKey: true })
    await box.trigger('keydown', { key: 'End' })
    await box.trigger('keydown', { key: 'Enter' })
    await settle()
    expect(address(w)).toBe('Sales')
  })

  it('a new name typed with a range selected defines it over the selection, in the workbook’s scope, as a single change', async () => {
    const { source, save } = spied()
    const w = await mountSheet(source)
    await goTo(w, 'A1:A3')
    await goTo(w, 'Total')
    expect(save).toHaveBeenCalledTimes(1)
    expect(save.mock.calls[0]![0].at(-1)).toEqual({ name: 'Total', ref: 'S!$A$1:$A$3' })
    expect(address(w)).toBe('Total')
    expect((await source.names!()).map((n) => n.name)).toContain('Total')
  })

  it('a name the source refuses gets Excel’s message; a read-only sheet defines nothing', async () => {
    const { source, save } = spied()
    const w = await mountSheet(source)
    await goTo(w, 'a b')
    expect(save).toHaveBeenCalledTimes(1)
    expect(message(w)).toBe('You must enter a valid reference or type a valid name for the selection.')
    const ro = spied()
    const r = await mountSheet(ro.source, { readOnly: true })
    await goTo(r, 'Total')
    expect(ro.save).not.toHaveBeenCalled()
    expect(message(r)).toBe('Not an address: try B12, B2:D9 or Sheet!C7')
  })

  it('without the source’s names there is no arrow, and a name is not an address', async () => {
    const source: SheetSource = { ...memorySource(book()), names: undefined, setNames: undefined }
    const w = await mountSheet(source)
    expect($('[data-xs-namebox-arrow]')).toBeNull()
    await goTo(w, 'Sales')
    expect(message(w)).toBe('Not an address: try B12, B2:D9 or Sheet!C7')
    expect(api(w).nameManager()).toBe(false)
    expect(api(w).defineName()).toBe(false)
  })
})

describe('reading the names', () => {
  it('once per source, and again after setNames, undo, redo and a reload event', async () => {
    const { source, read } = spied()
    const undo = vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [] }))
    const src: SheetSource = Object.assign(source, { undo, redo: undo })
    const w = await mountSheet(src)
    expect(read).toHaveBeenCalledTimes(1)
    await goTo(w, 'B2')
    expect(read).toHaveBeenCalledTimes(1)
    await api(w).undo()
    await settle()
    expect(read).toHaveBeenCalledTimes(2)
    await api(w).redo()
    await settle()
    expect(read).toHaveBeenCalledTimes(3)
    source.emit({ kind: 'reload' })
    await settle()
    expect(read).toHaveBeenCalledTimes(4)
    await goTo(w, 'Newer')
    // the sheet is read again after the change, and the names with it
    expect(read).toHaveBeenCalledTimes(5)
    const other = spied()
    await w.setProps({ source: other.source })
    await settle()
    expect(other.read).toHaveBeenCalledTimes(1)
    expect(address(w)).toBe('A1')
  })
})

describe('Define Name', () => {
  it('opens New Name for the selection: Workbook scope, Refers to the selection with its sheet; OK adds the name', async () => {
    const { source, save } = spied()
    const w = await mountSheet(source)
    await goTo(w, 'A2:B3')
    expect(api(w).defineName()).toBe(true)
    await settle()
    expect($('[data-xs-name-dialog] .xs-modal-title')!.textContent).toBe('New Name')
    expect(input('[data-xs-name-refers]').value).toBe('=S!$A$2:$B$3')
    expect($<HTMLSelectElement>('[data-xs-name-scope]')!.value).toBe('')
    expect(document.activeElement).toBe(input('[data-xs-name-field]'))
    await setInput('[data-xs-name-field]', 'Block')
    const scope = $<HTMLSelectElement>('[data-xs-name-scope]')!
    scope.value = 'Other sheet'
    scope.dispatchEvent(new Event('change'))
    await settle()
    $<HTMLFormElement>('[data-xs-name-dialog]')!.requestSubmit()
    await settle()
    expect(save.mock.calls[0]![0].at(-1)).toEqual({ name: 'Block', ref: 'S!$A$2:$B$3', scope: 'Other sheet' })
    expect($('[data-xs-name-dialog]')).toBeNull()
  })

  it('a refusal shows Excel’s alert with the source’s message, and the dialog stays open; Cancel closes it', async () => {
    const { source, save } = spied()
    const w = await mountSheet(source)
    api(w).defineName()
    await settle()
    await setInput('[data-xs-name-field]', 'Sales')
    $<HTMLFormElement>('[data-xs-name-dialog]')!.requestSubmit()
    await settle()
    expect(save).toHaveBeenCalledTimes(1)
    expect($('[data-xs-name-alert="refused"]')!.textContent).toContain('The name Sales already exists')
    await click('[data-xs-name-answer="ok"]')
    expect($('[data-xs-name-alert]')).toBeNull()
    expect($('[data-xs-name-dialog]')).not.toBeNull()
    await click('[data-xs-name-cancel]')
    expect($('[data-xs-name-dialog]')).toBeNull()
  })

  it('is in the cell menu where names can be defined, and not on a read-only sheet', async () => {
    const w = await mountSheet(memorySource(book()))
    await w.find('[data-cell="2:2"]').trigger('contextmenu', { button: 2, clientX: 100, clientY: 100 })
    await settle()
    await click('[data-xs-menu="defineName"]')
    expect(input('[data-xs-name-refers]').value).toBe('=S!$B$2')
    const r = await mountSheet(memorySource(book()), { readOnly: true })
    await r.find('[data-cell="2:2"]').trigger('contextmenu', { button: 2, clientX: 100, clientY: 100 })
    await settle()
    expect($('[data-xs-menu="defineName"]')).toBeNull()
    expect(api(r).defineName()).toBe(false)
  })

  it('is not offered while a cell is being edited, as Excel greys it', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, '=')
    expect(api(w).defineName()).toBe(false)
    expect(api(w).nameManager()).toBe(false)
  })
})

describe('the Name Manager', () => {
  const rows = () => [...document.querySelectorAll<HTMLElement>('[data-xs-name-row]')].map((r) => [...r.children].map((c) => c.textContent))
  const selected = () => $('[data-xs-name-row][aria-selected=true]')?.dataset.xsNameRow

  it('Ctrl+F3 opens it with every name by name: its Refers To with = and its scope; the first one selected', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, 'F3', { ctrlKey: true })
    await settle()
    expect($('[data-xs-name-manager] .xs-modal-title')!.textContent).toBe('Name Manager')
    expect(rows()).toEqual([
      ['cols', '=S!$B:$C', 'Workbook'],
      ['Far', "='Other sheet'!$D$4", 'Workbook'],
      ['Own', "='Other sheet'!$A$1:$A$2", 'Other sheet'],
      ['Rate', '=0.2', 'Workbook'],
      ['Sales', '=S!$B$2:$C$3', 'Workbook'],
    ])
    expect(selected()).toBe('cols')
    expect(input('[data-xs-name-ref]').value).toBe('=S!$B:$C')
    const list = $('[data-xs-names-list]')!
    list.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }))
    await settle()
    expect(selected()).toBe('Far')
    await click('[data-xs-name-close]')
    expect($('[data-xs-name-manager]')).toBeNull()
  })

  it('Edit changes a name and its reference, keeping its scope; Delete asks first', async () => {
    const { source, save } = spied()
    const w = await mountSheet(source)
    api(w).nameManager()
    await settle()
    $('[data-xs-name-row="Own"]')!.click()
    await settle()
    await click('[data-xs-name-edit]')
    expect($('[data-xs-name-dialog] .xs-modal-title')!.textContent).toBe('Edit Name')
    expect($<HTMLSelectElement>('[data-xs-name-scope]')!.disabled).toBe(true)
    await setInput('[data-xs-name-field]', 'Mine')
    await setInput('[data-xs-name-refers]', "='Other sheet'!$A$1:$A$5")
    $<HTMLFormElement>('[data-xs-name-dialog]')!.requestSubmit()
    await settle()
    expect(save.mock.calls[0]![0].find((n) => n.name === 'Mine')).toEqual({ name: 'Mine', ref: "'Other sheet'!$A$1:$A$5", scope: 'Other sheet' })
    expect(save.mock.calls[0]![0].some((n) => n.name === 'Own')).toBe(false)
    expect(selected()).toBe('Mine')
    await click('[data-xs-name-delete]')
    expect($('[data-xs-name-alert="delete"]')!.textContent).toContain('Are you sure you want to delete the name Mine?')
    await click('[data-xs-name-answer="cancel"]')
    expect(save).toHaveBeenCalledTimes(1)
    await click('[data-xs-name-delete]')
    await click('[data-xs-name-answer="ok"]')
    expect(save).toHaveBeenCalledTimes(2)
    expect(rows().map((r) => r[0])).toEqual(['cols', 'Far', 'Rate', 'Sales'])
  })

  it('Refers to is edited in the box at the bottom: Commit replaces it, Cancel restores it, Close asks to save an edit', async () => {
    const { source, save } = spied()
    const w = await mountSheet(source)
    api(w).nameManager()
    await settle()
    $('[data-xs-name-row="Rate"]')!.click()
    await settle()
    expect($<HTMLButtonElement>('[data-xs-name-ref-commit]')!.disabled).toBe(true)
    await setInput('[data-xs-name-ref]', '=0.25')
    await click('[data-xs-name-ref-cancel]')
    expect(input('[data-xs-name-ref]').value).toBe('=0.2')
    await setInput('[data-xs-name-ref]', '=0.25')
    await click('[data-xs-name-ref-commit]')
    expect(save.mock.calls[0]![0].find((n) => n.name === 'Rate')).toEqual({ name: 'Rate', ref: '0.25' })
    expect(rows().find((r) => r[0] === 'Rate')![1]).toBe('=0.25')
    await setInput('[data-xs-name-ref]', '=0.3')
    await click('[data-xs-name-close]')
    expect($('[data-xs-name-alert="save"]')).not.toBeNull()
    await click('[data-xs-name-answer="yes"]')
    expect(save).toHaveBeenCalledTimes(2)
    expect($('[data-xs-name-manager]')).toBeNull()
  })

  it('a refused Commit shows the source’s message and keeps the edit', async () => {
    const { source } = spied()
    source.failNext({ ok: false, reason: 'refused', message: 'The formula has a problem' })
    const w = await mountSheet(source)
    api(w).nameManager()
    await settle()
    await setInput('[data-xs-name-ref]', '=SUM(')
    await click('[data-xs-name-ref-commit]')
    expect($('[data-xs-name-alert="refused"]')!.textContent).toContain('The formula has a problem')
    await click('[data-xs-name-answer="ok"]')
    expect(input('[data-xs-name-ref]').value).toBe('=SUM(')
  })

  it('on a read-only sheet it only lists the names, in Russian too', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true, messages: ruMessages })
    expect(api(w).nameManager()).toBe(true)
    await settle()
    expect($('[data-xs-name-manager] .xs-modal-title')!.textContent).toBe('Диспетчер имен')
    expect($('[data-xs-name-new]')).toBeNull()
    expect(input('[data-xs-name-ref]').readOnly).toBe(true)
    expect(rows()[0]).toEqual(['cols', '=S!$B:$C', 'Книга'])
  })
})

describe('names in Formula AutoComplete', () => {
  const shown = () => [...document.querySelectorAll('.xs-fl-item')].map((li) => li.textContent!.trim())

  it('lists the names visible on the sheet among the functions, alphabetically, with their own icon; Tab inserts a name without a parenthesis', async () => {
    const w = await mountSheet(memorySource(book(NAMES, ['RANK', 'RATE', 'ROUND', 'SALES.TAX', 'SUM'])))
    await key(w, '=')
    await type(w, '=r')
    await settle()
    // a name spelt as a function follows it
    expect(shown()).toEqual(['RANK', 'RATE', 'Rate', 'ROUND'])
    expect($('[data-xs-name="Rate"] .xs-fl-icon')).not.toBeNull()
    expect($('[data-xs-name="Rate"]')!.getAttribute('aria-label')).toBe('Rate, defined name')
    await editorKey(w, 'ArrowDown')
    await editorKey(w, 'ArrowDown')
    await editorKey(w, 'Tab')
    await settle()
    expect((w.find('.xs-editor').element as HTMLTextAreaElement).value).toBe('=Rate')
    // a sheet's own name is offered on its sheet only
    await type(w, '=Rate+o')
    await settle()
    expect(shown()).toEqual([])
  })

  it('names are offered without the source’s functionNames too', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, '=')
    await type(w, '=s')
    await settle()
    expect(shown()).toEqual(['Sales'])
  })
})

describe('names over localSource, in the workbook’s locale', () => {
  it('a name defined in the Name Box calculates in a formula; undo takes it away, and the Name Box shows the address again', async () => {
    const source = localSource({ ...localBook(), sheets: [{ name: 'Лист1', cells: { A1: { value: 2 }, B1: { formula: 'Двойка*3' } } }] })
    const w = await mountSheet(source, { locale: source.locale })
    await until(() => cellText(w, 1, 2) === '#ИМЯ?')
    await goTo(w, 'A1')
    await goTo(w, 'Двойка')
    await until(() => cellText(w, 1, 2) === '6')
    expect(await source.names()).toEqual([{ name: 'Двойка', ref: 'Лист1!$A$1' }])
    expect(address(w)).toBe('Двойка')
    await api(w).undo()
    await until(() => address(w) === 'A1')
    expect(await source.names()).toEqual([])
  })

  it('a Refers to that does not parse in the locale is refused with its message', async () => {
    const source = localSource(localBook())
    const w = await mountSheet(source, { locale: source.locale })
    api(w).defineName()
    await settle()
    await setInput('[data-xs-name-field]', 'Плохое')
    await setInput('[data-xs-name-refers]', '=СУММ(')
    $<HTMLFormElement>('[data-xs-name-dialog]')!.requestSubmit()
    await until(() => !!$('[data-xs-name-alert="refused"]'))
    expect($('[data-xs-name-alert] p:last-of-type')!.textContent).not.toBe('')
    expect(await source.names()).toEqual([])
  })
})
