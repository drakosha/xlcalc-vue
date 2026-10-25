// Excel's Go To (F5, Ctrl+G): the Reference box with the location the last jump left, the list of the last four, OK as the Name Box, and Excel's alert for a bad reference.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { absoluteRef } from '../src/core/goto'
import { memorySource, ruMessages, type MemoryBook, type SheetViewApi } from '../src'
import { active, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  sheets: [
    { name: 'S', cells: { A1: 1, A2: { value: 'place', link: { location: "'Sheet 2'!B2" } } } },
    { name: 'Sheet 2', cells: { B2: 'here' } },
  ],
})
const dialog = (w: VueWrapper) => w.find('[data-xs-goto]')
const box = (w: VueWrapper) => w.find('[data-xs-goto-ref]').element as HTMLInputElement
const items = (w: VueWrapper) => w.findAll('[data-xs-goto-item]').map((i) => i.text())
const tab = (w: VueWrapper) => w.find('.xs-tab-on').text()
async function open(w: VueWrapper): Promise<void> {
  await key(w, 'F5')
  await settle()
}
async function ok(w: VueWrapper, text?: string): Promise<void> {
  if (text !== undefined) await w.find('[data-xs-goto-ref]').setValue(text)
  await dialog(w).trigger('submit')
  await settle()
}

describe('absoluteRef', () => {
  it('makes every part of a selection absolute, as Excel lists it', () => {
    expect(absoluteRef('B2')).toBe('$B$2')
    expect(absoluteRef('B2:D5')).toBe('$B$2:$D$5')
    expect(absoluteRef('B:D')).toBe('$B:$D')
    expect(absoluteRef('2:5')).toBe('$2:$5')
  })
})

describe('Go To', () => {
  it('opens with F5 and with Ctrl+G on any layout, focused on an empty Reference box the first time', async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    expect(dialog(w).exists()).toBe(true)
    expect(document.activeElement).toBe(box(w))
    expect(box(w).value).toBe('')
    expect(items(w)).toEqual([])
    await dialog(w).trigger('keydown', { key: 'Escape' })
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(document.activeElement).toBe(w.find('.xs-scroller').element)
    // Russian layout: the key is the letter on G's place
    await key(w, 'п', { code: 'KeyG', ctrlKey: true })
    await settle()
    expect(dialog(w).exists()).toBe(true)
  })

  it('OK goes to the reference as the Name Box does, and the next time offers the location it left, so F5 and Enter go back', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B3')
    await open(w)
    await ok(w, 'C5')
    expect(dialog(w).exists()).toBe(false)
    expect(active(w)).toBe('5:3')
    await open(w)
    expect(box(w).value).toBe('S!$B$3')
    expect(items(w)).toEqual(['S!$B$3'])
    await ok(w)
    expect(active(w)).toBe('3:2')
    await open(w)
    expect(box(w).value).toBe('S!$C$5')
    expect(items(w)).toEqual(['S!$C$5', 'S!$B$3'])
  })

  it('goes to another sheet and to a range, and keeps the last four locations, newest first, each once', async () => {
    const w = await mountSheet(memorySource(book()))
    for (const place of ["'Sheet 2'!B2", 'S!A3:B4', 'C1', 'D1', 'E1', 'C1']) {
      await open(w)
      await ok(w, place)
    }
    expect(tab(w)).toBe('S')
    expect(active(w)).toBe('1:3')
    await open(w)
    expect(items(w)).toEqual(['S!$E$1', 'S!$D$1', 'S!$C$1', 'S!$A$3:$B$4'])
    await ok(w, "'Sheet 2'!$B$2")
    expect(tab(w)).toBe('Sheet 2')
    expect(active(w)).toBe('2:2')
  })

  it('a click on a location puts it into the Reference box, the arrows step through the list, and a double click goes there', async () => {
    const w = await mountSheet(memorySource(book()))
    for (const place of ['C3', 'D4']) {
      await open(w)
      await ok(w, place)
    }
    await open(w)
    expect(items(w)).toEqual(['S!$C$3', 'S!$A$1'])
    await w.findAll('[data-xs-goto-item]')[1]!.trigger('click')
    expect(box(w).value).toBe('S!$A$1')
    await dialog(w).trigger('keydown', { key: 'ArrowUp' })
    expect(box(w).value).toBe('S!$C$3')
    expect(w.findAll('[data-xs-goto-item]')[0]!.attributes('aria-selected')).toBe('true')
    await w.findAll('[data-xs-goto-item]')[0]!.trigger('dblclick')
    await settle()
    expect(dialog(w).exists()).toBe(false)
    expect(active(w)).toBe('3:3')
  })

  it("a text that is not a reference brings up Excel's alert, after which the dialog stays with the text selected", async () => {
    const w = await mountSheet(memorySource(book()), { messages: ruMessages })
    await open(w)
    expect(dialog(w).find('.xs-modal-title').text()).toBe('Переход')
    await ok(w, 'B2+C1')
    const alert = document.querySelector('[data-xs-alert="notice"]')
    expect(alert?.querySelector('[data-xs-alert-message]')?.textContent).toBe('Недопустимая ссылка.')
    expect(document.activeElement).toBe(alert?.querySelector('[data-xs-alert-button="ok"]'))
    alert!.querySelector<HTMLElement>('[data-xs-alert-button="ok"]')!.click()
    await settle()
    expect(document.querySelector('[data-xs-alert]')).toBeNull()
    expect(dialog(w).exists()).toBe(true)
    expect(document.activeElement).toBe(box(w))
    expect(box(w).value).toBe('B2+C1')
    expect(active(w)).toBe('1:1')
    await ok(w, 'Nowhere!A1')
    expect(document.querySelector('[data-xs-alert="notice"]')).not.toBeNull()
  })

  it('a hyperlink into the workbook leaves its location in the list too, and the read-only sheet goes as well', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'A2')
    await w.find('[data-cell="2:1"]').find('button.xs-link').trigger('click')
    await settle()
    expect(tab(w)).toBe('Sheet 2')
    expect(active(w)).toBe('2:2')
    await open(w)
    expect(box(w).value).toBe('S!$A$2')
    await ok(w)
    expect(tab(w)).toBe('S')
    expect(active(w)).toBe('2:1')
  })

  it('opens from the host through goTo()', async () => {
    const w = await mountSheet(memorySource(book()))
    expect((w.vm as unknown as SheetViewApi).goTo()).toBe(true)
    await settle()
    expect(dialog(w).exists()).toBe(true)
    await w.find('[data-xs-dialog-cancel]').trigger('click')
    await settle()
    expect(dialog(w).exists()).toBe(false)
  })
})
