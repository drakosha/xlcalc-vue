// The formula bar's height, as in Excel: one line by default, a button and Ctrl+Shift+U expand and collapse it, and the choice persists as the cell changes.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { enMessages, memorySource, ruMessages, type MemoryBook } from '../src'
import { barField, goTo, key, mountSheet, pressBar } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'one\ntwo\nthree\nfour\nfive\nsix', A2: 'single' } }] })
const toggle = (w: VueWrapper) => w.find('[data-xs-bar-toggle]')
const expanded = (w: VueWrapper) => w.find('.xs-fx').classes().includes('xs-fx-open')

describe('formula bar height', () => {
  it('shows a multi-line cell collapsed to one line, with the button to expand it', async () => {
    const w = await mountSheet(memorySource(book()))
    expect(barField(w).value).toBe('one\ntwo\nthree\nfour\nfive\nsix')
    expect(expanded(w)).toBe(false)
    expect(toggle(w).attributes('aria-expanded')).toBe('false')
    expect(toggle(w).attributes('aria-label')).toBe(enMessages.expandBar)
    expect(toggle(w).attributes('title')).toBe(enMessages.expandBar)
  })

  it('the button expands and collapses it, and the labels follow', async () => {
    const w = await mountSheet(memorySource(book()), { messages: ruMessages })
    expect(toggle(w).attributes('aria-label')).toBe(ruMessages.expandBar)
    await toggle(w).trigger('click')
    expect(expanded(w)).toBe(true)
    expect(toggle(w).attributes('aria-expanded')).toBe('true')
    expect(toggle(w).attributes('aria-label')).toBe(ruMessages.collapseBar)
    await toggle(w).trigger('click')
    expect(expanded(w)).toBe(false)
    expect(toggle(w).attributes('aria-expanded')).toBe('false')
  })

  it('Ctrl+Shift+U on the grid toggles it, with any layout; Ctrl+U alone does not', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, 'U', { ctrlKey: true, shiftKey: true, code: 'KeyU' })
    expect(expanded(w)).toBe(true)
    await key(w, 'Г', { ctrlKey: true, shiftKey: true, code: 'KeyU' })
    expect(expanded(w)).toBe(false)
    await key(w, 'u', { ctrlKey: true, code: 'KeyU' })
    expect(expanded(w)).toBe(false)
  })

  it("Ctrl+Shift+U in the formula bar's input toggles it, and the edit continues", async () => {
    const w = await mountSheet(memorySource(book()))
    await pressBar(w, 3)
    const editing = () => w.find('.xs-bar').classes().includes('xs-bar-editing')
    expect(editing()).toBe(true)
    await w.find('.xs-fx-input').trigger('keydown', { key: 'U', ctrlKey: true, shiftKey: true, code: 'KeyU' })
    expect(expanded(w)).toBe(true)
    expect(editing()).toBe(true)
  })

  it('stays expanded as the active cell changes, until toggled again', async () => {
    const w = await mountSheet(memorySource(book()))
    await toggle(w).trigger('click')
    await goTo(w, 'A2')
    expect(barField(w).value).toBe('single')
    expect(expanded(w)).toBe(true)
    await goTo(w, 'A1')
    expect(expanded(w)).toBe(true)
  })
})
