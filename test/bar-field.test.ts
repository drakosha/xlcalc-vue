// The formula bar is always a field, as Excel's: it holds the active cell's text, a press starts the edit with the caret or the selection where the mouse put it, and a read-only sheet lets the text be selected only.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { active, address, barField, cellText, editorOpen, goTo, key, mountSheet, pressBar, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 1, A2: 2, B1: { value: 3, formula: 'SUM(A1:A2)' }, C1: 'text' } }] })
const editingBar = (w: VueWrapper) => w.find('.xs-bar').classes().includes('xs-bar-editing')

describe('the formula bar field', () => {
  it('holds the active cell’s text outside an edit and follows the cell, its changes and a cancel', async () => {
    const w = await mountSheet(memorySource(book()))
    expect(barField(w).value).toBe('1')
    await goTo(w, 'B1')
    expect(barField(w).value).toBe('=SUM(A1:A2)')
    // One field, no buttons over it.
    expect(w.findAll('.xs-fx button')).toHaveLength(0)
    // A range: the active cell's text, no address before it.
    await goTo(w, 'B1:C2')
    expect(barField(w).value).toBe('=SUM(A1:A2)')
    await goTo(w, 'C1')
    await key(w, 'x')
    await type(w, 'typed')
    expect(barField(w).value).toBe('typed')
    await w.find('.xs-editor').trigger('keydown', { key: 'Escape' })
    await settle()
    expect(barField(w).value).toBe('text')
    await key(w, 'x')
    await type(w, 'new')
    await w.find('.xs-editor').trigger('keydown', { key: 'Enter' })
    await settle()
    await goTo(w, 'C1')
    expect(barField(w).value).toBe('new')
  })

  it('is read-only on a sheet without edit: the text can be selected, no edit starts', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B1')
    const el = barField(w)
    expect(el.readOnly).toBe(true)
    expect(el.value).toBe('=SUM(A1:A2)')
    await pressBar(w, 1, 4)
    expect(editingBar(w)).toBe(false)
    expect([el.selectionStart, el.selectionEnd]).toEqual([1, 4])
    await w.find('.xs-fx-input').trigger('keydown', { key: 'x' })
    await w.find('.xs-fx-input').trigger('keydown', { key: 'F2' })
    await settle()
    expect(editingBar(w)).toBe(false)
    expect(w.find('.xs-editor:not(.xs-editor-idle)').exists()).toBe(false)
  })

  it('a click starts the edit in the bar with the caret where it was put, not at the end', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B1')
    await pressBar(w, 5)
    const el = barField(w)
    expect(editingBar(w)).toBe(true)
    expect(document.activeElement).toBe(el)
    expect([el.selectionStart, el.selectionEnd]).toEqual([5, 5])
  })

  it('a drag starts the edit with the dragged text selected', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B1')
    await pressBar(w, 5, 10)
    const el = barField(w)
    expect(editingBar(w)).toBe(true)
    expect([el.selectionStart, el.selectionEnd]).toEqual([5, 10])
  })

  it('a click on the bar moves a field open in the cell there with its draft, the caret where it was put', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'C1')
    await key(w, 'F2')
    await settle()
    await type(w, 'textual')
    await pressBar(w, 2)
    const el = barField(w)
    expect(editingBar(w)).toBe(true)
    expect(editorOpen(w)).toBe(false)
    expect(el.value).toBe('textual')
    expect([el.selectionStart, el.selectionEnd]).toEqual([2, 2])
  })

  it('Enter after an edit in the bar writes the cell and moves down, as before', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'C1')
    await pressBar(w, 4)
    const el = barField(w)
    el.value = 'te-xt'
    await w.find('.xs-fx-input').trigger('input')
    await w.find('.xs-fx-input').trigger('keydown', { key: 'Enter' })
    await settle()
    expect(cellText(w, 1, 3)).toBe('te-xt')
    expect(active(w)).toBe('2:3')
    expect(barField(w).value).toBe('')
  })

  it('a plain click on a reference puts the caret in it and starts the edit, as Excel', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B1')
    await pressBar(w, 6)
    expect(editingBar(w)).toBe(true)
    expect([barField(w).selectionStart, barField(w).selectionEnd]).toEqual([6, 6])
    expect(address(w)).toBe('B1')
  })

  it('Ctrl+click or Cmd+click on a reference outside an edit goes to its range; elsewhere it is a click', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B1')
    expect(barField(w).title).toContain('+click a reference to go to it')
    await pressBar(w, 6, 6, { ctrlKey: true })
    expect(editingBar(w)).toBe(false)
    expect(address(w)).toBe('A1:A2')
    await goTo(w, 'B1')
    await pressBar(w, 9, 9, { metaKey: true })
    expect(editingBar(w)).toBe(false)
    expect(address(w)).toBe('A1:A2')
    await goTo(w, 'B1')
    await pressBar(w, 2, 2, { ctrlKey: true })
    expect(editingBar(w)).toBe(true)
    expect(address(w)).toBe('B1')
  })

  it('on a read-only sheet Ctrl+click on a reference goes to its range too', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B1')
    await pressBar(w, 5, 5, { ctrlKey: true })
    expect(address(w)).toBe('A1:A2')
    expect(editingBar(w)).toBe(false)
  })

  it('colours the references outside an edit, and while one is typed by the draft in the colours of its frames on the grid', async () => {
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'B1')
    const refs = () => w.findAll('.xs-fx-layer .xs-ref')
    expect(w.find('.xs-fx-layer').text()).toBe('=SUM(A1:A2)')
    expect(refs().map((r) => r.text())).toEqual(['A1:A2'])
    await goTo(w, 'D1')
    await key(w, '=')
    await type(w, '=A1+C1*A1')
    await settle()
    expect(w.find('.xs-fx-layer').text()).toBe('=A1+C1*A1')
    expect(refs().map((r) => r.text())).toEqual(['A1', 'C1', 'A1'])
    const color = (i: number) => (refs()[i]!.element as HTMLElement).style.color
    expect(color(0)).not.toBe('')
    expect(color(0)).toBe(color(2))
    expect(color(0)).not.toBe(color(1))
    expect((w.find('[data-cell="1:1"]').element as HTMLElement).style.boxShadow).toContain(color(0))
    expect((w.find('[data-cell="1:3"]').element as HTMLElement).style.boxShadow).toContain(color(1))
  })
})
