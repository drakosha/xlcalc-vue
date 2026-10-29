// Excel's notes shown all the time: a note with `shown` drawn in its box next to its cell with a connector, the cell menu's
// Show/Hide Note through the source's `note`, an edited note staying shown, and Show All Notes as a view setting.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, ruMessages, type SheetViewApi, type SheetViewState } from '../src'
import type { MemoryBook } from '../src/contract/memory'
import { mountSheet, settle, until } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: 'Total',
        B2: { value: 'Due', note: { author: 'Ann', text: 'On Friday', shown: true } },
        C4: { value: 'Owner', note: { text: 'Bob' } },
      },
    },
  ],
})
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).map((e) => e[0] as SheetViewState | null).at(-1)
const shown = (w: VueWrapper) => w.findAll('[data-xs-note-shown]').map((b) => b.attributes('data-xs-note-shown'))
const items = () => [...document.querySelectorAll<HTMLElement>('[data-xs-menu]')].map((b) => b.dataset.xsMenu)
async function context(w: VueWrapper, cell: string): Promise<void> {
  await w.find(`[data-cell="${cell}"]`).trigger('contextmenu', { button: 2, clientX: 200, clientY: 100 })
  await settle()
}
async function pick(id: string): Promise<void> {
  document.querySelector<HTMLElement>(`[data-xs-menu="${id}"]`)!.click()
  await settle()
}

describe('notes shown all the time', () => {
  it('draws a shown note in its box to the right of its cell, with a line from the cell’s corner, and not a hidden one', async () => {
    const w = await mountSheet(memorySource(book()))
    expect(shown(w)).toEqual(['2:2'])
    const box = w.find('[data-xs-note-shown="2:2"]')
    expect(box.find('.xs-note-author').text()).toBe('Ann:')
    expect(box.find('.xs-note-text').text()).toBe('On Friday')
    // the line runs from the cell's top right corner to the box, which starts to the right of the cell
    const line = w.find('[data-xs-note-line="2:2"]')
    const left = parseFloat((box.element as HTMLElement).style.left)
    expect(Number(line.attributes('x2'))).toBe(left)
    expect(Number(line.attributes('x1'))).toBeLessThan(left)
    expect(Number(line.attributes('y1'))).toBeLessThanOrEqual(parseFloat((box.element as HTMLElement).style.top) + 24)
  })

  it("the cell menu's Show/Hide Note shows a note and hides it again through `note`, after Edit Note and Delete Note", async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const w = await mountSheet(source)
    await context(w, '4:3')
    expect(items().slice(-5, -2)).toEqual(['editNote', 'deleteNote', 'showNote'])
    await pick('showNote')
    expect(note).toHaveBeenLastCalledWith('S', 'C4', { text: 'Bob', shown: true })
    expect(shown(w)).toEqual(['2:2', '4:3'])
    await context(w, '2:2')
    await pick('showNote')
    expect(note).toHaveBeenLastCalledWith('S', 'B2', { author: 'Ann', text: 'On Friday' })
    // a repaint that closely follows another is read once they stop
    await until(() => shown(w).length === 1)
    expect(shown(w)).toEqual(['4:3'])
  })

  it('keeps a note shown when it is edited, and does not draw it twice meanwhile', async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const w = await mountSheet(source)
    await context(w, '2:2')
    await pick('editNote')
    expect(shown(w)).toEqual([])
    await w.find('[data-xs-note-text]').setValue('On Monday')
    await w.find('[data-xs-note-text]').trigger('keydown', { key: 'Escape' })
    await settle()
    expect(note).toHaveBeenLastCalledWith('S', 'B2', { author: 'Ann', text: 'On Monday', shown: true })
    expect(shown(w)).toEqual(['2:2'])
  })

  it('does not peek at a shown note under the mouse, which is in view already', async () => {
    const w = await mountSheet(memorySource(book()))
    w.find('[data-cell="2:2"]').element.dispatchEvent(new PointerEvent('pointermove', { bubbles: true, pointerType: 'mouse' }))
    await settle()
    expect(w.find('[data-xs-note-peek]').exists()).toBe(false)
  })

  it('showNotes is Show All Notes: every note in view, a view setting that sends nothing', async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const w = await mountSheet(source)
    expect(api(w).showNotes()).toBe(true)
    await settle()
    expect(shown(w)).toEqual(['2:2', '4:3'])
    expect(lastState(w)).toMatchObject({ allNotes: true })
    expect(api(w).showNotes(false)).toBe(false)
    await settle()
    expect(shown(w)).toEqual(['2:2'])
    expect(lastState(w)).not.toHaveProperty('allNotes')
    expect(note).not.toHaveBeenCalled()
  })

  it("names Show/Hide Note in the interface's language", async () => {
    const w = await mountSheet(memorySource(book()), { messages: ruMessages })
    await context(w, '2:2')
    expect(document.querySelector('[data-xs-menu="showNote"]')!.textContent).toBe('Показать или скрыть примечание')
  })
})
