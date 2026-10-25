// Excel's Insert Hyperlink dialog (Ctrl+K, Edit Hyperlink), the cell menu's hyperlink and note items, and the note box: each change
// goes to the source's `link` or `note`, a changed Text to Display to `edit` first, and the windows over it are read again.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi } from '../src'
import { cellText, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: {
        A1: 'Total',
        A2: { value: 'Report', link: { href: 'https://example.com/r', tooltip: 'The report' } },
        A3: { value: 'Data', link: { location: "'Q 2'!B3" } },
        B1: { value: 3, formula: 'SUM(C1:C2)' },
        C1: { value: 'Due', note: { author: 'Ann', text: 'On Friday' } },
      },
    },
    { name: 'Q 2', cells: { B3: 7 } },
    { name: 'Hidden', hidden: true, cells: {} },
  ],
})
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const items = () => [...document.querySelectorAll<HTMLElement>('[data-xs-menu]')].map((b) => b.dataset.xsMenu)
async function context(w: VueWrapper, cell: string): Promise<void> {
  await w.find(`[data-cell="${cell}"]`).trigger('contextmenu', { button: 2, clientX: 200, clientY: 100 })
  await settle()
}
async function pick(id: string): Promise<void> {
  document.querySelector<HTMLElement>(`[data-xs-menu="${id}"]`)!.click()
  await settle()
}
const field = (w: VueWrapper, name: string) => w.find(`[data-xs-link-${name}]`).element as HTMLInputElement
async function typeIn(w: VueWrapper, name: string, text: string): Promise<void> {
  await w.find(`[data-xs-link-${name}]`).setValue(text)
}
async function ok(w: VueWrapper): Promise<void> {
  await w.find('[data-xs-link]').trigger('submit')
  await settle()
}
/** A memory source whose `edit` and `link` calls are logged in order. */
function logged() {
  const source = memorySource(book())
  const log: string[] = []
  const edit = vi.spyOn(source, 'edit').mockImplementation(async (entries) => {
    log.push(`edit ${entries.map((e) => `${e.cell}=${e.text}`).join(',')}`)
    return { ok: true, changes: [] }
  })
  const link = vi.spyOn(source, 'link')
  link.mockImplementation(async (_sheet, range, l) => {
    log.push(`link ${range} ${JSON.stringify(l)}`)
    return { ok: true, changes: [] }
  })
  return { source, log, edit, link }
}

describe('Insert Hyperlink', () => {
  it('Ctrl+K opens it with the cell’s text; OK writes the address through `link`, and the windows are read again', async () => {
    const source = memorySource(book())
    const link = vi.spyOn(source, 'link')
    const edit = vi.spyOn(source, 'edit')
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source)
    await key(w, 'k', { ctrlKey: true })
    await settle()
    expect(w.find('[data-xs-link] .xs-modal-title').text()).toBe('Insert Hyperlink')
    expect(field(w, 'text').value).toBe('Total')
    await typeIn(w, 'address', 'https://example.com/total')
    await typeIn(w, 'tip', 'All of it')
    const before = window.mock.calls.length
    await ok(w)
    expect(link).toHaveBeenCalledWith('S', 'A1', { href: 'https://example.com/total', tooltip: 'All of it' })
    expect(edit).not.toHaveBeenCalled()
    expect(w.find('[data-xs-link]').exists()).toBe(false)
    expect(window.mock.calls.length).toBeGreaterThan(before)
    expect(w.find('[data-cell="1:1"] a').attributes('href')).toBe('https://example.com/total')
  })

  it('on an empty cell Text to Display follows the address, and is entered before the link, as a separate step', async () => {
    const { source, log } = logged()
    const w = await mountSheet(source)
    await goTo(w, 'D4')
    expect(api(w).insertLink()).toBe(true)
    await settle()
    await typeIn(w, 'address', 'https://example.com')
    expect(field(w, 'text').value).toBe('https://example.com')
    await ok(w)
    expect(log).toEqual(['edit D4=https://example.com', 'link D4 {"href":"https://example.com"}'])
  })

  it('a changed Text to Display is entered first; a typed one no longer follows the address', async () => {
    const { source, log } = logged()
    const w = await mountSheet(source)
    await goTo(w, 'D4')
    api(w).insertLink()
    await settle()
    await typeIn(w, 'text', 'Home')
    await typeIn(w, 'address', 'https://example.com')
    expect(field(w, 'text').value).toBe('Home')
    await ok(w)
    expect(log).toEqual(['edit D4=Home', 'link D4 {"href":"https://example.com"}'])
  })

  it('Place in This Document lists the visible sheets and writes the reference on the chosen one; a bad reference is refused', async () => {
    const { source, log } = logged()
    const w = await mountSheet(source)
    api(w).insertLink()
    await settle()
    await w.find('[data-xs-link-to="place"]').trigger('change')
    expect(w.findAll('[data-xs-link-place]').map((p) => p.text())).toEqual(['S', 'Q 2'])
    await w.find('[data-xs-link-place="sheet:Q 2"]').trigger('click')
    await typeIn(w, 'ref', 'B3+1')
    await ok(w)
    expect(w.find('[data-xs-link] [role=alert]').text()).toBe('Reference isn’t valid.')
    expect(log).toEqual([])
    await typeIn(w, 'ref', 'b3:c4')
    await ok(w)
    expect(log).toEqual(['link A1 {"location":"\'Q 2\'!B3:C4"}'])
  })

  it('several cells or a formula show <<Selection in Document>>, which the link leaves alone', async () => {
    const { source, log } = logged()
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    api(w).insertLink()
    await settle()
    expect(field(w, 'text').disabled).toBe(true)
    expect(field(w, 'text').value).toBe('<<Selection in Document>>')
    await w.find('[data-xs-dialog-cancel]').trigger('click')
    await goTo(w, 'D1:D3')
    api(w).insertLink()
    await settle()
    expect(field(w, 'text').disabled).toBe(true)
    await typeIn(w, 'address', 'https://example.com')
    await ok(w)
    expect(log).toEqual(['link D1:D3 {"href":"https://example.com"}'])
  })

  it('on a cell with a link it is Edit Hyperlink, filled from the link, with Remove Link', async () => {
    const { source, log } = logged()
    const w = await mountSheet(source)
    await goTo(w, 'A3')
    await key(w, 'k', { ctrlKey: true })
    await settle()
    expect(w.find('[data-xs-link] .xs-modal-title').text()).toBe('Edit Hyperlink')
    expect((w.find('[data-xs-link-to="place"]').element as HTMLInputElement).checked).toBe(true)
    expect(w.find('[data-xs-link-place="sheet:Q 2"]').attributes('aria-selected')).toBe('true')
    expect(field(w, 'ref').value).toBe('B3')
    await w.find('[data-xs-link-remove]').trigger('click')
    await settle()
    expect(log).toEqual(['link A3 null'])
    await goTo(w, 'A2')
    api(w).insertLink()
    await settle()
    expect(field(w, 'address').value).toBe('https://example.com/r')
    expect(field(w, 'tip').value).toBe('The report')
  })
})

describe('the cell menu’s hyperlink items', () => {
  it('Link on a cell without one; Edit, Open and Remove Hyperlink on a cell with one', async () => {
    const { source, log } = logged()
    const w = await mountSheet(source)
    await context(w, '1:1')
    expect(items().slice(-2)).toEqual(['newNote', 'link'])
    await pick('link')
    expect(w.find('[data-xs-link]').exists()).toBe(true)
    await w.find('[data-xs-dialog-cancel]').trigger('click')
    await context(w, '3:1')
    expect(items().slice(-3)).toEqual(['editLink', 'openLink', 'removeLink'])
    await pick('openLink')
    expect(w.find('.xs-tab-on').text()).toContain('Q 2')
    expect(w.find('.xs-address').element).toHaveProperty('value', 'B3')
    await goTo(w, 'S!A2')
    await context(w, '2:1')
    await pick('removeLink')
    expect(log).toEqual(['link A2 null'])
  })

  it('a right click on the link itself opens this menu; the click of a long press’s release on an item chooses nothing', async () => {
    const w = await mountSheet(memorySource(book()))
    await w.find('[data-cell="2:1"] a').trigger('contextmenu', { button: 2, clientX: 200, clientY: 100 })
    await settle()
    expect(items().slice(-3)).toEqual(['editLink', 'openLink', 'removeLink'])
    const item = document.querySelector<HTMLElement>('[data-xs-menu="editLink"]')!
    item.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 1 }))
    await settle()
    expect(w.find('[data-xs-link]').exists()).toBe(false)
    item.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    item.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 1 }))
    await settle()
    expect(w.find('[data-xs-link]').exists()).toBe(true)
  })

  it('a read-only sheet only opens links: no Ctrl+K, no insertLink()', async () => {
    const source = memorySource(book())
    const open = vi.spyOn(window, 'open').mockImplementation(() => null)
    const w = await mountSheet(source, { readOnly: true })
    await context(w, '2:1')
    expect(items()).toEqual(['copy', 'precedents', 'openLink'])
    await pick('openLink')
    expect(open).toHaveBeenCalledWith('https://example.com/r', '_blank', 'noopener')
    expect(api(w).insertLink()).toBe(false)
    await key(w, 'k', { ctrlKey: true })
    await settle()
    expect(w.find('[data-xs-link]').exists()).toBe(false)
  })

  it('Open Hyperlink goes to the host when it listens', async () => {
    const w = await mountSheet(memorySource(book()), { hostLinks: true })
    await context(w, '2:1')
    await pick('openLink')
    expect(w.emitted('openLink')?.[0]?.[0]).toMatchObject({ href: 'https://example.com/r', sheet: 'S', cell: 'A2' })
  })
})

describe('notes', () => {
  const box = (w: VueWrapper) => w.find('[data-xs-note-box]')
  const text = (w: VueWrapper) => w.find('[data-xs-note-text]')

  it('New Note opens the box at the cell with the author in bold; Esc saves it through `note`', async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source, { noteAuthor: 'Bob' })
    await context(w, '1:1')
    await pick('newNote')
    expect(box(w).exists()).toBe(true)
    expect(w.find('.xs-note-author').text()).toBe('Bob:')
    expect(document.activeElement).toBe(text(w).element)
    await text(w).setValue('Check the total')
    const before = window.mock.calls.length
    await text(w).trigger('keydown', { key: 'Escape' })
    await settle()
    expect(box(w).exists()).toBe(false)
    expect(note).toHaveBeenCalledWith('S', 'A1', { author: 'Bob', text: 'Check the total' })
    expect(window.mock.calls.length).toBeGreaterThan(before)
    expect(w.find('[data-cell="1:1"]').attributes('title')).toContain('Check the total')
  })

  it('Shift+F2 edits the note of a cell that has one; a press outside saves the change', async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const w = await mountSheet(source, { noteAuthor: 'Bob' })
    await goTo(w, 'C1')
    await key(w, 'F2', { shiftKey: true })
    await settle()
    expect(w.find('.xs-editor-idle').exists()).toBe(true)
    expect((text(w).element as HTMLTextAreaElement).value).toBe('On Friday')
    expect(w.find('.xs-note-author').text()).toBe('Ann:')
    await text(w).setValue('On Monday')
    document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
    await settle()
    expect(note).toHaveBeenCalledWith('S', 'C1', { author: 'Ann', text: 'On Monday' })
  })

  it('an unchanged note, or an empty new one without an author, sends nothing', async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const w = await mountSheet(source)
    await goTo(w, 'C1')
    await key(w, 'F2', { shiftKey: true })
    await settle()
    await text(w).trigger('keydown', { key: 'Escape' })
    await goTo(w, 'D1')
    await key(w, 'F2', { shiftKey: true })
    await settle()
    expect(w.find('.xs-note-author').exists()).toBe(false)
    await text(w).trigger('keydown', { key: 'Escape' })
    await settle()
    expect(note).not.toHaveBeenCalled()
  })

  it('the menu has Edit Note and Delete Note on a cell with one; Delete Note removes it', async () => {
    const source = memorySource(book())
    const note = vi.spyOn(source, 'note')
    const w = await mountSheet(source)
    await context(w, '1:3')
    expect(items().slice(-3)).toEqual(['editNote', 'deleteNote', 'link'])
    await pick('deleteNote')
    expect(note).toHaveBeenCalledWith('S', 'C1', null)
    expect(w.find('[data-cell="1:3"]').attributes('title') ?? '').not.toContain('On Friday')
    expect(cellText(w, 1, 3)).toBe('Due')
  })

  it('without `note`, or read-only, there are no note items and Shift+F2 edits the cell as F2', async () => {
    const source = memorySource(book())
    delete (source as { note?: unknown }).note
    const w = await mountSheet(source)
    await context(w, '1:1')
    expect(items()).not.toContain('newNote')
    await w.find('.xs-scroller').trigger('keydown', { key: 'Escape' })
    await key(w, 'F2', { shiftKey: true })
    await settle()
    expect(box(w).exists()).toBe(false)
    expect(w.find('.xs-editor-idle').exists()).toBe(false)
  })
})
