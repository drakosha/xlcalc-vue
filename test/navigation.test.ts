// Jumps and the bar: the address box, history back, the flash, `focus`, `state`, the cell link, formula references and
// their highlight, selection by headers with Shift, totals and copying; and the phone: a narrow frame, tap and double tap.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { fakeSource, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const book = (): MemoryBook => ({
  sheets: [
    { name: 'One', cells: { A1: 1, A2: 2, A3: 3, B1: { value: 6, formula: 'SUM(A1:A3)' }, B2: { value: 4, formula: "A1+'Two sheet'!B2" }, C1: 'x' } },
    { name: 'Two sheet', cells: { B2: 'far', D40: 'end' } },
  ],
})
const active = (w: VueWrapper) => w.find('.xs-cell.xs-active').attributes('data-cell')
const sel = (w: VueWrapper) => w.findAll('.xs-cell.xs-sel').map((c) => c.attributes('data-cell'))
async function goTo(w: VueWrapper, address: string): Promise<void> {
  const input = w.find('.xs-address')
  await input.trigger('focus')
  await input.setValue(address)
  await input.trigger('keydown', { key: 'Enter' })
  await settle()
}

describe('address box and history', () => {
  it('goes to a range on another sheet, flashes it, and back returns', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B2')
    await goTo(w, "'Two sheet'!B2:C3")
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    expect(active(w)).toBe('2:2')
    expect(sel(w).sort()).toEqual(['2:2', '2:3', '3:2', '3:3'])
    expect(w.findAll('.xs-flash').length).toBe(4)
    expect((w.find('.xs-address').element as HTMLInputElement).value).toBe('B2:C3')
    await w.find('.xs-bar .xs-tool').trigger('click')
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('One')
    expect(active(w)).toBe('2:2')
  })

  it('a sheet alone switches the tab; a bad address says so', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true, messages: { addressBad: 'Bad address' } })
    await goTo(w, "'Two sheet'")
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    await goTo(w, 'Nope!A1')
    expect(w.find('.xs-message').text()).toBe('Bad address')
    await goTo(w, 'B2+C1')
    expect(w.find('.xs-message').text()).toBe('Bad address')
  })

  it('takes `focus` given before the book loaded and every new focus after', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true, focus: { ref: "'Two sheet'!D40" } })
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    expect(active(w)).toBe('40:4')
    await w.setProps({ focus: { ref: 'One!C1' } })
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('One')
    expect(active(w)).toBe('1:3')
  })

  it('reports `state` and copies the link to the cell', async () => {
    const states: unknown[] = []
    const writeText = vi.fn(async () => {})
    Object.defineProperty(navigator, 'clipboard', { value: { writeText }, configurable: true })
    const w = await mountSheet(memorySource(book()), { readOnly: true, onState: (s: unknown) => states.push(s), cellUrl: (ref: string) => `https://host/book?at=${ref}` })
    await goTo(w, 'B1')
    expect(states.at(-1)).toEqual({ sheet: 'One', ref: 'B1', formula: 'SUM(A1:A3)', value: '6' })
    await goTo(w, "'Two sheet'!B2")
    expect(states.at(-1)).toEqual({ sheet: 'Two sheet', ref: 'B2', value: 'far' })
    await w.find('.xs-tools .xs-tool').trigger('click')
    await settle()
    expect(writeText).toHaveBeenCalledWith("https://host/book?at='Two sheet'!B2")
    w.unmount()
    expect(states.at(-1)).toBeNull()
  })
})

describe('formula bar', () => {
  it('shows the formula with references; a click selects the range, on another sheet too', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B1')
    const refs = w.findAll('.xs-ref')
    expect(refs.map((r) => r.text())).toEqual(['A1:A3'])
    expect(w.find('.xs-fx-text').text()).toContain('→')
    // The formula's range is highlighted on the grid.
    expect((w.find('[data-cell="2:1"]').element as HTMLElement).style.backgroundImage).toContain('linear-gradient')
    await refs[0]!.trigger('click')
    await settle()
    expect(sel(w).sort()).toEqual(['1:1', '2:1', '3:1'])
    await goTo(w, 'B2')
    await w.findAll('.xs-ref')[1]!.trigger('click')
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    expect(active(w)).toBe('2:2')
  })
})

describe('selection', () => {
  it('by column and row headers, with Shift; totals over loaded cells, the count beyond them', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await w.findAll('.xs-head')[0]!.trigger('click')
    await w.findAll('.xs-head')[1]!.trigger('click', { shiftKey: true })
    expect(w.find('.xs-stats').text()).toBe('Count: 5 · Sum: 16 · Average: 3.2 · Min: 1 · Max: 6')
    await w.findAll('.xs-num')[1]!.trigger('click')
    expect(sel(w)).toEqual(['2:1', '2:2', '2:3'])
    await goTo(w, 'A1:XFD1048576')
    expect(w.find('.xs-stats').text()).toBe('Count: 6 · Sum: 16 · Average: 3.2 · Min: 1 · Max: 6 · Numbers: 5')
  })

  it('a selection larger than the loaded windows shows only its size', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 5000, maxCol: 3 }])
    const w = await mountSheet(f.source, { readOnly: true })
    await w.findAll('.xs-head')[0]!.trigger('click')
    expect(w.find('.xs-stats').text()).toBe('Cells: 5,000')
  })

  it('copies the selection as TSV', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'A1:C2')
    const data: Record<string, string> = {}
    const e = new Event('copy', { bubbles: true, cancelable: true }) as ClipboardEvent
    Object.defineProperty(e, 'clipboardData', { value: { setData: (t: string, v: string) => (data[t] = v) } })
    w.find('.xs-scroller').element.dispatchEvent(e)
    expect(data['text/plain']).toBe('1\t6\tx\n2\t4\t')
  })

  it('keys move as in the the app: arrows, Ctrl+arrows, Shift extends, Ctrl+A selects all', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    const key = (k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
    await key('ArrowRight')
    await key('ArrowDown')
    expect(active(w)).toBe('2:2')
    await key('ArrowDown', { shiftKey: true })
    expect(sel(w).sort()).toEqual(['2:2', '3:2'])
    await key('ArrowDown', { ctrlKey: true })
    expect(active(w)).toBe('3:2')
    await key('Home', { ctrlKey: true })
    expect(active(w)).toBe('1:1')
    await key('a', { ctrlKey: true })
    expect(sel(w).length).toBe(9)
  })
})

describe('keys on an editable sheet', () => {
  it('Ctrl jumps to the data edge first, then to the end of the grid', async () => {
    const w = await mountSheet(memorySource(book()))
    const key = (k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
    await key('ArrowDown', { ctrlKey: true })
    expect(active(w)).toBe('3:1')
    await key('ArrowDown', { ctrlKey: true })
    await settle()
    expect((w.find('.xs-address').element as HTMLInputElement).value).toBe('A100')
    await key('End', { ctrlKey: true })
    expect(active(w)).toBe('3:3')
  })
})

describe('a phone', () => {
  it('a narrow frame narrows the numbers; a tap selects, a double tap and a tap on the bar open the field', async () => {
    const width = vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-scroller') ? 390 : 0
    })
    const w = await mountSheet(memorySource(book()))
    expect(w.find('.xs-sheet').classes()).toContain('xs-narrow')
    expect(Number.parseInt((w.find('col').element as HTMLElement).style.width)).toBeLessThan(56)
    const cell = w.find('[data-cell="3:1"]')
    await cell.trigger('pointerdown', { pointerType: 'touch', button: 0 })
    await cell.trigger('click')
    expect(active(w)).toBe('3:1')
    expect(w.find('.xs-editor').classes()).toContain('xs-editor-idle')
    await cell.trigger('pointerdown', { pointerType: 'touch', button: 0 })
    await cell.trigger('click')
    await settle()
    expect(w.find('.xs-editor').classes()).not.toContain('xs-editor-idle')
    expect(document.activeElement).toBe(w.find('.xs-editor').element)
    await w.find('.xs-editor').trigger('keydown', { key: 'Escape' })
    await w.find('.xs-fx').trigger('click')
    await settle()
    expect(document.activeElement).toBe(w.find('.xs-fx-input').element)
    expect((w.find('.xs-fx-input').element as HTMLTextAreaElement).value).toBe('3')
    width.mockRestore()
  })
})
