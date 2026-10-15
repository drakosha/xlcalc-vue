// Jumps and the bar: the address box, history back, the flash, `focus`, `state`, the cell link, formula references and
// their highlight, selection by headers with Shift, totals and copying; and the phone: a narrow frame, tap and double tap.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { barField, fakeSource, mountSheet, pressBar, settle } from './helpers'

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
const addressOf = (w: VueWrapper) => (w.find('.xs-address').element as HTMLInputElement).value
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

  it('the flash stays on its sheet: another sheet shown while it lasts has none', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'A1:B3')
    expect(w.findAll('.xs-flash').length).toBe(6)
    await w.findAll('.xs-tab').find((t) => t.text() === 'Two sheet')!.trigger('click')
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    expect(w.findAll('.xs-flash').length).toBe(0)
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
  it('shows the formula with its references coloured; Ctrl+click on one selects the range, on another sheet too', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B1')
    expect(w.findAll('.xs-fx-layer .xs-ref').map((r) => r.text())).toEqual(['A1:A3'])
    // As Excel's formula bar: the text the cell opens with, no value after it.
    expect(barField(w).value).toBe('=SUM(A1:A3)')
    // The formula's range is highlighted on the grid.
    expect((w.find('[data-cell="2:1"]').element as HTMLElement).style.backgroundImage).toContain('linear-gradient')
    await pressBar(w, 7, 7, { ctrlKey: true })
    expect(sel(w).sort()).toEqual(['1:1', '2:1', '3:1'])
    await goTo(w, 'B2')
    await pressBar(w, 8, 8, { ctrlKey: true })
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    expect(active(w)).toBe('2:2')
  })
})

describe('selection', () => {
  it('by column and row headers, with Shift; totals over loaded cells, the count beyond them', async () => {
    const source = memorySource(book())
    delete (source as { stats?: unknown }).stats
    const w = await mountSheet(source, { readOnly: true })
    await w.findAll('.xs-head')[0]!.trigger('click')
    await w.findAll('.xs-head')[1]!.trigger('click', { shiftKey: true })
    expect(w.find('.xs-stats').text()).toBe('Count: 5 · Sum: 16 · Average: 3.2 · Min: 1 · Max: 6')
    await w.findAll('.xs-num')[1]!.trigger('click')
    expect(sel(w)).toEqual(['2:1', '2:2', '2:3'])
    await goTo(w, 'A1:XFD1048576')
    expect(w.find('.xs-stats').text()).toBe('Count: 6 · Sum: 16 · Average: 3.2 · Min: 1 · Max: 6 · Numbers: 5')
  })

  it('the address box writes a range over every row or column as Excel does: B:B, 2:2', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 1_048_576, maxCol: 16_384 }])
    const w = await mountSheet(f.source, { readOnly: true })
    const address = () => (w.find('.xs-address').element as HTMLInputElement).value
    await w.findAll('.xs-head')[1]!.trigger('click')
    expect(address()).toBe('B:B')
    await w.findAll('.xs-num')[1]!.trigger('click')
    expect(address()).toBe('2:2')
  })

  it('whole columns and rows of a small sheet stay whole, as Excel writes them: B:D, 2:5, the sheet, after the grid grows too', async () => {
    const states: { ref: string }[] = []
    const f = fakeSource([{ name: 'S', maxRow: 10, maxCol: 5 }])
    const w = await mountSheet(f.source, { onState: (s: { ref: string } | null) => s && states.push(s) })
    const ref = () => states.at(-1)!.ref
    await w.findAll('.xs-head')[1]!.trigger('click')
    expect([ref(), addressOf(w)]).toEqual(['B:B', 'B:B'])
    await w.findAll('.xs-head')[3]!.trigger('click', { shiftKey: true })
    expect(ref()).toBe('B:D')
    await w.findAll('.xs-num')[1]!.trigger('click')
    await w.findAll('.xs-num')[4]!.trigger('click', { shiftKey: true })
    expect(ref()).toBe('2:5')
    await w.find('.xs-corner').trigger('click')
    expect(ref()).toBe('1:1048576')
    await goTo(w, 'C:C')
    expect(ref()).toBe('C:C')
    expect(sel(w)).toContain('20:3')
    // Data grows past the grid's rows: the column stays whole.
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: 'A400', text: 'x', value: 'x' }] })
    await settle()
    expect(ref()).toBe('C:C')
  })

  it('a selection larger than the loaded windows shows only its size', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 5000, maxCol: 3 }])
    const w = await mountSheet(f.source, { readOnly: true })
    await w.findAll('.xs-head')[0]!.trigger('click')
    expect(w.find('.xs-stats').text()).toBe('Cells: 5,000')
  })

  it('copies the selection as Excel puts plain text on the clipboard', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'A1:C2')
    const data: Record<string, string> = {}
    const e = new Event('copy', { bubbles: true, cancelable: true }) as ClipboardEvent
    Object.defineProperty(e, 'clipboardData', { value: { setData: (t: string, v: string) => (data[t] = v) } })
    w.find('.xs-scroller').element.dispatchEvent(e)
    expect(data['text/plain']).toBe('1\t6\tx\r\n2\t4\t\r\n')
  })

  it('keys move as in a grid: arrows, Ctrl+arrows, Shift extends, Ctrl+A selects all', async () => {
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
  it('Control and an arrow go where the source\'s end says, as Excel\'s Range.End; Shift extends there', async () => {
    const src = memorySource(book())
    const end = vi.fn(async (_sheet: string, cell: string, dir: string) => (dir === 'down' ? 'A40' : cell === 'A40' ? 'F40' : 'A1'))
    const w = await mountSheet({ ...src, end })
    const key = (k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
    await key('ArrowDown', { ctrlKey: true })
    await settle()
    expect(end).toHaveBeenLastCalledWith('One', 'A1', 'down')
    expect(active(w)).toBe('40:1')
    await key('ArrowRight', { ctrlKey: true, shiftKey: true })
    await settle()
    expect(end).toHaveBeenLastCalledWith('One', 'A40', 'right')
    expect(active(w)).toBe('40:1')
    expect(addressOf(w)).toBe('A40:F40')
  })

  it('Enter and Tab move the active cell inside a selection of several cells, which stays, round past its end, as Excel', async () => {
    const w = await mountSheet(memorySource(book()))
    const key = (k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
    await goTo(w, 'B2:C3')
    const all = ['2:2', '2:3', '3:2', '3:3']
    const steps: string[] = []
    for (let i = 0; i < 4; i++) {
      await key('Enter')
      steps.push(active(w)!)
      expect(sel(w).sort()).toEqual(all)
    }
    // down the column, then the top of the next, then round to the first
    expect(steps).toEqual(['3:2', '2:3', '3:3', '2:2'])
    await key('Tab')
    expect(active(w)).toBe('2:3')
    await key('Tab')
    expect(active(w)).toBe('3:2')
    // back with Shift: up the column; left along the row and round to the last cell
    await key('Enter', { shiftKey: true })
    expect(active(w)).toBe('2:2')
    await key('Tab', { shiftKey: true })
    expect(active(w)).toBe('3:3')
    expect(sel(w).sort()).toEqual(all)
    // a value typed and entered goes to the active cell, and Enter goes on inside the selection
    await key('7')
    await w.find('.xs-editor').trigger('keydown', { key: 'Enter' })
    await settle()
    expect(w.find('[data-cell="3:3"]').text()).toBe('7')
    expect(active(w)).toBe('2:2')
    expect(sel(w).sort()).toEqual(all)
    // an arrow drops the selection to the active cell's move
    await key('ArrowRight')
    expect(sel(w)).toEqual(['2:3'])
  })

  it('an arrow without Shift moves from the active cell, not from the selection\'s far corner, as Excel', async () => {
    const w = await mountSheet(memorySource(book()))
    const key = (k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
    await goTo(w, 'B2')
    await key('ArrowDown', { shiftKey: true })
    await key('ArrowRight', { shiftKey: true })
    expect(sel(w).sort()).toEqual(['2:2', '2:3', '3:2', '3:3'])
    await key('ArrowDown')
    expect(active(w)).toBe('3:2')
    expect(sel(w)).toEqual(['3:2'])
    // a column's letter, then Down: the cell under the column's first
    await w.findAll('.xs-head')[2]!.trigger('click')
    await key('ArrowDown')
    expect(active(w)).toBe('2:3')
  })

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

  it("the grid ends at Excel's last row and column", async () => {
    const w = await mountSheet(memorySource(book()))
    const key = (k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
    await goTo(w, 'XFD1048576')
    expect(active(w)).toBe('1048576:16384')
    await key('ArrowDown')
    await key('ArrowRight')
    await key('ArrowDown', { ctrlKey: true })
    await key('ArrowRight', { ctrlKey: true })
    await settle()
    expect(active(w)).toBe('1048576:16384')
  })
})

describe('a phone', () => {
  it('the formula bar colours a formula\'s references, and a press with a key held does not go to them', async () => {
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-scroller') ? 390 : 0
    })
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B1')
    expect(barField(w).value).toBe('=SUM(A1:A3)')
    expect(barField(w).title).toBe('')
    const ref = w.find('.xs-fx-layer .xs-ref')
    expect(ref.text()).toBe('A1:A3')
    expect((ref.element as HTMLElement).style.color).not.toBe('')
    await pressBar(w, 7, 7, { ctrlKey: true })
    expect(active(w)).toBe('1:2')
  })

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
    await pressBar(w, 1)
    expect(document.activeElement).toBe(w.find('.xs-fx-input').element)
    expect((w.find('.xs-fx-input').element as HTMLTextAreaElement).value).toBe('3')
    width.mockRestore()
  })
})
