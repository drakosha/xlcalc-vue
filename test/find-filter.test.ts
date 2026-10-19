// 0.2.0 in the sheet: search with jumps through the matches, the funnel (display rows, the header, storage per sheet), totals
// from the source's stats, and the wait for a recalculated workbook with its retries.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { SheetPending, SheetSource, SheetWindow, WindowRequest } from '../src/contract/source'
import { gridWindow, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
  localStorage.clear()
})

/** Fake timers for the pauses of search, totals and retries; setImmediate stays real, since flushPromises needs it. */
const fakeTimers = () => vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })
const active = (w: VueWrapper) => w.find('.xs-cell.xs-active').attributes('data-cell')
const numbers = (w: VueWrapper) => w.findAll('.xs-num').map((n) => n.text())
const key = (w: VueWrapper, k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
async function goTo(w: VueWrapper, address: string): Promise<void> {
  const input = w.find('.xs-address')
  await input.trigger('focus')
  await input.setValue(address)
  await input.trigger('keydown', { key: 'Enter' })
  await settle()
}

describe('search', () => {
  const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'apple', B3: 'pineapple', C10: 'Apple pie', D20: 'pear' } }] })

  it('asks the source, jumps to the first hit, steps with Enter and Shift+Enter, tints hits; Esc closes', async () => {
    fakeTimers()
    const source = memorySource(book())
    const find = vi.spyOn(source, 'find')
    const w = await mountSheet(source, { readOnly: true })
    await goTo(w, 'D20')
    await key(w, 'f', { ctrlKey: true })
    await settle()
    expect(w.find('.xs-find').exists()).toBe(true)
    await w.find('.xs-find-input').setValue('APPLE')
    await vi.advanceTimersByTimeAsync(300)
    await settle()
    expect(find).toHaveBeenLastCalledWith({ sheet: 'S', query: 'APPLE', lookIn: 'formulas' }, expect.any(AbortSignal))
    expect(active(w)).toBe('1:1')
    expect(w.find('.xs-find-status').text()).toBe('1 of 3')
    expect(w.findAll('.xs-cell.xs-hit').map((c) => c.attributes('data-cell'))).toEqual(['1:1', '3:2', '10:3'])
    await w.find('.xs-find-input').trigger('keydown', { key: 'Enter' })
    expect(active(w)).toBe('3:2')
    expect(w.find('.xs-cell.xs-hit-now').attributes('data-cell')).toBe('3:2')
    await w.find('.xs-find-input').trigger('keydown', { key: 'Enter', shiftKey: true })
    expect(active(w)).toBe('1:1')
    await w.find('.xs-find-input').trigger('keydown', { key: 'Escape' })
    expect(w.find('.xs-find').exists()).toBe(false)
    expect(w.findAll('.xs-cell.xs-hit')).toHaveLength(0)
  })

  it('there is no search without `find`', async () => {
    const source = memorySource(book())
    delete (source as Partial<SheetSource>).find
    const w = await mountSheet(source, { readOnly: true })
    expect(w.find('[data-xs-find]').exists()).toBe(false)
    await key(w, 'f', { ctrlKey: true })
    expect(w.find('.xs-find').exists()).toBe(false)
  })
})

describe('the display filter', () => {
  const book = (): MemoryBook => ({
    sheets: [
      { name: 'S', cells: { A1: 'Name', B1: 'Status', A2: 'n2', B2: 'a', A3: 'n3', B3: 'b', A4: 'n4', B4: 'a', A5: 'n5', B5: 'b', A6: 'n6', B6: 'a' } },
      { name: 'T', cells: { A1: 'other' } },
    ],
  })
  async function funnel(w: VueWrapper, col: number): Promise<void> {
    await w.findAll('.xs-head')[col - 1]!.find('.xs-funnel').trigger('click')
    await settle()
  }
  async function uncheck(w: VueWrapper, text: string): Promise<void> {
    const label = w.findAll('.xs-value').find((l) => l.find('.xs-value-text').text() === text)!
    await label.find('input').trigger('change')
  }

  it('applies a column condition: the header stays, rows go by display index, numbers stay the sheet’s', async () => {
    const source = memorySource(book())
    const visibleRows = vi.spyOn(source, 'visibleRows')
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source, { readOnly: true })
    await funnel(w, 2)
    expect(w.find('.xs-popover-title').text()).toBe('Status (B)')
    expect(w.findAll('.xs-value').map((v) => v.text())).toEqual(['Select all2', 'a3', 'b2'])
    await uncheck(w, 'b')
    await w.find('[data-xs-apply]').trigger('click')
    await settle()
    expect(visibleRows).toHaveBeenLastCalledWith('S', { head: 1, cols: { 2: { values: ['a'] } } }, expect.any(AbortSignal))
    const asked = window.mock.calls.at(-1)![0] as WindowRequest
    expect(asked).toMatchObject({ sheet: 'S', row: 1, filter: { head: 1, cols: { 2: { values: ['a'] } } } })
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    expect(w.find('[data-cell="4:1"]').text()).toBe('n4')
    expect(w.find('[data-cell="3:1"]').exists()).toBe(false)
    expect(w.find('[data-xs-filter-status]').text()).toBe('3 of 5 rows shown, filter on B')
    expect(w.findAll('.xs-head')[1]!.classes()).toContain('xs-filtered')
    // Keys move along the visible rows.
    await goTo(w, 'A2')
    await key(w, 'ArrowDown')
    expect(active(w)).toBe('4:1')
  })

  it('offers no Freeze Panes under the display filter, whose rows on screen are not the sheet\'s', async () => {
    const source = memorySource(book())
    const freeze = vi.fn(async () => ({ ok: true as const, changes: [] }))
    const w = await mountSheet({ ...source, freeze })
    const menuIds = async () => {
      await w.find('[data-cell="2:1"]').trigger('contextmenu', { button: 2, clientX: 40, clientY: 40 })
      await settle()
      const ids = [...document.querySelectorAll('[data-xs-menu]')].map((e) => e.getAttribute('data-xs-menu'))
      await key(w, 'Escape')
      await settle()
      return ids
    }
    expect(await menuIds()).toContain('freeze')
    await funnel(w, 1)
    await w.find('.xs-popover button[title="Sort descending"]').trigger('click')
    await settle()
    expect(await menuIds()).not.toContain('freeze')
    expect(await (w.vm as unknown as { freezePanes: () => Promise<unknown> }).freezePanes()).toBeNull()
    expect(freeze).not.toHaveBeenCalled()
  })

  it('sorts from the funnel; clearing brings the sheet’s order back', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await funnel(w, 1)
    await w.find('.xs-popover button[title="Sort descending"]').trigger('click')
    await settle()
    expect(numbers(w)).toEqual(['1', '6', '5', '4', '3', '2'])
    expect(w.findAll('.xs-head')[0]!.find('.xs-sortmark').text()).toBe('↓')
    await w.find('[data-xs-filter-clear]').trigger('click')
    await settle()
    expect(numbers(w)).toEqual(['1', '2', '3', '4', '5', '6'])
    expect(w.find('[data-xs-filter-status]').exists()).toBe(false)
  })

  it('remembers the conditions by sheet under storageKey', async () => {
    const source = memorySource(book())
    const w = await mountSheet(source, { readOnly: true, storageKey: 'k' })
    await funnel(w, 2)
    await uncheck(w, 'a')
    await w.find('[data-xs-apply]').trigger('click')
    await settle()
    expect(JSON.parse(localStorage.getItem('k:filter:S')!)).toEqual({ cols: { 2: { values: ['b'] } }, sort: null })
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    expect(w.find('[data-xs-filter-status]').exists()).toBe(false)
    await w.findAll('.xs-tab')[0]!.trigger('click')
    await settle()
    expect(numbers(w)).toEqual(['1', '3', '5'])
    w.unmount()
    const again = await mountSheet(memorySource(book()), { readOnly: true, storageKey: 'k' })
    expect(numbers(again)).toEqual(['1', '3', '5'])
  })

  it('a jump to a row the filter hides clears it, with a message', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await funnel(w, 2)
    await uncheck(w, 'b')
    await w.find('[data-xs-apply]').trigger('click')
    await settle()
    await goTo(w, 'B3')
    expect(w.find('[data-xs-filter-status]').exists()).toBe(false)
    expect(numbers(w)).toEqual(['1', '2', '3', '4', '5', '6'])
    expect(active(w)).toBe('3:2')
    expect(w.find('.xs-message').text()).toBe('The filter is cleared: it hid row 3')
  })

  it('a jump to a shown row past the count of shown rows lands on that row', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await funnel(w, 2)
    await uncheck(w, 'b')
    await w.find('[data-xs-apply]').trigger('click')
    await settle()
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    await goTo(w, 'A6:B6')
    expect([active(w), (w.find('.xs-address').element as HTMLInputElement).value]).toEqual(['6:1', 'A6:B6'])
    expect(w.find('[data-xs-filter-status]').exists()).toBe(true)
  })

  it('there is no funnel without `visibleRows` and `columnValues`', async () => {
    const source = memorySource(book())
    delete (source as Partial<SheetSource>).visibleRows
    const w = await mountSheet(source, { readOnly: true })
    expect(w.find('.xs-funnel').exists()).toBe(false)
  })
})

describe('totals from the source', () => {
  it('asks `stats` a moment after the selection settles, with the range and the filter', async () => {
    fakeTimers()
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, B2: 2 } }] })
    const stats = vi.spyOn(source, 'stats').mockResolvedValue({ cells: 99, numbers: 1, sum: 5, min: 5, max: 5 })
    const w = await mountSheet(source, { readOnly: true })
    await goTo(w, 'A1:B2')
    // Before the response, the loaded windows are counted.
    expect(w.find('.xs-stats').text()).toBe('Count: 2 · Sum: 3 · Average: 1.5 · Min: 1 · Max: 2')
    await vi.advanceTimersByTimeAsync(200)
    await settle()
    expect(stats).toHaveBeenCalledTimes(1)
    expect(stats).toHaveBeenCalledWith('S', 'A1:B2', undefined, expect.any(AbortSignal))
    expect(w.find('.xs-stats').text()).toBe('Count: 99 · Sum: 5 · Average: 5 · Numbers: 1')
  })

  it('without `stats` the loaded windows count', async () => {
    fakeTimers()
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, B2: 2 } }] })
    delete (source as Partial<SheetSource>).stats
    const w = await mountSheet(source, { readOnly: true })
    await goTo(w, 'A1:B2')
    await vi.advanceTimersByTimeAsync(200)
    await settle()
    expect(w.find('.xs-stats').text()).toBe('Count: 2 · Sum: 3 · Average: 1.5 · Min: 1 · Max: 2')
  })
})

describe('a book being recomputed', () => {
  /** Sheets whose windows return `pending` while `busy(sheet)` says so. */
  function waiting(busy: (sheet: string) => boolean) {
    const calls: WindowRequest[] = []
    const source: SheetSource = {
      sheets: async () => [
        { name: 'One', maxRow: 5, maxCol: 3 },
        { name: 'Two', maxRow: 5, maxCol: 3 },
      ],
      window: async (req): Promise<SheetWindow | SheetPending> => {
        calls.push(req)
        if (busy(req.sheet)) return { pending: { elapsedMs: 3000 } }
        return gridWindow({ ...req, rows: Math.min(req.rows, 5), cols: Math.min(req.cols, 3) }, (r, c) => `${req.sheet}${r}${c}`)
      },
    }
    return { source, calls }
  }

  it('waits with a stopwatch, asks again every second and draws the window once ready', async () => {
    fakeTimers()
    let polls = 0
    const { source, calls } = waiting(() => ++polls <= 2)
    const w = await mountSheet(source, { readOnly: true })
    expect(w.find('[data-xs-wait]').exists()).toBe(true)
    expect(w.find('.xs-wait-seconds').text()).toBe('3 s')
    await vi.advanceTimersByTimeAsync(1000)
    await settle()
    expect(calls).toHaveLength(2)
    expect(w.find('.xs-wait-seconds').text()).toBe('4 s')
    await vi.advanceTimersByTimeAsync(1000)
    await settle()
    expect(calls).toHaveLength(3)
    expect(w.find('[data-xs-wait]').exists()).toBe(false)
    expect(w.find('[data-cell="2:3"]').text()).toBe('One23')
  })

  it('stops asking for the previous sheet when another is picked', async () => {
    fakeTimers()
    const { source, calls } = waiting((sheet) => sheet === 'One')
    const w = await mountSheet(source, { readOnly: true })
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    expect(w.find('[data-xs-wait]').exists()).toBe(false)
    const before = calls.filter((c) => c.sheet === 'One').length
    await vi.advanceTimersByTimeAsync(5000)
    await settle()
    expect(calls.filter((c) => c.sheet === 'One')).toHaveLength(before)
    expect(w.find('[data-cell="1:1"]').text()).toBe('Two11')
  })

  it('gives up after the limit of polls', async () => {
    fakeTimers()
    const { source } = waiting(() => true)
    const w = await mountSheet(source, { readOnly: true, messages: { waitTimeout: 'Too long' } })
    for (let i = 0; i < 205; i++) await vi.advanceTimersByTimeAsync(1000)
    await settle()
    expect(w.find('.xs-fallback').text()).toContain('Too long')
  })
})
