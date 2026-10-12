// Excel's Find and Replace in the sheet's bar: the options go to the source's request, Find Next goes on from the active
// cell, a selection of several cells is where the search looks, Replace and Replace All go through the source's
// `replace` as an edit goes and the grid shows what they wrote.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { FindRequest, SheetSource } from '../src/source'
import { active, cellText, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

/** Fake timers for the pause before the source is asked; setImmediate stays real, flushPromises needs it. */
const fakeTimers = () => vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })
const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'apple', B3: 'pineapple', C10: 'Apple pie', D20: 'pear' } }] })
const lastRequest = (spy: { mock: { calls: unknown[][] } }) => spy.mock.calls.at(-1)![0] as FindRequest
const status = (w: VueWrapper) => w.find('.xs-find-status').text()

/** Open the bar with Ctrl+F, type a query and let the pause pass. */
async function search(w: VueWrapper, query: string): Promise<void> {
  await key(w, 'f', { ctrlKey: true })
  await settle()
  await w.find('.xs-find-input').setValue(query)
  await vi.advanceTimersByTimeAsync(300)
  await settle()
}
async function pause(): Promise<void> {
  await vi.advanceTimersByTimeAsync(300)
  await settle()
}

describe('Find', () => {
  it('asks with Excel’s options: in formulas by default, match case, the entire cell, values, by columns', async () => {
    fakeTimers()
    const source = memorySource(book())
    const find = vi.spyOn(source, 'find')
    const w = await mountSheet(source)
    await search(w, 'apple')
    expect(lastRequest(find)).toEqual({ sheet: 'S', query: 'apple', lookIn: 'formulas' })
    await w.find('[data-xs-find-options]').trigger('click')
    await w.find('[data-xs-find-case]').setValue(true)
    await w.find('[data-xs-find-whole]').setValue(true)
    await w.find('[data-xs-find-look]').setValue('values')
    await w.find('[data-xs-find-order]').setValue('columns')
    await pause()
    expect(lastRequest(find)).toEqual({ sheet: 'S', query: 'apple', matchCase: true, wholeCell: true, lookIn: 'values', byColumns: true })
    expect(w.findAll('.xs-cell.xs-hit').map((c) => c.attributes('data-cell'))).toEqual(['1:1'])
  })

  it('goes on from the active cell in the order asked, as Find Next does, and wraps round', async () => {
    fakeTimers()
    const w = await mountSheet(memorySource(book()))
    await goTo(w, 'A5')
    await search(w, 'apple')
    expect(active(w)).toBe('10:3')
    expect(status(w)).toBe('3 of 3')
    await w.find('.xs-find-input').trigger('keydown', { key: 'Enter' })
    expect(active(w)).toBe('1:1')
    await goTo(w, 'A5')
    await w.find('[data-xs-find-options]').trigger('click')
    await w.find('[data-xs-find-order]').setValue('columns')
    await pause()
    // by columns B3 comes after A5
    expect(active(w)).toBe('3:2')
    await w.find('.xs-find-input').trigger('keydown', { key: 'Enter', shiftKey: true })
    expect(active(w)).toBe('1:1')
  })

  it('from a hit goes to the next one in the source\'s order, as Find gives it by blocks on a filtered sheet', async () => {
    fakeTimers()
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'x', C1: 'x', A3: 'x', C3: 'x' } }] })
    // Excel's Find on a filtered sheet with row 2 and column B hidden gives the blocks' cells first: A1, C1, A3, C3 by columns too
    vi.spyOn(source, 'find').mockResolvedValue({ hits: [{ row: 1, col: 1, text: 'x' }, { row: 1, col: 3, text: 'x' }, { row: 3, col: 1, text: 'x' }, { row: 3, col: 3, text: 'x' }], total: 4 })
    const w = await mountSheet(source)
    await search(w, 'x')
    expect(active(w)).toBe('1:3')
    await w.find('[data-xs-find-options]').trigger('click')
    await w.find('[data-xs-find-order]').setValue('columns')
    await pause()
    await goTo(w, 'C1')
    const input = w.find('.xs-find-input')
    await input.trigger('keydown', { key: 'Enter' })
    expect(active(w)).toBe('3:1')
    await input.trigger('keydown', { key: 'Enter' })
    expect(active(w)).toBe('3:3')
    await input.trigger('keydown', { key: 'Enter' })
    expect(active(w)).toBe('1:1')
    await input.trigger('keydown', { key: 'Enter', shiftKey: true })
    expect(active(w)).toBe('3:3')
  })

  it('looks within a selection of more than one cell, as Excel does, until the person drops it', async () => {
    fakeTimers()
    const source = memorySource(book())
    const find = vi.spyOn(source, 'find')
    const w = await mountSheet(source)
    await goTo(w, 'A1:B5')
    await search(w, 'apple')
    expect(lastRequest(find).range).toBe('A1:B5')
    expect(w.findAll('.xs-cell.xs-hit').map((c) => c.attributes('data-cell'))).toEqual(['1:1', '3:2'])
    expect(w.find('[data-xs-find-scope]').text()).toContain('A1:B5')
    await w.find('[data-xs-find-scope] button').trigger('click')
    await pause()
    expect(lastRequest(find).range).toBeUndefined()
    expect(w.find('[data-xs-find-scope]').exists()).toBe(false)
  })
})

describe('Replace', () => {
  async function openReplace(w: VueWrapper, query: string, by: string): Promise<void> {
    await search(w, query)
    await w.find('[data-xs-find-replace]').trigger('click')
    await w.find('[data-xs-replace-input]').setValue(by)
  }

  it('Replace All goes through the source in formulas, the grid shows what it wrote, the bar says how many', async () => {
    fakeTimers()
    const source = memorySource(book())
    const replace = vi.spyOn(source, 'replace')
    const w = await mountSheet(source)
    await search(w, 'apple')
    await w.find('[data-xs-find-options]').trigger('click')
    await w.find('[data-xs-find-look]').setValue('values')
    // Excel's Replace looks in formulas only, whatever Find was set to
    await w.find('[data-xs-find-replace]').trigger('click')
    await w.find('[data-xs-replace-input]').setValue('kiwi')
    expect((w.find('[data-xs-find-look]').element as HTMLSelectElement).value).toBe('formulas')
    expect((w.find('[data-xs-find-look]').element as HTMLSelectElement).disabled).toBe(true)
    await pause()
    await w.find('[data-xs-replace-all]').trigger('click')
    await settle()
    expect(replace).toHaveBeenCalledWith({ sheet: 'S', query: 'apple', lookIn: 'formulas' }, 'kiwi')
    expect([cellText(w, 1, 1), cellText(w, 3, 2), cellText(w, 10, 3)]).toEqual(['kiwi', 'pinekiwi', 'kiwi pie'])
    expect(status(w)).toBe('Replaced: 3')
    expect(w.findAll('.xs-cell.xs-hit')).toHaveLength(0)
  })

  it('Replace replaces the active cell when it is a hit and goes to the next; elsewhere it only finds', async () => {
    fakeTimers()
    const source = memorySource(book())
    const replace = vi.spyOn(source, 'replace')
    const w = await mountSheet(source)
    await goTo(w, 'D20')
    await openReplace(w, 'apple', 'kiwi')
    expect(active(w)).toBe('1:1')
    await goTo(w, 'D20')
    await w.find('[data-xs-replace-one]').trigger('click')
    await settle()
    expect(replace).not.toHaveBeenCalled()
    expect(active(w)).toBe('1:1')
    await w.find('[data-xs-replace-one]').trigger('click')
    await settle()
    expect(replace).toHaveBeenCalledWith({ sheet: 'S', query: 'apple', lookIn: 'formulas', range: 'A1' }, 'kiwi')
    expect(cellText(w, 1, 1)).toBe('kiwi')
    expect(cellText(w, 3, 2)).toBe('pineapple')
    expect(active(w)).toBe('3:2')
    expect(status(w)).toBe('1 of 2')
  })

  it('a refusal shows in the sheet’s message line and changes nothing', async () => {
    fakeTimers()
    const source = memorySource(book())
    vi.spyOn(source, 'replace').mockResolvedValue({ ok: false, reason: 'refused', message: 'The sheet is protected', replaced: 0 })
    const w = await mountSheet(source)
    await openReplace(w, 'apple', 'kiwi')
    await w.find('[data-xs-replace-all]').trigger('click')
    await settle()
    expect(w.find('.xs-message').text()).toBe('The sheet is protected')
    expect(cellText(w, 1, 1)).toBe('apple')
    // Find Next went on from A1, the active cell, to B3
    expect(status(w)).toBe('2 of 3')
  })

  it('there is no replacing on a read-only sheet or without the source’s `replace`', async () => {
    fakeTimers()
    const ro = await mountSheet(memorySource(book()), { readOnly: true })
    await search(ro, 'apple')
    expect(ro.find('[data-xs-find-replace]').exists()).toBe(false)
    const source = memorySource(book())
    delete (source as Partial<SheetSource>).replace
    const w = await mountSheet(source)
    await search(w, 'apple')
    expect(w.find('[data-xs-find-replace]').exists()).toBe(false)
  })
})
