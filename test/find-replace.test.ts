// Excel's Find and Replace in the sheet's bar: the options go into the source's request, Find Next continues from the
// active cell, a selection of several cells limits the search, and Replace and Replace All go through the source's
// `replace` as an edit does, with the grid showing what they wrote.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { FindRequest, SheetSource } from '../src/contract/source'
import { active, cellText, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

/** Fake timers for the pause before the source is queried; setImmediate stays real, since flushPromises needs it. */
const fakeTimers = () => vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval', 'Date'] })
const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'apple', B3: 'pineapple', C10: 'Apple pie', D20: 'pear' } }] })
const lastRequest = (spy: { mock: { calls: unknown[][] } }) => spy.mock.calls.at(-1)![0] as FindRequest
const status = (w: VueWrapper) => w.find('.xs-find-status').text()

/** Opens the bar with Ctrl+F, types a query and lets the pause pass. */
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
  it("queries with Excel's options: in formulas by default, Match case, Match entire cell contents, values, by columns", async () => {
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

  it('continues from the active cell in the chosen order, as Find Next does, and wraps around', async () => {
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

  it("from a match goes to the next one in the source's order, as Find returns them by blocks on a filtered sheet", async () => {
    fakeTimers()
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'x', C1: 'x', A3: 'x', C3: 'x' } }] })
    // Excel's Find on a filtered sheet with row 2 and column B hidden returns the blocks' cells first: A1, C1, A3, C3 by columns too
    vi.spyOn(source, 'find').mockResolvedValue({ hits: [{ row: 1, col: 1 }, { row: 1, col: 3 }, { row: 3, col: 1 }, { row: 3, col: 3 }], total: 4 })
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

  it("the Replace row's toggle recounts but leaves the active cell, as Excel's Find and Replace tabs do", async () => {
    fakeTimers()
    const source = memorySource(book())
    const find = vi.spyOn(source, 'find')
    const w = await mountSheet(source)
    await search(w, 'apple')
    await w.find('[data-xs-find-options]').trigger('click')
    await w.find('[data-xs-find-look]').setValue('values')
    await pause()
    await w.find('.xs-find-input').trigger('keydown', { key: 'Enter' })
    const at = active(w)
    const calls = find.mock.calls.length
    await w.find('[data-xs-find-replace]').trigger('click')
    await pause()
    // Replace searches in formulas: the bar requests again, and the active cell stays
    expect(find.mock.calls.length).toBeGreaterThan(calls)
    expect(lastRequest(find).lookIn).toBe('formulas')
    expect(active(w)).toBe(at)
  })

  it("goes beyond the matches the source returned, as Excel's Find Next reaches every one: the next ones from the active cell", async () => {
    fakeTimers()
    // 2,500 matches in column A; memorySource returns 1,000 at a time
    const cells: Record<string, string> = {}
    for (let r = 1; r <= 2500; r++) cells[`A${r}`] = 'x'
    const source = memorySource({ sheets: [{ name: 'S', cells }] })
    const find = vi.spyOn(source, 'find')
    const w = await mountSheet(source)
    const next = () => w.find('.xs-find-input').trigger('keydown', { key: 'Enter' })
    await goTo(w, 'B700')
    await search(w, 'x')
    // a new query from B700: the first match after it, A701, requested from there as Excel's Find After
    expect(active(w)).toBe('701:1')
    expect(lastRequest(find).after).toBe('B700')
    await goTo(w, 'A2500')
    await next()
    await settle()
    // wrapping around at the end to A1, requested from A2500
    expect(active(w)).toBe('1:1')
    // the last of the returned matches, A1000, then onward
    await goTo(w, 'A1000')
    await next()
    await settle()
    // beyond the 1,000 returned matches: the next ones are requested from the active cell
    expect(active(w)).toBe('1001:1')
    expect(lastRequest(find).after).toBe('A1000')
    await next()
    expect(active(w)).toBe('1002:1')
  })

  it('searches within a selection of more than one cell, as Excel does, until the user removes it', async () => {
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

  it('Replace All goes through the source in formulas, the grid shows what it wrote, and the bar reports how many', async () => {
    fakeTimers()
    const source = memorySource(book())
    const replace = vi.spyOn(source, 'replace')
    const w = await mountSheet(source)
    await search(w, 'apple')
    await w.find('[data-xs-find-options]').trigger('click')
    await w.find('[data-xs-find-look]').setValue('values')
    // Excel's Replace searches in formulas only, whatever Find was set to
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

  it('Replace replaces the active cell when it is a match and goes to the next; elsewhere it only finds', async () => {
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

  it("Replace All reports the cells it left unchanged, with the source's reason", async () => {
    fakeTimers()
    const source = memorySource(book())
    vi.spyOn(source, 'replace').mockResolvedValue({ ok: true, changes: [], replaced: 2, refused: [{ sheet: 'S', cell: 'E1', message: 'The formula does not read' }] })
    const w = await mountSheet(source)
    await key(w, 'h', { ctrlKey: true })
    await settle()
    await w.find('.xs-find-input').setValue('apple')
    await pause()
    await w.find('[data-xs-replace-all]').trigger('click')
    await settle()
    expect(status(w)).toBe('Replaced: 2')
    expect(w.find('.xs-message').text()).toBe('1 cell was left unchanged: The formula does not read')
  })

  it("a rejection appears in the sheet's message line and changes nothing", async () => {
    fakeTimers()
    const source = memorySource(book())
    vi.spyOn(source, 'replace').mockResolvedValue({ ok: false, reason: 'refused', message: 'The sheet is protected', replaced: 0 })
    const w = await mountSheet(source)
    await openReplace(w, 'apple', 'kiwi')
    await w.find('[data-xs-replace-all]').trigger('click')
    await settle()
    expect(w.find('.xs-message').text()).toBe('The sheet is protected')
    expect(cellText(w, 1, 1)).toBe('apple')
    // Find Next continued from A1, the active cell, to B3
    expect(status(w)).toBe('2 of 3')
  })

  it("there is no replacing on a read-only sheet or without the source's `replace`", async () => {
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
