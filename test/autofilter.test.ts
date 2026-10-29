// Excel's AutoFilter on the sheet: the buttons on its header row over its columns, the display filter over its rows, Data >
// Filter (Ctrl+Shift+L, `toggleAutoFilter`) through the source's `setAutoFilter`, and conditions that follow its columns.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, SheetOp, SheetSource } from '../src/contract/source'
import { mountSheet, settle } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
  localStorage.clear()
})

const OK: EditOutcome = { ok: true, changes: [] }
const numbers = (w: VueWrapper) => w.findAll('.xs-num').map((n) => n.text())
const key = (w: VueWrapper, k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
/** The AutoFilter's buttons by column number. */
const buttons = (w: VueWrapper) => w.findAll('[data-xs-af]').map((b) => Number(b.attributes('data-xs-af')))
async function goTo(w: VueWrapper, address: string): Promise<void> {
  const input = w.find('.xs-address')
  await input.trigger('focus')
  await input.setValue(address)
  await input.trigger('keydown', { key: 'Enter' })
  await settle()
}
async function uncheck(w: VueWrapper, text: string): Promise<void> {
  const label = w.findAll('.xs-value').find((l) => l.find('.xs-value-text').text() === text)!
  await label.find('input').trigger('change')
}
async function filterOut(w: VueWrapper, col: number, text: string): Promise<void> {
  await w.find(`[data-xs-af="${col}"]`).trigger('click')
  await settle()
  await uncheck(w, text)
  await w.find('[data-xs-apply]').trigger('click')
  await settle()
}

/** A title row, the AutoFilter A2:B6 (header Key, Status), two rows below it, and column C beside it with data of its own. */
const book = (autoFilter: string | null = 'A2:B6'): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: 'Title', A2: 'Key', B2: 'Status', A3: 'k3', B3: 'a', A4: 'k4', B4: 'b', A5: 'k5', B5: 'a', A6: 'k6', B6: 'b', A7: 'k7', B7: 'b', A8: 'k8', B8: 'a', C2: 'Note', C3: 'n' },
      ...(autoFilter ? { autoFilter: { range: autoFilter } } : {}),
    },
  ],
})

/** `memorySource` over `b` with `setAutoFilter` (it sets the sheet's AutoFilter in `b`) and, for a test that needs it, `structure`. */
function withAutoFilter(b: MemoryBook, structure?: (op: SheetOp) => void) {
  const setAutoFilter = vi.fn(async (sheet: string, range: string | null): Promise<EditOutcome> => {
    const s = b.sheets.find((x) => x.name === sheet)!
    if (range) s.autoFilter = { range }
    else delete s.autoFilter
    return OK
  })
  const source: SheetSource = { ...memorySource(b), setAutoFilter, ...(structure ? { structure: async (op: SheetOp) => (structure(op), OK) } : {}) }
  return { source, setAutoFilter }
}

describe("the AutoFilter's buttons and the display filter over its rows", () => {
  it('draws a button on each header cell of its columns and no funnels on the column letters', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    expect(buttons(w)).toEqual([1, 2])
    expect(w.find('.xs-head .xs-funnel').exists()).toBe(false)
  })

  it("filters the range's rows below the header; the rows above and below it stay", async () => {
    const source = memorySource(book())
    const visibleRows = vi.spyOn(source, 'visibleRows')
    const w = await mountSheet(source, { readOnly: true })
    await w.find('[data-xs-af="2"]').trigger('click')
    await settle()
    expect(w.find('.xs-popover-title').text()).toBe('Status (B)')
    expect(w.findAll('.xs-value').map((v) => v.text())).toEqual(['Select all2', 'a2', 'b2'])
    await uncheck(w, 'b')
    await w.find('[data-xs-apply]').trigger('click')
    await settle()
    expect(visibleRows).toHaveBeenLastCalledWith('S', { head: 2, last: 6, cols: { 2: { values: ['a'] } } }, expect.any(AbortSignal))
    expect(numbers(w)).toEqual(['1', '2', '3', '5', '7', '8'])
    expect(w.find('[data-xs-filter-status]').text()).toBe('2 of 4 rows shown, filter on B')
    expect(w.find('[data-xs-af="2"]').classes()).toContain('xs-on')
  })

  it("the ribbon's Filter and Sort work on the AutoFilter's columns only", async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    const api = w.vm as unknown as SheetViewApi
    await goTo(w, 'C3')
    expect([api.filterColumn(), api.sortColumn('asc')]).toEqual([false, false])
    await goTo(w, 'A4')
    expect(api.sortColumn('desc')).toBe(true)
    await settle()
    expect(numbers(w)).toEqual(['1', '2', '6', '5', '4', '3', '7', '8'])
    expect(api.filterColumn()).toBe(true)
    await settle()
    expect(w.find('.xs-popover-title').text()).toBe('Key (A)')
  })

  it("Alt+Down on the AutoFilter's header cell opens its list, as in Excel", async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await goTo(w, 'B2')
    await key(w, 'ArrowDown', { altKey: true })
    await settle()
    expect(w.find('.xs-popover-title').text()).toBe('Status (B)')
  })

  it('a source that cannot set an AutoFilter keeps the funnels on the letters of a sheet without one', async () => {
    const w = await mountSheet(memorySource(book(null)), { readOnly: true })
    expect(w.find('.xs-head .xs-funnel').exists()).toBe(true)
    expect(buttons(w)).toEqual([])
  })
})

describe('Data > Filter', () => {
  it('Ctrl+Shift+L puts an AutoFilter over the current region, and again removes it with its conditions', async () => {
    const b = book(null)
    const { source, setAutoFilter } = withAutoFilter(b)
    const states: SheetViewState[] = []
    const w = await mountSheet(source, { onState: (st: SheetViewState | null) => st && states.push(st) })
    expect(w.find('.xs-head .xs-funnel').exists()).toBe(false)
    expect(buttons(w)).toEqual([])
    expect(states.at(-1)).toMatchObject({ canAutoFilter: true })
    expect(states.at(-1)!.autoFilter).toBeUndefined()
    await goTo(w, 'B4')
    await key(w, 'L', { ctrlKey: true, shiftKey: true, code: 'KeyL' })
    await settle()
    // Title in A1 touches the data, so the region is A1:C8, and its column C has data
    expect(setAutoFilter).toHaveBeenLastCalledWith('S', 'A1:C8')
    expect(buttons(w)).toEqual([1, 2, 3])
    expect(states.at(-1)!.autoFilter).toBe('A1:C8')
    await filterOut(w, 2, 'b')
    expect(numbers(w)).toEqual(['1', '2', '3', '5', '8'])
    const api = w.vm as unknown as SheetViewApi
    expect(await api.toggleAutoFilter()).toEqual(OK)
    await settle()
    expect(setAutoFilter).toHaveBeenLastCalledWith('S', null)
    expect(buttons(w)).toEqual([])
    expect(numbers(w).slice(0, 8)).toEqual(['1', '2', '3', '4', '5', '6', '7', '8'])
    expect(w.find('[data-xs-filter-status]').exists()).toBe(false)
    expect(states.at(-1)!.autoFilter).toBeUndefined()
  })

  it('takes a selection of more than one cell down to the end of its data', async () => {
    const { source, setAutoFilter } = withAutoFilter(book(null))
    const w = await mountSheet(source)
    await goTo(w, 'A2:B2')
    expect(await (w.vm as unknown as SheetViewApi).toggleAutoFilter()).toEqual(OK)
    expect(setAutoFilter).toHaveBeenLastCalledWith('S', 'A2:B8')
  })

  it('reads on while the region reaches the edge of the cells read', async () => {
    const cells: Record<string, string> = { A1: 'Key', B1: 'Value' }
    for (let r = 2; r <= 700; r++) Object.assign(cells, { [`A${r}`]: `k${r}`, [`B${r}`]: String(r) })
    const { source, setAutoFilter } = withAutoFilter({ sheets: [{ name: 'S', cells }] })
    const w = await mountSheet(source)
    await goTo(w, 'B3')
    expect(await (w.vm as unknown as SheetViewApi).toggleAutoFilter()).toEqual(OK)
    expect(setAutoFilter).toHaveBeenLastCalledWith('S', 'A1:B700')
  })

  it("on an empty cell away from the data sends nothing and displays Excel's message", async () => {
    const { source, setAutoFilter } = withAutoFilter(book(null))
    const w = await mountSheet(source)
    await goTo(w, 'F12')
    expect(await (w.vm as unknown as SheetViewApi).toggleAutoFilter()).toBeNull()
    await settle()
    expect(setAutoFilter).not.toHaveBeenCalled()
    expect(w.find('.xs-message').text()).toBe('This can’t be applied to the selected range. Select a single cell in a range and try again.')
  })

  it('a read-only sheet cannot toggle it and keeps the funnels on the letters of a sheet without one', async () => {
    const { source, setAutoFilter } = withAutoFilter(book(null))
    const states: SheetViewState[] = []
    const w = await mountSheet(source, { readOnly: true, onState: (st: SheetViewState | null) => st && states.push(st) })
    expect(w.find('.xs-head .xs-funnel').exists()).toBe(true)
    expect(states.at(-1)!.canAutoFilter).toBeUndefined()
    await key(w, 'L', { ctrlKey: true, shiftKey: true, code: 'KeyL' })
    expect(await (w.vm as unknown as SheetViewApi).toggleAutoFilter()).toBeNull()
    expect(setAutoFilter).not.toHaveBeenCalled()
  })
})

describe('the conditions follow the AutoFilter', () => {
  it('columns inserted before it move its conditions with their columns', async () => {
    const b = book()
    const { source } = withAutoFilter(b, (op) => {
      if (op.op === 'insertColumns') b.sheets[0]!.autoFilter = { range: 'B2:C6' }
    })
    const visibleRows = vi.spyOn(source, 'visibleRows')
    const w = await mountSheet(source)
    await filterOut(w, 2, 'b')
    await (w.vm as unknown as SheetViewApi).structure({ op: 'insertColumns', sheet: 'S', column: 'A' })
    await settle()
    expect(buttons(w)).toEqual([2, 3])
    expect(visibleRows).toHaveBeenLastCalledWith('S', { head: 2, last: 6, cols: { 3: { values: ['a'] } } }, expect.any(AbortSignal))
  })

  it('an AutoFilter removed elsewhere (its header row deleted) takes its conditions along', async () => {
    const b = book()
    const source = memorySource(b)
    const w = await mountSheet(source, { readOnly: true })
    await filterOut(w, 2, 'b')
    expect(w.find('[data-xs-filter-status]').exists()).toBe(true)
    delete b.sheets[0]!.autoFilter
    source.emit({ kind: 'reload' })
    await settle()
    expect(buttons(w)).toEqual([])
    expect(w.find('[data-xs-filter-status]').exists()).toBe(false)
  })
  it('a range grown by data typed under it (`sheets`) applies at Reapply, as Excel filters the new rows then', async () => {
    const b = book()
    const source = memorySource(b)
    const visibleRows = vi.spyOn(source, 'visibleRows')
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source, { readOnly: true })
    await filterOut(w, 2, 'b')
    expect(visibleRows).toHaveBeenLastCalledWith('S', { head: 2, last: 6, cols: { 2: { values: ['a'] } } }, expect.any(AbortSignal))
    const lists = visibleRows.mock.calls.length
    const windows = window.mock.calls.length
    b.sheets[0]!.autoFilter = { range: 'A2:B8' }
    source.emit({ kind: 'sheets' })
    await settle()
    // the sheet knows the range, and the rows stay as filtered: no list and no window read again
    expect((w.emitted('state')!.at(-1)![0] as SheetViewState).autoFilter).toBe('A2:B8')
    expect(visibleRows.mock.calls.length).toBe(lists)
    expect(window.mock.calls.length).toBe(windows)
    expect(numbers(w).slice(0, 6)).toEqual(['1', '2', '3', '5', '7', '8'])
    await key(w, 'l', { ctrlKey: true, altKey: true })
    await settle()
    expect(visibleRows).toHaveBeenLastCalledWith('S', { head: 2, last: 8, cols: { 2: { values: ['a'] } } }, expect.any(AbortSignal))
    expect(numbers(w).slice(0, 5)).toEqual(['1', '2', '3', '5', '8'])
  })
})
