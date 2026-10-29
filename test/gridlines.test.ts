// Excel's View > Show: a sheet hides its gridlines or its row and column headings as its view says, and the sheet's
// `showGridlines` and `showHeadings` change them through the source's `show`, not an undo step, or in the browser without it.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, SheetInfo, SheetShow } from '../src/contract/source'
import { fakeSource, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const states = (w: VueWrapper) => (w.emitted('state') ?? []).map((e) => e[0] as SheetViewState | null)
const lastState = (w: VueWrapper) => states(w).at(-1)

/** A workbook whose sheets the test holds; with `withShow`, the source keeps View > Show as `./local` does. */
function book(sheets: SheetInfo[], withShow = true, refuse = false) {
  const f = fakeSource(sheets)
  f.source.sheets = async () => sheets.map((x) => ({ ...x }))
  const show = vi.fn(async (sheet: string, what: SheetShow): Promise<EditOutcome> => {
    if (refuse) return { ok: false, reason: 'refused', message: 'Not now' }
    const at = sheets.findIndex((x) => x.name === sheet)
    sheets[at] = { ...sheets[at]!, ...what }
    return { ok: true, changes: [] }
  })
  if (withShow) f.source.show = show
  return { ...f, show }
}

describe('gridlines and headings', () => {
  it("draws the sheet without gridlines and without row numbers and column letters as its view says, the frozen rows sticking at the grid's top", async () => {
    const src = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, B2: 2 }, gridlines: false, headings: false, freeze: { rows: 1, cols: 0 } }] })
    const w = await mountSheet(src)
    expect(w.find('.xs-table').classes()).toContain('xs-nogrid')
    expect(w.find('.xs-head').exists()).toBe(false)
    expect(w.find('.xs-num').exists()).toBe(false)
    expect(w.find('.xs-corner').exists()).toBe(false)
    expect((w.find('[data-cell="1:1"]').element as HTMLElement).style.top).toBe('0px')
    expect(lastState(w)).toMatchObject({ gridlines: false, headings: false })
  })

  it('shows both by default', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 }, freeze: { rows: 1, cols: 0 } }] }))
    expect(w.find('.xs-table').classes()).not.toContain('xs-nogrid')
    expect(w.find('.xs-head').exists()).toBe(true)
    expect(w.find('.xs-num').exists()).toBe(true)
    expect((w.find('[data-cell="1:1"]').element as HTMLElement).style.top).toBe('24px')
    expect(lastState(w)).not.toHaveProperty('gridlines')
    expect(lastState(w)).not.toHaveProperty('headings')
  })

  it("keeps list mode's rows without numbers under its column titles", async () => {
    const src = memorySource({ sheets: [{ name: 'L', cells: { A1: 'x' }, rowNumbers: false, columns: [{ title: 'Name' }] }] })
    const w = await mountSheet(src)
    expect(w.find('.xs-num').exists()).toBe(false)
    expect(w.find('.xs-head').text()).toBe('Name')
  })

  it('showGridlines and showHeadings toggle them through the source’s show, one setting a call and no undo step', async () => {
    const b = book([{ name: 'S', maxRow: 3, maxCol: 3 }])
    const w = await mountSheet(b.source)
    expect(api(w).showGridlines()).toBe(false)
    await settle()
    expect(b.show.mock.calls).toEqual([['S', { gridlines: false }]])
    expect(w.find('.xs-table').classes()).toContain('xs-nogrid')
    expect(lastState(w)).toMatchObject({ gridlines: false })
    expect(api(w).showHeadings(false)).toBe(false)
    await settle()
    expect(b.show.mock.calls[1]).toEqual(['S', { headings: false }])
    expect(w.find('.xs-head').exists()).toBe(false)
    // setting what is already set sends nothing
    expect(api(w).showHeadings(false)).toBe(false)
    expect(api(w).showGridlines(true)).toBe(true)
    await settle()
    expect(b.show).toHaveBeenCalledTimes(3)
    expect(w.find('.xs-table').classes()).not.toContain('xs-nogrid')
    // the source keeps them: read again, the sheet stays as set
    b.emit({ kind: 'reload' })
    await settle()
    expect(w.find('.xs-head').exists()).toBe(false)
    expect(w.find('.xs-table').classes()).not.toContain('xs-nogrid')
  })

  it("puts a setting the source refused back and displays the source's message", async () => {
    const b = book([{ name: 'S', maxRow: 3, maxCol: 3 }], true, true)
    const w = await mountSheet(b.source)
    api(w).showGridlines(false)
    await settle()
    expect(w.find('.xs-table').classes()).not.toContain('xs-nogrid')
    expect(w.text()).toContain('Not now')
  })

  it('without the source’s show, or on a read-only sheet, keeps them in the browser for each sheet, past a reload', async () => {
    for (const [withShow, readOnly] of [[false, false], [true, true]] as const) {
      const b = book([{ name: 'S', maxRow: 3, maxCol: 3 }, { name: 'T', maxRow: 3, maxCol: 3 }], withShow)
      const w = await mountSheet(b.source, { readOnly })
      api(w).showGridlines(false)
      await settle()
      b.emit({ kind: 'reload' })
      await settle()
      expect(w.find('.xs-table').classes()).toContain('xs-nogrid')
      await w.findAll('.xs-tab')[1]!.trigger('click')
      await settle()
      expect(w.find('.xs-table').classes()).not.toContain('xs-nogrid')
      expect(b.show).not.toHaveBeenCalled()
      w.unmount()
    }
  })
})

describe('a setting kept in the browser, then by the source', () => {
  it('the source wins once the sheet is editable, and the browser forgets what the source now keeps', async () => {
    localStorage.clear()
    const b = book([{ name: 'S', maxRow: 3, maxCol: 3 }])
    const w = await mountSheet(b.source, { readOnly: true, storageKey: 'look' })
    api(w).showGridlines(false)
    await settle()
    expect(b.show).not.toHaveBeenCalled()
    await w.setProps({ readOnly: false })
    await settle()
    expect(api(w).showGridlines(true)).toBe(true)
    await settle()
    expect(b.show).toHaveBeenCalledWith('S', { gridlines: true })
    b.emit({ kind: 'reload' })
    await settle()
    expect(lastState(w)?.gridlines).toBeUndefined()
    await w.setProps({ readOnly: true })
    b.emit({ kind: 'reload' })
    await settle()
    expect(lastState(w)?.gridlines).toBeUndefined()
  })
})

describe('outline bars without headings', () => {
  const sheet = (headings: boolean, rows: boolean): SheetInfo =>
    ({ name: 'S', maxRow: 6, maxCol: 6, ...(headings ? {} : { headings: false }), outline: { rows: rows ? [{ first: 2, last: 3, level: 1, summary: 4, collapsed: false }] : [], columns: [{ first: 2, last: 3, level: 1, summary: 4, collapsed: false }] } }) as unknown as SheetInfo
  const levels = (w: VueWrapper, axis: string) => w.findAll(`.xs-ol-levels-${axis} [data-xs-outline-level]`).length
  for (const rows of [false, true])
    it(`keeps the bars and drops every level button, as Excel for Mac (${rows ? 'row and column groups' : 'column groups'})`, async () => {
      const on = await mountSheet(fakeSource([sheet(true, rows)]).source)
      await settle()
      expect([levels(on, 'columns'), levels(on, 'rows')]).toEqual([2, rows ? 2 : 0])
      on.unmount()
      const off = await mountSheet(fakeSource([sheet(false, rows)]).source)
      await settle()
      expect([levels(off, 'columns'), levels(off, 'rows'), off.find('.xs-ol-columns').exists(), off.find('.xs-ol-rows').exists()]).toEqual([0, 0, true, rows])
      // the first cell starts at the grid's left edge unless the row bar takes it
      expect(off.find('.xs-head').exists()).toBe(false)
    })
})
