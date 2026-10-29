// Excel's View > Split: the split placed from the active cell as Excel places it, kept through the source's `split` (not an
// undo step), reported to a ribbon, exclusive with frozen panes; `./local` reads it, keeps it and writes it to the file.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { SheetViewApi, SheetViewState } from '../src'
import type { EditOutcome, SheetInfo, SheetSplit } from '../src/contract/source'
import { splitPlace } from '../src/core/split'
import { openXlsx, saveXlsx } from '../src/local'
import { closeAll, source } from './local-fixture'
import { fakeSource, goTo, mountSheet, settle } from './helpers'

afterEach(async () => {
  document.body.innerHTML = ''
  await closeAll()
})

const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null

function book(info: Partial<SheetInfo> = {}, refuse = false) {
  const sheets: SheetInfo[] = [{ name: 'S', maxRow: 50, maxCol: 12, ...info }]
  const f = fakeSource(sheets)
  const split = vi.fn(async (sheet: string, panes: SheetSplit | null): Promise<EditOutcome> => {
    if (refuse) return { ok: false, reason: 'refused', message: 'Not now' }
    const { split: _, ...rest } = sheets[0]!
    sheets[0] = panes ? { ...rest, split: panes } : rest
    return { ok: true, changes: [] }
  })
  f.source.split = split
  return { ...f, split }
}

describe("Excel's place for a split", () => {
  it('goes above and left of the active cell from the window’s first cell, one line in its first row or column, the middle from that cell', () => {
    const middle = () => ({ rows: 9, cols: 5 })
    expect(splitPlace({ d: 5, col: 3 }, { d: 1, col: 1 }, middle)).toEqual({ rows: 4, cols: 2 })
    expect(splitPlace({ d: 5, col: 1 }, { d: 1, col: 1 }, middle)).toEqual({ rows: 4, cols: 0 })
    expect(splitPlace({ d: 1, col: 3 }, { d: 1, col: 1 }, middle)).toEqual({ rows: 0, cols: 2 })
    expect(splitPlace({ d: 1, col: 1 }, { d: 1, col: 1 }, middle)).toEqual({ rows: 9, cols: 5 })
    expect(splitPlace({ d: 25, col: 4 }, { d: 21, col: 4 }, middle)).toEqual({ rows: 4, cols: 0 })
  })
})

describe('View > Split on the sheet', () => {
  it('splits above and left of the active cell in the file’s points, keeps it through the source, and removes it again', async () => {
    const b = book()
    const w = await mountSheet(b.source)
    await goTo(w, 'C5')
    expect(api(w).split()).toBe(true)
    await settle()
    const got = b.split.mock.calls[0]![1]!
    expect(got).toMatchObject({ topLeft: 'C5', activePane: 'bottomRight' })
    // the headings counted in: the row numbers' 56 px and the letters' 24 px, then two columns and four rows, in points
    expect(got.x).toBeGreaterThan(56 * 0.75)
    expect(got.y).toBe((24 + 4 * 20) * 0.75)
    expect(lastState(w)?.split).toBe(true)
    expect(api(w).split()).toBe(false)
    await settle()
    expect(b.split).toHaveBeenLastCalledWith('S', null)
    expect(lastState(w)?.split).toBeUndefined()
  })

  it('puts one line for a cell in the first column, and four panes at the middle for the first cell', async () => {
    const b = book()
    const w = await mountSheet(b.source)
    await goTo(w, 'A5')
    api(w).split(true)
    expect(b.split.mock.calls[0]![1]).toEqual({ y: (24 + 4 * 20) * 0.75, topLeft: 'A5', activePane: 'bottomLeft' })
    api(w).split(false)
    await goTo(w, 'A1')
    api(w).split(true)
    const middle = b.split.mock.calls[2]![1]!
    expect(middle.x).toBeGreaterThan(0)
    expect(middle.y).toBeGreaterThan(0)
    expect(middle.activePane).toBeUndefined()
    expect(middle.topLeft).not.toBe('A1')
  })

  it('is not offered with frozen panes, which take its place, or without the source’s split; a refusal puts it back', async () => {
    const frozen = book({ freeze: { rows: 1, cols: 0 } })
    const w = await mountSheet(frozen.source)
    await goTo(w, 'C5')
    expect(api(w).split()).toBe(false)
    expect(frozen.split).not.toHaveBeenCalled()

    const refused = book({}, true)
    const r = await mountSheet(refused.source)
    await goTo(r, 'C5')
    expect(api(r).split()).toBe(true)
    await settle()
    expect(lastState(r)?.split).toBeUndefined()
    expect(r.find('.xs-message').text()).toBe('Not now')
  })
})

describe('./local', () => {
  const splitBook = (): BookInput => ({ sheets: [{ name: 'S', cells: { A1: { value: 1 } }, view: { split: { x: 120, y: 60, topLeft: 'F20', activePane: 'bottomRight' as const } } }, { name: 'T', cells: {}, frozen: { rows: 1 }, view: { split: { y: 60 } } }] })

  it("lists a sheet's split, but not under frozen panes, and keeps a split set or removed in the file it saves", async () => {
    const s = source(splitBook())
    const list = await s.sheets()
    expect(list[0]!.split).toEqual({ x: 120, y: 60, topLeft: 'F20', activePane: 'bottomRight' })
    expect(list[1]!.split).toBeUndefined()
    const undo = s.history.undo
    expect(await s.split!('S', { y: 45, topLeft: 'A4', activePane: 'bottomLeft' })).toEqual({ ok: true, changes: [] })
    expect(s.history.undo).toBe(undo)
    expect((await s.sheets())[0]!.split).toEqual({ y: 45, topLeft: 'A4', activePane: 'bottomLeft' })
    const back = await openXlsx((await saveXlsx(s)).data)
    expect(back.book.sheets[0]!.view?.split).toEqual({ y: 45, topLeft: 'A4', activePane: 'bottomLeft' })
    await s.split!('S', null)
    expect((await s.sheets())[0]!.split).toBeUndefined()
    expect((await s.book()).sheets[0]!.view?.split).toBeUndefined()
  })
})
