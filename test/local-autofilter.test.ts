// @vitest-environment node
// localSource's AutoFilter: the sheet's range from xlcalc's SheetInput.autoFilter, Data > Filter through one `setLayout`
// (a single undo step), Filter off showing the rows the file's filter hid, and the library moving it with rows and columns.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { SourceEvent } from '../src/contract/source'
import { closeAll, source } from './local-fixture'

afterEach(closeAll)

/** A list with its AutoFilter A1:B5, filtered in the file: row 3 hidden by it, row 8 (below it) hidden by hand. */
const filteredBook = (): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: { value: 'Key' }, B1: { value: 'Status' }, A2: { value: 'k2' }, B2: { value: 'a' }, A3: { value: 'k3' }, B3: { value: 'b' }, A4: { value: 'k4' }, B4: { value: 'a' }, A5: { value: 'k5' }, B5: { value: 'a' }, A8: { value: 'x' } },
      autoFilter: { range: 'A1:B5' },
      filtered: true,
      rows: { '3': { hidden: true }, '8': { hidden: true } },
    },
    { name: 'T', cells: { A1: { value: 1 } } },
  ],
})
const sheetOf = async (s: ReturnType<typeof source>) => (await s.book()).sheets[0]!

describe('localSource: the AutoFilter', () => {
  it("lists a sheet's AutoFilter, and none for a sheet without one", async () => {
    const list = await source(filteredBook()).sheets()
    expect(list[0]!.autoFilter).toEqual({ range: 'A1:B5' })
    expect(list[1]!.autoFilter).toBeUndefined()
  })

  it('puts an AutoFilter on a sheet in one undo step, and reloads', async () => {
    const s = source(filteredBook())
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const undo = s.history.undo
    expect(await s.setAutoFilter('T', 'A1')).toMatchObject({ ok: true })
    expect((await s.sheets())[1]!.autoFilter).toEqual({ range: 'A1' })
    expect(s.history.undo).toBe(undo + 1)
    expect(events).toContainEqual({ kind: 'reload' })
    await s.undo()
    expect((await s.sheets())[1]!.autoFilter).toBeUndefined()
  })

  it("removes it as Excel's Filter button does: the rows its filter hid are shown, others stay hidden, in one undo step", async () => {
    const s = source(filteredBook())
    const undo = s.history.undo
    expect(await s.setAutoFilter('S', null)).toMatchObject({ ok: true })
    const sheet = await sheetOf(s)
    expect(sheet.autoFilter).toBeUndefined()
    expect(sheet.filtered).toBeFalsy()
    expect(sheet.rows?.['3']?.hidden).toBeFalsy()
    expect(sheet.rows?.['8']?.hidden).toBe(true)
    expect(s.history.undo).toBe(undo + 1)
    await s.undo()
    const back = await sheetOf(s)
    expect([back.autoFilter, back.filtered, back.rows?.['3']?.hidden]).toEqual([{ range: 'A1:B5' }, true, true])
  })

  it('refuses an AutoFilter over a table, with the library\'s message', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 'h' }, A2: { value: 1 } }, tables: [{ name: 'List', range: 'A1:A2', columns: [{ name: 'h' }] }] }] })
    const got = await s.setAutoFilter('S', 'A1:A2')
    expect(got).toMatchObject({ ok: false })
    expect((await s.sheets())[0]!.autoFilter).toBeUndefined()
  })

  it('rows inserted inside it grow it, as the library moves it', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 'Key' }, A2: { value: 1 } }, autoFilter: { range: 'A1:B5' } }] })
    await s.structure({ op: 'insertRows', sheet: 'S', row: 3, count: 2 })
    expect((await s.sheets())[0]!.autoFilter).toEqual({ range: 'A1:B7' })
  })
})
