// @vitest-environment node
// localSource's links to other workbooks (xlcalc 0.62.0's setExternalBooks): a host that read another workbook gives the
// values its links need; the workbook is calculated again, in one step of undo, and typed text finds a new link by its name.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { SourceEvent } from '../src/contract/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

const linked = (): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { formula: "'[1]Sheet'!$I$8*2" } } }],
  externalBooks: [{ target: 'Report.xlsx', sheets: [{ name: 'Sheet', cells: { I8: 42 } }] }],
})

describe('localSource: links to other workbooks', () => {
  it("gives a link its values read again: the formulas over it show them, sent as changes, in one step of undo", async () => {
    const s = source(linked())
    expect((await win(s, 'S', 1, 1, 1, 1)).values).toEqual([[84]])
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const got = await s.setExternalBooks([{ target: 'Report.xlsx', sheets: [{ name: 'Sheet', cells: { I8: 50 } }] }])
    expect(got).toEqual({ ok: true, changes: [{ sheet: 'S', cell: 'A1', text: '100', value: 100 }] })
    expect(events).toEqual([{ kind: 'changed', changes: got.ok ? got.changes : [] }])
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    await s.undo()
    expect((await win(s, 'S', 1, 1, 1, 1)).values).toEqual([[84]])
    expect((await s.book()).externalBooks?.[0]?.sheets[0]?.cells).toEqual({ I8: 42 })
  })

  it('adds a link after the others, which typed text then finds by its file name', async () => {
    const s = source(linked())
    await s.setExternalBooks([...(await s.book()).externalBooks!, { target: 'Rates.xlsx', sheets: [{ name: 'R', cells: { A1: 3 } }] }])
    const got = await s.edit([{ sheet: 'S', cell: 'B1', text: "='[Rates.xlsx]R'!A1+1" }])
    // displayed with the file's name, as Excel displays a closed workbook's reference; stored by the link's number
    expect(got.ok && got.changes[0]).toMatchObject({ cell: 'B1', value: 4, formula: '[Rates.xlsx]R!A1+1' })
    expect((await s.book()).sheets[0]!.cells.B1).toEqual({ formula: "'[2]R'!A1+1" })
  })

  it('reads a link marked open as Excel reads an open workbook: a typed formula sees a cell missing from its values as blank, not #REF!', async () => {
    const typed = async (open: boolean) => {
      const s = source(linked())
      if (open) await s.setExternalBooks([{ ...(await s.book()).externalBooks![0]!, open: true }])
      const got = await s.edit([{ sheet: 'S', cell: 'B1', text: "='[Report.xlsx]Sheet'!J9+1" }])
      return got.ok && got.changes[0]?.value
    }
    // a closed workbook: typing a formula over it marks the link, as Excel does when it cannot read the file
    expect(await typed(false)).toEqual({ error: 'REF' })
    expect(await typed(true)).toBe(1)
  })

  it("finds and shows a link stored with %20 by the name a person reads, as Excel shows it", async () => {
    const s = source({ sheets: [{ name: 'S', cells: {} }], externalBooks: [{ target: 'Report%202025.xlsx', sheets: [{ name: 'Sheet', cells: { I8: 42 } }] }] })
    const got = await s.edit([{ sheet: 'S', cell: 'B1', text: "='[Report 2025.xlsx]Sheet'!I8+1" }])
    expect(got.ok && got.changes[0]).toMatchObject({ cell: 'B1', value: 43, formula: "'[Report 2025.xlsx]Sheet'!I8+1" })
    expect((await s.book()).sheets[0]!.cells.B1).toEqual({ formula: "'[1]Sheet'!I8+1" })
  })

  it('rejects links the workbook would not take, with the session message', async () => {
    const got = await source(linked()).setExternalBooks([{ target: 'Report.xlsx', sheets: [{ name: 'Sheet', cells: { I8: [1] as never } }] }])
    expect(got).toMatchObject({ ok: false, reason: 'refused' })
  })
})
