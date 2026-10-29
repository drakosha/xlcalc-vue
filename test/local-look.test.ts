// @vitest-environment node
// localSource and the sheet's look beside its cells: tab colors through `setLayout`, a single undo step, kept in the workbook;
// View > Show, which the session has no edit for, kept by the source on the sheets it lists and the workbook it saves.
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetWindow, SourceEvent } from '../src/contract/source'
import { openXlsx, saveXlsx } from '../src/local'
import { closeAll, source } from './local-fixture'

afterEach(closeAll)

const two = () => source({ sheets: [{ name: 'A', cells: { A1: { value: 1 } }, tabColor: '#00B050' }, { name: 'B', cells: {} }] })

describe('localSource: tab colors', () => {
  it("lists a sheet's tab color from the workbook", async () => {
    const list = await two().sheets()
    expect(list[0]!.tabColor).toBe('#00B050')
    expect(list[1]!.tabColor).toBeUndefined()
  })

  it('sets and removes a tab color in one undo step each, keeps it in the workbook, and tells subscribers to read the sheets again', async () => {
    const s = two()
    const events: SourceEvent['kind'][] = []
    s.subscribe((e) => events.push(e.kind))
    expect(await s.tabColor('B', '#FF0000')).toMatchObject({ ok: true })
    expect(events).toContain('reload')
    expect((await s.sheets())[1]!.tabColor).toBe('#FF0000')
    expect((await s.book()).sheets[1]!.tabColor).toBe('#FF0000')
    expect(s.history.undo).toBe(1)
    expect(await s.tabColor('A', null)).toMatchObject({ ok: true })
    expect((await s.sheets())[0]!.tabColor).toBeUndefined()
    await s.undo()
    expect((await s.sheets())[0]!.tabColor).toBe('#00B050')
    await s.undo()
    expect((await s.sheets())[1]!.tabColor).toBeUndefined()
  })
})

describe("localSource: View > Show", () => {
  const book = () => source({ sheets: [{ name: 'A', cells: { A1: { value: 1 } }, view: { gridlines: false } }, { name: 'B', cells: { A1: { value: 2 } } }] })

  it("lists a sheet's gridlines and headings from the workbook's view", async () => {
    const list = await book().sheets()
    expect(list[0]!.gridlines).toBe(false)
    expect(list[0]!.headings).toBeUndefined()
    expect(list[1]!.gridlines).toBeUndefined()
  })

  it('keeps what show sets without an undo step, in the sheets it lists and the workbook it saves', async () => {
    const s = book()
    expect(await s.show('B', { headings: false })).toMatchObject({ ok: true })
    expect(await s.show('A', { gridlines: true })).toMatchObject({ ok: true })
    expect(s.history.undo).toBe(0)
    const list = await s.sheets()
    expect(list[1]!.headings).toBe(false)
    expect(list[0]!.gridlines).toBeUndefined()
    expect((await s.book()).sheets[1]!.view).toMatchObject({ headings: false })
    expect((await s.snapshot()).book.sheets[0]!.view?.gridlines).not.toBe(false)
    // written to the file and read back
    const file = await saveXlsx(s)
    const read = await openXlsx(file.data)
    expect(read.book.sheets[1]!.view).toMatchObject({ headings: false })
    expect(read.book.sheets[0]!.view?.gridlines).not.toBe(false)
  })

  it('follows a sheet renamed, and renamed back by undo', async () => {
    const s = book()
    await s.show('B', { gridlines: false })
    await s.structure({ op: 'renameSheet', sheet: 'B', name: 'C' })
    expect((await s.sheets())[1]).toMatchObject({ name: 'C', gridlines: false })
    await s.undo()
    expect((await s.sheets())[1]).toMatchObject({ name: 'B', gridlines: false })
    await s.redo()
    expect((await s.sheets())[1]).toMatchObject({ name: 'C', gridlines: false })
  })

  it('brings a deleted sheet back with them on undo; a new sheet of its name has none', async () => {
    const s = book()
    await s.show('B', { gridlines: false, headings: false })
    await s.structure({ op: 'deleteSheet', sheet: 'B' })
    await s.undo()
    expect((await s.sheets())[1]).toMatchObject({ name: 'B', gridlines: false, headings: false })
    await s.redo()
    expect((await s.sheets()).map((x) => x.name)).toEqual(['A'])
    await s.structure({ op: 'insertSheet', name: 'B' })
    expect((await s.sheets())[1]).toMatchObject({ name: 'B' })
    expect((await s.sheets())[1]!.gridlines).toBeUndefined()
  })
})

describe('localSource: notes shown all the time', () => {
  it("gives a note's shown in its window and writes it with the note, to the file too", async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 1, note: { text: 'Always', shown: true } }, B1: { value: 2, note: { text: 'On hover' } } } }] })
    const w = (await s.window({ sheet: 'S', row: 1, rows: 1, col: 1, cols: 2 })) as SheetWindow
    expect(w.notes).toEqual({ A1: { text: 'Always', shown: true }, B1: { text: 'On hover' } })
    expect(await s.note('S', 'B1', { text: 'On hover', shown: true })).toMatchObject({ ok: true })
    expect(await s.note('S', 'A1', { text: 'Always' })).toMatchObject({ ok: true })
    const after = (await s.window({ sheet: 'S', row: 1, rows: 1, col: 1, cols: 2 })) as SheetWindow
    expect(after.notes).toEqual({ A1: { text: 'Always' }, B1: { text: 'On hover', shown: true } })
    const read = await openXlsx((await saveXlsx(s)).data)
    expect(read.book.sheets[0]!.cells.B1!.note).toMatchObject({ shown: true })
    expect(read.book.sheets[0]!.cells.A1!.note?.shown).toBeUndefined()
  })
})
