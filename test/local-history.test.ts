// @vitest-environment node
// localSource paste, Control+Return, undo and redo: text read as Excel reads it there, one step of the undo per call.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { SourceEvent } from '../src/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** A ru-RU sheet with numbers in A, a percent column C and a stop rule over D. */
const sheet = (): BookInput => ({
  locale: 'ru-RU',
  sheets: [{
    name: 'S',
    cells: { A1: { value: 2 }, A2: { value: 3 }, A3: { value: 4 }, D1: { value: 1 } },
    columns: { C: { numberFormat: '0%' } },
    validations: [{ range: 'D1:D9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'stop', error: 'Только больше нуля' }],
  }],
})

describe('localSource: paste', () => {
  it("keeps the apostrophe of '007 in the text, as Excel pastes it, and selects what it filled", async () => {
    const got = await source(sheet()).paste!('S', 'B1', "'007")
    expect(got).toEqual({ ok: true, select: { sheet: 'S', range: 'B1' }, changes: [{ sheet: 'S', cell: 'B1', text: "'007", value: "'007", formula: null, editText: "''007", numberFormat: null }] })
  })

  it('pastes 5 into a percent cell as 5, shown 500 %: a paste has no percent editor', async () => {
    const s = source(sheet())
    expect(await s.paste!('S', 'C1', '5')).toMatchObject({ ok: true, changes: [{ cell: 'C1', value: 5, text: '500%', editText: '500%' }] })
  })

  it('fills a rectangle from the cell, empty fields clearing theirs, the formulas in the locale', async () => {
    const s = source(sheet())
    const got = await s.paste!('S', 'A2', '=СУММ(A1;1,5)\t1 234,5\r\n\tx')
    expect(got).toMatchObject({ ok: true, select: { sheet: 'S', range: 'A2:B3' } })
    if (!got.ok) return
    const by = Object.fromEntries(got.changes.map((c) => [c.cell, c]))
    expect(by.A2).toMatchObject({ value: 3.5, formula: 'СУММ(A1;1,5)', editText: '=СУММ(A1;1,5)' })
    expect(by.B2).toMatchObject({ value: 1234.5, text: '1 234,50' })
    expect(by.A3).toMatchObject({ value: null, text: '', editText: '' })
    expect(by.B3).toMatchObject({ value: 'x' })
  })

  it('refuses a rectangle past the sheet', async () => {
    expect(await source(sheet()).paste!('S', 'XFD1', 'a\tb')).toMatchObject({ ok: false, reason: 'refused' })
  })
})

describe('localSource: Control+Return', () => {
  it('enters =A1*2 into B1:B3 as =A2*2 and =A3*2 below, as a copy moves it', async () => {
    const got = await source(sheet()).fill!('S', 'B1:B3', '=A1*2', 'B1')
    expect(got).toMatchObject({ ok: true, select: { sheet: 'S', range: 'B1:B3' } })
    if (!got.ok) return
    const by = Object.fromEntries(got.changes.map((c) => [c.cell, c]))
    expect(by.B1).toMatchObject({ formula: 'A1*2', value: 4, editText: '=A1*2' })
    expect(by.B2).toMatchObject({ formula: 'A2*2', value: 6 })
    expect(by.B3).toMatchObject({ formula: 'A3*2', value: 8 })
  })

  it('gives invalid at the place of a formula that does not read, with the active cell', async () => {
    const got = await source(sheet()).fill!('S', 'B1:B3', '=СУММ(A1;))', 'B2')
    expect(got).toMatchObject({ ok: false, reason: 'invalid', at: 10 })
  })

  it('gives the formula Excel offers for a typo as `correction`, as typing does', async () => {
    expect(await source(sheet()).fill!('S', 'B1:B3', '=(1+2', 'B2')).toMatchObject({ ok: false, reason: 'invalid', correction: '=(1+2)' })
  })

  it('refuses every cell when the active cell\'s stop rule does not take the value', async () => {
    const s = source(sheet())
    expect(await s.fill!('S', 'D1:D2', '-1', 'D1')).toEqual({
      ok: false,
      reason: 'refused',
      message: 'Только больше нуля',
      alert: { severity: 'stop', message: 'Только больше нуля', sheet: 'S', cell: 'D1' },
    })
    expect((await win(s, 'S', 1, 2, 4, 1)).values).toEqual([[1], [null]])
    expect(s.history).toEqual({ undo: 0, redo: 0 })
  })

  it('refuses a value a warning or an information rule does not take with its alert, and enters it after Yes or OK (`accept`)', async () => {
    const s = source({
      sheets: [{
        name: 'S',
        cells: {},
        validations: [
          { range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' },
          { range: 'B1:B9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'information' },
        ],
      }],
    })
    expect(await s.fill!('S', 'A1:A2', '-1', 'A1')).toMatchObject({ ok: false, reason: 'refused', alert: { severity: 'warning', cell: 'A1' } })
    expect(await s.fill!('S', 'B1:B2', '-2', 'B2')).toMatchObject({ ok: false, reason: 'refused', alert: { severity: 'information', cell: 'B2' } })
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect(await s.fill!('S', 'A1:A2', '-1', 'A1', { accept: true })).toMatchObject({ ok: true, select: { sheet: 'S', range: 'A1:A2' } })
    expect(await s.fill!('S', 'B1:B2', '-2', 'B2', { accept: true })).toMatchObject({ ok: true })
    expect((await win(s, 'S', 1, 2, 1, 2)).values).toEqual([[-1, -2], [-1, -2]])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
  })
})

describe('localSource: undo and redo', () => {
  it('takes a paste of several cells back in one undo and does it again in one redo', async () => {
    const s = source(sheet())
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.paste!('S', 'A1', '7\t8\r\n9\t10')
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    const back = await s.undo!()
    expect(back).toMatchObject({ ok: true, select: { sheet: 'S', range: 'A1:B2' } })
    expect((await win(s, 'S', 1, 2, 1, 2)).values).toEqual([[2, null], [3, null]])
    expect(s.history).toEqual({ undo: 0, redo: 1 })
    expect(await s.redo!()).toMatchObject({ ok: true, select: { sheet: 'S', range: 'A1:B2' } })
    expect((await win(s, 'S', 1, 2, 1, 2)).values).toEqual([[7, 8], [9, 10]])
    expect(events.map((e) => e.kind)).toEqual(['changed', 'changed', 'changed'])
  })

  it('gives the edited cells their formula and edit text back', async () => {
    const s = source(sheet())
    await s.edit([{ sheet: 'S', cell: 'B1', text: '=A1*10' }])
    await s.edit([{ sheet: 'S', cell: 'B1', text: '15.03.2026' }])
    const back = await s.undo!()
    expect(back).toMatchObject({ ok: true, changes: [{ cell: 'B1', value: 20, formula: 'A1*10', editText: '=A1*10' }] })
  })

  it('undoes a structure operation by a reload', async () => {
    const s = source(sheet())
    await s.structure({ op: 'insertRows', sheet: 'S', row: 1 })
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.undo!()).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect((await win(s, 'S', 1, 1, 1, 1)).values).toEqual([[2]])
    expect(await s.find({ sheet: 'S', query: '4' })).toMatchObject({ hits: [{ row: 3, col: 1 }] })
  })

  it('answers ok with no changes when there is nothing to undo or redo', async () => {
    const s = source(sheet())
    expect(await s.undo!()).toEqual({ ok: true, changes: [] })
    expect(await s.redo!()).toEqual({ ok: true, changes: [] })
  })

  it('leaves the history as it was after a value a stop rule refused, as Excel never took it', async () => {
    const s = source(sheet())
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.edit([{ sheet: 'S', cell: 'D1', text: '-5' }])).toMatchObject({ ok: false, reason: 'refused' })
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect(await s.redo!()).toEqual({ ok: true, changes: [] })
    expect((await win(s, 'S', 1, 1, 4, 1)).values).toEqual([[1]])
    expect(events).toEqual([])
    // a step undone before a refusal stays to redo: the refused value wrote nothing that would forget it
    await s.edit([{ sheet: 'S', cell: 'B1', text: '1' }])
    await s.undo!()
    expect(await s.fill!('S', 'D1:D2', '-1', 'D1')).toMatchObject({ ok: false, reason: 'refused' })
    expect(await s.edit([{ sheet: 'S', cell: 'D2', text: '0' }])).toMatchObject({ ok: false, reason: 'refused' })
    expect(s.history).toEqual({ undo: 0, redo: 1 })
    expect(await s.redo!()).toMatchObject({ ok: true, changes: [{ cell: 'B1', value: 1 }] })
    expect(s.history).toEqual({ undo: 1, redo: 0 })
  })
})
