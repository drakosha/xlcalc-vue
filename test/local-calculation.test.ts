// @vitest-environment node
// localSource's calculation options as Excel for Mac 16.114 behaved (measured 29.10): in Manual an edit shows itself and
// leaves its dependents as they were until Calculate Now (F9) or Calculate Sheet (Shift+F9); the option stays with the workbook.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { SourceEvent } from '../src/contract/source'
import { saveXlsx, openXlsx } from '../src/local'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** Sheet One: A1 1, B1 =A1*2, C1 =B1+1; sheet Two: A1 =One!A1*100, B1 =A1+1. */
const twoSheets = (calculation?: BookInput['calculation']): BookInput => ({
  ...(calculation ? { calculation } : {}),
  sheets: [
    { name: 'One', cells: { A1: { value: 1 }, B1: { formula: 'A1*2' }, C1: { formula: 'B1+1' } } },
    { name: 'Two', cells: { A1: { formula: 'One!A1*100' }, B1: { formula: 'A1+1' } } },
  ],
})
const row1 = async (s: ReturnType<typeof source>, sheet: string, cols = 3) => (await win(s, sheet, 1, 1, 1, cols)).values[0]

describe('localSource: calculation options', () => {
  it('reads the option from the workbook: automatic without one', async () => {
    expect(await source(twoSheets()).calculation()).toEqual({ mode: 'automatic', pending: false })
    expect(await source(twoSheets('manual')).calculation()).toEqual({ mode: 'manual', pending: false })
    expect(await source(twoSheets('autoNoTable')).calculation()).toEqual({ mode: 'autoNoTable', pending: false })
  })

  it('in Manual an edit shows the edited cell and leaves its dependents, until Calculate Now', async () => {
    const s = source(twoSheets('manual'))
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const got = await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    expect(got.ok && got.changes.map((c) => [c.cell, c.value])).toEqual([['A1', 5]])
    expect(await row1(s, 'One')).toEqual([5, 2, 3])
    expect(await row1(s, 'Two', 2)).toEqual([100, 101])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: true })
    // the edit's own cells first, then that cells wait
    expect(events.map((e) => e.kind)).toEqual(['changed', 'calculation'])
    expect(events[1]).toEqual({ kind: 'calculation', calculation: { mode: 'manual', pending: true } })
    // the next edit keeps the values shown as of the last calculation, not of the last edit
    await s.edit([{ sheet: 'One', cell: 'A1', text: '6' }])
    expect(await row1(s, 'One')).toEqual([6, 2, 3])

    const calc = await s.calculate()
    expect(calc.ok && calc.changes.map((c) => `${c.sheet}!${c.cell}=${c.value}`).sort()).toEqual(['One!B1=12', 'One!C1=13', 'Two!A1=600', 'Two!B1=601'])
    expect(events.at(-2)).toMatchObject({ kind: 'changed' })
    expect(events.at(-1)).toEqual({ kind: 'calculation', calculation: { mode: 'manual', pending: false } })
    expect(await row1(s, 'One')).toEqual([6, 12, 13])
  })

  it('a formula entered in Manual shows its value, as Excel calculates the entered cell', async () => {
    const s = source(twoSheets('manual'))
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    const got = await s.edit([{ sheet: 'One', cell: 'D1', text: '=A1*10' }])
    expect(got.ok && got.changes.map((c) => [c.cell, c.value])).toEqual([['D1', 50]])
    // later edits hold the entered formula too
    await s.edit([{ sheet: 'One', cell: 'A1', text: '7' }])
    expect((await row1(s, 'One', 4))?.[3]).toBe(50)
  })

  it('Calculate Sheet shows the waiting cells of that sheet only', async () => {
    const s = source(twoSheets('manual'))
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    const got = await s.calculate('Two')
    expect(got.ok && got.changes.map((c) => `${c.cell}=${c.value}`).sort()).toEqual(['A1=500', 'B1=501'])
    expect(await row1(s, 'Two', 2)).toEqual([500, 501])
    expect(await row1(s, 'One')).toEqual([5, 2, 3])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: true })
  })

  it('undo in Manual restores the edited cell and leaves the dependents too', async () => {
    const s = source(twoSheets('manual'))
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    await s.calculate()
    await s.undo()
    expect(await row1(s, 'One')).toEqual([1, 10, 11])
  })

  it('switching to Automatic calculates what waits; to Manual holds from then on; neither is a step of undo', async () => {
    const s = source(twoSheets())
    expect(await s.setCalculation('manual')).toEqual({ ok: true, changes: [] })
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    expect(await row1(s, 'One')).toEqual([5, 2, 3])
    const auto = await s.setCalculation('automatic')
    expect(auto.ok && auto.changes.map((c) => `${c.sheet}!${c.cell}`).sort()).toEqual(['One!B1', 'One!C1', 'Two!A1', 'Two!B1'])
    expect(await row1(s, 'One')).toEqual([5, 10, 11])
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    await s.edit([{ sheet: 'One', cell: 'A1', text: '6' }])
    expect(await row1(s, 'One')).toEqual([6, 12, 13])
  })

  it('the workbook keeps the option it was set to: book(), and a saved file', async () => {
    const s = source(twoSheets())
    await s.setCalculation('manual')
    expect((await s.book()).calculation).toBe('manual')
    await s.setCalculation('autoNoTable')
    expect((await s.snapshot()).book.calculation).toBe('autoNoTable')
    await s.setCalculation('automatic')
    expect((await s.book()).calculation).toBeUndefined()
  })

  it('saveXlsx in Manual calculates first, as Excel recalculates before saving, and writes the option', async () => {
    const s = source(twoSheets('manual'))
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    const saved = await saveXlsx(s)
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: false })
    expect(await row1(s, 'One')).toEqual([5, 10, 11])
    const back = await openXlsx(saved.data)
    expect(back.book.calculation).toBe('manual')
  })

  it('in Automatic an edit shows its dependents and Calculate changes nothing', async () => {
    const s = source(twoSheets())
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    expect(await row1(s, 'One')).toEqual([5, 10, 11])
    expect(await s.calculate()).toEqual({ ok: true, changes: [] })
  })

  it('a formula another command removes shows its new state in Manual: Clear All, Delete Table', async () => {
    const s = source({
      calculation: 'manual',
      sheets: [{ name: 'S', cells: { A1: { value: 'N' }, B1: { value: 'D' }, A2: { value: 1 }, B2: { formula: 'T[[#This Row],[N]]*2' } }, tables: [{ name: 'T', range: 'A1:B2', columns: [{ name: 'N' }, { name: 'D', formula: 'T[[#This Row],[N]]*2' }] }] }],
    })
    await s.edit([{ sheet: 'S', cell: 'A2', text: '5' }])
    expect((await win(s, 'S', 2, 1, 1, 2)).values[0]).toEqual([5, 2])
    expect(await s.clear('S', ['B2'], 'all')).toMatchObject({ ok: true })
    expect((await win(s, 'S', 2, 1, 1, 2)).values[0]).toEqual([5, null])
    await s.undo()
    expect((await win(s, 'S', 2, 1, 1, 2)).values[0]).toEqual([5, 10])
    await s.edit([{ sheet: 'S', cell: 'A2', text: '6' }])
    expect(await s.table({ op: 'deleteTable', table: 'T' })).toMatchObject({ ok: true })
    expect((await win(s, 'S', 2, 1, 1, 2)).values[0]).toEqual([null, null])
  })

  it('in Manual a change of how a formula cell looks shows at once and makes nothing wait: a number format, a narrow column, Clear Formats, a note', async () => {
    const s = source({ calculation: 'manual', sheets: [{ name: 'S', cells: { A1: { value: 0.5 }, B1: { formula: 'A1*2', numberFormat: '0.00%' }, C1: { formula: 'A1*1E+11' } } }] })
    const text = async () => (await win(s, 'S', 1, 1, 1, 3)).text[0]
    expect(await text()).toEqual(['0.5', '100.00%', '5E+10'])
    await s.format('S', [{ range: 'C1', numberFormat: '0.0' }])
    expect((await text())?.[2]).toBe('50000000000.0')
    await s.layout({ sheet: 'S', columns: 'C', width: 2 })
    expect((await text())?.[2]).toMatch(/^#+$/)
    await s.clear('S', ['B1'], 'formats')
    expect((await text())?.[1]).toBe('1')
    await s.note('S', 'B1', { text: 'check' })
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: false })
  })

  it('in Manual a constant cleared by Clear Contents shows empty, after an edit made it known', async () => {
    const s = source(twoSheets('manual'))
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    await s.clear('One', ['A1'], 'contents')
    expect(await row1(s, 'One')).toEqual([null, 2, 3])
  })

  it('a structure operation in Manual shows the calculated values, the cells having moved', async () => {
    const s = source(twoSheets('manual'))
    await s.edit([{ sheet: 'One', cell: 'A1', text: '5' }])
    expect(await s.structure({ op: 'insertRows', sheet: 'One', row: 1 })).toMatchObject({ ok: true })
    expect((await win(s, 'One', 2, 1, 1, 3)).values[0]).toEqual([5, 10, 11])
    expect(await s.calculation()).toEqual({ mode: 'manual', pending: false })
  })
})
