// @vitest-environment node
// localSource's Find and Replace as Excel's: in the text a cell opens with (formulas in the book's locale, values by the
// format they show) or in the shown text, kept with the edits and dropped on a reload; Replace All in one step of undo.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import type { FindRequest, SourceEvent } from '../src/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

/** Hits as `row:col text`. */
async function found(s: ReturnType<typeof source>, req: Partial<FindRequest> & { sheet: string; query: string }): Promise<string[]> {
  return (await s.find(req)).hits.map((h) => `${h.row}:${h.col} ${h.text}`)
}

/** Numbers by their formats, a date by its column's, a text that reads as a number, a formula. */
const formats = (): BookInput => ({
  locale: 'ru-RU',
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 46096 },
        B1: { value: 1250.5, numberFormat: '#,##0.00' },
        C1: { value: 0.05, numberFormat: '0%' },
        D1: { value: '007' },
        E1: { formula: 'SUM(B1:C1)' },
        F1: { value: 'x' },
      },
      columns: { A: { numberFormat: 14 } },
    },
  ],
})

describe('localSource: find', () => {
  it('looks in formulas by default, as Excel: the text each cell opens with, formulas in the book’s locale', async () => {
    const s = source()
    expect(await found(s, { sheet: 'Лист1', query: 'С' })).toEqual(['1:2 с примечанием', '2:2 ссылка', '3:1 =СУММ(A1:A2)', '3:2 на лист', '4:2 место', '5:4 строка'])
    expect(await found(s, { sheet: 'Лист1', query: 'С', lookIn: 'values' })).toEqual(['1:2 с примечанием', '2:2 ссылка', '3:2 на лист', '4:2 место', '5:4 строка'])
    expect(await found(s, { sheet: 'Таблица', query: '[@' })).toEqual(['2:3 =[@Сумма]*2', '3:3 =[@Сумма]*2'])
  })

  it('reads a value as its edit text in formulas and as shown in values: dates by the column’s format, numbers whole', async () => {
    const s = source(formats())
    expect(await found(s, { sheet: 'S', query: '15.03' })).toEqual(['1:1 15.03.2026'])
    expect(await found(s, { sheet: 'S', query: '1250,5' })).toEqual(['1:2 1250,5'])
    expect(await found(s, { sheet: 'S', query: '1250,5', lookIn: 'values' })).toEqual(['1:5 1250,55'])
    expect((await s.find({ sheet: 'S', query: '250,50', lookIn: 'values' })).hits.map((h) => [h.row, h.col])).toEqual([[1, 2]])
    expect(await found(s, { sheet: 'S', query: '5%' })).toEqual(['1:3 5%'])
    expect(await found(s, { sheet: 'S', query: "'0" })).toEqual(["1:4 '007"])
    expect(await found(s, { sheet: 'S', query: '=сумм(*)', wholeCell: true })).toEqual(['1:5 =СУММ(B1:C1)'])
  })

  it('takes the range, the order by columns and the display filter', async () => {
    const s = source()
    expect(await found(s, { sheet: 'Список', query: '*', range: 'A3:B4' })).toEqual(['3:1 а', '3:2 30', '4:1 б', '4:2 нет'])
    expect(await found(s, { sheet: 'Список', query: '*', range: 'A3:B4', byColumns: true })).toEqual(['3:1 а', '4:1 б', '3:2 30', '4:2 нет'])
    const filter = { head: 1, cols: { 1: { values: ['б'] } } }
    await s.visibleRows('Список', filter)
    expect(await found(s, { sheet: 'Список', query: '?', wholeCell: true, filter })).toEqual(['2:1 б', '2:2 5', '4:1 б'])
  })

  it('keeps the edit texts with edits, drops them on a reload, and takes an undo back', async () => {
    const s = source(formats())
    expect(await found(s, { sheet: 'S', query: 'x' })).toEqual(['1:6 x'])
    await s.edit([{ sheet: 'S', cell: 'F1', text: 'y' }, { sheet: 'S', cell: 'F2', text: '=F1&"x"' }])
    expect(await found(s, { sheet: 'S', query: 'x' })).toEqual(['2:6 =F1&"x"'])
    // a format changes how a number opens: the source reloads and reads the texts again
    await s.format('S', [{ range: 'B1', numberFormat: '0%' }])
    expect(await found(s, { sheet: 'S', query: '125050%' })).toEqual(['1:2 125050%'])
    await s.undo()
    expect(await found(s, { sheet: 'S', query: '1250,5' })).toEqual(['1:2 1250,5'])
  })
})

describe('localSource: replace', () => {
  it('replaces in formulas in one step of undo, tells subscribers and gives the changes and the count', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const before = s.history.undo
    const out = await s.replace({ sheet: 'Лист1', query: 'a2', lookIn: 'values' }, 'A1')
    expect(out).toMatchObject({ ok: true, replaced: 1 })
    expect(out.ok && out.changes.find((c) => c.cell === 'A3')).toMatchObject({ text: '1', value: 1, formula: 'СУММ(A1:A1)', editText: '=СУММ(A1:A1)' })
    expect(s.history.undo).toBe(before + 1)
    expect(events).toEqual([{ kind: 'changed', changes: out.ok ? out.changes : [] }])
    expect(await found(s, { sheet: 'Лист1', query: 'a1:a1' })).toEqual(['3:1 =СУММ(A1:A1)'])
    await s.undo()
    expect((await win(s, 'Лист1', 3, 1, 1, 1)).formulas).toEqual({ A3: 'СУММ(A1:A2)' })
  })

  it('writes each new text as typed: the apostrophe keeps a text, numbers in the locale, the columns as they are', async () => {
    const s = source(formats())
    // 15.03.2026 → 16.03.2026 by the column's date format, 1250,5 → 1260,6; a percent cell is left out, as typing `6%`
    // into it gives the text `6%%` (the library's percent editor of fresh typing, see the report)
    expect(await s.replace({ sheet: 'S', query: '5', range: 'A1:B1' }, '6')).toMatchObject({ ok: true, replaced: 2 })
    expect(await s.replace({ sheet: 'S', query: '0', range: 'D1' }, '1')).toMatchObject({ ok: true, replaced: 1 })
    const w = await win(s, 'S', 1, 1, 1, 6)
    expect(w.values[0]!.slice(0, 4)).toEqual([46097, 1260.6, 0.05, '117'])
    expect(await s.replace({ sheet: 'S', query: 'x', wholeCell: true }, '123456789012345')).toMatchObject({ ok: true, replaced: 1 })
    expect((await s.book()).sheets[0]!.columns?.F).toBeUndefined()
  })

  it('reads cells replaced far apart back in windows of the cell limit, so subscribers get changes, not a reload', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 'x' }, C50: { value: 'x' }, A9000: { value: 'x' } } }] }, { maxCells: 10 })
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const out = await s.replace({ sheet: 'S', query: 'x' }, 'y')
    expect(out.ok && out.changes.map((c) => `${c.cell} ${c.text}`)).toEqual(['A1 y', 'C50 y', 'A9000 y'])
    expect(events.map((e) => e.kind)).toEqual(['changed'])
  })

  it('finds nothing to replace in shown text, and refuses a formula the replacement breaks with nothing written', async () => {
    const s = source(formats())
    const before = s.history.undo
    expect(await s.replace({ sheet: 'S', query: '1 250', lookIn: 'values' }, '2')).toEqual({ ok: true, changes: [], replaced: 0 })
    const out = await s.replace({ sheet: 'S', query: ':', lookIn: 'formulas' }, '::')
    expect(out).toMatchObject({ ok: false, reason: 'invalid', replaced: 0, cells: [{ sheet: 'S', cell: 'E1' }] })
    expect(s.history.undo).toBe(before)
    expect((await win(s, 'S', 1, 1, 5, 1)).formulas).toEqual({ E1: 'СУММ(B1:C1)' })
  })
})
