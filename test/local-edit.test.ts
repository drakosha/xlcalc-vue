// @vitest-environment node
// localSource edits: typed text read in the book's locale, refusals, what the edited cells show after.
import { afterEach, describe, expect, it } from 'vitest'
import type { SourceEvent } from '../src/source'
import { BOLD, closeAll, source, styleOf, win } from './local-fixture'

afterEach(closeAll)

describe('localSource: typed edits', () => {
  it('reads 1,5 as a number in ru-RU and sends the dependent cell', async () => {
    const got = await source().edit([{ sheet: 'Лист1', cell: 'A1', text: '1,5' }])
    expect(got.ok).toBe(true)
    if (!got.ok) return
    expect(got.changes).toContainEqual({ sheet: 'Лист1', cell: 'A1', text: '1,5', value: 1.5, formula: null, editText: '1,5', numberFormat: null })
    expect(got.changes.find((c) => c.cell === 'A3')).toMatchObject({ value: 3.5, text: '3,5' })
  })

  it('reads 15.03.2026 as a date with the date format', async () => {
    const s = source()
    // B is 20 wide: in a column of the default width the date shows as ### (the library measures it so)
    const got = await s.edit([{ sheet: 'Лист1', cell: 'B4', text: '15.03.2026' }])
    expect(got).toMatchObject({ ok: true, changes: [{ cell: 'B4', value: 46096, text: '15.03.2026', numberFormat: 'dd.mm.yyyy' }] })
  })

  it('reads 15.03.2026 as a short date and 1 234,5 as a grouped number, widening the column the date does not fit', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.edit([{ sheet: 'Список', cell: 'C2', text: '15.03.2026' }, { sheet: 'Список', cell: 'C3', text: '1 234,5' }])).toMatchObject({ ok: true })
    const cells = (await s.book()).sheets[2]!.cells
    expect(cells.C2).toEqual({ value: 46096, numberFormat: 14 })
    expect(cells.C3).toEqual({ value: 1234.5, numberFormat: 4 })
    // typing widens a column of the default width, as Excel does: the windows are read again
    expect((await win(s, 'Список', 2, 2, 3, 1)).text).toEqual([['15.03.2026'], ['1 234,50']])
    expect(events).toEqual([{ kind: 'reload' }])
  })

  it('gives each edited cell the text it opens with now and its formula as shown', async () => {
    const got = await source().edit([
      { sheet: 'Список', cell: 'C2', text: '15.03.2026' },
      { sheet: 'Список', cell: 'C3', text: '5%' },
      { sheet: 'Список', cell: 'C4', text: "'007" },
      { sheet: 'Список', cell: 'C5', text: '=СУММ(B2:B3)' },
    ])
    expect(got.ok).toBe(true)
    if (!got.ok) return
    const by = Object.fromEntries(got.changes.map((c) => [c.cell, c]))
    expect(by.C2).toMatchObject({ editText: '15.03.2026', formula: null })
    expect(by.C3).toMatchObject({ editText: '5%', formula: null })
    expect(by.C4).toMatchObject({ editText: "'007", text: '007', formula: null })
    expect(by.C5).toMatchObject({ editText: '=СУММ(B2:B3)', formula: 'СУММ(B2:B3)', value: 35 })
  })

  it('refuses a value a stop rule of data validation does not take and keeps the cell; a warning or an information rule keeps the value', async () => {
    const s = source({
      sheets: [{
        name: 'S',
        cells: { A1: { value: 5 }, B1: { value: 5 }, C1: { value: 5 } },
        validations: [
          { range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'stop', error: 'Только больше нуля' },
          { range: 'B1:B9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' },
          { range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'information' },
        ],
      }],
    })
    expect(await s.edit([{ sheet: 'S', cell: 'A1', text: '-2' }])).toEqual({ ok: false, reason: 'refused', message: 'Только больше нуля', cells: [{ sheet: 'S', cell: 'A1' }] })
    expect(s.history).toEqual({ undo: 0, redo: 0 })
    expect(await s.edit([{ sheet: 'S', cell: 'B1', text: '-3' }])).toMatchObject({ ok: true, changes: [{ cell: 'B1', value: -3 }] })
    expect(await s.edit([{ sheet: 'S', cell: 'C1', text: '-4' }])).toMatchObject({ ok: true, changes: [{ cell: 'C1', value: -4 }] })
    expect((await win(s, 'S', 1, 1, 1, 3)).values).toEqual([[5, -3, -4]])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
    // several cells at once: a stop rule of any of them refuses all, a warning's cell coming first too
    expect(await s.edit([{ sheet: 'S', cell: 'B2', text: '-5' }, { sheet: 'S', cell: 'A2', text: '-6' }])).toEqual({ ok: false, reason: 'refused', message: 'Только больше нуля', cells: [{ sheet: 'S', cell: 'A2' }] })
    expect((await win(s, 'S', 2, 1, 1, 2)).values).toEqual([[null, null]])
    expect(s.history).toEqual({ undo: 2, redo: 0 })
  })

  it('clears a selection as Excel\'s Delete: the content and the hyperlink go, the style, format and note stay', async () => {
    const s = source()
    const got = await s.edit([
      { sheet: 'Лист1', cell: 'B1', text: null },
      { sheet: 'Лист1', cell: 'B2', text: null },
      { sheet: 'Лист1', cell: 'C1', text: null },
    ])
    expect(got).toMatchObject({ ok: true })
    const cells = (await s.book()).sheets[0]!.cells
    expect(cells.B1).toEqual({ note: { author: 'Аня', text: 'Проверить' } })
    expect(cells.B2).toBeUndefined()
    expect(cells.C1).toEqual({ style: 1 })
  })

  it('reloads after a Delete that takes a hyperlink and after its undo: a cell update carries no note or link', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.edit([{ sheet: 'Лист1', cell: 'B2', text: null }])
    expect(events).toEqual([{ kind: 'reload' }])
    expect((await win(s, 'Лист1', 1, 3, 1, 3)).links).toEqual({ B3: { location: "'Таблица'!A1" } })
    await s.undo()
    expect(events).toEqual([{ kind: 'reload' }, { kind: 'reload' }])
    expect((await win(s, 'Лист1', 2, 1, 2, 1)).links).toEqual({ B2: { href: 'https://example.org/', tooltip: 'Пример' } })
    // a cell whose note stays and a cell without either are cell updates
    await s.edit([{ sheet: 'Лист1', cell: 'B1', text: null }, { sheet: 'Лист1', cell: 'A1', text: null }])
    expect(events.at(-1)).toMatchObject({ kind: 'changed' })
  })

  it('renames a table column by its header cell, and the formulas over it are shown anew', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.edit([{ sheet: 'Таблица', cell: 'B1', text: 'Итог' }])).toMatchObject({ ok: true })
    expect(events.at(-1)).toEqual({ kind: 'reload' })
    expect((await win(s, 'Таблица', 1, 3, 1, 5)).formulas.E1).toBe('СУММ(Продажи[Итог])')
  })

  it('gives invalid with the position for a formula that does not parse, and changes nothing', async () => {
    const s = source()
    const got = await s.edit([{ sheet: 'Лист1', cell: 'A3', text: '=СУММ(A1:A2' }])
    expect(got).toMatchObject({ ok: false, reason: 'invalid', at: 11, cells: [{ sheet: 'Лист1', cell: 'A3' }] })
    expect((await win(s, 'Лист1', 3, 1, 1, 1)).formulas.A3).toBe('СУММ(A1:A2)')
  })

  it('returns the edited formula in the locale, percent, boolean and apostrophe text', async () => {
    const s = source()
    const got = await s.edit([
      { sheet: 'Лист1', cell: 'C3', text: '=СУММ(A1:A2)*2' },
      { sheet: 'Лист1', cell: 'C4', text: '5%' },
      { sheet: 'Лист1', cell: 'C5', text: 'ИСТИНА' },
      { sheet: 'Лист1', cell: 'C6', text: "'007" },
    ])
    expect(got.ok).toBe(true)
    if (!got.ok) return
    const by = Object.fromEntries(got.changes.map((c) => [c.cell, c]))
    expect(by.C3).toMatchObject({ value: 6, formula: 'СУММ(A1:A2)*2' })
    expect(by.C4).toMatchObject({ value: 0.05, text: '5%', formula: null })
    expect(by.C5).toMatchObject({ value: true, text: 'ИСТИНА' })
    expect(by.C6).toMatchObject({ value: '007', text: '007' })
  })

  it('keeps the cell style, note and link through an edit, and a clear keeps them too', async () => {
    const s = source()
    await s.edit([
      { sheet: 'Лист1', cell: 'B1', text: 'новый' },
      { sheet: 'Лист1', cell: 'B2', text: null },
      { sheet: 'Лист1', cell: 'C1', text: '7' },
    ])
    const w = await win(s, 'Лист1', 1, 2, 1, 3)
    expect(w.text[0]![1]).toBe('новый')
    expect(w.text[1]![1]).toBe('')
    expect(w.notes?.B1).toEqual({ author: 'Аня', text: 'Проверить' })
    expect(w.links?.B2).toEqual({ href: 'https://example.org/', tooltip: 'Пример' })
    expect(styleOf(w, 'C1')).toEqual(BOLD)
  })

  it('refuses what the book refuses: a formula in a table header', async () => {
    const got = await source().edit([{ sheet: 'Таблица', cell: 'B1', text: '=1+1' }])
    expect(got).toMatchObject({ ok: false, reason: 'refused' })
  })

  it('grows the sheet extent with an edit past it', async () => {
    const s = source()
    await s.edit([{ sheet: 'Список', cell: 'D9', text: '1' }])
    expect((await s.sheets())[2]).toMatchObject({ maxRow: 9, maxCol: 4 })
  })
})
