// @vitest-environment node
// localSource edits: typed text read in the book's locale, refusals, what the edited cells show after.
import { afterEach, describe, expect, it } from 'vitest'
import { BOLD, closeAll, source, styleOf, win } from './local-fixture'

afterEach(closeAll)

describe('localSource: typed edits', () => {
  it('reads 1,5 as a number in ru-RU and sends the dependent cell', async () => {
    const got = await source().edit([{ sheet: 'Лист1', cell: 'A1', text: '1,5' }])
    expect(got.ok).toBe(true)
    if (!got.ok) return
    expect(got.changes).toContainEqual({ sheet: 'Лист1', cell: 'A1', text: '1,5', value: 1.5, formula: null })
    expect(got.changes.find((c) => c.cell === 'A3')).toMatchObject({ value: 3.5, text: '3,5' })
  })

  it('reads 15.03.2026 as a date with the date format', async () => {
    const s = source()
    // B is 20 wide: in a column of the default width the date shows as ### (the library measures it so)
    const got = await s.edit([{ sheet: 'Лист1', cell: 'B4', text: '15.03.2026' }])
    expect(got).toMatchObject({ ok: true, changes: [{ cell: 'B4', value: 46096, text: '15.03.2026' }] })
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
