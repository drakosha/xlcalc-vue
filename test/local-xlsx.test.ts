// @vitest-environment node
// xlsx in and out of the local source: a book written and read back, a Blob, the reader's issues grouped.
import { afterEach, describe, expect, it } from 'vitest'
import { groupIssues, openXlsx, saveXlsx } from '../src/local'
import { BOLD, closeAll, source, styleOf, win } from './local-fixture'

afterEach(closeAll)

describe('xlsx in and out', () => {
  it('writes the book and reads it back with the same values, styles, notes and frozen panes', async () => {
    const s = source()
    await s.edit([{ sheet: 'Лист1', cell: 'A1', text: '1,5' }])
    const { data, issues } = await saveXlsx(s)
    expect(data.byteLength).toBeGreaterThan(1000)
    expect(issues).toEqual([])
    const back = await openXlsx(data)
    expect(back.book.locale).toBe('ru-RU')
    const t = source(back.book)
    const w = await win(t, 'Лист1', 1, 5, 1, 7)
    expect(w.text.map((r) => r[0])).toEqual(['1,5', '2', '3,5', '', ''])
    expect(w.formulas.A3).toBe('СУММ(A1:A2)')
    expect(styleOf(w, 'C1')).toMatchObject(BOLD)
    expect(w.notes?.B1).toEqual({ author: 'Аня', text: 'Проверить' })
    expect((await t.sheets())[0]!.freeze).toEqual({ rows: 1, cols: 1 })
    expect((await win(t, 'Таблица', 2, 1, 3, 1)).formulas.C2).toBe('[@Сумма]*2')
  })

  it('opens a Blob as a file input gives it', async () => {
    const { data } = await saveXlsx(source())
    const back = await openXlsx(new Blob([data as BlobPart]))
    expect(back.book.sheets.map((x) => x.name)).toEqual(['Лист1', 'Таблица', 'Список'])
    expect(back.issues).toEqual([])
  })

  it('groups what the reader left out by kind, for a notice', () => {
    expect(
      groupIssues([
        { code: 'drawing', message: 'a chart', sheet: 'A' },
        { code: 'conditionalFormat', message: 'a rule', sheet: 'A', cells: ['B2:B9'] },
        { code: 'drawing', message: 'a picture', sheet: 'B' },
        { code: 'drawing', message: 'a chart', sheet: 'A' },
      ]),
    ).toEqual([
      { code: 'drawing', count: 3, sheets: ['A', 'B'], message: 'a chart' },
      { code: 'conditionalFormat', count: 1, sheets: ['A'], message: 'a rule' },
    ])
  })
})
