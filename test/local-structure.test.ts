// @vitest-environment node
// localSource changes of the book's shape and look: structure operations, widths and heights, formats, frozen panes.
import { afterEach, describe, expect, it } from 'vitest'
import type { SourceEvent } from '../src/source'
import { BOLD, closeAll, ROW_FILL, source, styleOf, win } from './local-fixture'

afterEach(closeAll)

describe('localSource: structure, layout, format', () => {
  it('inserts a row: formulas move, the window reads the moved book, subscribers get reload', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const got = await s.structure({ op: 'insertRows', sheet: 'Лист1', row: 2 })
    expect(got.ok).toBe(true)
    const w = await win(s, 'Лист1', 1, 6, 1, 7)
    expect(w.formulas).toEqual({ A4: 'СУММ(A1:A3)' })
    expect(w.rowHeights).toEqual({ 3: 30 })
    expect(w.notes).toEqual({ B1: { author: 'Аня', text: 'Проверить' } })
    expect(styleOf(w, 'A6')).toEqual(ROW_FILL)
    expect((await s.sheets())[0]).toMatchObject({ maxRow: 6 })
    expect(events).toEqual([{ kind: 'reload' }])
  })

  it('renames a sheet and refuses a taken name with the library message', async () => {
    const s = source()
    expect(await s.structure({ op: 'renameSheet', sheet: 'Список', name: 'Люди' })).toMatchObject({ ok: true })
    expect((await s.sheets()).map((x) => x.name)).toEqual(['Лист1', 'Таблица', 'Люди'])
    const refused = await s.structure({ op: 'renameSheet', sheet: 'Люди', name: 'таблица' })
    expect(refused).toMatchObject({ ok: false, reason: 'refused' })
    // the library's message, in the book's locale
    expect(refused.ok ? '' : refused.message).toMatch(/таблица/)
  })

  it('sets a column width and a row height in the book', async () => {
    const s = source()
    expect(await s.layout({ sheet: 'Лист1', columns: 'C', width: 15 })).toMatchObject({ ok: true })
    expect(await s.layout({ sheet: 'Лист1', columns: 'E', width: 9 })).toMatchObject({ ok: true })
    expect(await s.layout({ sheet: 'Лист1', rows: '3:4', height: 0 })).toMatchObject({ ok: true })
    const w = await win(s, 'Лист1', 1, 5, 1, 5)
    expect(w.widths).toEqual([null, 20, 15, null, 9])
    expect(w.rowHeights).toEqual({ 2: 30, 3: 0, 4: 0 })
    const b = await s.book()
    expect(b.sheets[0]!.columns?.C).toEqual({ width: 15 })
  })

  it('formats a range: a new style in the table, a number format, and null removes them', async () => {
    const s = source()
    const italic = { font: { italic: true } }
    let got = await s.format('Лист1', 'A1:A2', { style: italic, numberFormat: '0.00' })
    expect(got.ok).toBe(true)
    let w = await win(s, 'Лист1', 1, 3, 1, 1)
    expect(styleOf(w, 'A1')).toEqual(italic)
    expect(styleOf(w, 'A2')).toEqual(italic)
    expect(w.text.map((r) => r[0])).toEqual(['1,00', '2,00', '3'])
    expect(w.formulas.A3).toBe('СУММ(A1:A2)')
    got = await s.format('Лист1', 'A1', { style: null, numberFormat: null })
    expect(got.ok).toBe(true)
    w = await win(s, 'Лист1', 1, 2, 1, 1)
    expect(styleOf(w, 'A1')).toBeUndefined()
    expect(w.text[0]![0]).toBe('1')
    expect(styleOf(w, 'A2')).toEqual(italic)
  })

  it('formats empty cells and whole columns, and keeps notes and formulas', async () => {
    const s = source()
    await s.format('Лист1', 'H1:H2', { style: BOLD })
    await s.format('Лист1', 'B:B', { style: ROW_FILL })
    const w = await win(s, 'Лист1', 1, 9, 1, 8)
    expect(styleOf(w, 'H1')).toEqual(BOLD)
    expect(styleOf(w, 'B1')).toEqual(ROW_FILL)
    expect(styleOf(w, 'B9')).toEqual(ROW_FILL)
    expect(w.notes?.B1).toBeDefined()
    expect(w.links?.B2).toBeDefined()
  })

  it('freezes panes in the book', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.freeze('Список', { rows: 1, cols: 0 })
    await s.freeze('Лист1', null)
    const list = await s.sheets()
    expect(list[2]!.freeze).toEqual({ rows: 1, cols: 0 })
    expect(list[0]!.freeze).toBeUndefined()
    expect(events).toEqual([{ kind: 'reload' }, { kind: 'reload' }])
  })
})
