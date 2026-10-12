// @vitest-environment node
// localSource over a ru-RU book in the same thread: sheets, windows, search, the display filter, totals, events.
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetFilter, SourceEvent } from '../src/source'
import { BOLD, book, closeAll, COL_FILL, ROW_FILL, source, styleOf, win } from './local-fixture'

/** Excel's width of a column without one, as xlcalc gives it. */
const DEFAULT_WIDTH = 8.83203125

afterEach(closeAll)

describe('localSource: sheets and windows', () => {
  it('lists the sheets with their extent and frozen panes', async () => {
    const list = await source().sheets()
    expect(list.map((x) => x.name)).toEqual(['Лист1', 'Таблица', 'Список'])
    expect(list[0]).toMatchObject({ maxRow: 5, maxCol: 7, freeze: { rows: 1, cols: 1 } })
    expect(list[1]).toMatchObject({ maxRow: 3, maxCol: 5 })
    expect(list[1]!.freeze).toBeUndefined()
  })

  it('takes in the extent the cells an array spilled into', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { formula: 'SEQUENCE(5,2)' } } }] })
    expect((await s.sheets())[0]).toMatchObject({ maxRow: 5, maxCol: 2 })
  })

  it('gives text, values and formulas in the book locale', async () => {
    const w = await win(source(), 'Лист1', 1, 5, 1, 7)
    expect(w.text[2]![0]).toBe('3')
    expect(w.values[0]![0]).toBe(1)
    expect(w.formulas).toEqual({ A3: 'СУММ(A1:A2)' })
    expect(w.text[0]![1]).toBe('с примечанием')
  })

  it('drops the table name in a cell of its own table, as Excel shows it', async () => {
    const w = await win(source(), 'Таблица', 1, 3, 1, 5)
    expect(w.formulas.C2).toBe('[@Сумма]*2')
    expect(w.formulas.E1).toBe('СУММ(Продажи[Сумма])')
    expect(w.text[1]![2]).toBe('20')
    expect(w.tables).toEqual([{ name: 'Продажи', range: 'A1:C3', headerRow: 1 }])
  })

  it('resolves styles: the cell own, else its row, else its column', async () => {
    const w = await win(source(), 'Лист1', 1, 5, 1, 7)
    expect(styleOf(w, 'C1')).toEqual(BOLD)
    expect(styleOf(w, 'D1')).toEqual(COL_FILL)
    expect(styleOf(w, 'A5')).toEqual(ROW_FILL)
    expect(styleOf(w, 'D5')).toEqual(ROW_FILL)
    expect(styleOf(w, 'A1')).toBeUndefined()
  })

  it('gives notes, links, widths, heights and merges', async () => {
    const w = await win(source(), 'Лист1', 1, 5, 1, 7)
    expect(w.notes).toEqual({ B1: { author: 'Аня', text: 'Проверить' } })
    // a place with an address is a place in what the address points to, not in this book
    expect(w.links).toEqual({ B2: { href: 'https://example.org/', tooltip: 'Пример' }, B3: { location: "'Таблица'!A1" }, B4: { href: 'https://example.org/page#top' } })
    // the library's widths: Excel's default where a column has none, 0 for a hidden one
    expect(w.widths).toEqual([DEFAULT_WIDTH, 20, DEFAULT_WIDTH, DEFAULT_WIDTH, 0, DEFAULT_WIDTH, DEFAULT_WIDTH])
    expect(w.rowHeights).toEqual({ 2: 30 })
    expect(w.merges).toEqual(['F1:G2'])
  })

  it('gives the text each cell opens with, as Excel\'s formula bar shows it', async () => {
    const s = source({
      locale: 'ru-RU',
      sheets: [{ name: 'S', cells: { A1: { value: { date: '2026-03-15' }, numberFormat: 14 }, A2: { value: 0.05, numberFormat: '0%' }, A3: { value: '007' }, A4: { formula: 'SUM(A1:A2)' } }, columns: { B: { style: 0 } } }],
      styles: [{ font: { bold: true } }],
    })
    const w = await win(s, 'S', 1, 5, 1, 2)
    // an empty cell, styled by its column or not, opens empty and is left out
    expect(w.editText).toEqual({ A1: '15.03.2026', A2: '5%', A3: "'007", A4: '=СУММ(A1:A2)' })
  })

  it('gives the number format each cell shows by, as a code, its row\'s and column\'s on an empty cell too', async () => {
    const s = source({
      locale: 'ru-RU',
      sheets: [{ name: 'S', cells: { A1: { value: 46096, numberFormat: 14 }, A2: { value: 0.05, numberFormat: '0%' }, A3: { value: 5 }, B1: { value: 2 } }, columns: { B: { numberFormat: '0.00' } }, rows: { '4': { numberFormat: '0%' } } }],
    })
    const w = await win(s, 'S', 1, 4, 1, 2)
    // as Excel shows them: an empty cell by its column's format, by its row's where the two cross
    expect(w.numberFormats).toEqual({ A1: 'dd.mm.yyyy', A2: '0%', B1: '0.00', B2: '0.00', B3: '0.00', A4: '0%', B4: '0%' })
  })

  it('keeps a window within maxCells', async () => {
    const s = source(book(), { maxCells: 10 })
    const w = await win(s, 'Лист1', 1, 50, 1, 5)
    expect(w.text.length).toBe(2)
    expect(s.maxCells).toBe(10)
  })
})

describe('localSource: search, display filter, totals', () => {
  it('finds shown text without regard to case, in row order', async () => {
    const s = source()
    expect(await s.find({ sheet: 'Лист1', query: 'С' })).toEqual({
      hits: [
        { row: 1, col: 2, text: 'с примечанием' },
        { row: 2, col: 2, text: 'ссылка' },
        { row: 3, col: 2, text: 'на лист' },
        { row: 4, col: 2, text: 'место' },
        { row: 5, col: 4, text: 'строка' },
      ],
      total: 5,
    })
    expect(await s.find({ sheet: 'Лист1', query: '  ' })).toEqual({ hits: [], total: 0 })
  })

  it('filters and sorts the display, the head stays on top', async () => {
    const s = source()
    expect(await s.visibleRows('Список', { head: 1 })).toEqual({ rows: null, total: 6 })
    expect(await s.visibleRows('Список', { head: 1, cols: { 1: { values: ['б'] } } })).toEqual({ rows: [1, 2, 4], total: 6 })
    expect(await s.visibleRows('Список', { head: 1, cols: { 1: { text: 'А' } } })).toEqual({ rows: [1, 3], total: 6 })
    // numbers by value before text, empty cells last
    expect(await s.visibleRows('Список', { head: 1, sort: { col: 2, dir: 'asc' } })).toEqual({ rows: [1, 6, 2, 5, 3, 4], total: 6 })
    expect(await s.visibleRows('Список', { head: 1, sort: { col: 1, dir: 'desc' } })).toEqual({ rows: [1, 6, 2, 4, 3, 5], total: 6 })
  })

  it('gives a filtered window by display index', async () => {
    const s = source()
    const filter: SheetFilter = { head: 1, cols: { 1: { values: ['б'] } } }
    await s.visibleRows('Список', filter)
    const w = await win(s, 'Список', 2, 5, 1, 2, filter)
    expect(w.row).toBe(2)
    expect(w.text).toEqual([['б', '5'], ['б', 'нет']])
  })

  it('gives a sorted window in display order', async () => {
    const s = source()
    const filter: SheetFilter = { head: 1, sort: { col: 1, dir: 'desc' } }
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 6, 2, 4, 3, 5])
    const w = await win(s, 'Список', 2, 3, 1, 2, filter)
    expect(w.text).toEqual([['в', '1,5'], ['б', '5'], ['б', 'нет']])
  })

  it('gives under a filter the merged areas of one row among the rows shown, as Excel shows them', async () => {
    const s = source({
      sheets: [{ name: 'S', cells: { A1: { value: 'k' }, A2: { value: 'x' }, A3: { value: 'y' }, A4: { value: 'x' } }, merges: ['C2:D2', 'C3:D4', 'E3:F3', 'B4:B5'] }],
    })
    const filter: SheetFilter = { head: 1, cols: { 1: { values: ['x'] } } }
    expect((await s.visibleRows('S', filter)).rows).toEqual([1, 2, 4])
    const w = await win(s, 'S', 1, 3, 1, 6, filter)
    expect(w.merges).toEqual(['C2:D2'])
    expect((await win(s, 'S', 1, 5, 1, 6)).merges.sort()).toEqual(['B4:B5', 'C2:D2', 'C3:D4', 'E3:F3'])
  })

  it('lists a column values under the other columns conditions', async () => {
    const s = source()
    expect(await s.columnValues('Список', 1, { head: 1 })).toEqual({
      values: [{ text: 'а', count: 1 }, { text: 'б', count: 2 }, { text: 'в', count: 1 }, { text: '', count: 1 }],
      more: 0,
    })
    expect(await s.columnValues('Список', 2, { head: 1, cols: { 1: { values: ['б'] }, 2: { values: ['5'] } } })).toEqual({
      values: [{ text: '5', count: 1 }, { text: 'нет', count: 1 }],
      more: 0,
    })
  })

  it('totals a range, under a filter only its visible rows', async () => {
    const s = source()
    expect(await s.stats('Список', 'A2:B6')).toEqual({ cells: 9, numbers: 4, sum: 43.5, min: 1.5, max: 30 })
    const filter: SheetFilter = { head: 1, cols: { 1: { values: ['б'] } } }
    await s.visibleRows('Список', filter)
    expect(await s.stats('Список', 'B:B', filter)).toEqual({ cells: 3, numbers: 1, sum: 5, min: 5, max: 5 })
  })

  it('searches after an edit with the new text', async () => {
    const s = source()
    await s.edit([{ sheet: 'Список', cell: 'A5', text: 'бета' }])
    expect((await s.find({ sheet: 'Список', query: 'бет' })).hits).toEqual([{ row: 5, col: 1, text: 'бета' }])
  })
})

describe('localSource: explain, events, closing', () => {
  it('explains a cell as text', async () => {
    const text = await source().explain('Лист1', 'A3')
    expect(text).toContain('A1')
    expect(text).toContain('СУММ')
  })

  it('applies edits one after another: the second reads the cell as the first left it', async () => {
    const s = source()
    const [a, b] = await Promise.all([s.format('Лист1', [{ range: 'A1', style: BOLD }]), s.edit([{ sheet: 'Лист1', cell: 'A1', text: '9' }])])
    expect(a.ok && b.ok).toBe(true)
    const w = await win(s, 'Лист1', 1, 3, 1, 1)
    expect(w.text[0]![0]).toBe('9')
    expect(styleOf(w, 'A1')).toEqual(BOLD)
  })

  it('tells subscribers of edits made through it', async () => {
    const s = source()
    const events: SourceEvent[] = []
    const off = s.subscribe((e) => events.push(e))
    await s.edit([{ sheet: 'Лист1', cell: 'A2', text: '5' }])
    off()
    await s.edit([{ sheet: 'Лист1', cell: 'A2', text: '6' }])
    expect(events).toHaveLength(1)
    expect(events[0]).toMatchObject({ kind: 'changed', changes: [{ cell: 'A2', value: 5 }, { cell: 'A3', value: 6 }] })
  })

  it('fails the opening of a book in a locale xlcalc has no module for, with the library message', async () => {
    const s = source({ locale: 'de-DE', sheets: [{ name: 'S', cells: {} }] })
    await expect(s.ready).rejects.toThrow(/^book\.locale: unknown locale "de-DE" \(loaded: .*; the package's: en-US, ru-RU\)$/)
  })

  it('rejects every call when the book cannot be opened', async () => {
    const s = source({ sheets: [{ name: 'a/b', cells: {} }] })
    await expect(s.ready).rejects.toThrow()
    await expect(s.sheets()).rejects.toThrow()
  })
})
