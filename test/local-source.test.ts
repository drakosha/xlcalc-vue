// @vitest-environment node
// localSource over a ru-RU workbook on the same thread: sheets, windows, search, the display filter, totals, events.
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetFilter, SourceEvent } from '../src/contract/source'
import type { CellView } from 'xlcalc'
import { SheetValues } from '../src/local/query'
import { BOLD, book, closeAll, COL_FILL, ROW_FILL, source, styleOf, watched, win } from './local-fixture'

/** Excel's width of a column without one, as xlcalc returns it. */
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
    // a location with an address is a location in what the address points to, not in this workbook
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
    // an empty cell, styled by its column or not, opens empty and is omitted
    expect(w.editText).toEqual({ A1: '15.03.2026', A2: '5%', A3: "'007", A4: '=СУММ(A1:A2)' })
  })

  it('gives the number format each cell shows by, as a code, its row\'s and column\'s on an empty cell too', async () => {
    const s = source({
      locale: 'ru-RU',
      sheets: [{ name: 'S', cells: { A1: { value: 46096, numberFormat: 14 }, A2: { value: 0.05, numberFormat: '0%' }, A3: { value: 5 }, B1: { value: 2 } }, columns: { B: { numberFormat: '0.00' } }, rows: { '4': { numberFormat: '0%' } } }],
    })
    const w = await win(s, 'S', 1, 4, 1, 2)
    // as Excel displays them: an empty cell by its column's format, by its row's where the two intersect
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
    expect(await s.find({ sheet: 'Лист1', query: 'С', lookIn: 'values' })).toEqual({
      hits: [
        { row: 1, col: 2 },
        { row: 2, col: 2 },
        { row: 3, col: 2 },
        { row: 4, col: 2 },
        { row: 5, col: 4 },
      ],
      total: 5,
    })
    expect(await s.find({ sheet: 'Лист1', query: '  ' })).toEqual({ hits: [], total: 0 })
  })

  it('filters and sorts the display, the head stays on top', async () => {
    const s = source()
    expect(await s.visibleRows('Список', { head: 1 })).toEqual({ rows: null })
    expect(await s.visibleRows('Список', { head: 1, cols: { 1: { values: ['б'] } } })).toEqual({ rows: [1, 2, 4] })
    expect(await s.visibleRows('Список', { head: 1, cols: { 1: { text: 'А' } } })).toEqual({ rows: [1, 3] })
    // numbers by value before text, empty cells last
    expect(await s.visibleRows('Список', { head: 1, sort: { col: 2, dir: 'asc' } })).toEqual({ rows: [1, 6, 2, 5, 3, 4] })
    expect(await s.visibleRows('Список', { head: 1, sort: { col: 1, dir: 'desc' } })).toEqual({ rows: [1, 6, 2, 4, 3, 5] })
  })

  it('sorts mixed values in Excel’s order, as memorySource: numbers, text in any case by en-US, FALSE, TRUE, errors, empty cells last either way', async () => {
    const cells = { A1: { value: 'k' }, A2: { value: 'b' }, A3: { value: 2 }, A4: { value: true }, A5: { formula: '1/0' }, A7: { value: 1 }, A8: { value: 'a' }, A9: { value: 'Item10' }, A10: { value: 'Item2' }, A11: { value: false } }
    const s = source({ sheets: [{ name: 'S', cells }] })
    expect((await s.visibleRows('S', { head: 1, sort: { col: 1, dir: 'asc' } })).rows).toEqual([1, 7, 3, 8, 2, 9, 10, 11, 4, 5, 6])
    expect((await s.visibleRows('S', { head: 1, sort: { col: 1, dir: 'desc' } })).rows).toEqual([1, 5, 4, 11, 10, 9, 2, 8, 3, 7, 6])
  })

  it('sorts text in the book’s locale: ru-RU’s order is en-US’s, Latin before Cyrillic, as Excel for Windows sorts', async () => {
    const cells = { A1: { value: 'h' }, A2: { value: 'б' }, A3: { value: 'b' }, A4: { value: 'А' }, A5: { value: 'a' } }
    const s = source({ locale: 'ru-RU', sheets: [{ name: 'S', cells }] })
    expect((await s.visibleRows('S', { head: 1, sort: { col: 1, dir: 'asc' } })).rows).toEqual([1, 5, 3, 4, 2])
    expect((await s.columnValues('S', 1, { head: 1 })).values.map((v) => v.text)).toEqual(['a', 'b', 'А', 'б'])
  })

  it('reads only the columns a filter, a sort or a funnel asks for, once, and keeps them up to date with edits', async () => {
    const { s, calls, count } = watched()
    const reads = () => calls.filter((c) => c.method === 'getCells').map((c) => c.args[1])
    const filter: SheetFilter = { head: 1, cols: { 1: { values: ['б'] } } }
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4])
    expect(reads()).toEqual([['A1:A6']])
    expect((await s.visibleRows('Список', { head: 1, sort: { col: 2, dir: 'asc' } })).rows).toEqual([1, 6, 2, 5, 3, 4])
    expect((await s.columnValues('Список', 1, filter)).values).toHaveLength(4)
    expect(reads()).toEqual([['A1:A6'], ['B1:B6']])
    await s.edit([{ sheet: 'Список', cell: 'A6', text: 'б' }])
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4, 6])
    // an edit beyond the sheet's end extends it, as Excel's used range grows
    await s.edit([{ sheet: 'Список', cell: 'A9', text: 'б' }])
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4, 6, 9])
    expect(await s.find({ sheet: 'Список', query: 'б', filter })).toMatchObject({ total: 4 })
    expect(reads()).toHaveLength(2)
    expect(count('getUsedRange')).toBe(1)
  })

  it('keeps a change that comes while its column is being read, and each kind of value', async () => {
    let release: (cells: Record<string, CellView>) => void = () => {}
    const v = new SheetValues(4, () => new Promise((resolve) => (release = resolve)))
    const loading = v.load([1])
    v.set(2, 1, 'new', 'new')
    const cell = (text: string, value: CellView['value']) => ({ text, value }) as CellView
    release({ A1: cell('h', 'h'), A2: cell('old', 'old'), A3: cell('ИСТИНА', true), A4: cell('#Н/Д', { error: 'NA' }) })
    await loading
    expect([1, 2, 3, 4].map((r) => v.textAt(r, 1))).toEqual(['h', 'new', 'ИСТИНА', '#Н/Д'])
    expect([1, 2, 3, 4, 5].map((r) => v.valueAt(r, 1))).toEqual(['h', 'new', true, { error: 'NA' }, null])
    v.set(7, 1, '1,5', 1.5)
    expect([v.maxRow, v.textAt(7, 1), v.valueAt(7, 1)]).toEqual([7, '1,5', 1.5])
  })

  it('stops a column read with the signal of the request it is for; one stopped for one request is read again for another', async () => {
    const reads: { signal?: AbortSignal; release: (cells: Record<string, CellView>) => void }[] = []
    const v = new SheetValues(2, (_ranges, signal) => new Promise((resolve, reject) => {
      reads.push({ signal, release: resolve })
      signal?.addEventListener('abort', () => reject(new Error('aborted by the session')))
    }))
    const ctrl = new AbortController()
    const funnel = v.load([1], ctrl.signal)
    const list = v.load([1])
    expect(reads.map((r) => r.signal)).toEqual([ctrl.signal])
    ctrl.abort()
    await expect(funnel).rejects.toMatchObject({ name: 'AbortError' })
    await new Promise((r) => setTimeout(r))
    expect(reads.map((r) => r.signal)).toEqual([ctrl.signal, undefined])
    reads[1]!.release({ A1: { text: 'h', value: 'h' } as CellView })
    await list
    expect(v.textAt(1, 1)).toBe('h')
  })

  it('passes a request’s signal to the session’s getCells, and an aborted one rejects as an abort', async () => {
    const { s, calls } = watched()
    const ctrl = new AbortController()
    await s.columnValues('Список', 1, { head: 1 }, ctrl.signal)
    expect(calls.filter((c) => c.method === 'getCells').map((c) => (c.args[2] as { signal?: AbortSignal } | undefined)?.signal)).toEqual([ctrl.signal])
    ctrl.abort()
    await expect(s.columnValues('Список', 2, { head: 1 }, ctrl.signal)).rejects.toMatchObject({ name: 'AbortError' })
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

  it('totals a range under a filter in the session, with its rows, without reading the cells', async () => {
    const { s, calls, count } = watched()
    const filter: SheetFilter = { head: 1, cols: { 1: { values: ['б'] } } }
    await s.visibleRows('Список', filter)
    const reads = count('getCells')
    expect(await s.stats('Список', 'B:B', filter)).toEqual({ cells: 3, numbers: 1, sum: 5, min: 5, max: 5 })
    expect(count('getCells')).toBe(reads)
    expect(calls.filter((c) => c.method === 'summarizeRanges').map((c) => (c.args[1] as { rows?: number[] }).rows)).toEqual([[1, 2, 4]])
  })

  it('totals as the grid counts its windows: a value its format hides counts, a formula giving "" does not', async () => {
    const s = source({ sheets: [{ name: 'S', cells: { A1: { value: 'k' }, A2: { value: 5, numberFormat: ';;;' }, A3: { value: 3 }, A4: { formula: '""' }, A5: { formula: '1/0' } } }] })
    const want = { cells: 3, numbers: 2, sum: 8, min: 3, max: 5 }
    expect(await s.stats('S', 'A2:A5')).toEqual(want)
    const filter: SheetFilter = { head: 1, sort: { col: 1, dir: 'asc' } }
    await s.visibleRows('S', filter)
    expect(await s.stats('S', 'A2:A5', filter)).toEqual(want)
    // the rows displaying nothing: the hidden number counts, the "" does not
    const blanks: SheetFilter = { head: 1, cols: { 1: { values: [''] } } }
    expect((await s.visibleRows('S', blanks)).rows).toEqual([1, 2, 4])
    expect(await s.stats('S', 'A2:A5', blanks)).toEqual({ cells: 1, numbers: 1, sum: 5, min: 5, max: 5 })
  })

  it('totals a range in the session without reading its cells, and stops when aborted', async () => {
    const s = source()
    await s.ready
    const ctrl = new AbortController()
    ctrl.abort()
    await expect(s.stats('Список', 'A2:B6', undefined, ctrl.signal)).rejects.toMatchObject({ name: 'AbortError' })
  })

  it('searches after an edit with the new text', async () => {
    const s = source()
    await s.edit([{ sheet: 'Список', cell: 'A5', text: 'бета' }])
    expect((await s.find({ sheet: 'Список', query: 'бет' })).hits).toEqual([{ row: 5, col: 1 }])
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

describe('localSource: Control and an arrow', () => {
  it('goes where Excel\'s Range.End goes, by the session\'s getEnd', async () => {
    const cells: Record<string, { value: number }> = {}
    for (let r = 1; r <= 10; r++) cells[`A${r}`] = { value: r }
    for (let r = 1; r <= 500; r++) cells[`B${r}`] = { value: r }
    const s = source({ sheets: [{ name: 'S', cells }] })
    expect(await s.end('S', 'A1', 'down')).toBe('A10')
    expect(await s.end('S', 'A10', 'down')).toBe('A1048576')
    expect(await s.end('S', 'A5', 'right')).toBe('B5')
  })
})

describe('localSource: the clipboard text (xlcalc 0.34 copyText)', () => {
  it('copies values by their formats without the column’s clipping, CRLF rows, none at the end for a Mac', async () => {
    const s = source({
      locale: 'ru-RU',
      sheets: [{ name: 'S', cells: { A1: { value: { date: '2026-03-15' }, numberFormat: 'dd.mm.yyyy' }, B1: { value: 'a\tb' }, A2: { value: 2 } }, columns: { A: { width: 2 } } }],
    })
    const w = await win(s, 'S', 1, 1, 1, 1)
    expect(w.text[0]![0]).toMatch(/^#+$/)
    expect(await s.copyText!('S', 'A1:B2')).toBe('15.03.2026\t"a\tb"\r\n2\t\r\n')
    expect(await s.copyText!('S', 'A2', { trailingNewline: false })).toBe('2')
  })
})
