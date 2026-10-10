// @vitest-environment node
// localSource over a ru-RU book in the same thread: windows, typed edits, structure, format, layout, queries, events, xlsx.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import { groupIssues, localSource, openXlsx, saveXlsx, type LocalSource } from '../src/local'
import type { SheetFilter, SheetWindow, SourceEvent } from '../src/source'

const BOLD = { font: { bold: true } }
const ROW_FILL = { fill: { color: '#ffeeaa' } }
const COL_FILL = { fill: { color: '#ddeeff' } }

/** A small book in ru-RU: values, a sum, styles of cells, rows and columns, a note, a link, a table, a list to filter. */
const book = (): BookInput => ({
  locale: 'ru-RU',
  styles: [{}, BOLD, ROW_FILL, COL_FILL],
  sheets: [
    {
      name: 'Лист1',
      cells: {
        A1: { value: 1 },
        A2: { value: 2 },
        A3: { formula: 'SUM(A1:A2)' },
        B1: { value: 'с примечанием', note: { author: 'Аня', text: 'Проверить' } },
        B2: { value: 'ссылка', link: { url: 'https://example.org/', tooltip: 'Пример' } },
        B3: { value: 'на лист', link: { location: "'Таблица'!A1" } },
        C1: { value: 'жирный', style: 1 },
        D5: { value: 'строка' },
      },
      columns: { B: { width: 20 }, D: { style: 3 }, E: { hidden: true } },
      rows: { '2': { height: 30 }, '5': { style: 2 } },
      merges: ['F1:G2'],
      frozen: { rows: 1, columns: 1 },
    },
    {
      name: 'Таблица',
      cells: {
        A1: { value: 'Товар' },
        B1: { value: 'Сумма' },
        C1: { value: 'Двойная' },
        A2: { value: 'чай' },
        B2: { value: 10 },
        C2: { formula: 'Продажи[[#This Row],[Сумма]]*2' },
        A3: { value: 'кофе' },
        B3: { value: 20 },
        C3: { formula: 'Продажи[[#This Row],[Сумма]]*2' },
        E1: { formula: 'SUM(Продажи[Сумма])' },
      },
      tables: [{ name: 'Продажи', range: 'A1:C3', columns: [{ name: 'Товар' }, { name: 'Сумма' }, { name: 'Двойная', formula: 'Продажи[[#This Row],[Сумма]]*2' }] }],
    },
    {
      name: 'Список',
      cells: {
        A1: { value: 'Имя' },
        B1: { value: 'Число' },
        A2: { value: 'б' },
        B2: { value: 5 },
        A3: { value: 'а' },
        B3: { value: 30 },
        A4: { value: 'б' },
        B4: { value: 'нет' },
        B5: { value: 7 },
        A6: { value: 'в' },
        B6: { value: 1.5 },
      },
    },
  ],
})

let open: LocalSource[] = []
function source(b: BookInput = book()): LocalSource {
  const s = localSource(b, { calc: 'inline' })
  open.push(s)
  return s
}
afterEach(async () => {
  await Promise.all(open.map((s) => s.close()))
  open = []
})

async function win(s: LocalSource, sheet: string, row: number, rows: number, col: number, cols: number, filter?: SheetFilter): Promise<SheetWindow> {
  const w = await s.window({ sheet, row, rows, col, cols, ...(filter ? { filter } : {}) })
  if ('pending' in w) throw new Error('pending')
  return w
}
const styleOf = (w: SheetWindow, addr: string) => (w.styleIds[addr] === undefined ? undefined : w.styles[w.styleIds[addr]!])

describe('localSource: sheets and windows', () => {
  it('lists the sheets with their extent and frozen panes', async () => {
    const list = await source().sheets()
    expect(list.map((x) => x.name)).toEqual(['Лист1', 'Таблица', 'Список'])
    expect(list[0]).toMatchObject({ maxRow: 5, maxCol: 7, freeze: { rows: 1, cols: 1 } })
    expect(list[1]).toMatchObject({ maxRow: 3, maxCol: 5 })
    expect(list[1]!.freeze).toBeUndefined()
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
    expect(w.links).toEqual({ B2: { href: 'https://example.org/', tooltip: 'Пример' }, B3: { location: "'Таблица'!A1" } })
    expect(w.widths).toEqual([null, 20, null, null, 0, null, null])
    expect(w.rowHeights).toEqual({ 2: 30 })
    expect(w.merges).toEqual(['F1:G2'])
  })

  it('keeps a window within maxCells', async () => {
    const s = localSource(book(), { calc: 'inline', maxCells: 10 })
    open.push(s)
    const w = await win(s, 'Лист1', 1, 50, 1, 5)
    expect(w.text.length).toBe(2)
    expect(s.maxCells).toBe(10)
  })
})

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
    const got = await s.edit([{ sheet: 'Лист1', cell: 'A4', text: '15.03.2026' }])
    expect(got).toMatchObject({ ok: true, changes: [{ cell: 'A4', value: 46096, text: '15.03.2026' }] })
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
    expect(refused.ok ? '' : refused.message).toMatch(/name/i)
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
    let got = await s.format('Лист1', 'A1:A2', { style: italic, format: '0.00' })
    expect(got.ok).toBe(true)
    let w = await win(s, 'Лист1', 1, 3, 1, 1)
    expect(styleOf(w, 'A1')).toEqual(italic)
    expect(styleOf(w, 'A2')).toEqual(italic)
    expect(w.text.map((r) => r[0])).toEqual(['1,00', '2,00', '3'])
    expect(w.formulas.A3).toBe('СУММ(A1:A2)')
    got = await s.format('Лист1', 'A1', { style: null, format: null })
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

describe('localSource: search, display filter, totals', () => {
  it('finds shown text without regard to case, in row order', async () => {
    const s = source()
    expect(await s.find('Лист1', 'С')).toEqual({
      hits: [
        { row: 1, col: 2, text: 'с примечанием' },
        { row: 2, col: 2, text: 'ссылка' },
        { row: 3, col: 2, text: 'на лист' },
        { row: 5, col: 4, text: 'строка' },
      ],
      total: 4,
    })
    expect(await s.find('Лист1', '  ')).toEqual({ hits: [], total: 0 })
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
    expect((await s.find('Список', 'бет')).hits).toEqual([{ row: 5, col: 1, text: 'бета' }])
  })
})

describe('localSource: explain, events, closing', () => {
  it('explains a cell as text', async () => {
    const text = await source().explain('Лист1', 'A3')
    expect(text).toContain('A1')
    expect(text).toContain('SUM')
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

  it('rejects every call when the book cannot be opened', async () => {
    const s = localSource({ sheets: [{ name: 'a/b', cells: {} }] }, { calc: 'inline' })
    open.push(s)
    await expect(s.ready).rejects.toThrow()
    await expect(s.sheets()).rejects.toThrow()
  })
})

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
