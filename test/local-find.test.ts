// @vitest-environment node
// localSource's Find and Replace as Excel's, by xlcalc's `findCells`: in the text a cell opens with (formulas in the
// book's locale, values by the format they show) or in the shown text, hidden cells skipped as Find skips them, hits in
// Find's order; under the display filter only its rows. Replace All in one step of undo.
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AsyncSession, BookInput, CellInput } from 'xlcalc'
import { findHits } from '../src/local/find'
import type { FindRequest, SourceEvent } from '../src/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

type Query = Partial<FindRequest> & { query: string }

/** Hits as `row:col text`. */
async function found(s: ReturnType<typeof source>, req: Query & { sheet: string }): Promise<string[]> {
  return (await s.find(req)).hits.map((h) => `${h.row}:${h.col} ${h.text}`)
}
/** Hits on sheet S as `row:col`. */
async function at(s: ReturnType<typeof source>, req: Query): Promise<string[]> {
  return (await s.find({ sheet: 'S', ...req })).hits.map((h) => `${h.row}:${h.col}`)
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

/** Texts where Find's own rules show: a character outside the BMP, a text that reads as a number, Greek sigmas, a dotted capital I. */
const letters = (): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { value: 'a😀b' }, B1: { value: '007' }, C1: { value: 'ΣΑΣ' }, D1: { value: 'ς' }, E1: { value: 'İx' } } }],
})

/** `x` in every cell of A1:C3 with row 2 and column B hidden; `filtered`: the sheet has a filter's criteria. */
const hidden = (filtered: boolean): BookInput => ({
  sheets: [
    {
      name: 'S',
      cells: Object.fromEntries(['A', 'B', 'C'].flatMap((c) => [1, 2, 3].map((r): [string, CellInput] => [`${c}${r}`, { value: 'x' }]))),
      rows: { '2': { hidden: true } },
      columns: { B: { hidden: true } },
      ...(filtered ? { filtered: true } : {}),
    },
  ],
})

describe('localSource: find', () => {
  it('looks in formulas by default, as Excel: the text each cell opens with, formulas in the book’s locale', async () => {
    const s = source()
    expect(await found(s, { sheet: 'Лист1', query: 'С' })).toEqual(['1:2 с примечанием', '2:2 ссылка', '3:1 =СУММ(A1:A2)', '3:2 на лист', '4:2 место', '5:4 строка'])
    expect(await found(s, { sheet: 'Лист1', query: 'С', lookIn: 'values' })).toEqual(['1:2 с примечанием', '2:2 ссылка', '3:2 на лист', '4:2 место', '5:4 строка'])
    expect(await found(s, { sheet: 'Таблица', query: '[@' })).toEqual(['2:3 =Продажи[@Сумма]*2', '3:3 =Продажи[@Сумма]*2'])
  })

  it('gives as the hit’s text the shown text, or in formulas the formula in the book’s locale', async () => {
    const s = source(formats())
    expect(await found(s, { sheet: 'S', query: '1250,5' })).toEqual(['1:2 1 250,50'])
    expect(await found(s, { sheet: 'S', query: '1250,5', lookIn: 'values' })).toEqual(['1:5 1250,55'])
    expect(await found(s, { sheet: 'S', query: 'сумм' })).toEqual(['1:5 =СУММ(B1:C1)'])
  })

  it('reads a value as its edit text in formulas and as shown in values: dates by the column’s format, numbers whole', async () => {
    const s = source(formats())
    expect(await at(s, { query: '15.03' })).toEqual(['1:1'])
    expect(await at(s, { query: '1250,5' })).toEqual(['1:2'])
    expect(await at(s, { query: '1250,5', lookIn: 'values' })).toEqual(['1:5'])
    expect(await at(s, { query: '250,50', lookIn: 'values' })).toEqual(['1:2'])
    expect(await at(s, { query: '5%' })).toEqual(['1:3'])
    expect(await at(s, { query: '=сумм(*)', wholeCell: true })).toEqual(['1:5'])
  })

  it('reads the query as Excel’s Find: `?` is one UTF-16 unit, a text has no apostrophe, σ ς Σ are one letter, i is not İ', async () => {
    const s = source(letters())
    expect(await at(s, { query: 'a??b' })).toEqual(['1:1'])
    expect(await at(s, { query: 'a?b' })).toEqual([])
    expect(await at(s, { query: '007', wholeCell: true })).toEqual(['1:2'])
    expect(await at(s, { query: "'0" })).toEqual([])
    expect(await at(s, { query: 'σ' })).toEqual(['1:3', '1:4'])
    expect(await at(s, { query: 'σ', matchCase: true })).toEqual([])
    expect(await at(s, { query: 'i' })).toEqual([])
  })

  it('skips hidden cells as Find: in values always, in formulas on a sheet with a filter, there by the blocks hidden lines leave', async () => {
    const plain = source(hidden(false))
    expect(await at(plain, { query: 'x', lookIn: 'values' })).toEqual(['1:1', '1:3', '3:1', '3:3'])
    expect(await at(plain, { query: 'x', lookIn: 'values', byColumns: true })).toEqual(['1:1', '3:1', '1:3', '3:3'])
    expect(await at(plain, { query: 'x' })).toEqual(['1:1', '1:2', '1:3', '2:1', '2:2', '2:3', '3:1', '3:2', '3:3'])
    const filtered = source(hidden(true))
    expect(await at(filtered, { query: 'x' })).toEqual(['1:1', '1:3', '3:1', '3:3'])
    expect(await at(filtered, { query: 'x', byColumns: true })).toEqual(['1:1', '1:3', '3:1', '3:3'])
  })

  it('takes the range, the order by columns and the display filter', async () => {
    const s = source()
    expect(await found(s, { sheet: 'Список', query: '*', range: 'A3:B4' })).toEqual(['3:1 а', '3:2 30', '4:1 б', '4:2 нет'])
    expect(await found(s, { sheet: 'Список', query: '*', range: 'A3:B4', byColumns: true })).toEqual(['3:1 а', '4:1 б', '3:2 30', '4:2 нет'])
    const filter = { head: 1, cols: { 1: { values: ['б'] } } }
    await s.visibleRows('Список', filter)
    expect(await found(s, { sheet: 'Список', query: '?', wholeCell: true, filter })).toEqual(['2:1 б', '2:2 5', '4:1 б'])
  })

  it('counts under the display filter only its rows, over the library’s pages, in Find’s order', async () => {
    // 12,001 hits in column A, more than a page; the filter keeps the head and every third row
    const cells: Record<string, CellInput> = {}
    for (let r = 1; r <= 12_001; r++) {
      cells[`A${r}`] = { value: 'x' }
      cells[`B${r}`] = { value: r % 3 ? 'a' : 'b' }
    }
    const s = source({ sheets: [{ name: 'S', cells }] })
    const filter = { head: 1, cols: { 2: { values: ['b'] } } }
    await s.visibleRows('S', filter)
    const got = await s.find({ sheet: 'S', query: 'x', filter })
    expect(got.total).toBe(4001)
    expect(got.hits.length).toBe(500)
    expect(got.hits.slice(0, 3).map((h) => [h.row, h.col])).toEqual([[1, 1], [3, 1], [6, 1]])
    expect(got.hits.at(-1)).toMatchObject({ row: 1497, col: 1 })
  })

  it('starts a search over pages again when an edit between two pages leaves the cursor stale', async () => {
    const page = (cells: string[], cursor?: string) => ({ revision: 1, clock: {}, truncated: !!cursor, cells: cells.map((cell) => ({ sheet: 'S', cell })), ...(cursor ? { cursor } : {}) })
    const stale = Object.assign(new Error('the cursor belongs to another revision'), { name: 'CalcError', code: 'stale_revision' })
    const answers = [page(['A1', 'A2'], 'c1'), stale, page(['A1', 'A3'], 'c2'), page(['A4'])]
    const findCells = vi.fn(async () => {
      const next = answers.shift()!
      if (next instanceof Error) throw next
      return next
    })
    const got = await findHits({ findCells } as unknown as AsyncSession, { sheet: 'S', query: 'x' }, 'en-US', new Set([1, 3, 4]))
    expect(got).toEqual({ hits: [{ row: 1, col: 1, text: '' }, { row: 3, col: 1, text: '' }, { row: 4, col: 1, text: '' }], total: 3 })
    expect(findCells.mock.calls.map((c) => (c as unknown[])[2])).toMatchObject([{ maxCells: 10_000 }, { cursor: 'c1' }, { maxCells: 10_000 }, { cursor: 'c2' }])
  })

  it('finds the cells as they are after edits, a format and an undo', async () => {
    const s = source(formats())
    expect(await at(s, { query: 'x' })).toEqual(['1:6'])
    await s.edit([{ sheet: 'S', cell: 'F1', text: 'y' }, { sheet: 'S', cell: 'F2', text: '=F1&"x"' }])
    expect(await at(s, { query: 'x' })).toEqual(['2:6'])
    // a format changes how a number opens
    await s.format('S', [{ range: 'B1', numberFormat: '0%' }])
    expect(await at(s, { query: '125050%' })).toEqual(['1:2'])
    await s.undo()
    expect(await at(s, { query: '1250,5' })).toEqual(['1:2'])
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

  it('replaces only in the cells Find finds: a text without its apostrophe, and not where our own match finds nothing', async () => {
    const s = source(letters())
    expect(await s.replace({ sheet: 'S', query: '007', wholeCell: true }, '8')).toMatchObject({ ok: true, replaced: 1 })
    expect((await win(s, 'S', 1, 1, 2, 1)).values).toEqual([['8']])
    // `a?b` is not Find's hit; `a??b` is, but the replacement's own pattern reads `?` as a whole character (TODO(xlcalc))
    expect(await s.replace({ sheet: 'S', query: 'a?b' }, 'y')).toEqual({ ok: true, changes: [], replaced: 0 })
    expect(await s.replace({ sheet: 'S', query: 'a??b' }, 'y')).toEqual({ ok: true, changes: [], replaced: 0 })
    expect((await win(s, 'S', 1, 1, 1, 1)).values).toEqual([['a😀b']])
  })

  it('replaces hidden cells as Excel’s Replace finds them, in formulas: all on a sheet without a filter, the visible ones with it', async () => {
    expect(await source(hidden(false)).replace({ sheet: 'S', query: 'x', lookIn: 'values' }, 'y')).toMatchObject({ ok: true, replaced: 9 })
    const filtered = source(hidden(true))
    expect(await filtered.replace({ sheet: 'S', query: 'x', lookIn: 'values' }, 'y')).toMatchObject({ ok: true, replaced: 4 })
    expect((await win(filtered, 'S', 1, 3, 1, 3)).values).toEqual([['y', 'x', 'y'], ['x', 'x', 'x'], ['y', 'x', 'y']])
  })

  it('replaces under the display filter only in its rows', async () => {
    const s = source()
    const filter = { head: 1, cols: { 1: { values: ['б'] } } }
    await s.visibleRows('Список', filter)
    // 5 in B2 and 1,5 in B6, whose row the filter hides
    expect(await s.replace({ sheet: 'Список', query: '5', filter }, '6')).toMatchObject({ ok: true, replaced: 1 })
    expect((await win(s, 'Список', 2, 5, 2, 1)).values).toEqual([[6], [30], ['нет'], [7], [1.5]])
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
