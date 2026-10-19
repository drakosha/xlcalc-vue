// @vitest-environment node
// localSource's Find and Replace as in Excel, by xlcalc's `findCells`: in the text a cell opens with (formulas in the
// workbook's locale, values by the format they are displayed with) or in the displayed text, hidden cells skipped as Find
// skips them, matches in Find's order; under the display filter only its rows. Replace All as a single undo step.
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput, CellInput } from 'xlcalc'
import { cellAddress } from 'xlcalc/a1'
import type { FindRequest, SourceEvent } from '../src/contract/source'
import { closeAll, source, win } from './local-fixture'

afterEach(closeAll)

type Query = Partial<FindRequest> & { query: string }

/** Matches as `row:col text`, with the text Find read: the displayed text, or in formulas the text the cell opens with. */
async function found(s: ReturnType<typeof source>, req: Query & { sheet: string }): Promise<string[]> {
  const hits = (await s.find(req)).hits
  return Promise.all(
    hits.map(async (h) => {
      const w = await win(s, req.sheet, h.row, 1, h.col, 1)
      const shown = w.text[0]![0]!
      return `${h.row}:${h.col} ${req.lookIn === 'values' ? shown : (w.editText?.[cellAddress(h.row, h.col)] ?? shown)}`
    }),
  )
}
/** Matches on sheet S as `row:col`. */
async function at(s: ReturnType<typeof source>, req: Query): Promise<string[]> {
  return (await s.find({ sheet: 'S', ...req })).hits.map((h) => `${h.row}:${h.col}`)
}

/** Numbers by their formats, a date by its column's format, text that parses as a number, a formula. */
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

/** Texts where Find's own rules apply: a character outside the BMP, text that parses as a number, Greek sigmas, a dotted capital I. */
const letters = (): BookInput => ({
  sheets: [{ name: 'S', cells: { A1: { value: 'a😀b' }, B1: { value: '007' }, C1: { value: 'ΣΑΣ' }, D1: { value: 'ς' }, E1: { value: 'İx' } } }],
})

/** `x` in every cell of A1:C3 with row 2 and column B hidden; `filtered`: the sheet has filter criteria. */
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
    expect(await found(s, { sheet: 'Таблица', query: '[@' })).toEqual(['2:3 =[@Сумма]*2', '3:3 =[@Сумма]*2'])
  })

  it('reads the shown text, or in formulas the text the cell opens with', async () => {
    const s = source(formats())
    expect(await found(s, { sheet: 'S', query: '1250,5' })).toEqual(['1:2 1250,5'])
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

  it('counts under the display filter only its rows, in Find’s order', async () => {
    // 12,001 matches in column A, more than a page; the filter keeps the header and every third row
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

  it('goes on from a cell, as Find Next does: from the first hit after it, round past the end, the cell itself last', async () => {
    const s = source(hidden(false))
    const order = async (req: Partial<FindRequest>) => (await at(s, { ...req, query: 'x' })).join(' ')
    expect(await order({ after: 'B2' })).toBe('2:3 3:1 3:2 3:3 1:1 1:2 1:3 2:1 2:2')
    expect(await order({ after: 'B2', byColumns: true })).toBe('3:2 1:3 2:3 3:3 1:1 2:1 3:1 1:2 2:2')
    // a cell outside the searched range: from the range's first cell
    expect(await order({ after: 'E9', range: 'A1:B2' })).toBe('1:1 1:2 2:1 2:2')
  })

  it('finds the cells as they are after edits, a format and an undo', async () => {
    const s = source(formats())
    expect(await at(s, { query: 'x' })).toEqual(['1:6'])
    await s.edit([{ sheet: 'S', cell: 'F1', text: 'y' }, { sheet: 'S', cell: 'F2', text: '=F1&"x"' }])
    expect(await at(s, { query: 'x' })).toEqual(['2:6'])
    // a format changes how a number opens for editing
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
    // 15.03.2026 → 16.03.2026 by the column's date format, 1250,5 → 1260,6; a percentage cell is omitted, since typing `6%`
    // into it gives the text `6%%` (the library's automatic percentage entry for fresh typing, see the report)
    expect(await s.replace({ sheet: 'S', query: '5', range: 'A1:B1' }, '6')).toMatchObject({ ok: true, replaced: 2 })
    expect(await s.replace({ sheet: 'S', query: '0', range: 'D1' }, '1')).toMatchObject({ ok: true, replaced: 1 })
    const w = await win(s, 'S', 1, 1, 1, 6)
    expect(w.values[0]!.slice(0, 4)).toEqual([46097, 1260.6, 0.05, '117'])
    expect(await s.replace({ sheet: 'S', query: 'x', wholeCell: true }, '123456789012345')).toMatchObject({ ok: true, replaced: 1 })
    expect((await s.book()).sheets[0]!.columns?.F).toBeUndefined()
  })

  it('replaces in the cells Find finds, by Excel’s rules: a text without its apostrophe, `?` one UTF-16 unit', async () => {
    const s = source(letters())
    expect(await s.replace({ sheet: 'S', query: '007', wholeCell: true }, '8')).toMatchObject({ ok: true, replaced: 1 })
    expect((await win(s, 'S', 1, 1, 2, 1)).values).toEqual([['8']])
    expect(await s.replace({ sheet: 'S', query: 'a?b' }, 'y')).toEqual({ ok: true, changes: [], replaced: 0 })
    expect(await s.replace({ sheet: 'S', query: 'a??b' }, 'y')).toMatchObject({ ok: true, replaced: 1 })
    expect((await win(s, 'S', 1, 1, 1, 1)).values).toEqual([['y']])
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

  it('finds nothing to replace in shown text, and leaves a formula the replacement breaks as it was, with why', async () => {
    const s = source(formats())
    const before = s.history.undo
    expect(await s.replace({ sheet: 'S', query: '1 250', lookIn: 'values' }, '2')).toEqual({ ok: true, changes: [], replaced: 0 })
    const out = await s.replace({ sheet: 'S', query: ':', lookIn: 'formulas' }, '::')
    expect(out).toMatchObject({ ok: true, replaced: 0, refused: [{ sheet: 'S', cell: 'E1', message: expect.any(String) }] })
    expect(s.history.undo).toBe(before)
    expect((await win(s, 'S', 1, 1, 5, 1)).formulas).toEqual({ E1: 'СУММ(B1:C1)' })
  })
})
