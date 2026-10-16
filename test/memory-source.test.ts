// memorySource: the window's fields, edits without an engine, layout, refusals for tests, events, and the cell limit.
import { describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { FindRequest, SheetWindow, SourceEvent } from '../src/source'

const book = (): MemoryBook => ({
  styles: [{}, { font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      freeze: { rows: 1, cols: 0 },
      locked: [{ range: 'A1:A2', reason: 'script' }],
      cells: {
        A1: 'head',
        B2: { value: 1.5, text: '1,50', style: 1, color: '#ff0000', formula: 'A2*3', link: { href: 'https://example.org' }, note: { text: 'n' } },
        C3: true,
        D4: { error: 'DIV0' },
      },
      widths: { B: 20, 'C:D': 0 },
      heights: { '2': 30, '3:4': 0 },
      defaultRowHeight: 18,
      merges: ['E1:F2', 'H9:I9'],
      tables: [{ name: 'T', range: 'A1:B2', headerRow: 1 }],
      notice: 'computed',
    },
    { name: 'Hidden', hidden: true, cells: {} },
  ],
})

describe('memorySource', () => {
  it('lists sheets with their extent and SheetInfo fields', async () => {
    const src = memorySource(book())
    expect(await src.sheets()).toEqual([
      { name: 'S', maxRow: 9, maxCol: 9, freeze: { rows: 1, cols: 0 }, locked: [{ range: 'A1:A2', reason: 'script' }], hidden: undefined, columns: undefined, rowNumbers: undefined, totals: undefined },
      { name: 'Hidden', maxRow: 0, maxCol: 0, hidden: true, freeze: undefined, locked: undefined, columns: undefined, rowNumbers: undefined, totals: undefined },
    ])
  })

  it('gives a window with every display field', async () => {
    const w = (await memorySource(book()).window({ sheet: 'S', row: 1, rows: 4, col: 1, cols: 4 })) as SheetWindow
    expect(w.text).toEqual([
      ['head', '', '', ''],
      ['', '1,50', '', ''],
      ['', '', 'TRUE', ''],
      ['', '', '', '#DIV/0!'],
    ])
    expect(w.values[1]![1]).toBe(1.5)
    expect(w.values[3]![3]).toEqual({ error: 'DIV0' })
    expect(w.formulas).toEqual({ B2: 'A2*3' })
    expect(w.colors).toEqual({ B2: '#ff0000' })
    expect(w.styleIds).toEqual({ B2: 1 })
    expect(w.styles[1]).toEqual({ font: { bold: true } })
    expect(w.widths).toEqual([null, 20, 0, 0])
    expect(w.rowHeights).toEqual({ 2: 30, 3: 0, 4: 0 })
    expect(w.defaultRowHeight).toBe(18)
    expect(w.merges).toEqual([])
    expect(w.tables).toEqual([{ name: 'T', range: 'A1:B2', headerRow: 1 }])
    expect(w.links).toEqual({ B2: { href: 'https://example.org' } })
    expect(w.notes).toEqual({ B2: { text: 'n' } })
    expect(w.notice).toBe('computed')
    const right = (await memorySource(book()).window({ sheet: 'S', row: 1, rows: 3, col: 5, cols: 2 })) as SheetWindow
    expect(right.merges).toEqual(['E1:F2'])
  })

  it('cuts a window to the cell limit and refuses an unknown sheet', async () => {
    const src = memorySource(book(), { maxCells: 10 })
    const w = (await src.window({ sheet: 'S', row: 1, rows: 100, col: 1, cols: 5 })) as SheetWindow
    expect(w.text).toHaveLength(2)
    await expect(src.window({ sheet: 'Nope', row: 1, rows: 1, col: 1, cols: 1 })).rejects.toThrow('No sheet Nope')
  })

  it('edits without an engine: a formula stays a string, `1.5` is a number, the rest text', async () => {
    const src = memorySource(book())
    const out = await src.edit!([
      { sheet: 'S', cell: 'a5', text: '=SUM(A1:A3)' },
      { sheet: 'S', cell: 'B5', text: '-1.25' },
      { sheet: 'S', cell: 'C5', text: '1,5' },
      { sheet: 'S', cell: 'D5', text: "'007" },
      { sheet: 'S', cell: 'B2', text: null },
    ])
    expect(out).toEqual({
      ok: true,
      changes: [
        { sheet: 'S', cell: 'A5', text: '', value: null, formula: 'SUM(A1:A3)' },
        { sheet: 'S', cell: 'B5', text: '-1.25', value: -1.25, formula: null },
        { sheet: 'S', cell: 'C5', text: '1,5', value: '1,5', formula: null },
        { sheet: 'S', cell: 'D5', text: '007', value: '007', formula: null },
        { sheet: 'S', cell: 'B2', text: '', value: null, formula: null },
      ],
    })
    const w = (await src.window({ sheet: 'S', row: 2, rows: 1, col: 2, cols: 1 })) as SheetWindow
    // Clearing keeps the style, links and notes, as Delete in Excel keeps formatting.
    expect(w.styleIds).toEqual({ B2: 1 })
    expect(w.text).toEqual([['']])
  })

  it('answers the next edits with refusals set by failNext, then applies again', async () => {
    const src = memorySource(book())
    src.failNext({ ok: false, reason: 'conflict', message: 'c' }, { ok: false, reason: 'invalid', message: 'i', at: 3 })
    expect(await src.edit!([{ sheet: 'S', cell: 'A9', text: '1' }])).toEqual({ ok: false, reason: 'conflict', message: 'c' })
    expect(await src.edit!([{ sheet: 'S', cell: 'A9', text: '1' }])).toEqual({ ok: false, reason: 'invalid', message: 'i', at: 3 })
    expect((await src.edit!([{ sheet: 'S', cell: 'A9', text: '1' }])).ok).toBe(true)
    expect(await src.edit!([{ sheet: 'Nope', cell: 'A1', text: '1' }])).toMatchObject({ ok: false, reason: 'failed' })
  })

  it('keeps widths and heights set by layout', async () => {
    const src = memorySource(book())
    expect(await src.layout!({ sheet: 'S', columns: 'A:B', width: 12 })).toEqual({ ok: true, changes: [] })
    await src.layout!({ sheet: 'S', rows: '7', height: 40 })
    const w = (await src.window({ sheet: 'S', row: 7, rows: 1, col: 1, cols: 2 })) as SheetWindow
    expect(w.widths).toEqual([12, 12])
    expect(w.rowHeights).toEqual({ 7: 40 })
  })

  it('takes a line as a span of one (`C:C`, `5:5`) in the book and in layout, as `C` and `5`', async () => {
    const src = memorySource({ sheets: [{ name: 'S', cells: { A1: 1 }, widths: { 'C:C': 30 }, heights: { '5:5': 25 }, outline: { rows: { '3:3': 1 } } }] })
    await src.layout!({ sheet: 'S', columns: 'd:d', width: 9 })
    const w = (await src.window({ sheet: 'S', row: 1, rows: 5, col: 1, cols: 4 })) as SheetWindow
    expect(w.widths).toEqual([null, null, 30, 9])
    expect(w.rowHeights).toEqual({ 5: 25 })
    expect((await src.sheets())[0]!.outline?.rows).toEqual([{ from: 3, to: 3, level: 1, collapsed: false }])
  })

  it('sends events to subscribers; setCells changes and sends `changed`', async () => {
    const src = memorySource(book())
    const got: SourceEvent[] = []
    const off = src.subscribe!((e) => got.push(e))
    src.emit({ kind: 'reload' })
    const changes = src.setCells('S', { A1: 'other tab' })
    off()
    src.emit({ kind: 'reload' })
    expect(got).toEqual([{ kind: 'reload' }, { kind: 'changed', changes }])
    expect(changes).toEqual([{ sheet: 'S', cell: 'A1', text: 'other tab', value: 'other tab', formula: null }])
  })

  it('rejects with AbortError when the signal is aborted', async () => {
    const src = memorySource(book(), { delay: 20 })
    const ctrl = new AbortController()
    const p = src.window({ sheet: 'S', row: 1, rows: 1, col: 1, cols: 1 }, ctrl.signal)
    ctrl.abort()
    await expect(p).rejects.toMatchObject({ name: 'AbortError' })
  })

  describe('0.2.0 queries', () => {
    const table = (): MemoryBook => ({
      sheets: [
        {
          name: 'T',
          cells: {
            A1: 'Name', B1: 'Sum', C1: 'Note',
            A2: 'b', B2: 3, C2: 'x',
            A3: 'a', B3: 10,
            A4: 'c', B4: 'text', C4: 'x',
            A5: 'a', B5: -1, C5: 'Apple',
            A6: 'b', C6: 'apple pie',
          },
          merges: ['C2:D2', 'C3:C4'],
        },
      ],
    })

    it('visibleRows: the header first, conditions on values and substrings, Excel’s sort with empty cells last', async () => {
      const src = memorySource(table())
      expect(await src.visibleRows!('T', { head: 1 })).toEqual({ rows: null })
      expect((await src.visibleRows!('T', { head: 1, cols: { 1: { values: ['a', 'b'] } } })).rows).toEqual([1, 2, 3, 5, 6])
      expect((await src.visibleRows!('T', { head: 1, cols: { 3: { text: 'APP' } } })).rows).toEqual([1, 5, 6])
      expect((await src.visibleRows!('T', { head: 1, sort: { col: 2, dir: 'asc' } })).rows).toEqual([1, 5, 2, 3, 4, 6])
      expect((await src.visibleRows!('T', { head: 1, sort: { col: 2, dir: 'desc' } })).rows).toEqual([1, 4, 3, 2, 5, 6])
      // as ./local: a text of spaces asks for nothing, a condition's text is taken trimmed, a fractional head counts its whole rows
      expect((await src.visibleRows!('T', { head: 1, cols: { 3: { text: '   ' } } })).rows).toBeNull()
      expect((await src.visibleRows!('T', { head: 1.5, cols: { 3: { text: ' app ' } } })).rows).toEqual([1, 5, 6])
    })

    it('columnValues: counts over the rows the other columns leave, empty as ""', async () => {
      const src = memorySource(table())
      expect(await src.columnValues!('T', 1, { head: 1 })).toEqual({
        values: [
          { text: 'a', count: 2 },
          { text: 'b', count: 2 },
          { text: 'c', count: 1 },
        ],
        more: 0,
      })
      // Its own condition does not narrow the column's list; another column's does.
      expect((await src.columnValues!('T', 1, { head: 1, cols: { 1: { values: ['a'] }, 3: { values: ['x'] } } })).values).toEqual([
        { text: 'b', count: 1 },
        { text: 'c', count: 1 },
      ])
      expect((await src.columnValues!('T', 3, { head: 1 })).values.map((v) => v.text)).toEqual(['Apple', 'apple pie', 'x', ''])
    })

    it('find: a substring in any case, in row order, only shown rows under a filter', async () => {
      const src = memorySource(table())
      expect(await src.find!({ sheet: 'T', query: 'APP' })).toEqual({
        hits: [
          { row: 5, col: 3 },
          { row: 6, col: 3 },
        ],
        total: 2,
      })
      expect((await src.find!({ sheet: 'T', query: 'app', filter: { head: 1, cols: { 1: { values: ['a'] } } } })).hits.map((h) => h.row)).toEqual([5])
    })

    /** Formulas, a number shown by a format and a text that reads as a number. */
    const sums = (): MemoryBook => ({
      sheets: [
        {
          name: 'F',
          cells: { A1: 10, A2: { value: 1.5, text: '1,50' }, A3: { value: 11.5, text: '11,50', formula: 'SUM(A1:A2)' }, B1: 'sum of', B2: '007', B3: 'Sum' },
        },
      ],
    })

    it('find: Excel’s options, in formulas (the default) or in shown values, a range, by columns', async () => {
      const src = memorySource(sums())
      const at = async (req: Partial<FindRequest>) => (await src.find!({ sheet: 'F', query: '', ...req })).hits.map((h) => `${h.row}:${h.col}`)
      expect(await at({ query: 'sum' })).toEqual(['1:2', '3:1', '3:2'])
      expect(await at({ query: 'sum', lookIn: 'values' })).toEqual(['1:2', '3:2'])
      expect(await at({ query: '1.5' })).toEqual(['2:1'])
      expect(await at({ query: '1,5', lookIn: 'values' })).toEqual(['2:1', '3:1'])
      expect(await at({ query: "'0" })).toEqual(['2:2'])
      expect(await at({ query: 'sum', matchCase: true })).toEqual(['1:2'])
      expect(await at({ query: 'sum', wholeCell: true })).toEqual(['3:2'])
      expect(await at({ query: 's*', range: 'B2:B3', lookIn: 'values' })).toEqual(['3:2'])
      expect(await at({ query: '1', byColumns: true })).toEqual(['1:1', '2:1', '3:1'])
      expect(await at({ query: '' })).toEqual([])
    })

    it('replace: every hit in formulas whatever `lookIn` says, the texts read as typed, the changes and how many cells', async () => {
      const src = memorySource(sums())
      const out = await src.replace!({ sheet: 'F', query: 'a2', lookIn: 'values' }, 'A1')
      expect(out).toEqual({ ok: true, changes: [{ sheet: 'F', cell: 'A3', text: '', value: null, formula: 'SUM(A1:A1)' }], replaced: 1 })
      // the apostrophe keeps a text that reads as a number a text
      expect(await src.replace!({ sheet: 'F', query: '0' }, '1')).toMatchObject({ ok: true, replaced: 2 })
      const w = (await src.window({ sheet: 'F', row: 1, rows: 3, col: 1, cols: 2 })) as SheetWindow
      expect(w.values.map((r) => r[0])).toEqual([11, 1.5, null])
      expect(w.values[1]![1]).toBe('117')
      expect(await src.replace!({ sheet: 'F', query: 'nothing' }, 'x')).toEqual({ ok: true, changes: [], replaced: 0 })
    })

    it('stats: a rectangle’s totals, over shown rows under a filter', async () => {
      const src = memorySource(table())
      expect(await src.stats!('T', 'B2:B6', undefined)).toEqual({ cells: 4, numbers: 3, sum: 12, min: -1, max: 10 })
      expect(await src.stats!('T', 'B1:B6', { head: 1, cols: { 1: { values: ['a'] } } })).toEqual({ cells: 3, numbers: 2, sum: 9, min: -1, max: 10 })
    })

    it('a window under a filter goes by display index; only one-row merges on shown rows come', async () => {
      const src = memorySource(table())
      const w = (await src.window({ sheet: 'T', row: 2, rows: 10, col: 1, cols: 4, filter: { head: 1, cols: { 3: { values: ['x'] } } } })) as SheetWindow
      expect(w.row).toBe(2)
      expect(w.text.map((r) => r[0])).toEqual(['b', 'c'])
      expect(w.formulas).toEqual({})
      expect(w.merges).toEqual(['C2:D2'])
    })

    it('recompute: windows answer pending until the time is up', async () => {
      vi.useFakeTimers({ toFake: ['Date'] })
      const src = memorySource(table())
      src.recompute(2000)
      vi.advanceTimersByTime(500)
      expect(await src.window({ sheet: 'T', row: 1, rows: 1, col: 1, cols: 1 })).toEqual({ pending: { elapsedMs: 500 } })
      vi.advanceTimersByTime(1600)
      expect(((await src.window({ sheet: 'T', row: 1, rows: 1, col: 1, cols: 1 })) as SheetWindow).text).toEqual([['Name']])
      vi.useRealTimers()
    })
  })
})

describe('memorySource under a display filter', () => {
  it('reads windows by the list `visibleRows` last gave, so a row an edit took out stays where the sheet holds it until Reapply', async () => {
    const filter = { head: 1, cols: { 2: { values: ['a'] } } }
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 'Name', B1: 'Status', A2: 'n2', B2: 'a', A3: 'n3', B3: 'b', A4: 'n4', B4: 'a', A5: 'n5', B5: 'b', A6: 'n6', B6: 'a' } }] })
    const column = async () => ((await source.window({ sheet: 'S', row: 1, rows: 10, col: 1, cols: 1, filter })) as { text: string[][] }).text.map((r) => r[0])
    expect((await source.visibleRows!('S', filter)).rows).toEqual([1, 2, 4, 6])
    source.setCells('S', { B4: 'b' })
    expect(await column()).toEqual(['Name', 'n2', 'n4', 'n6'])
    expect((await source.visibleRows!('S', filter)).rows).toEqual([1, 2, 6])
    expect(await column()).toEqual(['Name', 'n2', 'n6'])
    // Another filter, or none asked before, reads the sheet as it is.
    expect(((await source.window({ sheet: 'S', row: 1, rows: 10, col: 1, cols: 1, filter: { head: 1, cols: { 2: { values: ['b'] } } } })) as { text: string[][] }).text.map((r) => r[0])).toEqual(['Name', 'n3', 'n4', 'n5'])
  })
})
