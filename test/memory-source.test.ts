// memorySource: the window's fields, edits without an engine, layout, refusals for tests, events, and the cell limit.
import { describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { SheetWindow, SourceEvent } from '../src/source'

const book = (): MemoryBook => ({
  styles: [{}, { font: { bold: true } }],
  sheets: [
    {
      name: 'S',
      freeze: { rows: 1, cols: 0 },
      locked: [{ range: 'A1:A2', reason: 'script' }],
      cells: {
        A1: 'head',
        B2: { value: 1.5, text: '1,50', style: 1, color: '#ff0000', formula: 'A2*3', link: { href: 'https://example.org' }, note: { text: 'n' }, bar: { ratio: 0.5, color: '#00f' } },
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
    expect(w.bars).toEqual({ B2: { ratio: 0.5, color: '#00f' } })
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
})
