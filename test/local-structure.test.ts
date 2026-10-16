// @vitest-environment node
// localSource changes of the book's shape and look: structure operations, widths and heights, formats, frozen panes.
import { afterEach, describe, expect, it } from 'vitest'
import type { LocalSource } from '../src/local'
import type { SheetOp, SourceEvent } from '../src/source'
import { BOLD, closeAll, ROW_FILL, source, styleOf, watched, win } from './local-fixture'

afterEach(closeAll)

/** A source over the fixture with its events, and the outcome of one operation. */
async function operate(op: SheetOp, prepare?: (s: LocalSource) => Promise<unknown>) {
  const s = source()
  await prepare?.(s)
  const events: SourceEvent[] = []
  s.subscribe((e) => events.push(e))
  const got = await s.structure(op)
  return { s, got, events }
}
const names = async (s: LocalSource) => (await s.sheets()).map((x) => x.name)

describe('localSource: structure operations, each answered by a reload', () => {
  it('undo and redo of a copy onto empty cells reload, as the step does not list the cells it empties', async () => {
    const { s, got } = await operate({ op: 'copyRange', from: { sheet: 'Лист1', range: 'A1:A2' }, to: { sheet: 'Лист1', range: 'H1' } })
    expect(got).toMatchObject({ ok: true })
    expect((await win(s, 'Лист1', 1, 2, 8, 1)).values).toEqual([[1], [2]])
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.undo()
    expect(events).toContainEqual({ kind: 'reload' })
    expect((await win(s, 'Лист1', 1, 2, 8, 1)).values).toEqual([[null], [null]])
    events.length = 0
    await s.redo()
    expect(events).toContainEqual({ kind: 'reload' })
    expect((await win(s, 'Лист1', 1, 2, 8, 1)).values).toEqual([[1], [2]])
  })

  it('insertRows: formulas move, the window reads the moved book', async () => {
    const { s, got, events } = await operate({ op: 'insertRows', sheet: 'Лист1', row: 2 })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 6, 1, 7)
    expect(w.formulas).toEqual({ A4: 'СУММ(A1:A3)' })
    expect(w.values.map((r) => r[0])).toEqual([1, null, 2, 3, null, null])
    expect(w.rowHeights).toEqual({ 3: 30 })
    expect(w.notes).toEqual({ B1: { author: 'Аня', text: 'Проверить' } })
    expect(styleOf(w, 'A6')).toEqual(ROW_FILL)
    expect((await s.sheets())[0]).toMatchObject({ maxRow: 6 })
  })

  it('deleteRows: the rows below move up, a range shrinks', async () => {
    const { s, got, events } = await operate({ op: 'deleteRows', sheet: 'Лист1', row: 1 })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 2, 1, 1)
    expect(w.formulas).toEqual({ A2: 'СУММ(A1:A1)' })
    expect(w.values).toEqual([[2], [2]])
    expect((await s.sheets())[0]).toMatchObject({ maxRow: 4 })
  })

  it('insertColumns: formulas move right', async () => {
    const { s, got, events } = await operate({ op: 'insertColumns', sheet: 'Лист1', column: 'A' })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 3, 1, 2)
    expect(w.formulas).toEqual({ B3: 'СУММ(B1:B2)' })
    expect(w.values.map((r) => r[1])).toEqual([1, 2, 3])
    expect((await s.sheets())[0]).toMatchObject({ maxCol: 8 })
  })

  it('deleteColumns: formulas move left', async () => {
    const { s, got, events } = await operate({ op: 'deleteColumns', sheet: 'Таблица', column: 'D' })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Таблица', 1, 1, 4, 1)
    expect(w.formulas).toEqual({ D1: 'СУММ(Продажи[Сумма])' })
    expect(w.values).toEqual([[30]])
    expect((await s.sheets())[1]).toMatchObject({ maxCol: 4 })
  })

  it('moveRange: the cells go, search finds them at their new place', async () => {
    const { s, got, events } = await operate({ op: 'moveRange', from: { sheet: 'Список', range: 'A3:B3' }, to: { sheet: 'Список', cell: 'D3' } }, (s) => s.find({ sheet: 'Список', query: 'а' }))
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    expect((await win(s, 'Список', 3, 1, 1, 5)).text).toEqual([['', '', '', 'а', '30']])
    expect((await s.find({ sheet: 'Список', query: 'а' })).hits).toEqual([{ row: 3, col: 4 }])
  })

  it('copyRange: a formula is copied with its references moved', async () => {
    const { s, got, events } = await operate({ op: 'copyRange', from: { sheet: 'Лист1', range: 'A3' }, to: { sheet: 'Лист1', range: 'C3' } })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 3, 1, 1, 3)
    expect(w.formulas).toEqual({ A3: 'СУММ(A1:A2)', C3: 'СУММ(C1:C2)' })
    expect(w.values).toEqual([[3, 'на лист', 0]])
  })

  it('renameSheet: the formulas name it anew; a taken name is refused with the library message', async () => {
    const { s, got, events } = await operate({ op: 'renameSheet', sheet: 'Список', name: 'Люди' }, (s) => s.edit([{ sheet: 'Лист1', cell: 'H1', text: '=Список!B2*2' }]))
    expect(got).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await names(s)).toEqual(['Лист1', 'Таблица', 'Люди'])
    // the sheet in quotes only where Excel puts them, in the formula as shown and the text the cell opens with
    const w = await win(s, 'Лист1', 1, 1, 8, 1)
    expect(w.formulas).toEqual({ H1: 'Люди!B2*2' })
    expect(w.editText).toEqual({ H1: '=Люди!B2*2' })
    const refused = await s.structure({ op: 'renameSheet', sheet: 'Люди', name: 'таблица' })
    expect(refused).toMatchObject({ ok: false, reason: 'refused' })
    // the library's message, in the book's locale
    expect(refused.ok ? '' : refused.message).toMatch(/таблица/)
  })

  it('deleteSheet: the sheet goes, references to it become #ССЫЛКА!', async () => {
    const { s, got, events } = await operate({ op: 'deleteSheet', sheet: 'Список' }, (s) => s.edit([{ sheet: 'Лист1', cell: 'H1', text: '=Список!B2*2' }]))
    expect(got).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await names(s)).toEqual(['Лист1', 'Таблица'])
    expect((await win(s, 'Лист1', 1, 1, 8, 1)).text).toEqual([['#ССЫЛКА!']])
  })

  it('insertSheet: an empty sheet before another', async () => {
    const { s, got, events } = await operate({ op: 'insertSheet', name: 'Новый', before: 'Таблица' })
    expect(got).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await names(s)).toEqual(['Лист1', 'Новый', 'Таблица', 'Список'])
    expect((await s.sheets())[1]).toMatchObject({ maxRow: 0, maxCol: 0 })
    expect(await s.edit([{ sheet: 'Новый', cell: 'A1', text: '=Таблица!B2' }])).toMatchObject({ ok: true, changes: [{ cell: 'A1', value: 10 }] })
  })

  it('moveSheet: the order changes, the values do not', async () => {
    const { s, got, events } = await operate({ op: 'moveSheet', sheet: 'Список', before: 'Лист1' })
    expect(got).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await names(s)).toEqual(['Список', 'Лист1', 'Таблица'])
    expect((await win(s, 'Список', 2, 1, 1, 2)).values).toEqual([['б', 5]])
  })

  it('hideSheet and unhideSheet: the list says which sheets are hidden', async () => {
    const { s, got, events } = await operate({ op: 'hideSheet', sheet: 'Таблица' })
    expect(got).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect((await s.sheets()).map((x) => !!x.hidden)).toEqual([false, true, false])
    expect(await s.structure({ op: 'unhideSheet', sheet: 'Таблица' })).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }, { kind: 'reload' }])
    expect((await s.sheets()).map((x) => !!x.hidden)).toEqual([false, false, false])
  })
})

describe('localSource: widths, heights, formats, frozen panes', () => {
  it('sets a column width and a row height in the book', async () => {
    const s = source()
    expect(await s.layout({ sheet: 'Лист1', columns: 'C', width: 15 })).toMatchObject({ ok: true })
    expect(await s.layout({ sheet: 'Лист1', columns: 'E', width: 9 })).toMatchObject({ ok: true })
    expect(await s.layout({ sheet: 'Лист1', rows: '3:4', height: 0 })).toMatchObject({ ok: true })
    const w = await win(s, 'Лист1', 1, 5, 1, 5)
    expect(w.widths).toEqual([8.83203125, 20, 15, 8.83203125, 9])
    expect(w.rowHeights).toEqual({ 2: 30, 3: 0, 4: 0 })
    const b = await s.book()
    expect(b.sheets[0]!.columns?.C).toEqual({ width: 15 })
  })

  it('hides and shows lines keeping their size, and fits columns to what they show, each by a reload', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.layout({ sheet: 'Лист1', columns: 'B', hidden: true })).toMatchObject({ ok: true })
    expect(await s.layout({ sheet: 'Лист1', rows: '2', hidden: true })).toMatchObject({ ok: true })
    let w = await win(s, 'Лист1', 1, 3, 1, 3)
    expect(w.widths[1]).toBe(0)
    expect(w.rowHeights[2]).toBe(0)
    await s.layout({ sheet: 'Лист1', columns: 'B', hidden: false })
    await s.layout({ sheet: 'Лист1', rows: '2', hidden: false })
    w = await win(s, 'Лист1', 1, 3, 1, 3)
    // the width and the height they had come back, as Excel shows a line again
    expect(w.widths[1]).toBe(20)
    expect(w.rowHeights[2]).toBe(30)
    await s.layout({ sheet: 'Лист1', columns: 'B', fit: true })
    const b = await s.book()
    expect(b.sheets[0]!.columns?.B).toMatchObject({ bestFit: true })
    expect(b.sheets[0]!.columns!.B!.width).toBeLessThan(20)
    // a width set by hand after the fit is no longer the fitted one
    await s.layout({ sheet: 'Лист1', columns: 'B', width: 12 })
    expect((await s.book()).sheets[0]!.columns?.B).toEqual({ width: 12 })
    expect(events.slice(0, 5)).toEqual([{ kind: 'reload' }, { kind: 'reload' }, { kind: 'reload' }, { kind: 'reload' }, { kind: 'reload' }])
  })

  it('a line break typed wraps the cell and fits its row (xlcalc 0.33); Fit brings a row set by hand back to its text', async () => {
    const s = source()
    expect(await s.edit([{ sheet: 'Лист1', cell: 'F8', text: 'один\nдва' }])).toMatchObject({ ok: true })
    let b = await s.book()
    const fitted = b.sheets[0]!.rows?.['8']
    expect(fitted).toMatchObject({ bestFit: true })
    expect(fitted!.height!).toBeGreaterThan(15)
    const w = await win(s, 'Лист1', 8, 1, 6, 1)
    expect(styleOf(w, 'F8')?.alignment?.wrap).toBe(true)
    expect(w.rowHeights[8]).toBe(fitted!.height)
    expect(await s.layout({ sheet: 'Лист1', rows: '8', height: 60 })).toMatchObject({ ok: true })
    expect(await s.layout({ sheet: 'Лист1', rows: '8', fit: true })).toMatchObject({ ok: true })
    b = await s.book()
    expect(b.sheets[0]!.rows?.['8']).toEqual(fitted)
  })

  it('formats a range: a new style in the table, a number format, and null removes them', async () => {
    const s = source()
    const italic = { font: { italic: true } }
    let got = await s.format('Лист1', [{ range: 'A1:A2', style: italic, numberFormat: '0.00' }])
    expect(got.ok).toBe(true)
    let w = await win(s, 'Лист1', 1, 3, 1, 1)
    expect(styleOf(w, 'A1')).toEqual(italic)
    expect(styleOf(w, 'A2')).toEqual(italic)
    expect(w.text.map((r) => r[0])).toEqual(['1,00', '2,00', '3'])
    expect(w.formulas.A3).toBe('СУММ(A1:A2)')
    got = await s.format('Лист1', [{ range: 'A1', style: null, numberFormat: null }])
    expect(got.ok).toBe(true)
    w = await win(s, 'Лист1', 1, 2, 1, 1)
    expect(styleOf(w, 'A1')).toBeUndefined()
    expect(w.text[0]![0]).toBe('1')
    expect(styleOf(w, 'A2')).toEqual(italic)
  })

  it('formats several ranges in one step of undo, as outside borders over a selection', async () => {
    const s = source()
    const line = { style: 'thin' as const }
    const got = await s.format('Список', [
      { range: 'D2:E2', style: { border: { top: line } } },
      { range: 'D3:E3', style: { border: { bottom: line } } },
      { range: 'D2:D3', style: { border: { left: line } } },
      { range: 'E2:E3', style: { border: { right: line } } },
    ])
    expect(got.ok).toBe(true)
    let w = await win(s, 'Список', 2, 2, 4, 2)
    expect(styleOf(w, 'D2')).toEqual({ border: { top: line, left: line } })
    expect(styleOf(w, 'E3')).toEqual({ border: { bottom: line, right: line } })
    expect(s.history).toEqual({ undo: 1, redo: 0 })
    await s.undo()
    w = await win(s, 'Список', 2, 2, 4, 2)
    expect(['D2', 'E2', 'D3', 'E3'].map((a) => styleOf(w, a))).toEqual([undefined, undefined, undefined, undefined])
  })

  it('formats empty cells and whole columns, and keeps notes and formulas', async () => {
    const s = source()
    await s.format('Лист1', [{ range: 'H1:H2', style: BOLD }])
    await s.format('Лист1', [{ range: 'B:B', style: ROW_FILL }])
    const w = await win(s, 'Лист1', 1, 9, 1, 8)
    expect(styleOf(w, 'H1')).toEqual(BOLD)
    expect(styleOf(w, 'B1')).toEqual(ROW_FILL)
    expect(styleOf(w, 'B9')).toEqual(ROW_FILL)
    expect(w.notes?.B1).toBeDefined()
    expect(w.links?.B2).toBeDefined()
  })

  it('formats a bold row across a filled column as Excel does: their crossing cells get both', async () => {
    const s = source()
    const FILL = { fill: { color: '#ddeeff' } }
    expect(await s.format('Список', [{ range: 'C:C', style: FILL }])).toMatchObject({ ok: true })
    expect(await s.format('Список', [{ range: '2:2', style: { font: { bold: true } } }])).toMatchObject({ ok: true })
    const w = await win(s, 'Список', 1, 3, 2, 3)
    expect(styleOf(w, 'C2')).toEqual({ font: { bold: true }, fill: { color: '#ddeeff' } })
    expect(styleOf(w, 'B2')).toEqual({ font: { bold: true } })
    expect(styleOf(w, 'C3')).toEqual(FILL)
    expect(styleOf(w, 'B3')).toBeUndefined()
  })

  it('formats over each cell\'s look: bold keeps a cell italic, rows 2:3 at once', async () => {
    const s = source()
    await s.format('Список', [{ range: 'A2', style: { font: { italic: true } } }])
    await s.format('Список', [{ range: '2:3', style: { font: { bold: true } } }])
    const w = await win(s, 'Список', 2, 2, 1, 1)
    expect(styleOf(w, 'A2')).toEqual({ font: { italic: true, bold: true } })
    expect(styleOf(w, 'A3')).toEqual({ font: { bold: true } })
  })

  it('reloads after a number format, which changes what cells show and open with; a look alone repaints its ranges', async () => {
    const { s, count } = watched()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.format('Лист1', [{ range: 'A1:A2', numberFormat: '0%' }])
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 2, 1, 1)
    expect(w.numberFormats).toMatchObject({ A1: '0%', A2: '0%' })
    expect(w.editText).toMatchObject({ A1: '100%', A2: '200%' })
    // a look over cells and over whole rows: the windows over them are read again in place, the values the filter read stay
    const filter = { head: 1, cols: { 1: { values: ['б'] } } }
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4])
    await s.format('Список', [{ range: 'A1', style: BOLD }, { range: '3:4', style: ROW_FILL }])
    expect(events.at(-1)).toEqual({ kind: 'repaint', ranges: [{ sheet: 'Список', range: 'A1' }, { sheet: 'Список', range: '3:4' }] })
    expect(styleOf(await win(s, 'Список', 1, 1, 1, 1), 'A1')).toEqual(BOLD)
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4])
    expect(count('getCells')).toBe(1)
  })

  it('freezes panes in the book, the values the filter read kept', async () => {
    const { s, count } = watched()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    const filter = { head: 1, sort: { col: 1, dir: 'asc' as const } }
    const rows = (await s.visibleRows('Список', filter)).rows
    await s.freeze('Список', { rows: 1, cols: 0 })
    expect((await s.visibleRows('Список', filter)).rows).toEqual(rows)
    expect(count('getCells')).toBe(1)
    await s.freeze('Лист1', null)
    const list = await s.sheets()
    expect(list[2]!.freeze).toEqual({ rows: 1, cols: 0 })
    expect(list[0]!.freeze).toBeUndefined()
    expect(events).toEqual([{ kind: 'reload' }, { kind: 'reload' }])
  })
})
