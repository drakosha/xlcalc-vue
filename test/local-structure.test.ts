// @vitest-environment node
// localSource changes to the workbook's shape and formatting: structure operations, widths and heights, formats, frozen panes.
import { afterEach, describe, expect, it } from 'vitest'
import type { LocalSource } from '../src/local'
import type { SheetOp, SourceEvent } from '../src/contract/source'
import { BOLD, book, closeAll, ROW_FILL, source, styleOf, watched, win } from './local-fixture'

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

describe('localSource: structure operations, each followed by a reload', () => {
  it('undo and redo of a copy onto empty cells reload, since the step does not list the cells it empties', async () => {
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

  it('insertRows: formulas are adjusted, and the window reads the changed workbook', async () => {
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

  it('deleteRows: the rows below move up, and a range shrinks', async () => {
    const { s, got, events } = await operate({ op: 'deleteRows', sheet: 'Лист1', row: 1 })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 2, 1, 1)
    expect(w.formulas).toEqual({ A2: 'СУММ(A1:A1)' })
    expect(w.values).toEqual([[2], [2]])
    expect((await s.sheets())[0]).toMatchObject({ maxRow: 4 })
  })

  it('insertColumns: formulas shift right', async () => {
    const { s, got, events } = await operate({ op: 'insertColumns', sheet: 'Лист1', column: 'A' })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 3, 1, 2)
    expect(w.formulas).toEqual({ B3: 'СУММ(B1:B2)' })
    expect(w.values.map((r) => r[1])).toEqual([1, 2, 3])
    expect((await s.sheets())[0]).toMatchObject({ maxCol: 8 })
  })

  it('deleteColumns: formulas shift left', async () => {
    const { s, got, events } = await operate({ op: 'deleteColumns', sheet: 'Таблица', column: 'D' })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Таблица', 1, 1, 4, 1)
    expect(w.formulas).toEqual({ D1: 'СУММ(Продажи[Сумма])' })
    expect(w.values).toEqual([[30]])
    expect((await s.sheets())[1]).toMatchObject({ maxCol: 4 })
  })

  it('moveRange: the cells move, and search finds them at their new location', async () => {
    const { s, got, events } = await operate({ op: 'moveRange', from: { sheet: 'Список', range: 'A3:B3' }, to: { sheet: 'Список', cell: 'D3' } }, (s) => s.find({ sheet: 'Список', query: 'а' }))
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    expect((await win(s, 'Список', 3, 1, 1, 5)).text).toEqual([['', '', '', 'а', '30']])
    expect((await s.find({ sheet: 'Список', query: 'а' })).hits).toEqual([{ row: 3, col: 4 }])
  })

  it('copyRange: a formula is copied with its references adjusted', async () => {
    const { s, got, events } = await operate({ op: 'copyRange', from: { sheet: 'Лист1', range: 'A3' }, to: { sheet: 'Лист1', range: 'C3' } })
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 3, 1, 1, 3)
    expect(w.formulas).toEqual({ A3: 'СУММ(A1:A2)', C3: 'СУММ(C1:C2)' })
    expect(w.values).toEqual([[3, 'на лист', 0]])
  })

  it("renameSheet: formulas use the new name; a name already in use is rejected with the library's message", async () => {
    const { s, got, events } = await operate({ op: 'renameSheet', sheet: 'Список', name: 'Люди' }, (s) => s.edit([{ sheet: 'Лист1', cell: 'H1', text: '=Список!B2*2' }]))
    expect(got).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    expect(await names(s)).toEqual(['Лист1', 'Таблица', 'Люди'])
    // the sheet name in quotes only where Excel puts them, in the formula as displayed and in the text the cell opens with
    const w = await win(s, 'Лист1', 1, 1, 8, 1)
    expect(w.formulas).toEqual({ H1: 'Люди!B2*2' })
    expect(w.editText).toEqual({ H1: '=Люди!B2*2' })
    const refused = await s.structure({ op: 'renameSheet', sheet: 'Люди', name: 'таблица' })
    expect(refused).toMatchObject({ ok: false, reason: 'refused' })
    // the library's message, in the workbook's locale
    expect(refused.ok ? '' : refused.message).toMatch(/таблица/)
  })

  it('deleteSheet: the sheet is removed, and references to it become #ССЫЛКА!', async () => {
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

  it('hideSheet and unhideSheet: the list shows which sheets are hidden', async () => {
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
  it('sets a column width and a row height in the workbook', async () => {
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

  it('hides and unhides rows and columns keeping their size, and fits columns to their contents, each followed by a reload', async () => {
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
    // their previous width and height come back, as Excel displays a row or column again
    expect(w.widths[1]).toBe(20)
    expect(w.rowHeights[2]).toBe(30)
    await s.layout({ sheet: 'Лист1', columns: 'B', fit: true })
    const b = await s.book()
    expect(b.sheets[0]!.columns?.B).toMatchObject({ bestFit: true })
    expect(b.sheets[0]!.columns!.B!.width).toBeLessThan(20)
    // a width set manually after the fit is no longer the fitted one
    await s.layout({ sheet: 'Лист1', columns: 'B', width: 12 })
    expect((await s.book()).sheets[0]!.columns?.B).toEqual({ width: 12 })
    expect(events.slice(0, 5)).toEqual([{ kind: 'reload' }, { kind: 'reload' }, { kind: 'reload' }, { kind: 'reload' }, { kind: 'reload' }])
  })

  it('a typed line break wraps the cell and fits its row (xlcalc 0.33); fitting returns a manually sized row to its text', async () => {
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

  it('formats several ranges as a single undo step, like outside borders over a selection', async () => {
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

  it('formats empty cells and entire columns, and keeps notes and formulas', async () => {
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

  it('formats a bold row across a filled column as Excel does: their intersecting cells get both', async () => {
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

  it("formats on top of each cell's formatting: bold keeps a cell italic, rows 2:3 at once", async () => {
    const s = source()
    await s.format('Список', [{ range: 'A2', style: { font: { italic: true } } }])
    await s.format('Список', [{ range: '2:3', style: { font: { bold: true } } }])
    const w = await win(s, 'Список', 2, 2, 1, 1)
    expect(styleOf(w, 'A2')).toEqual({ font: { italic: true, bold: true } })
    expect(styleOf(w, 'A3')).toEqual({ font: { bold: true } })
  })

  it('reloads after a number format, which changes how cells are displayed and what they open with; formatting alone repaints its ranges', async () => {
    const { s, count } = watched()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.format('Лист1', [{ range: 'A1:A2', numberFormat: '0%' }])
    expect(events).toEqual([{ kind: 'reload' }])
    const w = await win(s, 'Лист1', 1, 2, 1, 1)
    expect(w.numberFormats).toMatchObject({ A1: '0%', A2: '0%' })
    expect(w.editText).toMatchObject({ A1: '100%', A2: '200%' })
    // formatting over cells and over entire rows: the windows over them are re-read in place, and the filter's cached values stay
    const filter = { head: 1, cols: { 1: { values: ['б'] } } }
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4])
    await s.format('Список', [{ range: 'A1', style: BOLD }, { range: '3:4', style: ROW_FILL }])
    expect(events.at(-1)).toEqual({ kind: 'repaint', ranges: [{ sheet: 'Список', range: 'A1' }, { sheet: 'Список', range: '3:4' }] })
    expect(styleOf(await win(s, 'Список', 1, 1, 1, 1), 'A1')).toEqual(BOLD)
    expect((await s.visibleRows('Список', filter)).rows).toEqual([1, 2, 4])
    expect(count('getCells')).toBe(1)
  })

  it("gives the workbook's default font, Calibri 11 without one, as xlcalc measures", async () => {
    expect((await source().sheets())[0]!.font).toEqual({ name: 'Calibri', size: 11 })
    expect((await source({ ...book(), font: { name: 'Arial', size: 10 } }).sheets())[0]!.font).toEqual({ name: 'Arial', size: 10 })
  })

  it('a larger font fits its row to it (xlcalc 0.59): the sheets and windows are read again', async () => {
    const s = source()
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    await s.format('Список', [{ range: 'A3', style: { font: { size: 24 } } }])
    expect(events.at(-1)).toEqual({ kind: 'reload' })
    expect((await win(s, 'Список', 3, 1, 1, 1)).rowHeights[3]).toBeGreaterThan(20)
  })

  it("freezes panes in the workbook, keeping the filter's cached values", async () => {
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

  it("the fill handle puts its copies into rows hidden by a filter too, as Excel's does (xlcalc 0.51)", async () => {
    const s = source({ sheets: [{ name: 'S', filtered: true, cells: { A1: { value: 'x' } }, rows: { '3': { hidden: true } } }] })
    expect(await s.structure({ op: 'fillRange', sheet: 'S', range: 'A1', to: 'A1:A4' })).toMatchObject({ ok: true })
    expect((await win(s, 'S', 1, 4, 1, 1)).values.map((r) => r[0])).toEqual(['x', 'x', 'x', 'x'])
  })
})
