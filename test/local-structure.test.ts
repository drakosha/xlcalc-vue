// @vitest-environment node
// localSource changes of the book's shape and look: structure operations, widths and heights, formats, frozen panes.
import { afterEach, describe, expect, it } from 'vitest'
import type { LocalSource } from '../src/local'
import type { SheetOp, SourceEvent } from '../src/source'
import { BOLD, closeAll, ROW_FILL, source, styleOf, win } from './local-fixture'

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
    const { s, got, events } = await operate({ op: 'moveRange', from: { sheet: 'Список', range: 'A3:B3' }, to: { sheet: 'Список', cell: 'D3' } }, (s) => s.find('Список', 'а'))
    expect(got.ok).toBe(true)
    expect(events).toEqual([{ kind: 'reload' }])
    expect((await win(s, 'Список', 3, 1, 1, 5)).text).toEqual([['', '', '', 'а', '30']])
    expect((await s.find('Список', 'а')).hits).toEqual([{ row: 3, col: 4, text: 'а' }])
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
    // Excel shows Люди!B2*2; xlcalc 0.19 keeps the stored text's quotes in formulaLocal (a request to the library)
    expect((await win(s, 'Лист1', 1, 1, 8, 1)).formulas).toEqual({ H1: "'Люди'!B2*2" })
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
