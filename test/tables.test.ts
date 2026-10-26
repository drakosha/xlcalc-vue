// Excel's Create Table (Ctrl+T, Ctrl+L, insertTable) through the local source: the range guessed from the data around the
// active cell, My table has headers, column names from the header cells, a header row of Column1… inserted above the data
// without it, Table1 and Table2; Convert to Range after Excel's question, and Delete Table.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { BookInput } from 'xlcalc'
import { memorySource, type SheetViewApi } from '../src'
import { absoluteRange, columnNames, currentRegion } from '../src/core/tables'
import { localSource, type LocalSource } from '../src/local'
import { goTo, key, mountSheet, settle } from './helpers'

let open: LocalSource[] = []
afterEach(async () => {
  await Promise.all(open.map((s) => s.close()))
  open = []
  document.body.innerHTML = ''
})

function local(book: BookInput): LocalSource {
  const s = localSource(book, { calc: 'inline' })
  open.push(s)
  return s
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const dialog = (w: VueWrapper) => w.find('[data-xs-table-dialog]')
const range = (w: VueWrapper) => (w.find('[data-xs-table-range]').element as HTMLInputElement).value
const headers = (w: VueWrapper) => (w.find('[data-xs-table-headers]').element as HTMLInputElement).checked
async function ok(w: VueWrapper): Promise<void> {
  await w.find('[data-xs-dialog-ok]').trigger('click')
  await settle(16)
}

describe('the Create Table helpers', () => {
  it('the current region grows over rows and columns with a non-empty cell next to it, corners included', () => {
    const filled = new Set(['1:1', '1:2', '2:1', '3:3', '6:6'])
    const area = { top: 1, left: 1, bottom: 10, right: 10 }
    // C3 touches the corner of A1:B2 diagonally, so the region takes it in; F6 stands apart
    expect(currentRegion({ row: 1, col: 1 }, area, (r, c) => filled.has(`${r}:${c}`))).toEqual({ top: 1, left: 1, bottom: 3, right: 3 })
    expect(currentRegion({ row: 6, col: 6 }, area, (r, c) => filled.has(`${r}:${c}`))).toEqual({ top: 6, left: 6, bottom: 6, right: 6 })
    expect(currentRegion({ row: 9, col: 9 }, area, () => false)).toEqual({ top: 9, left: 9, bottom: 9, right: 9 })
  })

  it('names columns as Excel does: Column1 for an empty header, a number after a name already used', () => {
    expect(columnNames(['Name', '', 'Qty', 'qty', ''], 'Column')).toEqual(['Name', 'Column1', 'Qty', 'qty2', 'Column2'])
    expect(columnNames(['', 'Column1'], 'Column')).toEqual(['Column2', 'Column1'])
    expect(columnNames(['', ''], 'Столбец')).toEqual(['Столбец1', 'Столбец2'])
  })

  it('writes the range as the dialog shows it', () => {
    expect(absoluteRange({ top: 1, left: 1, bottom: 5, right: 3 })).toBe('=$A$1:$C$5')
    expect(absoluteRange({ top: 2, left: 2, bottom: 2, right: 2 })).toBe('=$B$2')
  })
})

describe('Create Table', () => {
  const data = (): BookInput => ({
    locale: 'en-US',
    sheets: [
      { name: 'S', cells: { A1: { value: 'Name' }, B1: { value: 'Qty' }, C1: { value: 'Qty' }, A2: { value: 'tea' }, B2: { value: 1 }, C2: { value: 2 }, A3: { value: 'milk' }, B3: { value: 3 }, C3: { value: 4 }, E1: { value: 'apart' } } },
      { name: 'T', cells: { A1: { value: 5 }, B1: { value: 6 }, A2: { value: 7 }, B2: { value: 8 } } },
    ],
  })

  it('Ctrl+T proposes the current region with headers; OK names the columns from the header cells, Table1 then Table2', async () => {
    const s = local(data())
    const w = await mountSheet(s, { style: 'height: 400px' })
    await settle(12)
    await goTo(w, 'B2')
    await key(w, 't', { ctrlKey: true })
    await settle(12)
    expect(dialog(w).exists()).toBe(true)
    expect(range(w)).toBe('=$A$1:$C$3')
    expect(headers(w)).toBe(true)
    await ok(w)
    expect(dialog(w).exists()).toBe(false)
    const book = await s.book()
    expect(book.sheets[0]!.tables).toEqual([expect.objectContaining({ name: 'Table1', range: 'A1:C3', columns: [{ name: 'Name' }, { name: 'Qty' }, { name: 'Qty2' }] })])
    expect(book.sheets[0]!.cells.C1).toMatchObject({ value: 'Qty2' })
    // the table is drawn with the sheet's table colors: its header row
    expect(w.find('[data-cell="1:1"]').classes()).toContain('xs-thead')
    // Ctrl+L on the other sheet: numbers in the first row, so no headers are guessed
    await goTo(w, 'T!A1')
    await key(w, 'l', { ctrlKey: true })
    await settle(12)
    expect(range(w)).toBe('=$A$1:$B$2')
    expect(headers(w)).toBe(false)
    await ok(w)
    expect((await s.book()).sheets[1]!.tables).toEqual([expect.objectContaining({ name: 'Table2', range: 'A1:B3' })])
  })

  it('without My table has headers, a header row of Column1, Column2 goes above the data, which moves down', async () => {
    const s = local(data())
    const w = await mountSheet(s, { style: 'height: 400px' })
    await settle(12)
    await goTo(w, 'T!A1')
    expect(await api(w).insertTable()).toBe(true)
    await settle()
    await ok(w)
    const book = await s.book()
    const t = book.sheets[1]!
    expect(t.tables).toEqual([expect.objectContaining({ name: 'Table1', range: 'A1:B3', columns: [{ name: 'Column1' }, { name: 'Column2' }] })])
    expect([t.cells.A1?.value, t.cells.B1?.value, t.cells.A2?.value, t.cells.B3?.value]).toEqual(['Column1', 'Column2', 5, 8])
  })

  it('a typed range is taken with or without $ and =; one Excel would not take keeps the dialog with its message', async () => {
    const s = local(data())
    const w = await mountSheet(s, { style: 'height: 400px' })
    await settle(12)
    await api(w).insertTable()
    await settle()
    await w.find('[data-xs-table-range]').setValue('=nonsense')
    await ok(w)
    expect(dialog(w).find('.xs-dialog-error').text()).toBe("The reference isn't valid.")
    await w.find('[data-xs-table-range]').setValue('A1:B3')
    await ok(w)
    expect((await s.book()).sheets[0]!.tables?.[0]).toMatchObject({ range: 'A1:B3', columns: [{ name: 'Name' }, { name: 'Qty' }] })
  })

  it('Convert to Range asks first and keeps the cells; Delete Table removes the table with its cells', async () => {
    const book = data()
    book.sheets[0]!.tables = [{ name: 'Sales', range: 'A1:C3', columns: [{ name: 'Name' }, { name: 'Qty' }, { name: 'Qty2' }] }]
    book.sheets[0]!.cells.C1 = { value: 'Qty2' }
    const s = local(book)
    const w = await mountSheet(s, { style: 'height: 400px' })
    await settle(12)
    await goTo(w, 'B2')
    const asked = api(w).convertToRange()
    await settle()
    expect(dialog(w).text()).toContain('convert the table to a normal range')
    await w.find('[data-xs-dialog-cancel]').trigger('click')
    expect(await asked).toBeNull()
    expect((await s.book()).sheets[0]!.tables).toHaveLength(1)
    const yes = api(w).convertToRange()
    await settle()
    await w.find('[data-xs-dialog-ok]').trigger('click')
    expect(await yes).toMatchObject({ ok: true })
    await settle(12)
    expect((await s.book()).sheets[0]!.tables ?? []).toEqual([])
    expect((await s.book()).sheets[0]!.cells.A2).toMatchObject({ value: 'tea' })
    expect(await api(w).deleteTable()).toBeNull()
  })

  it('Delete Table removes the table at the active cell with its cells', async () => {
    const book = data()
    book.sheets[1]!.tables = [{ name: 'Nums', range: 'A1:B2', headers: false, columns: [{ name: 'Column1' }, { name: 'Column2' }] }]
    const s = local(book)
    const w = await mountSheet(s, { style: 'height: 400px' })
    await settle(12)
    await goTo(w, 'T!A2')
    expect(await api(w).deleteTable()).toMatchObject({ ok: true })
    const t = (await s.book()).sheets[1]!
    expect([t.tables ?? [], t.cells.A2]).toEqual([[], undefined])
  })

  it('a read-only sheet and a source without `table` have no table commands', async () => {
    const ro = await mountSheet(local(data()), { readOnly: true })
    await settle(12)
    expect(await api(ro).insertTable()).toBe(false)
    const mem = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 'x' } }] }))
    expect(await api(mem).insertTable()).toBe(false)
    await key(mem, 't', { ctrlKey: true })
    await settle()
    expect(dialog(mem).exists()).toBe(false)
  })
})
