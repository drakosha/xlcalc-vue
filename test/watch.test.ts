// Excel's Watch Window: Add Watch for each cell of the selection, Excel's columns with the value and the formula, values
// following edits, events and reloads, Delete Watch, a double click going to the cell, and the list as the host's view state.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi, type SheetWatch } from '../src'
import type { WindowRequest } from '../src/contract/source'
import { WATCH_MAX } from '../src/core/watch'
import { active, editorKey, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({
  sheets: [
    { name: 'One', cells: { A1: 1, B1: 2, A2: 3, B2: { value: 5, formula: 'A1+A2+1' } } },
    { name: 'Two sheet', cells: { C3: 'far' } },
  ],
})
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lists = (w: VueWrapper) => (w.emitted('watches') ?? []).map(([list]) => list as SheetWatch[])
/** The Watch Window's rows as Excel's columns: Book, Sheet, Name, Cell, Value, Formula. */
const rows = (w: VueWrapper) => w.findAll('[data-xs-watch-row]').map((tr) => tr.findAll('td').map((td) => td.text()))

async function open(w: VueWrapper): Promise<void> {
  api(w).watchWindow()
  await settle()
}

describe('Watch Window', () => {
  it("Add Watch adds each cell of the selection with Excel's columns; a cell already watched is not added again", async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    expect(w.findAll('[data-xs-watch] th').map((th) => th.text())).toEqual(['Book', 'Sheet', 'Name', 'Cell', 'Value', 'Formula'])
    await goTo(w, 'A1:B2')
    await w.find('[data-xs-watch-add]').trigger('click')
    await settle()
    expect(rows(w)).toEqual([
      ['', 'One', '', 'A1', '1', ''],
      ['', 'One', '', 'B1', '2', ''],
      ['', 'One', '', 'A2', '3', ''],
      ['', 'One', '', 'B2', '5', '=A1+A2+1'],
    ])
    await goTo(w, 'B2:B3')
    await w.find('[data-xs-watch-add]').trigger('click')
    await settle()
    expect(rows(w).map((r) => r[3])).toEqual(['A1', 'B1', 'A2', 'B2', 'B3'])
    expect(lists(w).map((l) => l.length)).toEqual([4, 5])
    expect(lists(w)[0]![3]).toEqual({ sheet: 'One', cell: 'B2' })
  })

  it("values follow the sheet's own edits, the source's events and reloads, whatever sheet is displayed", async () => {
    const source = memorySource(book())
    const w = await mountSheet(source, { watches: [{ sheet: 'One', cell: 'B2' }, { sheet: 'Two sheet', cell: 'C3' }] })
    await open(w)
    expect(rows(w).map((r) => r[4])).toEqual(['5', 'far'])
    await goTo(w, 'B2')
    await key(w, '7')
    await type(w, '7')
    await editorKey(w, 'Enter')
    await settle()
    expect(rows(w)[0]!.slice(4)).toEqual(['7', ''])
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    source.setCells('One', { B2: '=A1*10' })
    await settle()
    expect(rows(w)[0]![5]).toBe('=A1*10')
    source.setCells('Two sheet', { C3: 'near' })
    source.emit({ kind: 'reload' })
    await settle()
    expect(rows(w)[1]![4]).toBe('near')
  })

  it('reads the values only while open, with one window per sheet over its watched cells', async () => {
    const source = memorySource(book())
    const asked: WindowRequest[] = []
    const window = source.window.bind(source)
    source.window = (req, signal) => {
      asked.push(req)
      return window(req, signal)
    }
    const w = await mountSheet(source, { watches: [{ sheet: 'One', cell: 'A1' }, { sheet: 'One', cell: 'B2' }, { sheet: 'Two sheet', cell: 'C3' }] })
    asked.length = 0
    source.setCells('One', { A1: '4' })
    await settle()
    expect(asked).toEqual([])
    await open(w)
    expect(asked).toEqual([
      { sheet: 'One', row: 1, col: 1, rows: 2, cols: 2 },
      { sheet: 'Two sheet', row: 3, col: 3, rows: 1, cols: 1 },
    ])
    expect(rows(w).map((r) => r[4])).toEqual(['4', '5', 'far'])
  })

  it('Delete Watch removes the chosen rows (Ctrl adds a row, Shift a range), and so does the Delete key', async () => {
    const list = ['A1', 'B1', 'A2', 'B2'].map((cell) => ({ sheet: 'One', cell }))
    const w = await mountSheet(memorySource(book()), { watches: list })
    await open(w)
    expect(w.find('[data-xs-watch-delete]').attributes('disabled')).toBeDefined()
    const tr = (i: number) => w.findAll('[data-xs-watch-row]')[i]!
    await tr(0).trigger('click')
    await tr(2).trigger('click', { ctrlKey: true, metaKey: true })
    await w.find('[data-xs-watch-delete]').trigger('click')
    await settle()
    expect(rows(w).map((r) => r[3])).toEqual(['B1', 'B2'])
    await tr(0).trigger('click')
    await tr(1).trigger('click', { shiftKey: true })
    await w.find('[data-xs-watch-list]').trigger('keydown', { key: 'Delete' })
    await settle()
    expect(rows(w)).toEqual([])
    expect(lists(w).map((l) => l.map((x) => x.cell))).toEqual([['B1', 'B2'], []])
  })

  it('a double click goes to the watched cell on its sheet; the host replaces the list, and the window closes', async () => {
    const w = await mountSheet(memorySource(book()), { watches: [{ sheet: 'Two sheet', cell: 'C3' }] })
    await open(w)
    await w.find('[data-xs-watch-row]').trigger('dblclick')
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('Two sheet')
    expect(active(w)).toBe('3:3')
    await w.setProps({ watches: [{ sheet: 'One', cell: 'A2' }] })
    await settle()
    expect(rows(w)).toEqual([['', 'One', '', 'A2', '3', '']])
    expect(lists(w)).toEqual([])
    expect(api(w).watchWindow(false)).toBe(false)
    await settle()
    expect(w.find('[data-xs-watch]').exists()).toBe(false)
    await open(w)
    await w.find('[data-xs-watch-close]').trigger('click')
    await settle()
    expect(w.find('[data-xs-watch]').exists()).toBe(false)
  })

  it(`watches at most ${WATCH_MAX} cells, with a message`, async () => {
    const w = await mountSheet(memorySource(book()))
    await open(w)
    await goTo(w, 'A1:J120')
    await w.find('[data-xs-watch-add]').trigger('click')
    await settle()
    expect(lists(w)[0]!.length).toBe(WATCH_MAX)
    expect(w.find('.xs-message').text()).toBe(`Only the first ${WATCH_MAX} cells are watched`)
  })
})
