// List mode: a host's list on the sheet. Sortable titles sort through the host, drill-down cells and the first column
// open what they lead to by a click and Enter, the totals row stays at the bottom, and `reload` reads the host's new order.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type MemoryBook, type SheetInfo, type SheetSource } from '../src'
import { gridWindow, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

const td = (w: VueWrapper, r: number, c: number) => w.find(`[data-cell="${r}:${c}"]`)
const active = (w: VueWrapper) => w.find('.xs-cell.xs-active').attributes('data-cell')
const key = (w: VueWrapper, k: string, extra: Record<string, unknown> = {}) => w.find('.xs-scroller').trigger('keydown', { key: k, ...extra })
const head = (w: VueWrapper, c: number) => w.find(`th.xs-head[data-col="${c}"]`)

/** A view of three rows under a frozen header row: Name opens rows, Sum is sortable and drills down, totals at the bottom. */
const book = (over: Partial<MemoryBook['sheets'][number]> = {}): MemoryBook => ({
  sheets: [
    {
      name: 'View',
      rowNumbers: false,
      freeze: { rows: 1, cols: 1 },
      columns: [{ title: 'Name' }, { title: 'Sum', align: 'right', sortable: true }, { title: 'Note' }],
      totals: ['Total', '60', null],
      cells: {
        A1: 'Name',
        B1: 'Sum',
        A2: 'a',
        B2: { value: 10, drill: true },
        C2: 'x',
        A3: 'b',
        B3: { value: 20, drill: true },
        A4: 'c',
        B4: { value: 30, drill: true, text: '' },
      },
      ...over,
    },
  ],
})

describe('sortable titles', () => {
  it('a click on a sortable title asks the host to sort, ascending first and descending on the column already ascending', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    expect(head(w, 2).classes()).toContain('xs-sortable')
    expect(head(w, 2).find('.xs-sorthint').exists()).toBe(true)
    await head(w, 2).trigger('click')
    expect(w.emitted('sort')).toEqual([[{ col: 2, dir: 'asc' }]])
    await w.setProps({ sort: { col: 2, dir: 'asc' } })
    expect(head(w, 2).find('.xs-list-sort').text()).toBe('↑')
    expect(head(w, 2).attributes('aria-sort')).toBe('ascending')
    await head(w, 2).trigger('click')
    expect(w.emitted('sort')![1]).toEqual([{ col: 2, dir: 'desc' }])
    await w.setProps({ sort: { col: 2, dir: 'desc' } })
    expect(head(w, 2).find('.xs-list-sort').text()).toBe('↓')
  })

  it('a title that is not sortable selects its column as before, and the titles take their alignment', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await head(w, 3).trigger('click')
    expect(w.emitted('sort')).toBeUndefined()
    expect(td(w, 2, 3).classes()).toContain('xs-sel')
    expect((head(w, 2).find('.xs-head-label').element as HTMLElement).style.textAlign).toBe('right')
    expect((head(w, 1).find('.xs-head-label').element as HTMLElement).style.textAlign).toBe('')
  })
})

describe('drill-down and rows', () => {
  it('a drill-down cell is marked, and a click or Enter on it sends cellDrill; an empty one is plain', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    expect(td(w, 2, 2).classes()).toContain('xs-drill')
    expect(td(w, 4, 2).classes()).not.toContain('xs-drill')
    await td(w, 3, 2).trigger('pointerdown', { button: 0 })
    await td(w, 3, 2).trigger('click')
    expect(w.emitted('cellDrill')).toEqual([[{ sheet: 'View', cell: 'B3' }]])
    await key(w, 'Enter')
    expect(w.emitted('cellDrill')![1]).toEqual([{ sheet: 'View', cell: 'B3' }])
    expect(active(w)).toBe('3:2')
    expect(w.emitted('rowOpen')).toBeUndefined()
  })

  it('with openRows the first column opens its row by a click, Enter on any other cell too; the header row does not', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true, openRows: true })
    expect(td(w, 2, 1).classes()).toContain('xs-rowlink')
    expect(td(w, 2, 1).attributes('title')).toBe('Open the row')
    expect(td(w, 1, 1).classes()).not.toContain('xs-rowlink')
    await td(w, 3, 1).trigger('click')
    expect(w.emitted('rowOpen')).toEqual([[{ sheet: 'View', row: 3 }]])
    await td(w, 2, 3).trigger('pointerdown', { button: 0 })
    await key(w, 'Enter')
    expect(w.emitted('rowOpen')![1]).toEqual([{ sheet: 'View', row: 2 }])
    expect(active(w)).toBe('2:3')
    await td(w, 1, 3).trigger('pointerdown', { button: 0 })
    await key(w, 'Enter')
    expect(w.emitted('rowOpen')).toHaveLength(2)
  })

  it('without openRows a click on the first column only selects, and Enter moves down as in a sheet', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    expect(td(w, 2, 1).classes()).not.toContain('xs-rowlink')
    await td(w, 2, 1).trigger('pointerdown', { button: 0 })
    await td(w, 2, 1).trigger('click')
    await key(w, 'Enter')
    expect(w.emitted('rowOpen')).toBeUndefined()
    expect(active(w)).toBe('3:1')
  })

  it('Shift and a click on a drill-down cell extend the selection and open nothing', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await td(w, 2, 2).trigger('click', { shiftKey: true })
    expect(w.emitted('cellDrill')).toBeUndefined()
  })
})

describe('the totals row', () => {
  it('shows the totals by column at the bottom, aligned as the titles, and nothing without them', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    const cells = w.findAll('td.xs-total')
    expect(cells.map((c) => c.text()).slice(0, 3)).toEqual(['Total', '60', ''])
    expect((cells[1]!.element as HTMLElement).style.textAlign).toBe('right')
    expect(cells[0]!.classes()).toContain('xs-fz-col')
    w.unmount()
    const plain = await mountSheet(memorySource(book({ totals: undefined })), { readOnly: true })
    expect(plain.find('tfoot').exists()).toBe(false)
  })
})

describe('the host sorts', () => {
  it('reload after the host changed its order reads the windows again in the new order', async () => {
    const names = ['b', 'c', 'a']
    let order = names.slice()
    let emit: (e: { kind: 'reload' }) => void = () => {}
    const info: SheetInfo = { name: 'L', maxRow: 3, maxCol: 1, rowNumbers: false, columns: [{ title: 'Name', sortable: true }] }
    const source: SheetSource = {
      sheets: async () => [info],
      window: async (req) => gridWindow(req, (r) => order[r - 1] ?? ''),
      subscribe: (l) => ((emit = l as typeof emit), () => {}),
    }
    const w = await mountSheet(source, {
      readOnly: true,
      onSort: (s: { dir: 'asc' | 'desc' }) => {
        order = names.slice().sort()
        if (s.dir === 'desc') order.reverse()
        emit({ kind: 'reload' })
      },
    })
    expect([1, 2, 3].map((r) => td(w, r, 1).text())).toEqual(['b', 'c', 'a'])
    await head(w, 1).trigger('click')
    await settle()
    expect([1, 2, 3].map((r) => td(w, r, 1).text())).toEqual(['a', 'b', 'c'])
  })
})
