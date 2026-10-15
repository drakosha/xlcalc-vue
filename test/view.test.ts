// Drawing: the book's styles, General alignment, merges, hidden rows, columns and sheets, links, notes, tables, notice,
// overflow, the freeze lock and its memory, and the column width with the mouse.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { mount } from '@vue/test-utils'
import { SheetView, memorySource, type MemoryBook, type SheetViewApi } from '../src'
import { colPx, widthFromPx } from '../src/core/geometry'
import { cellText, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

const td = (w: VueWrapper, r: number, c: number) => w.find(`[data-cell="${r}:${c}"]`)
const css = (w: VueWrapper, r: number, c: number) => (td(w, r, c).element as HTMLElement).style
/** A side's line as `width style color`, from the longhands the cell sets. */
const side = (w: VueWrapper, r: number, c: number, s: 'Top' | 'Right' | 'Bottom' | 'Left') => {
  const st = css(w, r, c) as unknown as Record<string, string>
  return [st[`border${s}Width`], st[`border${s}Style`], st[`border${s}Color`]].filter(Boolean).join(' ')
}
const box = (w: VueWrapper, r: number, c: number) => (td(w, r, c).find('.xs-in').element as HTMLElement).style

describe('cell style', () => {
  it('draws font, fill and borders, and aligns General by the value', async () => {
    const book: MemoryBook = {
      styles: [
        {},
        { font: { bold: true, italic: true, underline: 'double', color: '#112233', size: 22 }, fill: { color: '#ffff00' } },
        { border: { bottom: { style: 'thick', color: '#ff0000' }, right: { style: 'dashed' } } },
        { border: { left: { style: 'medium', color: '#00ff00' } } },
        { alignment: { horizontal: 'center', vertical: 'top', wrap: true, indent: 1 } },
      ],
      sheets: [{ name: 'S', cells: { A1: { value: 'styled', style: 1 }, A2: { value: 5, style: 2 }, B2: { value: 'x', style: 3 }, A3: 7, B3: 'text', C3: true, A4: { value: 'c', style: 4 } } }],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    const a1 = css(w, 1, 1)
    expect(a1.fontWeight).toBe('700')
    expect(a1.fontStyle).toBe('italic')
    expect(a1.color).toBe('#112233')
    expect(a1.fontSize).toBe('25px')
    expect(a1.backgroundColor).toBe('#ffff00')
    expect(side(w, 2, 1, 'Bottom')).toBe('3px solid #ff0000')
    // One line between A2 and B2: A2's own right side wins over B2's left.
    expect(side(w, 2, 1, 'Right')).toBe('1px dashed var(--xs-border-auto)')
    // B2's own left side is drawn by its neighbour A2 (which has a right side of its own here), never by B2 itself.
    expect(side(w, 2, 2, 'Left')).toBe('')
    expect(box(w, 3, 1).textAlign).toBe('right')
    expect(box(w, 3, 2).textAlign).toBe('left')
    expect(box(w, 3, 3).textAlign).toBe('center')
    expect(box(w, 4, 1).textAlign).toBe('center')
    expect(box(w, 4, 1).paddingLeft).toBe('')
    expect(td(w, 4, 1).find('.xs-in').classes()).toEqual(expect.arrayContaining(['xs-wrap', 'xs-top']))
    expect(td(w, 3, 1).find('.xs-in').classes()).toContain('xs-bot')
    // The row's height holds the thick bottom line: the box is the rest.
    expect(box(w, 2, 1).height).toBe(`${20 - 2 - 3}px`)
  })

  it('takes the left neighbour line and hides grid lines under a fill', async () => {
    const book: MemoryBook = {
      styles: [{}, { border: { left: { style: 'medium', color: '#00ff00' } } }, { fill: { color: '#abcdef' } }],
      sheets: [{ name: 'S', cells: { A1: 'a', B1: { value: 'b', style: 1 }, A3: { value: 'f', style: 2 } } }],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(side(w, 1, 1, 'Right')).toBe('2px solid #00ff00')
    expect(side(w, 3, 1, 'Right')).toBe('1px solid #abcdef')
    expect(side(w, 2, 1, 'Bottom')).toBe('1px solid #abcdef')
  })

  it('shows the section colour, a data bar, a note and a table header', async () => {
    const book: MemoryBook = {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: -1, text: '(1)', color: '#ff0000', note: { author: 'Ann', text: 'Check it' } },
            B1: { value: 3, bar: { ratio: 0.25, color: '#638ec6' } },
            B3: 'Name',
            B4: 'x',
          },
          tables: [{ name: 'T', range: 'B3:C4', headerRow: 3 }],
          notice: 'Values computed elsewhere',
        },
      ],
    }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(css(w, 1, 1).color).toBe('#ff0000')
    expect(td(w, 1, 1).text()).toBe('(1)')
    expect(css(w, 1, 1).backgroundImage).toContain('var(--xs-note)')
    expect(td(w, 1, 1).attributes('title')).toContain('Note (Ann): Check it')
    // an old-style bar is drawn as an Excel bar of the window's appearance would be: from the left edge to its share
    expect(css(w, 1, 2).backgroundImage).toContain('linear-gradient(to right, transparent 0%, #638ec6 0%, color-mix(in srgb, #638ec6 18%, white) 25%, transparent 25%)')
    expect(td(w, 3, 2).classes()).toContain('xs-thead')
    expect(td(w, 4, 2).classes()).not.toContain('xs-thead')
    expect(css(w, 3, 2).boxShadow).toContain('var(--xs-table-line)')
    expect(w.find('.xs-notice').text()).toBe('Values computed elsewhere')
  })

  it('lets text overflow into empty neighbours, but not numbers or wrapped text', async () => {
    const book: MemoryBook = { styles: [{}, { alignment: { wrap: true } }], sheets: [{ name: 'S', cells: { A1: 'a long title', A2: 12345, A3: { value: 'wrapped', style: 1 }, A4: 'stops', B4: 'x' } }] }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(td(w, 1, 1).classes()).toContain('xs-spill')
    expect(td(w, 1, 1).find('.xs-ov').exists()).toBe(true)
    expect(td(w, 2, 1).classes()).not.toContain('xs-spill')
    expect(td(w, 3, 1).classes()).not.toContain('xs-spill')
    expect(td(w, 4, 1).classes()).not.toContain('xs-spill')
  })
})

describe('shape of the sheet', () => {
  it('draws a merge as one cell, skipping hidden rows and columns', async () => {
    const book: MemoryBook = { sheets: [{ name: 'S', cells: { A1: 'merged', E5: 'end' }, merges: ['A1:C3'], widths: { B: 0 }, heights: { '2': 0 } }] }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(td(w, 1, 1).attributes('colspan')).toBe('2')
    expect(td(w, 1, 1).attributes('rowspan')).toBe('2')
    expect(td(w, 1, 3).exists()).toBe(false)
    expect(td(w, 3, 1).exists()).toBe(false)
    expect(td(w, 4, 2).exists()).toBe(false)
    expect(w.findAll('.xs-num').map((n) => n.text())).toEqual(['1', '3', '4', '5'])
    expect(w.findAll('.xs-head').map((n) => n.text())).toEqual(['A', 'C', 'D', 'E'])
  })

  it('hides hidden sheets from the tabs and opens the first visible one', async () => {
    const book: MemoryBook = { sheets: [{ name: 'Hidden', hidden: true, cells: { A1: 'h' } }, { name: 'One', cells: { A1: 'one' } }, { name: 'Two', cells: { A1: 'two' } }] }
    const states: unknown[] = []
    const w = await mountSheet(memorySource(book), { readOnly: true, onState: (s: unknown) => states.push(s) })
    expect(w.findAll('.xs-tab').map((t) => t.text())).toEqual(['One', 'Two'])
    expect(td(w, 1, 1).text()).toBe('one')
    expect(states.at(-1)).toEqual({ sheet: 'One', ref: 'A1', value: 'one' })
  })

  it('takes column titles and no row numbers from SheetInfo (the ground for list mode)', async () => {
    const book: MemoryBook = { sheets: [{ name: 'L', rowNumbers: false, columns: [{ title: 'Name' }, { title: 'Sum' }], cells: { A1: 'a', B1: 1 } }] }
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(w.findAll('.xs-head').map((h) => h.text())).toEqual(['Name', 'Sum'])
    expect(w.find('.xs-num').exists()).toBe(false)
    expect(w.find('.xs-corner').exists()).toBe(false)
  })

  it('shows "the sheet is empty" for an empty read-only sheet, and a grid for an editable one', async () => {
    const book: MemoryBook = { sheets: [{ name: 'S', cells: {} }] }
    const ro = await mountSheet(memorySource(book), { readOnly: true, messages: { empty: 'Nothing here' } })
    expect(ro.find('.xs-empty').text()).toBe('Nothing here')
    ro.unmount()
    const ed = await mountSheet(memorySource(book))
    expect(ed.findAll('.xs-head').length).toBeGreaterThan(10)
    expect(ed.find('[data-cell="1:1"]').exists()).toBe(true)
  })
})

describe('props and slots', () => {
  it('a book of no sheets stops reading and shows an empty frame, not "Reading the book…" for ever', async () => {
    const w = await mountSheet(memorySource({ sheets: [] }))
    await settle()
    expect(w.find('.xs-loading').exists()).toBe(false)
    expect(w.findAll('.xs-tab')).toHaveLength(0)
  })

  it('formulaBar: false hides the bar', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }), { formulaBar: false })
    expect(w.find('.xs-bar').exists()).toBe(false)
    expect(w.find('[data-cell="1:1"]').text()).toBe('1')
  })

  it('the fallback slot gets the source error and a retry', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] })
    const sheets = source.sheets
    let fail = true
    source.sheets = (signal) => (fail ? Promise.reject(new Error('gone')) : sheets(signal))
    const w = mount(SheetView, {
      props: { source },
      slots: { fallback: `<template #fallback="{ failure, retry }"><button class="mine" @click="retry">{{ failure.text }}</button></template>` },
      attachTo: document.body,
    })
    await settle()
    expect(w.find('.mine').text()).toBe('gone')
    fail = false
    await w.find('.mine').trigger('click')
    await settle()
    expect(w.find('.mine').exists()).toBe(false)
    expect(w.find('[data-cell="1:1"]').text()).toBe('1')
  })
})

describe('links', () => {
  const book: MemoryBook = {
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 'site', link: { href: 'https://example.org/' } },
          A2: { value: 'place', link: { location: "'Other sheet'!B7" } },
          A3: { value: 'file', link: { target: 'file:1' } },
          A4: { value: 'bad', link: { href: 'javascript:alert(1)' } },
        },
      },
      { name: 'Other sheet', cells: { B7: 'here' } },
    ],
  }
  it('opens a web link as a link; a place in the book itself; a host target by `openLink`', async () => {
    const opened: unknown[] = []
    const w = await mountSheet(memorySource(book), { readOnly: true, onOpenLink: (l: unknown) => opened.push(l) })
    expect(td(w, 1, 1).find('a.xs-link').attributes('href')).toBe('https://example.org/')
    expect(td(w, 4, 1).find('a').exists()).toBe(false)
    await td(w, 3, 1).find('button.xs-link').trigger('click')
    expect(opened).toEqual([{ target: 'file:1', sheet: 'S', cell: 'A3' }])
    await td(w, 2, 1).find('button.xs-link').trigger('click')
    await settle()
    expect(w.find('.xs-tab-on').text()).toBe('Other sheet')
    expect(w.find('.xs-cell.xs-active').attributes('data-cell')).toBe('7:2')
  })
  it('without an `openLink` listener a host target or an unsafe link is only shown', async () => {
    const w = await mountSheet(memorySource(book), { readOnly: true })
    expect(td(w, 3, 1).find('.xs-deadlink').exists()).toBe(true)
    expect(td(w, 4, 1).find('.xs-deadlink').exists()).toBe(true)
  })
})

describe('freeze', () => {
  const book = (rows: number): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 'h', E80: 'end' }, freeze: { rows, cols: 1 } }] })
  it('sticks the frozen panes; the lock drops them and the choice is remembered by storageKey', async () => {
    const w = await mountSheet(memorySource(book(2)), { readOnly: true, storageKey: 'k' })
    expect(td(w, 1, 1).classes()).toContain('xs-fz-corner')
    expect(td(w, 2, 2).classes()).toContain('xs-fz-row')
    expect(td(w, 3, 1).classes()).toContain('xs-fz-col')
    await w.find('.xs-tools .xs-tool').trigger('click')
    expect(td(w, 1, 1).classes()).not.toContain('xs-fz-corner')
    expect(localStorage.getItem('k:freeze:S')).toBe('false')
    w.unmount()
    const again = await mountSheet(memorySource(book(2)), { readOnly: true, storageKey: 'k' })
    expect(td(again, 1, 1).classes()).not.toContain('xs-fz-corner')
  })
  it('drops by itself when the frozen rows eat over half the frame', async () => {
    const w = await mountSheet(memorySource(book(20)), { readOnly: true })
    expect(td(w, 1, 1).classes()).not.toContain('xs-fz-corner')
    expect(w.find('.xs-tools .xs-tool').exists()).toBe(true)
  })
})

describe('column width with the mouse', () => {
  async function drag(w: VueWrapper, col: number, dx: number): Promise<void> {
    const grip = w.findAll('.xs-head')[col - 1]!.find('.xs-grip')
    await grip.trigger('pointerdown', { clientX: 100, pointerId: 1 })
    await grip.trigger('pointermove', { clientX: 100 + dx, pointerId: 1 })
    await grip.trigger('pointerup', { clientX: 100 + dx, pointerId: 1 })
    await settle()
  }
  const colWidths = (w: VueWrapper) => w.findAll('col').map((c) => (c.element as HTMLElement).style.width)

  it('without `layout` the browser remembers it; a double click restores the book width', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { C3: 'x' } }] })
    delete (source as { layout?: unknown }).layout
    const w = await mountSheet(source, { storageKey: 'k' })
    await drag(w, 2, 100 - colPx(null))
    expect(colWidths(w)[2]).toBe('100px')
    expect(JSON.parse(localStorage.getItem('k:widths:S')!)).toEqual({ 2: 100 })
    await w.findAll('.xs-head')[1]!.find('.xs-grip').trigger('dblclick')
    expect(colWidths(w)[2]).toBe(`${colPx(null)}px`)
    expect(localStorage.getItem('k:widths:S')).toBeNull()
  })

  it('with `layout` the width goes to the source in its units; edgeRestoresWidth: a double click restores the book width', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { C3: 'x' } }] })
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source, { storageKey: 'k', edgeRestoresWidth: true })
    await drag(w, 2, 100 - colPx(null))
    expect(layout).toHaveBeenCalledWith({ sheet: 'S', columns: 'B', width: widthFromPx(100) })
    expect(localStorage.getItem('k:widths:S')).toBeNull()
    await w.findAll('.xs-head')[1]!.find('.xs-grip').trigger('dblclick')
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'B', width: widthFromPx(colPx(null)) })
  })

  it("a width set with the mouse shows the cells the source says it changed: a date that did not fit shows itself", async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A5: '#####' } }] })
    vi.spyOn(source, 'layout').mockResolvedValue({ ok: true, changes: [{ sheet: 'S', cell: 'A5', text: '15.03.2026', value: 46096 }] })
    const w = await mountSheet(source)
    expect(cellText(w, 5, 1)).toBe('#####')
    await drag(w, 1, 40)
    expect(cellText(w, 5, 1)).toBe('15.03.2026')
  })

  it('edgeRestoresWidth: after a structure operation a double click does not give a column the width remembered for the one it replaced', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { C3: 'x' }, widths: { C: 20 } }] })
    const layout = vi.spyOn(source, 'layout')
    source.structure = vi.fn(async () => ({ ok: true as const, changes: [] }))
    const w = await mountSheet(source, { edgeRestoresWidth: true })
    await drag(w, 3, 40)
    await (w.vm as unknown as SheetViewApi).structure({ op: 'insertColumns', sheet: 'S', column: 'B' })
    await settle()
    layout.mockClear()
    // C is now a new column; the remembered width was the old C's, which is D.
    await w.findAll('.xs-head')[2]!.find('.xs-grip').trigger('dblclick')
    await settle()
    expect(layout).not.toHaveBeenCalled()
  })

  it('a double click on an edge fits the column through `layout`, as Excel; every selected whole column when it is one of them', async () => {
    const source = memorySource({ sheets: [{ name: 'S', cells: { C3: 'x' } }] })
    const layout = vi.spyOn(source, 'layout')
    const sheets = vi.spyOn(source, 'sheets')
    const w = await mountSheet(source)
    const grip = (col: number) => w.findAll('.xs-head')[col - 1]!.find('.xs-grip')
    const reads = sheets.mock.calls.length
    await grip(2).trigger('dblclick')
    await settle()
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'B', fit: true })
    // The source tells nothing: the sheet reads its widths again itself.
    expect(sheets.mock.calls.length).toBe(reads + 1)
    await w.findAll('.xs-head')[1]!.trigger('click')
    await w.findAll('.xs-head')[3]!.trigger('click', { shiftKey: true })
    await grip(3).trigger('dblclick')
    await settle()
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'B:D', fit: true })
    await grip(5).trigger('dblclick')
    await settle()
    expect(layout).toHaveBeenLastCalledWith({ sheet: 'S', columns: 'E', fit: true })
    expect(layout).toHaveBeenCalledTimes(3)
  })

  it("a column without a width of its own is the library's default, 8.83203125 (53 points, 71 px)", async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 'a' }, widths: { B: 8.83203125 } }] }))
    expect(colWidths(w)[1]).toBe('71px')
    expect(colWidths(w)[2]).toBe('71px')
  })
})
