// Excel's Trace Dependents and Remove Arrows through the local source: arrows from the active cell to its dependents, a dashed
// arrow to a worksheet icon for other sheets, the next level on each call, arrows over frozen panes, and removal on an edit.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { memorySource, type SheetViewApi } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, key, mountSheet, settle, type } from './helpers'

let open: LocalSource[] = []
afterEach(async () => {
  await Promise.all(open.map((s) => s.close()))
  open = []
  document.body.innerHTML = ''
})

/** A1 feeds B1 and a cell on another sheet, B1 feeds C1, C1 feeds D1. */
function book(freeze?: { rows: number }) {
  const s = localSource(
    {
      locale: 'en-US',
      sheets: [
        { name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1*2' }, C1: { formula: 'B1+1' }, D1: { formula: 'C1+1' }, A5: { value: 3 } }, ...(freeze ? { frozen: freeze } : {}) },
        { name: 'Other', cells: { B2: { formula: 'S!A1+1' } } },
      ],
    },
    { calc: 'inline' },
  )
  open.push(s)
  return s
}
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const arrows = (w: VueWrapper) => w.findAll('[data-xs-arrow]').map((a) => a.attributes('data-xs-arrow'))

describe('Trace Dependents', () => {
  it('draws arrows to the direct dependents, then the next level on each call, and Remove Arrows clears them', async () => {
    const w = await mountSheet(book(), { style: 'height: 400px' })
    await settle(12)
    expect(await api(w).traceDependents()).toBe(true)
    await settle()
    expect(arrows(w).sort()).toEqual(['A1>*', 'A1>B1'])
    // the arrow to another sheet is dashed and ends at the worksheet icon
    expect(w.find('[data-xs-arrow="A1>*"]').classes()).toContain('xs-trace-sheet')
    expect(w.find('[data-xs-arrow="A1>*"] .xs-trace-icon').exists()).toBe(true)
    expect(w.find('[data-xs-arrow="A1>B1"]').classes()).toContain('xs-trace-cell')
    expect(await api(w).traceDependents()).toBe(true)
    await settle()
    expect(arrows(w).sort()).toEqual(['A1>*', 'A1>B1', 'B1>C1'])
    expect(await api(w).traceDependents()).toBe(true)
    expect(await api(w).traceDependents()).toBe(false)
    await settle()
    expect(arrows(w)).toHaveLength(4)
    expect(api(w).removeArrows()).toBe(true)
    await settle()
    expect(arrows(w)).toEqual([])
    expect(api(w).removeArrows()).toBe(false)
  })

  it('draws from the centre of the cell to the centre of its dependent, in the grid content', async () => {
    const w = await mountSheet(book(), { style: 'height: 400px' })
    await settle(12)
    await api(w).traceDependents()
    await settle()
    const line = w.find('[data-xs-arrow="A1>B1"] line')
    // Columns are 71 px after the 56 px row numbers, rows 20 px below the 24 px column letters.
    expect([line.attributes('x1'), line.attributes('y1'), line.attributes('x2'), line.attributes('y2')].map(Number)).toEqual([56 + 35.5, 24 + 10, 56 + 71 + 35.5, 24 + 10])
    expect(w.find('[data-xs-trace="body"]').exists()).toBe(true)
  })

  it('an arrow with an end in frozen panes goes over them and follows the scroll', async () => {
    const w = await mountSheet(book({ rows: 1 }), { style: 'height: 400px' })
    await settle(12)
    await api(w).traceDependents()
    await settle()
    expect(w.find('[data-xs-trace="frozen"] [data-xs-arrow="A1>B1"]').exists()).toBe(true)
    const scroller = w.find('.xs-scroller')
    scroller.element.scrollTop = 60
    await scroller.trigger('scroll')
    await settle()
    expect(Number(w.find('[data-xs-arrow="A1>B1"] line').attributes('y1'))).toBe(60 + 24 + 10)
  })

  it('an edit removes the arrows, as in Excel', async () => {
    const w = await mountSheet(book(), { style: 'height: 400px' })
    await settle(12)
    await api(w).traceDependents()
    await settle()
    expect(arrows(w).length).toBeGreaterThan(0)
    await goTo(w, 'A5')
    await key(w, 'F2')
    await type(w, '4')
    await w.find('.xs-editor').trigger('keydown', { key: 'Enter' })
    await settle(12)
    expect(arrows(w)).toEqual([])
  })

  it('a cell no formula refers to: no arrows and Excel’s message; a source without `dependents` traces nothing', async () => {
    const w = await mountSheet(book(), { style: 'height: 400px' })
    await settle(12)
    await goTo(w, 'A5')
    expect(await api(w).traceDependents()).toBe(false)
    await settle()
    expect(w.find('.xs-message').text()).toContain('found no formulas')
    const plain = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    expect(await api(plain).traceDependents()).toBe(false)
  })
})
