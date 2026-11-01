// A PivotTable's filter menu asked for while the PivotTable chunk loads: the chunk starts loading with the sheet's own list of
// PivotTables, and Alt+Down on a filter button's cell opens the menu under the button the chunk draws, focused, Esc closing it.
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SheetSource } from '../src'
import { localSource } from '../src/local'
import { goTo, key, mountSheet, settle, until } from './helpers'
import { addSales, salesBook } from './pivot-fixture'
import { lastState, q, sales } from './pivot-sheet'

const chunk = vi.hoisted(() => {
  let arrive = () => {}
  const arrived = new Promise<void>((resolve) => (arrive = resolve))
  return { loads: 0, arrived, arrive: () => arrive() }
})
vi.mock('../src/components/pivot-ui', async (original) => {
  chunk.loads++
  await chunk.arrived
  return original()
})
afterEach(() => vi.restoreAllMocks())

const menu = () => q('[data-pivot-filter-menu]')

describe('the PivotTable chunk while it loads', () => {
  it("starts loading with the sheet's list of PivotTables, before their definitions and cells are read", async () => {
    const s = localSource(salesBook(), { calc: 'inline' })
    expect(await s.structure(addSales({ rows: [0], values: [{ field: 3 }] }))).toMatchObject({ ok: true })
    // the PivotTable's definition and cells never come, the active cell is off it
    const never = () => new Promise<never>(() => {})
    const held: SheetSource = { ...s, pivotTable: never, pivotCells: never }
    const w = await mountSheet(held, { style: 'height: 500px', pivotFields: false })
    await until(() => lastState(w)?.sheet === 'Data')
    expect(chunk.loads).toBe(0)
    await goTo(w, 'R!A1')
    await until(() => lastState(w)?.sheet === 'R')
    expect(chunk.loads).toBe(1)
    w.unmount()
    await s.close()
  })

  it("Alt+Down on a filter button's cell opens the menu under the button once the chunk arrives, focused, and Esc closes it", async () => {
    const { w } = await sales({ rows: [0], values: [{ field: 3 }] }, 'R!A3')
    expect(menu()).toBeNull()
    expect(q('[data-pivot-button="A3"]')).toBeNull()
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      return this.matches('[data-pivot-button="A3"]') ? new DOMRect(200, 100, 16, 16) : new DOMRect(0, 0, 0, 0)
    })
    await key(w, 'ArrowDown', { altKey: true })
    await settle()
    expect(menu()).toBeNull()
    chunk.arrive()
    await until(() => !!menu())
    await settle()
    expect(q('[data-pivot-button="A3"]')).not.toBeNull()
    expect([menu()!.style.left, menu()!.style.top]).toEqual(['200px', '116px'])
    expect(document.activeElement).toBe(menu()!.querySelector('[data-pfm-first]'))
    document.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    await settle()
    expect(menu()).toBeNull()
    expect(document.activeElement).toBe(w.find('.xs-scroller').element)
  })
})
