// A PivotTable chunk that fails to load: the dialog that waited for it is dropped rather than popping up later, what the
// sheet shows (a PivotTable, moves over it) does not load it again, and the next request of the user's does.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, key, mountSheet, settle, until } from './helpers'
import { addSales, salesBook } from './pivot-fixture'

const chunk = vi.hoisted(() => ({ loads: 0 }))
vi.mock('../src/components/pivot-ui', async (original) => {
  if (++chunk.loads === 1) throw new Error('offline')
  return original()
})

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
const dialog = () => document.querySelector<HTMLElement>('[data-xs-pivot-dialog]')
const menu = () => document.querySelector<HTMLElement>('[data-pivot-filter-menu]')

describe('a PivotTable chunk that fails to load', () => {
  it('drops the dialog that waited for it, does not load again for a PivotTable or moves over it, and loads on the next request', async () => {
    const s = localSource(salesBook(), { calc: 'inline' })
    sources.push(s)
    const w = await mountSheet(s, { style: 'height: 400px', pivotFields: false })
    await goTo(w, 'Data!B2')
    await until(() => lastState(w)?.canInsertPivot === true)
    expect(await api(w).insertPivotTable()).toBe(true)
    await settle()
    expect(chunk.loads).toBe(1)
    expect(dialog()).toBeNull()
    // a PivotTable added meanwhile, and the active cell moved over it with the keys: no load, the keys stay on the grid
    expect(await s.structure(addSales({ rows: [0], values: [{ field: 3 }] }))).toMatchObject({ ok: true })
    // the sheet reads the workbook again with it first
    await settle()
    await goTo(w, 'R!B4')
    await until(() => lastState(w)?.canEditPivot === true)
    const grid = w.find('.xs-scroller').element as HTMLElement
    grid.focus()
    for (const k of ['ArrowDown', 'ArrowUp', 'ArrowLeft']) {
      await key(w, k)
      await settle()
    }
    expect(lastState(w)?.ref).toBe('A4')
    expect(chunk.loads).toBe(1)
    expect(document.activeElement).toBe(grid)
    // Alt+Down on Row Labels asks for the filter menu: the chunk loads, without the dropped dialog
    await key(w, 'ArrowUp')
    await settle()
    await key(w, 'ArrowDown', { altKey: true })
    await until(() => !!menu())
    expect(chunk.loads).toBe(2)
    expect(dialog()).toBeNull()
    expect(document.activeElement).toBe(menu()!.querySelector('[data-pfm-first]'))
  })
})
