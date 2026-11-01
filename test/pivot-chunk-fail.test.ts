// A PivotTable chunk that fails to load: the dialog that waited for it is dropped rather than popping up later, and the
// next need (here a PivotTable added to the workbook) loads it again.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, mountSheet, settle, until } from './helpers'
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

describe('a PivotTable chunk that fails to load', () => {
  it('drops the dialog that waited for it and loads again on the next need', async () => {
    const s = localSource(salesBook(), { calc: 'inline' })
    sources.push(s)
    const w = await mountSheet(s, { style: 'height: 400px' })
    await goTo(w, 'Data!B2')
    await until(() => lastState(w)?.canInsertPivot === true)
    expect(await api(w).insertPivotTable()).toBe(true)
    await settle()
    expect(chunk.loads).toBe(1)
    expect(dialog()).toBeNull()
    // a PivotTable added meanwhile loads the chunk again, without the dropped dialog
    expect(await s.structure(addSales({ rows: [0], values: [{ field: 3 }] }))).toMatchObject({ ok: true })
    // the sheet reads the workbook again with it first
    await settle()
    await goTo(w, 'R!B4')
    await until(() => !!document.querySelector('[data-pivot-button]'))
    expect(chunk.loads).toBe(2)
    await settle()
    expect(dialog()).toBeNull()
    await goTo(w, 'Data!B2')
    expect(await api(w).insertPivotTable()).toBe(true)
    await until(() => !!dialog())
    expect(document.activeElement).toBe(document.querySelector('[data-xs-pivot-source]'))
  })
})
