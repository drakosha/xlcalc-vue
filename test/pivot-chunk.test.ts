// The PivotTable, slicer and timeline parts come as a chunk of their own: a workbook without them never loads it, and a
// dialog asked for while it loads opens, focused, once it arrives. The chunk here waits for the test to let it in.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, mountSheet, settle, until } from './helpers'
import { salesBook } from './pivot-fixture'

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

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
const dialog = () => document.querySelector<HTMLElement>('[data-xs-pivot-dialog]')

describe('the PivotTable chunk', () => {
  it('is not loaded for a workbook without PivotTables, slicers and timelines; Insert PivotTable while it loads opens once it arrives, focused', async () => {
    const s = localSource(salesBook(), { calc: 'inline' })
    sources.push(s)
    const w = await mountSheet(s, { style: 'height: 400px' })
    await goTo(w, 'Data!B2')
    await until(() => lastState(w)?.canInsertPivot === true)
    expect(chunk.loads).toBe(0)
    expect(await api(w).insertPivotTable()).toBe(true)
    await settle()
    // the request starts the load and waits in the sheet's state, with nothing drawn yet
    expect(chunk.loads).toBe(1)
    expect(dialog()).toBeNull()
    chunk.arrive()
    await until(() => !!dialog())
    expect(document.activeElement).toBe(document.querySelector('[data-xs-pivot-source]'))
    expect(document.querySelector<HTMLInputElement>('[data-xs-pivot-source]')!.value).toBe('Data!$A$1:$E$5')
    // OK puts the PivotTable on a new sheet, whose filter buttons come from the chunk too
    document.querySelector<HTMLButtonElement>('[data-xs-dialog-ok]')!.click()
    await until(() => !!lastState(w)?.pivot)
    expect(dialog()).toBeNull()
    expect(chunk.loads).toBe(1)
  })
})
