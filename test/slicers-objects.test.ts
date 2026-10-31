// Slicers and timelines as objects of the sheet: drawn where the grid has no cells (a read-only sheet, beyond the columns
// drawn, a sheet of objects alone), Ctrl+Z on a box, the box of the last one inserted taking the keys, a cell made active
// leaving the object, the protection of the PivotTables' sheets, the display filter, offsets beyond their cell, and the
// timeline's level list and Shift+click.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import type { SheetViewApi, SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { goTo, mountSheet, settle, until } from './helpers'
import { slicerBook } from './slicer-fixture'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})
/** A column's width and a row's height in the grid's pixels without their own, and the row numbers' width and the letters' height. */
const COL = 71
const ROW = 20
const NUMS = 56
const HEAD = 24

async function open(book = slicerBook(), props: Record<string, unknown> = {}, at = 'Report!B4'): Promise<{ w: VueWrapper; s: LocalSource; api: SheetViewApi }> {
  const s = localSource(book, { calc: 'inline' })
  sources.push(s)
  const w = await mountSheet(s, props)
  await goTo(w, at)
  await until(() => !!document.querySelector('[data-xs-object]'))
  return { w, s, api: w.vm as unknown as SheetViewApi }
}
const q = <T extends Element = HTMLElement>(sel: string) => document.querySelector<T>(sel)
const box = (key: string) => q(`[data-xs-object="${key}"]`)
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
async function click(el: Element, init: MouseEventInit = {}): Promise<void> {
  el.dispatchEvent(new MouseEvent('click', { bubbles: true, ...init }))
  await settle()
}
async function press(el: Element, type: string, init: MouseEventInit = {}): Promise<void> {
  el.dispatchEvent(new MouseEvent(type, { bubbles: true, button: 0, ...init }))
  await settle()
}

describe('where the grid has no cells', () => {
  it('a read-only sheet draws the slicer and the timeline right of its data at their places and sizes', async () => {
    await open(slicerBook(), { readOnly: true })
    await until(() => !!box('timeline:Date'))
    expect(box('slicer:Region')!.style.left).toBe(`${NUMS + 3 * COL}px`)
    expect(box('slicer:Region')!.style.height).toBe(`${9 * ROW}px`)
    expect(box('timeline:Date')!.style.left).toBe(`${NUMS + 7 * COL}px`)
    expect(box('timeline:Date')!.style.width).toBe(`${6 * COL}px`)
  })

  it('an editable sheet draws a slicer beyond the columns it would draw, and a corner far below', async () => {
    const book = slicerBook()
    book.sheets[1]!.slicers![0]!.place = { from: 'AB3', to: 'AD300' }
    await open(book)
    expect(box('slicer:Region')!.style.left).toBe(`${NUMS + 27 * COL}px`)
    expect(box('slicer:Region')!.style.height).toBe(`${297 * ROW}px`)
  })

  it('a read-only sheet of a slicer alone, without cells, draws it', async () => {
    const book = slicerBook()
    const report = book.sheets[1]!
    book.sheets.push({ name: 'Board', cells: {}, slicers: report.slicers })
    report.slicers = []
    await open(book, { readOnly: true }, 'Board!A1')
    expect(box('slicer:Region')!.style.top).toBe(`${HEAD + 2 * ROW}px`)
  })

  it('an offset larger than its cell stops at the cell, as an anchor lies within its cell', async () => {
    const book = slicerBook()
    book.sheets[1]!.slicers![0]!.place = { from: 'D3', to: 'F12', fromOffset: [10 * 9525 * COL, 0] }
    book.sheets[1]!.columns = { D: { hidden: true } }
    await open(book)
    expect(box('slicer:Region')!.style.left).toBe(`${NUMS + 3 * COL}px`)
  })
})

describe('the box and the keys', () => {
  it('Ctrl+Z on the box undoes its choice, and Ctrl+Y does it again', async () => {
    await open()
    await click(q('[data-xs-slicer-item="1"]')!)
    await until(() => q('[data-xs-slicer-item="0"]')!.getAttribute('aria-selected') === 'false')
    box('slicer:Region')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', code: 'KeyZ', ctrlKey: true, bubbles: true, cancelable: true }))
    await until(() => q('[data-xs-slicer-item="0"]')!.getAttribute('aria-selected') === 'true')
    box('slicer:Region')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'y', code: 'KeyY', ctrlKey: true, bubbles: true, cancelable: true }))
    await until(() => q('[data-xs-slicer-item="0"]')!.getAttribute('aria-selected') === 'false')
  })

  it('Insert Slicers of two fields selects the last and gives it the keys, so Delete removes it', async () => {
    const { w, api } = await open()
    api.insertSlicer()
    await settle()
    await click(q('[data-xs-field="Kind"]')!)
    await click(q('[data-xs-field="Year"]')!)
    await click(q('[data-xs-dialog-ok]')!)
    await until(() => !!box('slicer:Kind') && !!box('slicer:Year') && lastState(w)?.object?.name === 'Year')
    await until(() => document.activeElement === box('slicer:Year'))
    document.activeElement!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true }))
    await until(() => !box('slicer:Year'))
    expect(box('slicer:Kind')).not.toBeNull()
  })

  it('an insert the source refuses selects nothing that appears later', async () => {
    const s = localSource(slicerBook(), { calc: 'inline' })
    sources.push(s)
    const structure = s.structure.bind(s)
    let refuse = true
    s.structure = (op) => (refuse && op.op === 'addSlicer' ? Promise.resolve({ ok: false, reason: 'refused', message: 'no' }) : structure(op))
    const w = await mountSheet(s)
    await goTo(w, 'Report!B4')
    await until(() => !!box('slicer:Region'))
    ;(w.vm as unknown as SheetViewApi).insertSlicer()
    await settle()
    await click(q('[data-xs-field="Kind"]')!)
    await click(q('[data-xs-dialog-ok]')!)
    await settle()
    q<HTMLButtonElement>('[data-xs-alert-button="ok"]')?.click()
    refuse = false
    await (w.vm as unknown as SheetViewApi).structure({ op: 'addSlicer', sheet: 'Report', slicer: {}, cache: { field: 'Kind', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } })
    await until(() => !!box('slicer:Kind'))
    expect(lastState(w)?.object).toBeUndefined()
  })

  it('a slicer with Disable resizing and moving has no handles', async () => {
    await open(slicerBook({ positionLocked: true }))
    await press(q('[data-xs-object="slicer:Region"] .xs-slicer-head')!, 'pointerdown')
    expect(document.querySelectorAll('[data-xs-handle]').length).toBe(0)
  })

  it('Insert Slicers puts them in the middle of the window in Excel\'s size, each next one lower and to the right', async () => {
    const { api } = await open()
    api.insertSlicer()
    await settle()
    await click(q('[data-xs-field="Kind"]')!)
    await click(q('[data-xs-field="Year"]')!)
    await click(q('[data-xs-dialog-ok]')!)
    await until(() => !!box('slicer:Kind') && !!box('slicer:Year'))
    const [kind, year] = [box('slicer:Kind')!, box('slicer:Year')!]
    expect([kind.style.width, kind.style.height]).toEqual(['192px', '265px'])
    expect(parseFloat(year.style.left) - parseFloat(kind.style.left)).toBeCloseTo(20, 0)
    expect(parseFloat(year.style.top) - parseFloat(kind.style.top)).toBeCloseTo(20, 0)
  })

  it('a cell made active by the Name Box leaves the slicer selected', async () => {
    const { w } = await open()
    await press(q('[data-xs-object="slicer:Region"] .xs-slicer-head')!, 'pointerdown')
    expect(lastState(w)?.object).toEqual({ kind: 'slicer', name: 'Region' })
    await goTo(w, 'Q20')
    expect(lastState(w)?.object).toBeUndefined()
  })
})

describe('protection and the display filter', () => {
  it("disables the items where the PivotTable's sheet does not allow PivotTables, whatever the slicer's own sheet allows", async () => {
    const book = slicerBook({ protection: {} })
    const report = book.sheets[1]!
    book.sheets.push({ name: 'Board', cells: { A1: { value: 'board' } }, slicers: report.slicers })
    report.slicers = []
    await open(book, {}, 'Board!A1')
    expect(q<HTMLButtonElement>('[data-xs-slicer-item="0"]')!.disabled).toBe(true)
  })

  it('no move or resize while the display filter orders the rows', async () => {
    const book = slicerBook()
    const report = book.sheets[1]!
    report.cells = { ...report.cells, P20: { value: 'n' }, P21: { value: 1 }, P22: { value: 3 }, P23: { value: 2 } }
    report.autoFilter = { range: 'P20:P23' }
    const { w, api } = await open(book)
    await goTo(w, 'P21')
    expect(api.sortColumn('desc')).toBe(true)
    await settle()
    await press(q('[data-xs-object="slicer:Region"] .xs-slicer-head')!, 'pointerdown')
    expect(lastState(w)?.object?.name).toBe('Region')
    expect(document.querySelectorAll('[data-xs-handle]').length).toBe(0)
  })
})

describe("the timeline's level list and Shift+click", () => {
  it('the level list closes on a press outside and on Esc', async () => {
    await open()
    await click(q('[data-xs-timeline-level]')!)
    expect(q('[data-xs-level="years"]')).not.toBeNull()
    await press(document.body, 'pointerdown')
    expect(q('[data-xs-level="years"]')).toBeNull()
    await click(q('[data-xs-timeline-level]')!)
    q('[data-xs-level="years"]')!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    await settle()
    expect(q('[data-xs-level="years"]')).toBeNull()
  })

  it('Shift+click extends from the dates chosen before, and from the period pressed after a change of level', async () => {
    const { s } = await open()
    await s.structure({ op: 'editTimeline', timeline: 'Date', cache: { selection: { start: '2024-04-01T00:00:00', end: '2024-06-30T00:00:00' } } })
    await until(() => q('[data-xs-period="1"]')?.getAttribute('aria-selected') === 'true')
    const grid = q('[data-xs-periods]')!
    await press(q('[data-xs-period="3"]')!, 'pointerdown', { shiftKey: true })
    await press(grid, 'pointerup')
    await until(() => q('[data-xs-period="3"]')?.getAttribute('aria-selected') === 'true')
    expect((await s.sheets())[1]!.timelines![0]!.selection).toEqual({ start: '2024-04-01T00:00:00', end: '2024-12-31T00:00:00' })
    // Q3 2024 pressed, then months: Shift+click on Jan 2025 extends from July 2024
    await press(q('[data-xs-period="2"]')!, 'pointerdown')
    await press(grid, 'pointerup')
    await click(q('[data-xs-timeline-level]')!)
    await click(q('[data-xs-level="months"]')!)
    await until(() => document.querySelectorAll('[data-xs-period]').length === 24)
    await press(q('[data-xs-period="12"]')!, 'pointerdown', { shiftKey: true })
    await press(q('[data-xs-periods]')!, 'pointerup')
    await until(() => q('[data-xs-period="12"]')?.getAttribute('aria-selected') === 'true')
    expect((await s.sheets())[1]!.timelines![0]!.selection).toEqual({ start: '2024-07-01T00:00:00', end: '2025-01-31T00:00:00' })
  })
})
