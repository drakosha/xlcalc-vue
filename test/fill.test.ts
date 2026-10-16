// Excel's fill handle: dragged from the selection's corner its frame grows along one axis cell by cell, and the release
// fills the frame from the selection through `structure` (fillRange, series with Ctrl); a double click fills down as far
// as the data of the column beside goes. There is no handle where the sheet cannot copy.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import { colPx } from '../src/core/geometry'
import { localSource } from '../src/local'
import type { EditOutcome, SheetOp } from '../src/source'
import { address, cellText, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
const book = (): MemoryBook => ({
  sheets: [
    {
      name: 'S',
      cells: { A1: 1, A2: 2, A3: 3, A4: 4, B1: { value: 10, formula: 'A1*10' }, E1: { value: 5, formula: 'F1' }, F1: 5, F2: 6, F3: 7 },
      locked: [{ range: 'H1:H9', reason: 'script' }],
    },
  ],
})
function withStructure(answer: (op: SheetOp) => EditOutcome = () => OK) {
  const source = memorySource(book())
  const structure = vi.fn(async (op: SheetOp) => answer(op))
  source.structure = structure
  return { source, structure }
}
/** A column without a width, in pixels. */
const COL = colPx(null)
/** The middle of a cell on the page: the scroller is at 0, 0; letters 24 px high, numbers 56 px wide, rows 20 px, columns `COL`. */
const at = (row: number, col: number) => ({ clientX: 56 + (col - 1) * COL + COL / 2, clientY: 24 + (row - 1) * 20 + 10 })
const handle = (w: VueWrapper) => w.find('[data-xs-fill-handle]')
const frame = (w: VueWrapper) => w.find('[data-xs-fill-frame]')
/** Drag the handle through points and let go at the last one, with the keys held at the release. */
async function drag(w: VueWrapper, points: { clientX: number; clientY: number }[], pointerType = 'mouse', keys: { ctrlKey?: boolean; altKey?: boolean } = {}): Promise<void> {
  const h = handle(w)
  const style = h.attributes('style') ?? ''
  const start = { clientX: parseFloat(/left: ([\d.]+)px/.exec(style)![1]!), clientY: parseFloat(/top: ([\d.]+)px/.exec(style)![1]!) }
  await h.trigger('pointerdown', { button: 0, pointerId: 1, pointerType, ...start })
  for (const p of points) await h.trigger('pointermove', { pointerId: 1, pointerType, ...p })
  await h.trigger('pointerup', { pointerId: 1, pointerType, ...keys, ...points.at(-1)! })
  await settle()
}
const ops = (structure: ReturnType<typeof withStructure>['structure']) => structure.mock.calls.map((c) => c[0])
const fill = (range: string, to: string, type?: 'series'): SheetOp => ({ op: 'fillRange', sheet: 'S', range, to, ...(type ? { type } : {}) })

describe('the fill handle', () => {
  it('sits at the selection’s bottom-right corner', async () => {
    const w = await mountSheet(withStructure().source)
    await goTo(w, 'B2:C3')
    expect(handle(w).attributes('style')).toContain(`top: ${24 + 3 * 20}px; left: ${56 + 3 * COL}px`)
  })

  it('dragged down three rows from a formula fills the frame from it; the selection becomes the frame', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    await drag(w, [at(2, 2), at(4, 2)])
    expect(ops(structure)).toEqual([fill('B1', 'B1:B4')])
    expect(address(w)).toBe('B1:B4')
  })

  it('grows along the axis the pointer went furthest: right, up, left, and down on a diagonal', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await drag(w, [at(2, 4)])
    await goTo(w, 'B5')
    await drag(w, [at(2, 2)])
    await goTo(w, 'C3')
    await drag(w, [at(3, 1)])
    await goTo(w, 'B2')
    await drag(w, [at(6, 3)])
    expect(ops(structure)).toEqual([fill('B2', 'B2:D2'), fill('B5', 'B2:B5'), fill('C3', 'A3:C3'), fill('B2', 'B2:B6')])
    expect(address(w)).toBe('B2:B6')
  })

  it('a selection of several rows or columns grows cell by cell, not by whole tiles of it, as in Excel', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1:A2')
    const h = handle(w)
    await h.trigger('pointerdown', { button: 0, pointerId: 1, ...at(2, 1) })
    await h.trigger('pointermove', { pointerId: 1, ...at(3, 1) })
    expect(frame(w).attributes('style')).toContain(`top: 24px; left: 56px; width: ${COL}px; height: 60px`)
    await h.trigger('pointerup', { pointerId: 1, ...at(3, 1) })
    await settle()
    await goTo(w, 'A1:A2')
    await drag(w, [at(5, 1)])
    await goTo(w, 'C4:D4')
    await drag(w, [at(4, 5)])
    await goTo(w, 'C4:D5')
    await drag(w, [at(4, 2)])
    expect(ops(structure)).toEqual([fill('A1:A2', 'A1:A3'), fill('A1:A2', 'A1:A5'), fill('C4:D4', 'C4:E4'), fill('C4:D5', 'B4:D5')])
    expect(address(w)).toBe('B4:D5')
  })

  it('released with Ctrl it fills a series (Excel’s drag with Ctrl), with Option too on a Mac', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1')
    await drag(w, [at(4, 1)], 'mouse', { ctrlKey: true })
    await goTo(w, 'A1')
    await drag(w, [at(4, 1)], 'mouse', { altKey: true })
    vi.spyOn(navigator, 'platform', 'get').mockReturnValue('MacIntel')
    await goTo(w, 'A1')
    await drag(w, [at(1, 3)], 'mouse', { altKey: true })
    expect(ops(structure)).toEqual([fill('A1', 'A1:A4', 'series'), fill('A1', 'A1:A4'), fill('A1', 'A1:C1', 'series')])
  })

  it('draws its frame while dragged; back on the selection there is none, and nothing is copied', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    const h = handle(w)
    await h.trigger('pointerdown', { button: 0, pointerId: 1, ...at(2, 2) })
    await h.trigger('pointermove', { pointerId: 1, ...at(4, 2) })
    expect(frame(w).attributes('style')).toContain(`top: ${24 + 20}px; left: ${56 + COL}px; width: ${COL}px; height: 60px`)
    expect(handle(w).classes()).toContain('xs-fill-dragging')
    await h.trigger('pointermove', { pointerId: 1, ...at(2, 2) })
    expect(frame(w).exists()).toBe(false)
    await h.trigger('pointerup', { pointerId: 1, ...at(2, 2) })
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(address(w)).toBe('B2')
  })

  it('measures the frame once a move: the edge scroll and the cell under the pointer share it', async () => {
    const { source } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    const h = handle(w)
    await h.trigger('pointerdown', { button: 0, pointerId: 1, ...at(2, 2) })
    const box = vi.spyOn(w.find('.xs-scroller').element, 'getBoundingClientRect')
    await h.trigger('pointermove', { pointerId: 1, ...at(4, 2) })
    expect(box).toHaveBeenCalledTimes(1)
    expect(frame(w).attributes('style')).toContain(`height: 60px`)
    await h.trigger('pointerup', { pointerId: 1, ...at(2, 2) })
  })

  it('Esc while dragged cancels the drag, as in Excel: no frame, nothing copied on release, the selection and the copy mark stay; the next drag fills', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1')
    const e = new Event('copy', { bubbles: true, cancelable: true })
    Object.defineProperty(e, 'clipboardData', { value: { setData: () => {} } })
    w.find('.xs-scroller').element.dispatchEvent(e)
    await goTo(w, 'B2')
    expect(w.find('[data-xs-ants]').exists()).toBe(true)
    const h = handle(w)
    const el = h.element as HTMLElement
    el.setPointerCapture = vi.fn()
    el.releasePointerCapture = vi.fn()
    const keys = vi.spyOn(document, 'addEventListener')
    const unkeys = vi.spyOn(document, 'removeEventListener')
    await h.trigger('pointerdown', { button: 0, pointerId: 1, ...at(2, 2) })
    await h.trigger('pointermove', { pointerId: 1, ...at(4, 2) })
    expect(frame(w).exists()).toBe(true)
    await key(w, 'Escape')
    expect(frame(w).exists()).toBe(false)
    expect(handle(w).classes()).not.toContain('xs-fill-dragging')
    expect(el.releasePointerCapture).toHaveBeenCalledWith(1)
    await h.trigger('pointermove', { pointerId: 1, ...at(5, 2) })
    expect(frame(w).exists()).toBe(false)
    await h.trigger('pointerup', { pointerId: 1, ...at(5, 2) })
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(address(w)).toBe('B2')
    expect(w.find('[data-xs-ants]').exists()).toBe(true)
    const added = keys.mock.calls.filter((c) => c[0] === 'keydown')
    expect(added).toHaveLength(1)
    expect(unkeys.mock.calls.filter((c) => c[0] === 'keydown' && c[1] === added[0]![1])).toHaveLength(1)
    await drag(w, [at(3, 2)])
    expect(ops(structure)).toEqual([fill('B2', 'B2:B3')])
    expect(address(w)).toBe('B2:B3')
    expect(keys.mock.calls.filter((c) => c[0] === 'keydown')).toHaveLength(2)
    expect(unkeys.mock.calls.filter((c) => c[0] === 'keydown')).toHaveLength(2)
  })

  it('a finger’s tap on its wide target picks the cell under it', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    const h = handle(w)
    await h.trigger('pointerdown', { button: 0, pointerId: 3, pointerType: 'touch', ...at(3, 3) })
    await h.trigger('pointerup', { pointerId: 3, pointerType: 'touch', ...at(3, 3) })
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(address(w)).toBe('C3')
  })

  it('a double click fills down as far as the column left of it has data, else the one right of it, cell by cell', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    await handle(w).trigger('dblclick')
    await settle()
    await goTo(w, 'E1')
    await handle(w).trigger('dblclick')
    await settle()
    await goTo(w, 'E1:E2')
    await handle(w).trigger('dblclick')
    await settle()
    expect(address(w)).toBe('E1:E3')
    await goTo(w, 'A4')
    await handle(w).trigger('dblclick')
    await settle()
    expect(ops(structure)).toEqual([fill('B1', 'B1:B4'), fill('E1', 'E1:E3'), fill('E1:E2', 'E1:E3')])
    expect(address(w)).toBe('A4')
  })

  it('a locked target is not sent; a refusal shows the source’s message', async () => {
    const { source, structure } = withStructure(() => ({ ok: false, reason: 'refused', message: 'Cannot change part of a merged cell' }))
    const w = await mountSheet(source)
    await goTo(w, 'G1')
    await drag(w, [at(1, 8)])
    expect(structure).not.toHaveBeenCalled()
    expect(w.find('.xs-message').text()).toContain('script')
    await goTo(w, 'A1')
    await drag(w, [at(3, 1)])
    expect(w.find('.xs-message').text()).toBe('Cannot change part of a merged cell')
    expect(address(w)).toBe('A1')
  })

  it('is not there on a read-only sheet, without structure, over whole columns, while a field is open, under the display filter', async () => {
    const ro = await mountSheet(withStructure().source, { readOnly: true })
    expect(handle(ro).exists()).toBe(false)
    document.body.innerHTML = ''
    const plain = await mountSheet(memorySource(book()))
    expect(handle(plain).exists()).toBe(false)
    document.body.innerHTML = ''
    const w = await mountSheet(withStructure().source)
    expect(handle(w).exists()).toBe(true)
    await w.findAll('.xs-head')[1]!.trigger('click')
    expect(handle(w).exists()).toBe(false)
    await goTo(w, 'C3')
    await key(w, 'F2')
    expect(handle(w).exists()).toBe(false)
    await key(w, 'Escape')
    w.find('.xs-editor').trigger('keydown', { key: 'Escape' })
    await settle()
    expect(handle(w).exists()).toBe(true)
    await w.findAll('.xs-head')[0]!.find('.xs-funnel').trigger('click')
    await settle()
    await w.find('.xs-popover button[title="Sort descending"]').trigger('click')
    await settle()
    expect(handle(w).exists()).toBe(false)
  })
})

describe('with the local source', () => {
  it('a formula dragged down three rows moves its references in each copy', async () => {
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells: { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 }, B1: { formula: 'A1*10' } } }] }, { calc: 'inline' })
    const states: { formula?: string }[] = []
    const w = await mountSheet(source, { onState: (st: { formula?: string } | null) => st && states.push(st) })
    await settle(12)
    await goTo(w, 'B1')
    await drag(w, [at(4, 2)])
    await settle(12)
    expect([2, 3, 4].map((r) => cellText(w, r, 2))).toEqual(['20', '30', '40'])
    expect(address(w)).toBe('B1:B4')
    await goTo(w, 'B4')
    expect(states.at(-1)!.formula).toBe('A4*10')
    await source.close()
  })

  it('goes on with series as Excel: 1, 2 to A5 is 3, 4, 5; Item1 to B3 is Item2, Item3; a date by days; Control+Z takes it back', async () => {
    const cells = { A1: { value: 1 }, A2: { value: 2 }, B1: { value: 'Товар1' }, C1: { value: 45688, numberFormat: 'm/d' } }
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells }] }, { calc: 'inline' })
    const w = await mountSheet(source)
    await settle(12)
    const column = (col: number, rows: number) => Array.from({ length: rows }, (_, i) => cellText(w, i + 1, col))
    await goTo(w, 'A1:A2')
    await drag(w, [at(5, 1)])
    await settle(12)
    expect(column(1, 5)).toEqual(['1', '2', '3', '4', '5'])
    expect(address(w)).toBe('A1:A5')
    await goTo(w, 'B1')
    await drag(w, [at(3, 2)])
    await settle(12)
    expect(column(2, 3)).toEqual(['Товар1', 'Товар2', 'Товар3'])
    await goTo(w, 'C1')
    await drag(w, [at(3, 3)])
    await settle(12)
    expect(column(3, 3)).toEqual(['1/31', '2/1', '2/2'])
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect(column(3, 3)).toEqual(['1/31', '', ''])
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect(column(2, 3)).toEqual(['Товар1', '', ''])
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect(column(1, 5)).toEqual(['1', '2', '', '', ''])
    await source.close()
  })

  it('a single number dragged copies it, with Ctrl it goes on by 1', async () => {
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells: { A1: { value: 7 }, B1: { value: 7 } } }] }, { calc: 'inline' })
    const w = await mountSheet(source)
    await settle(12)
    await goTo(w, 'A1')
    await drag(w, [at(3, 1)])
    await settle(12)
    await goTo(w, 'B1')
    await drag(w, [at(3, 2)], 'mouse', { ctrlKey: true })
    await settle(12)
    expect([1, 2, 3].map((r) => [cellText(w, r, 1), cellText(w, r, 2)])).toEqual([['7', '7'], ['7', '8'], ['7', '9']])
    await source.close()
  })
})
