// Excel's fill handle: dragged from the selection's corner its frame grows along one axis in whole tiles of the
// selection, and the release copies the selection over the cells added through `structure` (copyRange); a double click
// fills down as far as the data of the column beside goes. There is no handle where the sheet cannot copy.
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
/** Drag the handle through points and let go at the last one. */
async function drag(w: VueWrapper, points: { clientX: number; clientY: number }[], pointerType = 'mouse'): Promise<void> {
  const h = handle(w)
  const style = h.attributes('style') ?? ''
  const start = { clientX: parseFloat(/left: ([\d.]+)px/.exec(style)![1]!), clientY: parseFloat(/top: ([\d.]+)px/.exec(style)![1]!) }
  await h.trigger('pointerdown', { button: 0, pointerId: 1, pointerType, ...start })
  for (const p of points) await h.trigger('pointermove', { pointerId: 1, pointerType, ...p })
  await h.trigger('pointerup', { pointerId: 1, pointerType, ...points.at(-1)! })
  await settle()
}
const ops = (structure: ReturnType<typeof withStructure>['structure']) => structure.mock.calls.map((c) => c[0])
const copy = (from: string, to: string): SheetOp => ({ op: 'copyRange', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to } })

describe('the fill handle', () => {
  it('sits at the selection’s bottom-right corner', async () => {
    const w = await mountSheet(withStructure().source)
    await goTo(w, 'B2:C3')
    expect(handle(w).attributes('style')).toContain(`top: ${24 + 3 * 20}px; left: ${56 + 3 * COL}px`)
  })

  it('dragged down three rows from a formula copies it over them; the selection becomes both', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    await drag(w, [at(2, 2), at(4, 2)])
    expect(ops(structure)).toEqual([copy('B1', 'B2:B4')])
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
    expect(ops(structure)).toEqual([copy('B2', 'C2:D2'), copy('B5', 'B2:B4'), copy('C3', 'A3:B3'), copy('B2', 'B3:B6')])
  })

  it('a selection of several rows or columns grows by whole tiles of it, as copyRange tiles', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'A1:A2')
    await drag(w, [at(3, 1)])
    await goTo(w, 'A1:A2')
    await drag(w, [at(5, 1)])
    await goTo(w, 'C4:D4')
    await drag(w, [at(4, 5)])
    expect(ops(structure)).toEqual([copy('A1:A2', 'A3:A4'), copy('A1:A2', 'A3:A6'), copy('C4:D4', 'E4:F4')])
    expect(address(w)).toBe('C4:F4')
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

  it('a double click fills down as far as the column left of it has data, else the one right of it', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B1')
    await handle(w).trigger('dblclick')
    await settle()
    await goTo(w, 'E1')
    await handle(w).trigger('dblclick')
    await settle()
    await goTo(w, 'A4')
    await handle(w).trigger('dblclick')
    await settle()
    expect(ops(structure)).toEqual([copy('B1', 'B2:B4'), copy('E1', 'E2:E3')])
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
})
