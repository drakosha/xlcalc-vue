// Excel's Fill Down (Ctrl+D) and Fill Right (Ctrl+R): the selection's first row or column copied over the rest through
// `structure` (`copyRange`), the row above or the column to the left for a selection one line thick, one undo step.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type SheetViewApi } from '../src'
import type { EditOutcome, SheetOp } from '../src/contract/source'
import { localSource } from '../src/local'
import { address, cellText, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
function withStructure() {
  const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, B1: { value: 10, formula: 'A1*10' } }, locked: [{ range: 'H1:H9', reason: 'script' }] }] })
  const structure = vi.fn(async (_op: SheetOp) => OK)
  source.structure = structure
  return { source, structure }
}
const copy = (from: string, to: string): SheetOp => ({ op: 'copyRange', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to } })
const ops = (structure: ReturnType<typeof withStructure>['structure']) => structure.mock.calls.map((c) => c[0])
/** A key on the grid as the browser sends it, cancelable, to see whether the sheet took it from the browser. */
function press(w: VueWrapper, k: string, extra: Record<string, unknown> = {}): KeyboardEvent {
  const e = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...extra })
  w.find('.xs-scroller').element.dispatchEvent(e)
  return e
}

describe('Fill Down and Fill Right', () => {
  it('Ctrl+D copies the top row over the rows below it, Ctrl+R the left column over the columns to its right; the selection stays', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B2:C5')
    await key(w, 'd', { ctrlKey: true })
    await settle()
    await key(w, 'r', { ctrlKey: true })
    await settle()
    expect(ops(structure)).toEqual([copy('B2:C2', 'B3:C5'), copy('B2:B5', 'C2:C5')])
    expect(address(w)).toBe('B2:C5')
  })

  it('a selection one row tall takes the row above it, one column wide the column to its left', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B3')
    await key(w, 'd', { ctrlKey: true })
    await settle()
    await goTo(w, 'C1:C4')
    await key(w, 'r', { ctrlKey: true })
    await settle()
    await goTo(w, 'B4:D4')
    await key(w, 'd', { ctrlKey: true })
    await settle()
    expect(ops(structure)).toEqual([copy('B2', 'B3'), copy('B1:B4', 'C1:C4'), copy('B3:D3', 'B4:D4')])
  })

  it('entire columns and rows: from the first row or column of the sheet over the rest', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await w.find('.xs-num[data-row="12"]').trigger('click')
    await key(w, 'd', { ctrlKey: true })
    await settle()
    await w.findAll('.xs-head')[1]!.trigger('click')
    await key(w, 'd', { ctrlKey: true })
    await settle()
    expect(ops(structure)).toEqual([copy('11:11', '12:12'), copy('B1', 'B2:B1048576')])
  })

  it('nothing above row 1 or left of column A: nothing is sent, and the browser still does not get the key', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B1:D1')
    expect(press(w, 'd', { ctrlKey: true }).defaultPrevented).toBe(true)
    await goTo(w, 'A2:A5')
    expect(press(w, 'r', { ctrlKey: true }).defaultPrevented).toBe(true)
    await settle()
    expect(structure).not.toHaveBeenCalled()
  })

  it('Cmd and Control on a Mac, and the key by its position under another layout', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'B3')
    await key(w, 'd', { metaKey: true })
    await settle()
    await key(w, 'в', { code: 'KeyD', ctrlKey: true })
    await settle()
    await key(w, 'к', { code: 'KeyR', metaKey: true })
    await settle()
    expect(ops(structure)).toEqual([copy('B2', 'B3'), copy('B2', 'B3'), copy('A3', 'B3')])
  })

  it('a locked cell among the filled ones is reported and nothing is sent; Ctrl+Shift+D is not Fill Down', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    await goTo(w, 'H1:H3')
    await key(w, 'd', { ctrlKey: true })
    await settle()
    expect(w.find('.xs-message').text()).toContain('script')
    await goTo(w, 'B3')
    expect(press(w, 'D', { ctrlKey: true, shiftKey: true }).defaultPrevented).toBe(false)
    await settle()
    expect(structure).not.toHaveBeenCalled()
  })

  it('a read-only sheet leaves the keys to the browser; without `structure` the sheet keeps them and sends nothing', async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source, { readOnly: true })
    await goTo(w, 'B3')
    expect(press(w, 'd', { ctrlKey: true }).defaultPrevented).toBe(false)
    w.unmount()
    const plain = memorySource({ sheets: [{ name: 'S', cells: { B2: 1 } }] })
    const v = await mountSheet(plain)
    await goTo(v, 'B3')
    expect(press(v, 'r', { ctrlKey: true }).defaultPrevented).toBe(true)
    await settle()
    expect(structure).not.toHaveBeenCalled()
    expect(cellText(v, 3, 2)).toBe('')
  })

  it("the host's methods: fillDown and fillRight resolve to the outcome, null when nothing was sent", async () => {
    const { source, structure } = withStructure()
    const w = await mountSheet(source)
    const api = w.vm as unknown as SheetViewApi
    await goTo(w, 'C2:C4')
    expect(await api.fillDown()).toEqual(OK)
    expect(await api.fillRight()).toEqual(OK)
    await goTo(w, 'A1')
    expect(await api.fillDown()).toBeNull()
    expect(await api.fillRight()).toBeNull()
    expect(ops(structure)).toEqual([copy('C2', 'C3:C4'), copy('B2:B4', 'C2:C4')])
  })
})

describe('Fill Down in a workbook calculated in the browser', () => {
  it('shifts relative references as Excel does, keeps absolute ones, and Ctrl+Z undoes it in one step', async () => {
    const cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, A4: { value: 4 }, C1: { value: 100 }, B1: { formula: 'A1*10+$C$1' } }
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells }] }, { calc: 'inline' })
    const states: { formula?: string }[] = []
    const w = await mountSheet(source, { onState: (st: { formula?: string } | null) => st && states.push(st) })
    await settle(12)
    await goTo(w, 'B1:B4')
    await key(w, 'd', { ctrlKey: true })
    await settle(12)
    expect([1, 2, 3, 4].map((r) => cellText(w, r, 2))).toEqual(['110', '120', '130', '140'])
    expect(address(w)).toBe('B1:B4')
    await goTo(w, 'B4')
    expect(states.at(-1)!.formula).toBe('A4*10+$C$1')
    await goTo(w, 'C2')
    await key(w, 'r', { ctrlKey: true })
    await settle(12)
    expect(states.at(-1)!.formula).toBe('B2*10+$C$1')
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect([2, 3, 4].map((r) => cellText(w, r, 2))).toEqual(['', '', ''])
    await source.close()
  })
})
