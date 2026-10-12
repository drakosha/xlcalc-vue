// The source event `repaint`: the windows that meet its ranges are read again in place (no sheets, no empty frame, the display
// filter's rows kept), one request per window, edits that come meanwhile kept.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cellAddress } from 'xlcalc/a1'
import { memorySource, type MemoryBook } from '../src'
import type { SheetInfo, WindowRequest } from '../src/source'
import { cellText, fakeSource, gridWindow, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

type Fake = ReturnType<typeof fakeSource>
type Held = Fake['held'][number]

/** A held request answered with `text` in every cell, cut to the sheet as the fake source cuts it. */
function answer(f: Fake, p: Held, sheet: SheetInfo, text: string): void {
  f.held.splice(f.held.indexOf(p), 1)
  const req = p.req
  p.resolve(gridWindow({ ...req, rows: Math.min(req.rows, sheet.maxRow - req.row + 1), cols: Math.min(req.cols, sheet.maxCol - req.col + 1) }, () => text))
}
/** Answer every held request, and those their answers bring, until none is left. */
async function drain(f: Fake): Promise<void> {
  while (f.held.length) {
    f.release(f.held[0]!)
    await settle()
  }
}
const rect = (q: WindowRequest) => ({ row: q.row, rows: q.rows, col: q.col, cols: q.cols })
const loadingShown = (w: VueWrapper) => w.find('.xs-loading').exists()

describe('repaint in the grid', () => {
  it('reads again in place only the windows that meet its ranges: no sheets, the old window drawn until the new one comes', async () => {
    let version = 0
    const sheets: SheetInfo[] = [{ name: 'S', maxRow: 3000, maxCol: 40, freeze: { rows: 2, cols: 1 } }, { name: 'T', maxRow: 5, maxCol: 5 }]
    const f = fakeSource(sheets, { manual: true, cell: (sheet, r, c) => `${sheet}:${cellAddress(r, c)}:${version}` })
    const w = await mountSheet(f.source, { readOnly: true })
    await drain(f)
    const scroller = w.find('.xs-scroller')
    scroller.element.scrollTop = 1990 * 20
    await scroller.trigger('scroll')
    await drain(f)
    // The last request of each slot: the frozen rows, the frozen column, the body.
    const asked = f.requests.map(rect).reverse()
    const isTop = (q: ReturnType<typeof rect>) => q.row === 1 && q.rows === 2
    const isLeft = (q: ReturnType<typeof rect>) => q.col === 1 && q.cols === 1
    const top = asked.find(isTop)!
    const left = asked.find(isLeft)!
    const body = asked.find((q) => !isTop(q) && !isLeft(q))!
    expect(body.row).toBeGreaterThan(1000)
    expect(cellText(w, 1, 3)).toBe('S:C1:0')
    expect(cellText(w, 2000, 1)).toBe('S:A2000:0')
    expect(cellText(w, 2000, 3)).toBe('S:C2000:0')

    // The frozen rows meet C1:C2; the body and the frozen column do not.
    version = 1
    let before = f.requests.length
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'C1:C2' }] })
    await settle()
    expect(f.requests.slice(before).map(rect)).toEqual([top])
    // Until the answer the old window stays drawn, with no loading mark: nothing went empty.
    expect(cellText(w, 1, 3)).toBe('S:C1:0')
    expect(cellText(w, 2000, 3)).toBe('S:C2000:0')
    expect(loadingShown(w)).toBe(false)
    await drain(f)
    expect(cellText(w, 1, 3)).toBe('S:C1:1')
    expect(cellText(w, 2000, 3)).toBe('S:C2000:0')

    // Another sheet's ranges and a range away from every window read nothing.
    before = f.requests.length
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'T', range: 'A1:E5' }, { sheet: 'S', range: 'C100:C200' }] })
    await settle()
    expect(f.requests.length).toBe(before)

    // A2000 is in the body and the frozen column.
    version = 2
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A2000' }] })
    await settle()
    expect(f.requests.slice(before).map(rect).sort((a, b) => a.cols - b.cols)).toEqual([left, body])
    await drain(f)
    expect(cellText(w, 2000, 1)).toBe('S:A2000:2')
    expect(cellText(w, 2000, 3)).toBe('S:C2000:2')
    expect(cellText(w, 1, 3)).toBe('S:C1:1')
    expect(f.sheetsCalls).toBe(1)
  })

  it('sends one request per window: a second repaint aborts the first, and the first answer coming late is dropped', async () => {
    const sheet: SheetInfo = { name: 'S', maxRow: 100, maxCol: 5 }
    const f = fakeSource([sheet], { manual: true })
    const w = await mountSheet(f.source, { readOnly: true })
    await drain(f)
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1' }] })
    await settle()
    const first = f.held[0]!
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'B2' }] })
    await settle()
    expect(f.held).toHaveLength(2)
    const second = f.held[1]!
    expect(first.signal?.aborted).toBe(true)
    expect(second.signal?.aborted).toBe(false)
    expect(rect(second.req)).toEqual(rect(first.req))
    answer(f, second, sheet, 'second')
    await settle()
    expect(cellText(w, 1, 1)).toBe('second')
    answer(f, first, sheet, 'first')
    await settle()
    expect(cellText(w, 1, 1)).toBe('second')
    expect(f.held).toHaveLength(0)
  })

  it('under a display filter meets the ranges by the sheet rows the windows show', async () => {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: { 1: { values: ['x'] } }, sort: null }))
    const cells: MemoryBook['sheets'][number]['cells'] = { A1: 'head' }
    for (let r = 2; r <= 600; r++) cells[`A${r}`] = r === 500 ? 'x' : 'y'
    const source = memorySource({ sheets: [{ name: 'S', cells }] })
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source, { readOnly: true, storageKey: 'k' })
    expect(w.findAll('.xs-num').map((n) => n.text())).toEqual(['1', '500'])
    const before = window.mock.calls.length
    // Row 3 is hidden by the filter, though its display place 3 is near the window's two.
    source.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A2:A3' }] })
    await settle()
    expect(window.mock.calls.length).toBe(before)
    source.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A500' }] })
    await settle()
    expect(window.mock.calls.length).toBe(before + 1)
    expect(window.mock.calls.at(-1)![0]).toMatchObject({ row: 1, filter: { cols: { 1: { values: ['x'] } } } })
  })

  it('keeps a change that came between the request and the answer', async () => {
    const sheet: SheetInfo = { name: 'S', maxRow: 100, maxCol: 5 }
    const f = fakeSource([sheet], { manual: true })
    const w = await mountSheet(f.source, { readOnly: true })
    await drain(f)
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1:A3' }] })
    await settle()
    expect(f.held).toHaveLength(1)
    f.emit({ kind: 'changed', changes: [{ sheet: 'S', cell: 'C3', text: 'new', value: 'new' }] })
    await settle()
    expect(cellText(w, 3, 3)).toBe('new')
    // The answer was read before the change: the change goes over it.
    answer(f, f.held[0]!, sheet, 'read')
    await settle()
    expect(cellText(w, 3, 3)).toBe('new')
    expect(cellText(w, 3, 2)).toBe('read')
  })
})
