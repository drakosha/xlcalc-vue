// The source event `repaint`: the windows that overlap its ranges are re-read in place (no sheets, no empty viewport, the
// display filter's rows kept), one request per window, edits that arrive meanwhile kept; the local source sends it after an edit.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { BookInput } from 'xlcalc'
import { cellAddress } from 'xlcalc/a1'
import { memorySource, type MemoryBook } from '../src'
import { localSource } from '../src/local'
import type { SheetInfo, SourceEvent, WindowRequest } from '../src/contract/source'
import { cellText, editorKey, fakeSource, gridWindow, goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

type Fake = ReturnType<typeof fakeSource>
type Held = Fake['held'][number]

/** A held request answered with `text` in every cell, truncated to the sheet as the fake source truncates it. */
function answer(f: Fake, p: Held, sheet: SheetInfo, text: string): void {
  f.held.splice(f.held.indexOf(p), 1)
  const req = p.req
  p.resolve(gridWindow({ ...req, rows: Math.min(req.rows, sheet.maxRow - req.row + 1), cols: Math.min(req.cols, sheet.maxCol - req.col + 1) }, () => text))
}
/** Answers every held request, and those their responses trigger, until none is left. */
async function drain(f: Fake): Promise<void> {
  while (f.held.length) {
    f.release(f.held[0]!)
    await settle()
  }
}
const rect = (q: WindowRequest) => ({ row: q.row, rows: q.rows, col: q.col, cols: q.cols })
const loadingShown = (w: VueWrapper) => w.find('.xs-loading').exists()
/** Longer than the wait of a repaint that closely follows another. */
const pause = () => new Promise((r) => setTimeout(r, 100))

describe('repaint in the grid', () => {
  it('re-reads in place only the windows that overlap its ranges: no sheets, and the old window stays drawn until the new one arrives', async () => {
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

    // The frozen rows overlap C1:C2; the body and the frozen column do not.
    version = 1
    let before = f.requests.length
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'C1:C2' }] })
    await settle()
    expect(f.requests.slice(before).map(rect)).toEqual([top])
    // Until the response the old window stays drawn, with no loading indicator: nothing went empty.
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
    await pause()
    await settle()
    expect(f.requests.slice(before).map(rect).sort((a, b) => a.cols - b.cols)).toEqual([left, body])
    await drain(f)
    expect(cellText(w, 2000, 1)).toBe('S:A2000:2')
    expect(cellText(w, 2000, 3)).toBe('S:C2000:2')
    expect(cellText(w, 1, 3)).toBe('S:C1:1')
    expect(f.sheetsCalls).toBe(1)
  })

  it('sends one request per window: a second repaint aborts the first, and the first response arriving late is discarded', async () => {
    const sheet: SheetInfo = { name: 'S', maxRow: 100, maxCol: 5 }
    const f = fakeSource([sheet], { manual: true })
    const w = await mountSheet(f.source, { readOnly: true })
    await drain(f)
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1' }] })
    await settle()
    const first = f.held[0]!
    f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'B2' }] })
    await pause()
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

  it('reads the repaints from one task together; one that closely follows requests once they stop', async () => {
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'Date'] })
    try {
      const f = fakeSource([{ name: 'S', maxRow: 100, maxCol: 5 }])
      await mountSheet(f.source, { readOnly: true })
      let before = f.requests.length
      for (let i = 1; i <= 20; i++) f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: `A${i}` }] })
      await settle()
      expect(f.requests.length - before).toBe(1)
      vi.advanceTimersByTime(100)
      before = f.requests.length
      f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A1' }] })
      await settle()
      expect(f.requests.length - before).toBe(1)
      for (let i = 0; i < 4; i++) {
        vi.advanceTimersByTime(20)
        f.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'B2' }] })
        await settle()
      }
      expect(f.requests.length - before).toBe(1)
      vi.advanceTimersByTime(100)
      await settle()
      expect(f.requests.length - before).toBe(2)
    } finally {
      vi.useRealTimers()
    }
  })

  it('under a display filter overlaps the ranges by the sheet rows the windows display', async () => {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: { 1: { values: ['x'] } }, sort: null }))
    const cells: MemoryBook['sheets'][number]['cells'] = { A1: 'head' }
    for (let r = 2; r <= 600; r++) cells[`A${r}`] = r === 500 ? 'x' : 'y'
    const source = memorySource({ sheets: [{ name: 'S', cells }] })
    const window = vi.spyOn(source, 'window')
    const w = await mountSheet(source, { readOnly: true, storageKey: 'k' })
    expect(w.findAll('.xs-num').map((n) => n.text())).toEqual(['1', '500'])
    const before = window.mock.calls.length
    // Row 3 is hidden by the filter, though its display position 3 is near the window's two.
    source.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A2:A3' }] })
    await settle()
    expect(window.mock.calls.length).toBe(before)
    // A range taller than the window is overlapped by the window's rows: none of 3 to 499 is displayed.
    source.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A3:A499' }] })
    await settle()
    expect(window.mock.calls.length).toBe(before)
    source.emit({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'A500' }] })
    await settle()
    expect(window.mock.calls.length).toBe(before + 1)
    expect(window.mock.calls.at(-1)![0]).toMatchObject({ row: 1, filter: { cols: { 1: { values: ['x'] } } } })
  })

  it('keeps a change that arrived between the request and the response', async () => {
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
    // The response was read before the change: the change is applied over it.
    answer(f, f.held[0]!, sheet, 'read')
    await settle()
    expect(cellText(w, 3, 3)).toBe('new')
    expect(cellText(w, 3, 2)).toBe('read')
  })
})

describe('repaint from the local source', () => {
  const MIN = '#F8696B'
  const MAX = '#63BE7B'
  /** Names in A, numbers in B under a color scale; a funnel on B keeping 1 and 3 shows rows 1, 2, 4 and 6. */
  const book = (): BookInput => ({
    sheets: [
      {
        name: 'S',
        cells: {
          A1: { value: 'Name' }, B1: { value: 'Value' },
          A2: { value: 'n2' }, B2: { value: 1 },
          A3: { value: 'n3' }, B3: { value: 2 },
          A4: { value: 'n4' }, B4: { value: 1 },
          A5: { value: 'n5' }, B5: { value: 2 },
          A6: { value: 'n6' }, B6: { value: 3 },
        },
        conditionalFormats: [{ type: 'colorScale', ranges: ['B2:B6'], points: [{ value: { type: 'min' }, color: MIN }, { value: { type: 'max' }, color: MAX }] }],
      },
    ],
  })
  const filter = { head: 1, cols: { 2: { values: ['1', '3'] } } }
  const numbers = (w: VueWrapper) => w.findAll('.xs-num').map((n) => n.text())
  const fill = (w: VueWrapper, r: number, c: number) => getComputedStyle(w.find(`[data-cell="${r}:${c}"]`).element).backgroundColor

  it('sends changed and repaint after an edit in a workbook with rules, repaint after a formatting change and reload after a number format', async () => {
    const s = localSource(book(), { calc: 'inline' })
    const events: SourceEvent[] = []
    s.subscribe((e) => events.push(e))
    expect(await s.edit([{ sheet: 'S', cell: 'B4', text: '9' }])).toMatchObject({ ok: true })
    expect(events.map((e) => e.kind)).toEqual(['changed', 'repaint'])
    expect(events[0]).toMatchObject({ changes: [{ sheet: 'S', cell: 'B4', value: 9 }] })
    expect(events[1]).toEqual({ kind: 'repaint', ranges: [{ sheet: 'S', range: 'B2:B6' }] })
    events.length = 0
    expect(await s.format('S', [{ range: 'B1', style: { font: { bold: true } } }])).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'repaint', ranges: [{ sheet: 'S', range: 'B1' }] }])
    events.length = 0
    expect(await s.format('S', [{ range: 'B1', numberFormat: '0.00' }])).toMatchObject({ ok: true })
    expect(events).toEqual([{ kind: 'reload' }])
    await s.close()
  })

  it("keeps the source's display filter rows after an edit and an undo that only repaint, also with a rule that has a number format", async () => {
    for (const numberFormat of [undefined, '0.00']) {
      const b = book()
      if (numberFormat) b.sheets[0]!.conditionalFormats!.push({ type: 'cellIs', ranges: ['B2:B6'], operator: 'greaterThan', formulas: ['5'], style: { numberFormat } })
      const s = localSource(b, { calc: 'inline' })
      const rows = async () => {
        const got = await s.window({ sheet: 'S', row: 1, rows: 10, col: 1, cols: 1, filter })
        if ('pending' in got) throw new Error('pending')
        return got.text.map((line) => line[0])
      }
      expect((await s.visibleRows('S', filter)).rows).toEqual([1, 2, 4, 6])
      // B4 no longer matches the condition and B3 now does: the rows stay those the filter returned until it is reapplied.
      await s.edit([{ sheet: 'S', cell: 'B4', text: '9' }])
      await s.edit([{ sheet: 'S', cell: 'B3', text: '3' }])
      expect(await rows()).toEqual(['Name', 'n2', 'n4', 'n6'])
      await s.undo!()
      expect(await rows()).toEqual(['Name', 'n2', 'n4', 'n6'])
      expect((await s.visibleRows('S', filter)).rows).toEqual([1, 2, 6])
      await s.close()
    }
  })

  it('keeps a row that no longer matches the display filter after an edit, recolored, until the filter is reapplied', async () => {
    localStorage.setItem('k:filter:S', JSON.stringify({ cols: filter.cols, sort: null }))
    const s = localSource(book(), { calc: 'inline' })
    const w = await mountSheet(s, { storageKey: 'k' })
    await settle(12)
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    expect(fill(w, 4, 2)).toBe(MIN)
    expect(fill(w, 6, 2)).toBe(MAX)
    await goTo(w, 'B4')
    await key(w, '9')
    await type(w, '9')
    await editorKey(w, 'Enter')
    await settle(12)
    // As in Excel the row stays until Reapply; the scale is drawn over the new values, the edited cell's too.
    expect(numbers(w)).toEqual(['1', '2', '4', '6'])
    expect(cellText(w, 4, 2)).toBe('9')
    expect(fill(w, 4, 2)).toBe(MAX)
    expect(fill(w, 6, 2)).not.toBe(MAX)
    expect(fill(w, 2, 2)).toBe(MIN)
    // A reload reapplies the filter: the row is hidden.
    expect(await s.format('S', [{ range: 'A1', numberFormat: '@' }])).toMatchObject({ ok: true })
    await settle(12)
    expect(numbers(w)).toEqual(['1', '2', '6'])
    await s.close()
  })
})
