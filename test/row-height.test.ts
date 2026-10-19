// A row's height set with the mouse, as in Excel: the grip on a row number's bottom edge drags the row's height and the release
// sets it through `layout`, for every selected entire row when the dragged one is among them; no grip without `layout` or when read-only.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook } from '../src'
import type { LayoutChange } from '../src/contract/source'
import { goTo, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const book = (): MemoryBook => ({ sheets: [{ name: 'S', cells: { A1: 1, A2: 2, A3: 'line one\nline two', A4: 4 } }] })
const grip = (w: VueWrapper, row: number) => w.find(`th[data-row="${row}"] .xs-grip`)
const rowPx = (w: VueWrapper, row: number) => parseFloat((w.find(`th[data-row="${row}"]`).element.parentElement as HTMLElement).style.height)

async function dragGrip(w: VueWrapper, row: number, dy: number): Promise<void> {
  const g = grip(w, row)
  await g.trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', clientX: 20, clientY: 100 })
  await g.trigger('pointermove', { pointerId: 1, pointerType: 'mouse', clientX: 20, clientY: 100 + dy / 2 })
  await g.trigger('pointermove', { pointerId: 1, pointerType: 'mouse', clientX: 20, clientY: 100 + dy })
  await g.trigger('pointerup', { pointerId: 1, pointerType: 'mouse' })
  await settle()
}

describe('row height with the mouse', () => {
  it('the drag sets the row’s height in points through `layout`; the row is drawn at it', async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    expect(rowPx(w, 3)).toBe(20)
    await dragGrip(w, 3, 20)
    expect(layout.mock.calls.map((c) => c[0])).toEqual([{ sheet: 'S', rows: '3', height: 30 } satisfies LayoutChange])
    expect(rowPx(w, 3)).toBe(40)
    expect(rowPx(w, 2)).toBe(20)
  })

  it("a drag stops at Excel's highest row, 409.5 points", async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    await dragGrip(w, 3, 2000)
    expect(layout.mock.calls.map((c) => c[0])).toEqual([{ sheet: 'S', rows: '3', height: 409.5 }])
  })

  it('with whole rows selected, a drag on one of them sets them all', async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    await goTo(w, '2:4')
    await dragGrip(w, 3, 10)
    expect(layout.mock.calls.map((c) => c[0])).toEqual([{ sheet: 'S', rows: '2:4', height: 22.5 }])
    expect([rowPx(w, 2), rowPx(w, 3), rowPx(w, 4)]).toEqual([30, 30, 30])
  })

  it('a double click on the grip fits the row, or every selected whole row it is one of, through `layout`', async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    await grip(w, 3).trigger('dblclick')
    await goTo(w, '2:4')
    await grip(w, 4).trigger('dblclick')
    await settle()
    expect(layout.mock.calls.map((c) => c[0])).toEqual([
      { sheet: 'S', rows: '3', fit: true },
      { sheet: 'S', rows: '2:4', fit: true },
    ])
  })

  it('a press without a move changes nothing and selects nothing', async () => {
    const source = memorySource(book())
    const layout = vi.spyOn(source, 'layout')
    const w = await mountSheet(source)
    await dragGrip(w, 3, 0)
    await grip(w, 3).trigger('click')
    expect(layout).not.toHaveBeenCalled()
    expect((w.find('.xs-address').element as HTMLInputElement).value).toBe('A1')
  })

  it('no grip on a phone: a finger on a number snaps to it and the number’s long press opens no menu; the menu sets the height there', async () => {
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-scroller') ? 390 : 0
    })
    const w = await mountSheet(memorySource(book()))
    expect(w.find('.xs-sheet').classes()).toContain('xs-touch')
    expect(grip(w, 3).exists()).toBe(false)
  })

  it('no grip without `layout` or on a read-only sheet', async () => {
    const plain = memorySource(book())
    delete (plain as { layout?: unknown }).layout
    const w = await mountSheet(plain)
    expect(grip(w, 3).exists()).toBe(false)
    w.unmount()
    const ro = await mountSheet(memorySource(book()), { readOnly: true })
    expect(grip(ro, 3).exists()).toBe(false)
  })
})
