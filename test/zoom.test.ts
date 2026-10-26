// Excel's zoom: hit testing of the pointer at 50% and 200% (Point mode's drag, a column edge), Ctrl+wheel's 15% steps and a
// pinch, each sheet's own zoom from the host, the status bar and the Zoom dialog, Zoom to Selection, and what is drawn zoomed.
import type { VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { memorySource, type MemoryBook, type SheetViewApi, type SheetViewState, type SheetZoomEvent } from '../src'
import { colPx } from '../src/core/geometry'
import { clampZoom, sliderOfZoom, stepZoom, zoomOfSlider } from '../src/core/zoom'
import { goTo, key, mountSheet, settle, type } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.useRealTimers()
})

const book = (): MemoryBook => ({
  sheets: [
    { name: 'One', cells: { A1: 'a1', B2: 5, C3: { value: 10, formula: 'B2*2' } } },
    { name: 'Two', cells: { A1: 'two' } },
  ],
})
const COL = colPx(null)
const api = (w: VueWrapper) => w.vm as unknown as SheetViewApi
const lastState = (w: VueWrapper) => (w.emitted('state') as [SheetViewState | null][]).at(-1)![0]
const zooms = (w: VueWrapper) => (w.emitted('zoom') ?? []).map(([e]) => e as SheetZoomEvent)
const field = (w: VueWrapper) => w.find('.xs-editor').element as HTMLTextAreaElement
const scroller = (w: VueWrapper) => w.find('.xs-scroller')
/** A wheel event with Ctrl held (happy-dom's WheelEvent drops the modifier keys of its init). */
function ctrlWheel(deltaY: number, ctrlKey = true): WheelEvent {
  const e = new WheelEvent('wheel', { deltaY, bubbles: true, cancelable: true })
  Object.defineProperty(e, 'ctrlKey', { value: ctrlKey })
  return e
}

/** A cell's center on the page at zoom `z` with the grid scrolled by `top` unzoomed pixels: the scroller is at 0, 0; letters 24 px, numbers 56 px, rows 20 px. */
const at = (row: number, col: number, z: number, top = 0) => ({ clientX: (56 + (col - 1) * COL + COL / 2) * z, clientY: (24 + (row - 1) * 20 + 10 - top) * z })

async function scrollTo(w: VueWrapper, top: number, z: number): Promise<void> {
  const el = scroller(w).element as HTMLElement
  el.scrollTop = top * z
  await scroller(w).trigger('scroll')
  await settle()
}
/** A drag in Point mode from a cell through others, its pointer at zoomed page coordinates. */
async function pointDrag(w: VueWrapper, from: [number, number], to: [number, number], z: number, top = 0): Promise<void> {
  const td = w.find(`[data-cell="${from[0]}:${from[1]}"]`)
  await td.trigger('pointerdown', { button: 0, pointerId: 1, pointerType: 'mouse', ...at(from[0], from[1], z, top) })
  td.element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }))
  await scroller(w).trigger('pointermove', { pointerId: 1, pointerType: 'mouse', ...at(to[0], to[1], z, top) })
  await scroller(w).trigger('pointerup', { pointerId: 1, pointerType: 'mouse' })
  await td.trigger('click', { button: 0 })
  await settle()
}

describe('zoom: hit testing', () => {
  for (const z of [0.5, 2]) {
    it(`a pointer at ${z * 100}% finds the cell under it by the geometry, scrolled or not`, async () => {
      const w = await mountSheet(memorySource(book()))
      api(w).setZoom(z * 100)
      await settle()
      await key(w, '=')
      await type(w, '=')
      await pointDrag(w, [2, 2], [4, 3], z)
      expect(field(w).value).toBe('=B2:C4')
      await key(w, 'Escape')
      await w.find('.xs-editor').trigger('keydown', { key: 'Escape' })
      await settle()
      // scrolled by 200 unzoomed pixels: row 15 is where row 5 was
      const big = memorySource({ sheets: [{ name: 'S', cells: Object.fromEntries(Array.from({ length: 80 }, (_, i) => [`A${i + 1}`, i + 1])) }] })
      const v = await mountSheet(big)
      api(v).setZoom(z * 100)
      await settle()
      await scrollTo(v, 200, z)
      await goTo(v, 'C20')
      await key(v, '=')
      await type(v, '=')
      await pointDrag(v, [15, 1], [17, 2], z, 200)
      expect(field(v).value).toBe('=A15:B17')
    })
  }

  it('a column edge dragged at 200% widens the column by the distance in unzoomed pixels', async () => {
    const w = await mountSheet(memorySource(book()), { storageKey: 'zoom-widths' })
    api(w).setZoom(200)
    await settle()
    const grip = w.findAll('.xs-head')[1]!.find('.xs-grip')
    await grip.trigger('pointerdown', { clientX: 100, pointerId: 1 })
    await grip.trigger('pointermove', { clientX: 140, pointerId: 1 })
    await grip.trigger('pointerup', { clientX: 140, pointerId: 1 })
    await settle()
    expect((w.findAll('.xs-table col')[2]!.element as HTMLElement).style.width).toBe(`${COL + 20}px`)
  })
})

describe('zoom: wheel and pinch', () => {
  it('Ctrl+wheel steps by 15% within 10% to 400% and reports each notch; the wheel alone scrolls', async () => {
    const w = await mountSheet(memorySource(book()))
    const wheel = async (deltaY: number, ctrlKey = true) => {
      const e = ctrlWheel(deltaY, ctrlKey)
      scroller(w).element.dispatchEvent(e)
      await settle()
      return e
    }
    expect((await wheel(-100)).defaultPrevented).toBe(true)
    expect(lastState(w)?.zoom).toBe(115)
    await wheel(100)
    await wheel(100)
    expect(lastState(w)?.zoom).toBe(85)
    expect(zooms(w)).toEqual([
      { sheet: 'One', zoom: 115 },
      { sheet: 'One', zoom: 100 },
      { sheet: 'One', zoom: 85 },
    ])
    expect((await wheel(100, false)).defaultPrevented).toBe(false)
    for (let i = 0; i < 30; i++) await wheel(-100)
    expect(lastState(w)?.zoom).toBe(400)
  })

  it("a trackpad's pinch (small Ctrl+wheel deltas) zooms in proportion and reports once it ends", async () => {
    vi.useFakeTimers()
    const w = await mountSheet(memorySource(book()))
    for (let i = 0; i < 10; i++) scroller(w).element.dispatchEvent(ctrlWheel(-7))
    await settle()
    expect(lastState(w)?.zoom).toBe(Math.round(100 * Math.exp(0.7)))
    expect(zooms(w)).toEqual([])
    vi.advanceTimersByTime(300)
    expect(zooms(w)).toEqual([{ sheet: 'One', zoom: 201 }])
  })

  it('two fingers spread twice as far zoom to 200%, reported when they lift', async () => {
    const w = await mountSheet(memorySource(book()))
    const touch = (type: string, d: number | null) => {
      const e = new Event(type, { bubbles: true, cancelable: true })
      Object.defineProperty(e, 'touches', { value: d === null ? [] : [{ clientX: 100, clientY: 100 }, { clientX: 100 + d, clientY: 100 }] })
      scroller(w).element.dispatchEvent(e)
      return e
    }
    touch('touchstart', 80)
    const move = touch('touchmove', 160)
    await settle()
    expect(move.defaultPrevented).toBe(true)
    expect(lastState(w)?.zoom).toBe(200)
    expect(zooms(w)).toEqual([])
    touch('touchend', null)
    await settle()
    expect(zooms(w)).toEqual([{ sheet: 'One', zoom: 200 }])
  })
})

describe('zoom: per sheet, the host and the controls', () => {
  it("each sheet opens at the host's zoom for it, keeps the user's in the session, and setZoom is not reported", async () => {
    const w = await mountSheet(memorySource(book()), { zoom: { one: 150 } })
    expect(lastState(w)?.zoom).toBe(150)
    expect((w.find('.xs-zoom').element as HTMLElement).style.zoom).toBe('1.5')
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    expect(lastState(w)?.zoom).toBeUndefined()
    expect(w.find('.xs-zoom').attributes('style')).toBeUndefined()
    expect(api(w).setZoom(5)).toBe(10)
    expect(api(w).setZoom(1000)).toBe(400)
    await w.findAll('.xs-tab')[0]!.trigger('click')
    await settle()
    expect(lastState(w)?.zoom).toBe(150)
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    expect(lastState(w)?.zoom).toBe(400)
    expect(zooms(w)).toEqual([])
    // a new zoom from the host applies to the active sheet at once
    await w.setProps({ zoom: { one: 150, Two: 75 } })
    await settle()
    expect(lastState(w)?.zoom).toBe(75)
  })

  it("a sheet without the host's zoom opens at the workbook's saved zoom (SheetInfo.zoom), and the host's wins over it", async () => {
    const saved = book()
    saved.sheets[0]!.zoom = 125
    saved.sheets[1]!.zoom = 60
    const w = await mountSheet(memorySource(saved), { zoom: { Two: 200 } })
    expect(lastState(w)?.zoom).toBe(125)
    await w.findAll('.xs-tab')[1]!.trigger('click')
    await settle()
    expect(lastState(w)?.zoom).toBe(200)
    expect(zooms(w)).toEqual([])
  })

  it('the status bar: − and + go to the next multiple of 10%, the slider maps its halves to 10–100% and 100–400%', async () => {
    expect([stepZoom(100, 1), stepZoom(115, 1), stepZoom(115, -1), stepZoom(10, -1), stepZoom(400, 1)]).toEqual([110, 120, 110, 10, 400])
    expect([zoomOfSlider(0), zoomOfSlider(50), zoomOfSlider(100), zoomOfSlider(25), zoomOfSlider(75)]).toEqual([10, 100, 400, 55, 250])
    expect(sliderOfZoom(250)).toBe(75)
    expect([clampZoom(Number.NaN), clampZoom(57.4)]).toEqual([100, 57])
    const w = await mountSheet(memorySource(book()))
    await w.find('[data-xs-zoom-in]').trigger('click')
    await settle()
    expect(w.find('[data-xs-zoom-level]').text()).toBe('110%')
    const slider = w.find('[data-xs-zoom-slider]')
    ;(slider.element as HTMLInputElement).value = '25'
    await slider.trigger('input')
    expect(w.find('[data-xs-zoom-level]').text()).toBe('55%')
    await slider.trigger('change')
    expect(zooms(w)).toEqual([
      { sheet: 'One', zoom: 110 },
      { sheet: 'One', zoom: 55 },
    ])
  })

  it("the Zoom dialog: a preset, Custom within 10–400%, and Fit selection, Excel's Zoom to Selection", async () => {
    const w = await mountSheet(memorySource(book()))
    await w.find('[data-xs-zoom-level]').trigger('click')
    await settle()
    expect((w.find('[data-xs-zoom-choice="100"]').element as HTMLInputElement).checked).toBe(true)
    await w.find('[data-xs-zoom-choice="75"]').setValue(true)
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await settle()
    expect(lastState(w)?.zoom).toBe(75)
    expect(w.find('[data-xs-zoom-dialog]').exists()).toBe(false)
    await w.find('[data-xs-zoom-level]').trigger('click')
    await settle()
    expect((w.find('[data-xs-zoom-choice="75"]').element as HTMLInputElement).checked).toBe(true)
    await w.find('[data-xs-zoom-custom]').setValue('500')
    await w.find('[data-xs-zoom-choice="custom"]').setValue(true)
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await settle()
    expect(w.find('.xs-dialog-error').text()).toBe('Enter a number between 10 and 400')
    // A1:J20 is 710 × 400 px with the headers 766 × 424 px: in the 800 × 600 viewport it fits at 104%
    await w.find('[data-xs-dialog-cancel]').trigger('click')
    await goTo(w, 'A1:J20')
    await w.find('[data-xs-zoom-level]').trigger('click')
    await settle()
    await w.find('[data-xs-zoom-choice="fit"]').setValue(true)
    await w.find('[data-xs-dialog-ok]').trigger('click')
    await settle()
    expect(lastState(w)?.zoom).toBe(104)
    expect(api(w).setZoom('selection')).toBe(104)
    expect(zooms(w).map((z) => z.zoom)).toEqual([75, 104])
  })
})

describe('zoom: drawing', () => {
  it('zoomed out, gridlines draw one screen pixel wide and the cell boxes give it room; the viewport is measured in unzoomed pixels', async () => {
    const w = await mountSheet(memorySource(book()))
    const columns = w.findAll('th.xs-head').length
    api(w).setZoom(50)
    await settle()
    const box = w.find('.xs-zoom').element as HTMLElement
    expect(box.style.zoom).toBe('0.5')
    expect(box.style.getPropertyValue('--xs-line')).toBe('2px')
    expect((w.find('[data-cell="1:1"] .xs-in').element as HTMLElement).style.height).toBe(`${20 - 2 - 2}px`)
    // the viewport holds twice as many columns, and the phone layout follows the screen's width, not the unzoomed one
    expect(w.findAll('th.xs-head').length).toBeGreaterThan(columns)
    api(w).setZoom(200)
    await settle()
    expect(w.find('.xs-sheet').classes()).not.toContain('xs-narrow')
    expect(box.style.getPropertyValue('--xs-line')).toBe('')
    expect((w.find('[data-cell="1:1"] .xs-in').element as HTMLElement).style.height).toBe(`${20 - 2 - 1}px`)
  })

  it('the editor is drawn at the zoom; the fill handle keeps its size, placed in the zoomed pixels', async () => {
    const source = memorySource(book())
    source.structure = async () => ({ ok: true, changes: [] })
    const w = await mountSheet(source)
    api(w).setZoom(200)
    await settle()
    const handle = w.find('[data-xs-fill-handle]').element as HTMLElement
    expect(handle.style.zoom).toBe('0.5')
    expect(handle.style.top).toBe(`${(24 + 20) * 2}px`)
    expect(handle.style.left).toBe(`${(56 + COL) * 2}px`)
    await key(w, 'F2')
    await settle()
    expect((w.find('.xs-editor').element as HTMLElement).style.zoom).toBe('2')
  })
})
