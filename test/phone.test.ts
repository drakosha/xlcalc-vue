// The phone is a finger on a narrow sheet: a mouse in a narrow window keeps the desktop's sizes and rows, and only what needs room gives way to the width.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createApp, nextTick } from 'vue'
import { SheetToolbar, SheetView, memorySource, type MemoryBook, type SheetViewState, type ToolbarItem } from '../src'
import type { SheetSource } from '../src/contract/source'
import { fingerPointer, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const book = (): MemoryBook => ({ sheets: [{ name: 'One', cells: { A1: 'a', A2: 'b', A3: 'c' } }] })

/** The sheet `px` wide on screen. */
function sheetWidth(px: number): void {
  vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
    return this.classList.contains('xs-scroller') ? px : 0
  })
}
const classes = async (): Promise<string[]> => {
  const w = await mountSheet(memorySource(book()))
  return w.find('.xs-sheet').classes()
}

describe('the sheet', () => {
  it('listens for double clicks on its root, so WebKit reads a double tap as one and does not zoom the page', async () => {
    const add = vi.spyOn(HTMLElement.prototype, 'addEventListener')
    // Not `mount`: test-utils itself listens on the root for every DOM event.
    const app = createApp(SheetView, { source: memorySource(book()) })
    app.mount(document.body.appendChild(document.createElement('div')))
    await settle()
    const on = add.mock.contexts.filter((_, i) => add.mock.calls[i]![0] === 'dblclick') as Element[]
    expect(on.some((el) => el.classList.contains('xs-sheet'))).toBe(true)
    app.unmount()
  })

  it('keeps a landscape touch window compact at 734 px wide', async () => {
    vi.spyOn(window, 'matchMedia').mockImplementation((query: string) => ({ matches: ['(pointer: coarse)', '(max-height: 500px)'].includes(query), media: query, addEventListener: () => {}, removeEventListener: () => {} }) as unknown as MediaQueryList)
    sheetWidth(734)
    expect(await classes()).toContain('xs-phone')
  })
  it('with a mouse below 600 px gives way only to the width: no phone, no touch sizes', async () => {
    sheetWidth(590)
    const on = await classes()
    expect(on).toContain('xs-narrow')
    expect(on).not.toContain('xs-phone')
    expect(on).not.toContain('xs-touch')
  })

  it('with a finger below 600 px is a phone', async () => {
    fingerPointer()
    sheetWidth(590)
    const on = await classes()
    expect(on).toContain('xs-narrow')
    expect(on).toContain('xs-phone')
    expect(on).toContain('xs-touch')
  })

  it('with a finger at 600 px and wider has touch sizes without the phone', async () => {
    fingerPointer()
    sheetWidth(600)
    const on = await classes()
    expect(on).not.toContain('xs-narrow')
    expect(on).not.toContain('xs-phone')
    expect(on).toContain('xs-touch')
  })

  it('with a mouse below 600 px keeps the row grip and the zoom bar', async () => {
    sheetWidth(590)
    const w = await mountSheet(memorySource(book()))
    await settle()
    expect(w.find('th[data-row="1"] .xs-grip').exists()).toBe(true)
    expect(w.find('.xs-zoom-bar').exists()).toBe(true)
  })
})

describe("the filter's list", () => {
  const open = async () => {
    const w = await mountSheet(memorySource(book()))
    await w.findAll('.xs-head')[0]!.find('.xs-funnel').trigger('click')
    await settle()
    return w
  }

  it('with a mouse below 600 px opens below its column', async () => {
    sheetWidth(590)
    const w = await open()
    expect(w.find('.xs-popover').classes()).not.toContain('xs-popover-full')
    expect((w.find('.xs-popover').element as HTMLElement).style.top).not.toBe('')
  })

  it('with a finger below 600 px takes the full screen, from the funnel in the bar', async () => {
    fingerPointer()
    sheetWidth(590)
    const w = await mountSheet(memorySource(book()))
    await w.find('[data-xs-funnel]').trigger('click')
    await settle()
    expect(w.find('.xs-popover').classes()).toContain('xs-popover-full')
  })
})

describe('the toolbar', () => {
  const state: SheetViewState = { sheet: 'One', ref: 'A1' }
  const source = () =>
    ({
      sheets: async () => [{ name: 'One', maxRow: 3, maxCol: 1 }],
      window: async () => ({ pending: { elapsedMs: 0 } }),
      format: async () => ({ ok: true, changes: [] }),
      undo: async () => ({ ok: true, changes: [] }),
      redo: async () => ({ ok: true, changes: [] }),
    }) as unknown as SheetSource
  const bar = async (rows: number) => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(390)
    const w = mount(SheetToolbar, { props: { source: source(), state, rows, items: ['fontStyle'] }, attachTo: document.body })
    await nextTick()
    await nextTick()
    return w
  }

  it('with a mouse below 600 px keeps the desktop sizes and the rows it is given', async () => {
    const w = await bar(2)
    expect(w.find('.xs-toolbar').classes()).not.toContain('xs-touch')
    expect(w.find('.xs-tb-row').classes()).toContain('xs-tb-ribbon')
  })

  it('with a mouse below 600 px keeps Bold in the row, each block alone, when the Font group would not fit', async () => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(420)
    const w = mount(SheetToolbar, { props: { source: source(), state, items: ['font', 'fontStyle'] }, attachTo: document.body })
    await nextTick()
    await nextTick()
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['font', 'fontStyle'])
  })

  /** The row `room` px wide at this window, More 44 px; a block as `sizes` says, Excel's whole Font group (the block with Bold and the Font boxes) `fontGroup` px. */
  function widths(room: number, viewport: number, sizes: Record<string, number>, fontGroup: number): void {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(viewport)
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-tb-row') ? room : 0
    })
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      const key = this.dataset.tbBlock
      const whole = key === 'font' && this.querySelector('[data-tb="bold"]')
      const width = whole ? fontGroup : key ? (sizes[key] ?? 60) : this.dataset.tb === 'more' ? 44 : 0
      return { width, height: 30, top: 0, left: 0, right: width, bottom: 30, x: 0, y: 0, toJSON: () => ({}) } as DOMRect
    })
  }
  const inRow = async (items: ToolbarItem[]) => {
    const w = mount(SheetToolbar, { props: { source: source(), state, items }, attachTo: document.body })
    await settle()
    return w.findAll('.xs-tb-row > [data-tb-block]').map((b) => b.attributes('data-tb-block'))
  }
  const SIZES = { undo: 88, font: 290, fontStyle: 100, borders: 60, align: 150 }

  it('keeps Bold, Italic and Underline over the Font boxes when both do not fit', async () => {
    widths(420, 590, SIZES, 600)
    expect(await inRow(['undo', 'font', 'fontStyle', 'borders', 'align'])).toEqual(['undo', 'fontStyle', 'borders'])
  })

  it('at 600 px and wider splits Excel\'s Font group into blocks when it does not fit whole, so that Bold stays in the row', async () => {
    widths(684, 700, SIZES, 700)
    expect(await inRow(['undo', 'font', 'fontStyle', 'borders', 'align'])).toEqual(['undo', 'font', 'fontStyle', 'borders'])
  })

  it('keeps Excel\'s Font group whole where it fits', async () => {
    widths(1000, 1016, SIZES, 450)
    expect(await inRow(['undo', 'font', 'fontStyle', 'borders', 'align'])).toEqual(['undo', 'font', 'align'])
  })

  it('with a finger below 600 px is one row of touch sizes', async () => {
    fingerPointer()
    const w = await bar(2)
    expect(w.find('.xs-toolbar').classes()).toContain('xs-touch')
    expect(w.find('.xs-tb-row').classes()).not.toContain('xs-tb-ribbon')
  })
})
