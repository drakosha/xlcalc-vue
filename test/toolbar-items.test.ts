// The toolbar's `items`: the groups it draws and their order, the host's buttons drawn as its own, and in a narrow viewport
// the items that fit in the row, in order, with the rest, the host's buttons too, in the menu.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, nextTick } from 'vue'
import { SheetToolbar, TOOLBAR_GROUPS, type SheetViewApi, type SheetViewState, type ToolbarItem } from '../src'
import type { EditOutcome, SheetSource } from '../src/contract/source'
import { fitLines } from '../src/toolbar'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
const source = () =>
  ({
    sheets: async () => [{ name: 'S', maxRow: 10, maxCol: 5 }],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    format: vi.fn(async () => OK),
    undo: vi.fn(async () => OK),
    redo: vi.fn(async () => OK),
  }) as SheetSource
const state: SheetViewState = { sheet: 'S', ref: 'A1' }
const mountBar = (items?: readonly ToolbarItem[]) => mount(SheetToolbar, { props: { source: source(), state, ...(items ? { items } : {}) }, attachTo: document.body })
/** The hooks of the toolbar's buttons and boxes in the order drawn, the host's as `host:<id>`. */
const hooks = (w: VueWrapper, within = '') =>
  w.findAll(`${within} [data-tb], ${within} [data-tb-host]`).map((e) => e.attributes('data-tb') ?? `host:${e.attributes('data-tb-host')}`)

describe("SheetToolbar's items", () => {
  it('without items draws every group in the default order', async () => {
    const w = mountBar()
    expect(TOOLBAR_GROUPS[0]).toBe('undo')
    expect(hooks(w).slice(0, 10)).toEqual(['undo', 'redo', 'font-name', 'font-name-list', 'font-size', 'font-size-list', 'font-up', 'font-down', 'bold', 'italic'])
    expect(hooks(w)).toContain('borders')
  })

  it('draws only the groups named, in their order; an unknown or repeated name is skipped', async () => {
    const w = mountBar(['borders', 'fontStyle', 'nothing' as never, 'borders', 'undo'])
    await nextTick()
    expect(hooks(w)).toEqual(['borders', 'bold', 'italic', 'underline', 'undo', 'redo'])
    expect(w.findAll('.xs-tb-group')).toHaveLength(3)
  })

  it("draws the host's buttons as its own: an icon component or SVG markup, or the label; runs, disables and presses them", async () => {
    const run = vi.fn()
    const Icon = defineComponent(() => () => h('svg', { 'data-mine': '' }))
    const w = mountBar([
      'fontStyle',
      { id: 'share', label: 'Share', icon: Icon, run },
      { id: 'pin', label: 'Pin', icon: '<svg viewBox="0 0 24 24"><path d="M4 4h16"/></svg>', run, pressed: true },
      { id: 'save', label: 'Save', run, disabled: true },
      'undo',
    ])
    await nextTick()
    expect(hooks(w)).toEqual(['bold', 'italic', 'underline', 'host:share', 'host:pin', 'host:save', 'undo', 'redo'])
    expect(w.findAll('.xs-tb-group')).toHaveLength(3)
    const share = w.find('[data-tb-host="share"]')
    expect(share.attributes('data-tip')).toBe('Share')
    expect(share.find('svg[data-mine]').classes()).toContain('xs-icon')
    expect(share.attributes('aria-pressed')).toBeUndefined()
    const pin = w.find('[data-tb-host="pin"]')
    expect(pin.find('.xs-tb-svg path').exists()).toBe(true)
    expect(pin.attributes('aria-pressed')).toBe('true')
    expect(pin.classes()).toContain('xs-tool-on')
    const save = w.find('[data-tb-host="save"]')
    expect(save.text()).toBe('Save')
    expect(save.attributes('disabled')).toBeDefined()
    await share.trigger('click')
    await pin.trigger('click')
    expect(run).toHaveBeenCalledTimes(2)
  })

  /** A narrow toolbar whose row is `room` px wide, each block `block` px and the More button 44 px wide. */
  function narrowLayout(room: number, block: (key: string, el: HTMLElement) => number = () => 100, viewport = 390): void {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(viewport)
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-tb-row') ? room : 0
    })
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      const key = this.dataset.tbBlock
      const width = key ? block(key, this) : this.dataset.tb === 'more' ? 44 : 0
      return { width, height: 44, top: 0, left: 0, right: width, bottom: 44, x: 0, y: 0, toJSON: () => ({}) } as DOMRect
    })
  }

  it('in a narrow viewport keeps the items in the row while they fit beside More, in order, and moves the rest, the host’s buttons too, into the menu', async () => {
    narrowLayout(300)
    const run = vi.fn()
    const w = mountBar(['fontStyle', 'undo', 'borders', { id: 'share', label: 'Share', run }])
    await nextTick()
    await nextTick()
    expect(hooks(w, '.xs-tb-row')).toEqual(['bold', 'italic', 'underline', 'undo', 'redo', 'more'])
    const menu = w.find('[data-tb-menu]')
    expect((menu.element as HTMLElement).style.display).toBe('none')
    await w.find('[data-tb="more"]').trigger('click')
    expect((menu.element as HTMLElement).style.display).toBe('')
    expect(hooks(w, '[data-tb-menu]')).toEqual(['borders', 'host:share'])
    await w.find('[data-tb-host="share"]').trigger('click')
    expect(run).toHaveBeenCalledOnce()
    expect((menu.element as HTMLElement).style.display).toBe('none')
  })

  it('in a wide viewport too: what does not fit goes into More rather than onto a second line', async () => {
    narrowLayout(1000, () => 300, 1440)
    const w = mountBar(['undo', 'fontStyle', 'font', 'borders', 'align'])
    await nextTick()
    await nextTick()
    expect(w.find('.xs-toolbar').classes()).not.toContain('xs-tb-narrow')
    expect(hooks(w, '.xs-tb-row').at(-1)).toBe('more')
    expect(w.findAll('[data-tb-menu] > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['borders', 'align'])
  })

  it('with rows wraps onto that many lines, as the row wraps them, and moves only the rest into More, which ends the last line', async () => {
    narrowLayout(920, () => 300, 1440)
    const w = mount(SheetToolbar, { props: { source: source(), state, items: ['undo', 'fontStyle', 'font', 'borders', 'align', 'colors', 'number', 'merge'], rows: 2 }, attachTo: document.body })
    await nextTick()
    await nextTick()
    expect(w.find('.xs-tb-row').classes()).toContain('xs-tb-rows')
    // Three blocks a line; the second line keeps two beside More
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((b) => [b.attributes('data-tb-block'), b.attributes('data-tb-line') !== undefined])).toEqual([
      ['undo', false],
      ['fontStyle', false],
      ['font', false],
      ['borders', true],
      ['align', false],
    ])
    expect(hooks(w, '.xs-tb-row').at(-1)).toBe('more')
    expect(w.findAll('[data-tb-menu] > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['colors', 'number'])
    await w.setProps({ rows: 3 })
    await nextTick()
    await nextTick()
    expect(w.find('[data-tb-menu]').exists()).toBe(false)
    expect(w.findAll('.xs-tb-row > [data-tb-line]').map((b) => b.attributes('data-tb-block'))).toEqual(['borders', 'number'])
  })

  it('measures again when the sheet arrives, which draws its groups, and when the Number Format list loads', async () => {
    narrowLayout(300, (key, el) => (key === 'number' ? (el.querySelector('select') ? 250 : 50) : 100), 1440)
    const sheet = { formatPainter: vi.fn() } as unknown as SheetViewApi
    const w = mount(SheetToolbar, { props: { source: source(), state, items: ['undo', 'formatPainter', 'fontStyle', 'borders'] }, attachTo: document.body })
    await nextTick()
    await nextTick()
    expect(w.find('[data-tb-menu]').exists()).toBe(false)
    await w.setProps({ sheet })
    await nextTick()
    await nextTick()
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['undo', 'formatPainter'])
    await w.setProps({ items: ['undo', 'number'], locale: 'ru-RU' })
    await vi.waitFor(() => expect(w.find('[data-tb="number"]').exists()).toBe(true))
    await nextTick()
    await nextTick()
    expect(w.findAll('[data-tb-menu] > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['number'])
  })

  it('below 600 px keeps one row whatever rows says', async () => {
    narrowLayout(300, () => 100, 390)
    const w = mount(SheetToolbar, { props: { source: source(), state, items: ['undo', 'fontStyle', 'borders', 'align'], rows: 3 }, attachTo: document.body })
    await nextTick()
    await nextTick()
    expect(w.find('.xs-tb-row').classes()).not.toContain('xs-tb-rows')
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['undo', 'fontStyle'])
  })

  it('in a narrow viewport keeps every item in the row when all fit, with no More button', async () => {
    narrowLayout(380, (key) => (key === 'undo' ? 100 : 90))
    const run = vi.fn()
    const w = mountBar(['undo', { id: 'a', label: 'A', run }, 'borders', 'fontStyle'])
    await nextTick()
    await nextTick()
    expect(hooks(w)).toEqual(['undo', 'redo', 'host:a', 'borders', 'bold', 'italic', 'underline'])
    expect(w.find('[data-tb-menu]').exists()).toBe(false)
    expect(w.find('[data-tb="more"]').exists()).toBe(false)
  })
})

describe('fitLines', () => {
  it('packs the blocks onto lines as a wrapping row does: all when they fit, otherwise those that fit with More on the last line', () => {
    expect(fitLines([40, 40, 40], 100, 1, 30)).toEqual({ fit: 1, breaks: [] })
    expect(fitLines([40, 40, 20], 100, 1, 30)).toEqual({ fit: 3, breaks: [] })
    expect(fitLines([40, 40, 40, 40], 100, 2, 30)).toEqual({ fit: 4, breaks: [2] })
    // The last line leaves 30 px for More: 40 + 40 > 70
    expect(fitLines([40, 40, 40, 40, 40], 100, 2, 30)).toEqual({ fit: 3, breaks: [2] })
    expect(fitLines([60, 50, 50, 50, 10, 90], 100, 3, 30)).toEqual({ fit: 5, breaks: [1, 3] })
  })

  it('stops at a block wider than a line, rather than letting it overflow', () => {
    expect(fitLines([40, 120, 40], 100, 2, 30)).toEqual({ fit: 1, breaks: [] })
    expect(fitLines([120, 40], 100, 3, 30)).toEqual({ fit: 0, breaks: [] })
  })

  it('with Infinity wraps every block', () => {
    expect(fitLines([60, 60, 60, 60], 100, Infinity, 30)).toEqual({ fit: 4, breaks: [1, 2, 3] })
  })
})

describe("a host's button with a list", () => {
  it('opens its list on a click without `run`; split with `run`: the button runs, its arrow opens the list; an item runs and closes it', async () => {
    const run = vi.fn()
    const pick = vi.fn()
    const w = mountBar([
      { id: 'cf', label: 'Conditional Formatting', menu: [{ id: 'bars', label: 'Data Bars', run: pick }, { id: 'clear', label: 'Clear Rules', run: pick, disabled: true }] },
      { id: 'freeze', label: 'Freeze Panes', run, menu: [{ id: 'top', label: 'Freeze Top Row', run: pick }] },
    ])
    expect(w.find('[data-tb-host="cf"]').exists()).toBe(false)
    const cf = w.find('[data-tb-host-menu="cf"]')
    expect([cf.attributes('aria-haspopup'), cf.attributes('data-tip')]).toEqual(['menu', 'Conditional Formatting'])
    await cf.trigger('click')
    expect(w.findAll('[role="menu"] [data-tb-host-item]').map((b) => [b.text(), b.attributes('disabled') !== undefined])).toEqual([['Data Bars', false], ['Clear Rules', true]])
    await w.find('[data-tb-host-item="bars"]').trigger('click')
    expect(pick).toHaveBeenCalledOnce()
    expect(w.find('[role="menu"]').exists()).toBe(false)
    await w.find('[data-tb-host="freeze"]').trigger('click')
    expect(run).toHaveBeenCalledOnce()
    expect(w.find('[role="menu"]').exists()).toBe(false)
    await w.find('[data-tb-host-menu="freeze"]').trigger('click')
    expect(w.find('[data-tb-host-item="top"]').exists()).toBe(true)
  })
})

describe("SheetToolbar's ScreenTips", () => {
  const over = (w: VueWrapper, hook: string, pointerType = 'mouse') => w.find(hook).element.dispatchEvent(new PointerEvent('pointerover', { bubbles: true, pointerType }))
  const out = (w: VueWrapper, hook: string) => w.find(hook).element.dispatchEvent(new PointerEvent('pointerout', { bubbles: true, pointerType: 'mouse' }))
  const tipText = (w: VueWrapper) => (w.find('[data-tb-tip]').exists() ? w.find('[data-tb-tip]').text() : null)

  it("shows a button's name and key after a short rest of the mouse, the next button's at once, and none on a press or a touch", async () => {
    vi.useFakeTimers()
    try {
      const w = mountBar(['fontStyle', { id: 'share', label: 'Share', run: () => {} }])
      over(w, '[data-tb="bold"]')
      await nextTick()
      expect(tipText(w)).toBeNull()
      vi.advanceTimersByTime(400)
      await nextTick()
      expect(tipText(w)).toBe('Bold (Ctrl+B)')
      expect(w.find('[data-tb="bold"]').attributes('title')).toBeUndefined()
      out(w, '[data-tb="bold"]')
      over(w, '[data-tb-host="share"]')
      await nextTick()
      expect(tipText(w)).toBe('Share')
      w.find('[data-tb-host="share"]').element.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, pointerType: 'mouse' }))
      await nextTick()
      expect(tipText(w)).toBeNull()
      vi.advanceTimersByTime(2000)
      over(w, '[data-tb="italic"]', 'touch')
      vi.advanceTimersByTime(1000)
      await nextTick()
      expect(tipText(w)).toBeNull()
    } finally {
      vi.useRealTimers()
    }
  })
})
