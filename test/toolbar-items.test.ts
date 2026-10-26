// The toolbar's `items`: the groups it draws and their order, the host's buttons drawn as its own, and in a narrow viewport
// Undo and the font style in the row with everything else, the host's buttons too, in the menu.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h } from 'vue'
import { SheetToolbar, TOOLBAR_GROUPS, type SheetViewState, type ToolbarItem } from '../src'
import type { EditOutcome, SheetSource } from '../src/contract/source'

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
    expect(hooks(w).slice(0, 6)).toEqual(['undo', 'redo', 'bold', 'italic', 'underline', 'font-name'])
    expect(hooks(w)).toContain('borders')
  })

  it('draws only the groups named, in their order; an unknown or repeated name is skipped', async () => {
    const w = mountBar(['borders', 'fontStyle', 'nothing' as never, 'borders', 'undo'])
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
    expect(hooks(w)).toEqual(['bold', 'italic', 'underline', 'host:share', 'host:pin', 'host:save', 'undo', 'redo'])
    expect(w.findAll('.xs-tb-group')).toHaveLength(3)
    const share = w.find('[data-tb-host="share"]')
    expect(share.attributes('title')).toBe('Share')
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

  it('in a narrow viewport keeps Undo and the font style in the row and moves the rest, the host’s buttons too, into the menu', async () => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(390)
    const run = vi.fn()
    const w = mountBar(['autoSum', 'borders', { id: 'share', label: 'Share', run }, 'fontStyle', 'undo'])
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

  it('in a narrow viewport without anything for the menu has no More button', async () => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(390)
    const w = mountBar(['undo', 'fontStyle'])
    expect(hooks(w)).toEqual(['undo', 'redo', 'bold', 'italic', 'underline'])
    expect(w.find('[data-tb-menu]').exists()).toBe(false)
  })
})
