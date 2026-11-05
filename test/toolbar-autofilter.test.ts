// Excel's Data > Filter on the toolbar: where it is drawn, its state, and the AutoFilter it puts on and takes off through the sheet.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { h, ref } from 'vue'
import SheetToolbar from '../src/components/SheetToolbar.vue'
import SheetView from '../src/components/SheetView.vue'
import type { SheetSource } from '../src/contract/source'
import type { SheetViewApi, SheetViewState, ToolbarItem } from '../src/contract/view'
import { localSource } from '../src/local'
import { ruMessages } from '../src/messages'
import { TOOLBAR_GROUPS } from '../src/toolbar'
import { settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  localStorage.clear()
})

const state = { sheet: 'S', ref: 'A1', canAutoFilter: true } as SheetViewState
const source = (over: Partial<SheetSource> = {}) => ({ format: vi.fn(), setAutoFilter: vi.fn(), ...over }) as unknown as SheetSource
const sheetApi = () => ({ toggleAutoFilter: vi.fn(async () => ({ ok: true })) }) as unknown as SheetViewApi
const bar = (props: Record<string, unknown>) => mount(SheetToolbar, { props: { source: source(), state, ...props } as never, attachTo: document.body })
const button = (w: VueWrapper) => w.find('[data-tb="auto-filter"]')

describe('the toolbar\'s Filter button', () => {
  it('comes right after AutoSum in the default order', () => {
    expect(TOOLBAR_GROUPS.indexOf('autoFilter')).toBe(TOOLBAR_GROUPS.indexOf('autoSum') + 1)
  })

  it('is drawn with the sheet and the source\'s setAutoFilter only', () => {
    expect(button(bar({ sheet: sheetApi() })).exists()).toBe(true)
    expect(button(bar({ sheet: null })).exists()).toBe(false)
    expect(button(bar({ sheet: sheetApi(), source: source({ setAutoFilter: undefined }) })).exists()).toBe(false)
  })

  it('is left out by items that do not name it, and placed by those that do', () => {
    const items: ToolbarItem[] = ['autoFilter', 'autoSum']
    const w = bar({ sheet: sheetApi(), items })
    expect(w.findAll('[data-tb-block]').map((g) => g.attributes('data-tb-block'))).toEqual(['autoFilter', 'autoSum'])
    expect(button(bar({ sheet: sheetApi(), items: ['autoSum'] })).exists()).toBe(false)
  })

  it('is pressed while the sheet has an AutoFilter, disabled without canAutoFilter', () => {
    const on = button(bar({ sheet: sheetApi(), state: { ...state, autoFilter: 'A1:B3' } }))
    expect([on.attributes('aria-pressed'), on.classes().includes('xs-tool-on')]).toEqual(['true', true])
    const off = button(bar({ sheet: sheetApi() }))
    expect([off.attributes('aria-pressed'), off.classes().includes('xs-tool-on')]).toEqual(['false', false])
    expect(button(bar({ sheet: sheetApi(), state: { ...state, canAutoFilter: undefined } })).attributes('disabled')).toBeDefined()
  })

  it('calls the sheet once and reports done', async () => {
    const sheet = sheetApi()
    const w = bar({ sheet })
    await button(w).trigger('click')
    await settle()
    expect(sheet.toggleAutoFilter).toHaveBeenCalledTimes(1)
    expect(w.emitted('done')).toHaveLength(1)
  })

  it('takes Excel\'s name, in Russian too, and its label in the ribbon without the key', async () => {
    expect(button(bar({ sheet: sheetApi() })).attributes('aria-label')).toBe('Filter (Ctrl+Shift+L)')
    expect(button(bar({ sheet: sheetApi(), messages: ruMessages })).attributes('aria-label')).toBe('Фильтр (Ctrl+Shift+L)')
    const w = bar({ sheet: sheetApi(), rows: 2, items: ['autoFilter'] })
    expect(button(w).find('.xs-tb-label').text()).toBe('Filter')
  })

  it('puts an AutoFilter on and takes it off through a real sheet, one step of undo each', async () => {
    const s = localSource({ sheets: [{ name: 'S', cells: { A1: { value: 'Name' }, B1: { value: 'Sum' }, A2: { value: 'x' }, B2: { value: 1 }, A3: { value: 'y' }, B3: { value: 2 } } }] }, { calc: 'inline' })
    const sheet = ref<SheetViewApi | null>(null)
    const st = ref<SheetViewState | null>(null)
    const w = mount(
      {
        render: () =>
          h('div', [
            h(SheetToolbar, { source: s, state: st.value, sheet: sheet.value }),
            h(SheetView, { source: s, ref: (x: unknown) => (sheet.value = x as SheetViewApi | null), onState: (x: SheetViewState | null) => (st.value = x) }),
          ]),
      },
      { attachTo: document.body },
    )
    await vi.waitFor(() => expect(st.value?.canAutoFilter).toBe(true))
    await settle()
    expect(button(w).attributes('aria-pressed')).toBe('false')
    await button(w).trigger('click')
    await vi.waitFor(() => expect(st.value?.autoFilter).toBe('A1:B3'))
    await settle()
    expect(button(w).attributes('aria-pressed')).toBe('true')
    await w.find('[data-tb="undo"]').trigger('click')
    await vi.waitFor(() => expect(st.value?.autoFilter).toBeUndefined())
    await button(w).trigger('click')
    await vi.waitFor(() => expect(st.value?.autoFilter).toBe('A1:B3'))
    await settle()
    await button(w).trigger('click')
    await vi.waitFor(() => expect(st.value?.autoFilter).toBeUndefined())
    w.unmount()
    s.close()
  })
})
