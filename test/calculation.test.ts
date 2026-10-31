// Excel's calculation options on the sheet: in Manual the status bar shows Calculate while cells wait, F9 calculates the
// workbook and Shift+F9 the active sheet, the state reports the option, and the host sets it through the sheet; Data >
// Refresh All (Ctrl+Alt+F5, the toolbar's button) asks a source with refreshConnections again.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import type { BookInput } from 'xlcalc'
import { memorySource, SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import { cellText, editorKey, goTo, key, mountSheet, type, until } from './helpers'

let sources: LocalSource[] = []
afterEach(async () => {
  document.body.innerHTML = ''
  await Promise.all(sources.map((s) => s.close()))
  sources = []
})

const book = (calculation?: BookInput['calculation']): BookInput => ({
  ...(calculation ? { calculation } : {}),
  sheets: [
    { name: 'One', cells: { A1: { value: 1 }, B1: { formula: 'A1*2' } } },
    { name: 'Two', cells: { A1: { formula: 'One!A1*100' } } },
  ],
})
function open(calculation?: BookInput['calculation']): LocalSource {
  const s = localSource(book(calculation), { calc: 'inline' })
  sources.push(s)
  return s
}
const status = (w: VueWrapper) => w.find('[data-xs-calculate]')
const lastState = (w: VueWrapper) => (w.emitted('state') ?? []).at(-1)?.[0] as SheetViewState | null
async function enter(w: VueWrapper, place: string, text: string): Promise<void> {
  await goTo(w, place)
  await key(w, 'F2')
  await type(w, text)
  await editorKey(w, 'Enter')
}

describe('calculation options on the sheet', () => {
  it('in Manual an edit leaves the dependent as it was, the status bar shows Calculate, and F9 calculates', async () => {
    const w = await mountSheet(open('manual'))
    await until(() => cellText(w, 1, 2) === '2')
    expect(lastState(w)).toMatchObject({ calculation: 'manual' })
    expect(status(w).exists()).toBe(false)
    await enter(w, 'A1', '5')
    await until(() => status(w).exists())
    expect(cellText(w, 1, 1)).toBe('5')
    expect(cellText(w, 1, 2)).toBe('2')
    expect(status(w).text()).toBe('Calculate')
    expect(lastState(w)).toMatchObject({ calculation: 'manual', calculationPending: true })
    await key(w, 'F9')
    await until(() => cellText(w, 1, 2) === '10')
    await until(() => !status(w).exists())
    expect(lastState(w)?.calculationPending).toBeUndefined()
  })

  it('Shift+F9 calculates the active sheet only; a click on Calculate calculates the workbook', async () => {
    const s = open('manual')
    const calculate = vi.spyOn(s, 'calculate')
    const w = await mountSheet(s)
    await until(() => cellText(w, 1, 2) === '2')
    await enter(w, 'A1', '5')
    await until(() => status(w).exists())
    await key(w, 'F9', { shiftKey: true })
    await until(() => cellText(w, 1, 2) === '10')
    expect(calculate).toHaveBeenLastCalledWith('One')
    // sheet Two still waits
    expect(status(w).exists()).toBe(true)
    await status(w).trigger('click')
    await until(() => !status(w).exists())
    expect(calculate).toHaveBeenLastCalledWith()
  })

  it('the host sets the option through the sheet; a read-only sheet sends nothing', async () => {
    const s = open()
    const w = await mountSheet(s)
    await until(() => lastState(w)?.calculation === 'automatic')
    const api = w.vm as unknown as SheetViewApi
    expect(await api.setCalculation('manual')).toEqual({ ok: true, changes: [] })
    await until(() => lastState(w)?.calculation === 'manual')
    await w.setProps({ readOnly: true })
    expect(await api.setCalculation('automatic')).toBeNull()
    expect(await api.calculate()).toBeNull()
  })

  it('a source without calculation reports none and F9 does nothing', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    expect(lastState(w)?.calculation).toBeUndefined()
    await key(w, 'F9')
    expect(await (w.vm as unknown as SheetViewApi).calculate()).toBeNull()
  })
})

describe('Refresh All on the sheet', () => {
  /** A memory source whose refreshConnections answers with B1's new value. */
  function refreshing() {
    const source = memorySource({ sheets: [{ name: 'S', cells: { A1: 1, B1: 2 } }] })
    source.edit = vi.fn(async () => ({ ok: true as const, changes: [] }))
    source.format = vi.fn(async () => ({ ok: true as const, changes: [] }))
    source.refreshConnections = vi.fn(async () => ({ ok: true as const, changes: [{ sheet: 'S', cell: 'B1', text: '7', value: 7 }] }))
    return source
  }

  it('Ctrl+Alt+F5 and refreshAll() ask the source again; the state reports it for a ribbon', async () => {
    const source = refreshing()
    const w = await mountSheet(source)
    await until(() => cellText(w, 1, 2) === '2')
    expect(lastState(w)?.canRefreshAll).toBe(true)
    await key(w, 'F5', { ctrlKey: true, altKey: true })
    await until(() => cellText(w, 1, 2) === '7')
    expect(source.refreshConnections).toHaveBeenCalledWith()
    // F5 alone is still Go To
    expect(w.find('[data-xs-goto]').exists()).toBe(false)
    expect(await (w.vm as unknown as SheetViewApi).refreshAll()).toMatchObject({ ok: true })
    expect(source.refreshConnections).toHaveBeenCalledTimes(2)
  })

  it('a source without refreshConnections or a read-only sheet sends nothing', async () => {
    const w = await mountSheet(memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }))
    expect(lastState(w)?.canRefreshAll).toBeUndefined()
    await key(w, 'F5', { ctrlKey: true, altKey: true })
    expect(await (w.vm as unknown as SheetViewApi).refreshAll()).toBeNull()
    w.unmount()
    const source = refreshing()
    const ro = await mountSheet(source, { readOnly: true })
    expect(lastState(ro)?.canRefreshAll).toBeUndefined()
    await key(ro, 'F5', { ctrlKey: true, altKey: true })
    expect(await (ro.vm as unknown as SheetViewApi).refreshAll()).toBeNull()
    expect(source.refreshConnections).not.toHaveBeenCalled()
  })

  it("the toolbar's Refresh All shows with the source's refreshConnections and refreshes through the sheet", async () => {
    const source = refreshing()
    const Host = defineComponent(() => {
      const sheet = ref<SheetViewApi | null>(null)
      const state = ref<SheetViewState | null>(null)
      return () => [h(SheetToolbar, { source, sheet: sheet.value, state: state.value }), h(SheetView, { ref: sheet, source, style: 'height: 400px', onState: (x: SheetViewState | null) => (state.value = x) })]
    })
    const w = mount(Host, { attachTo: document.body })
    await until(() => w.find('[data-tb="refresh-all"]').exists() && w.find('[data-tb="refresh-all"]').attributes('disabled') === undefined)
    const button = w.find('[data-tb="refresh-all"]')
    expect(button.attributes('data-tip')).toBe('Refresh All (Ctrl+Alt+F5)')
    expect(button.attributes('data-tip-more')).toBe('Get the latest data by refreshing all sources in the workbook.')
    // the list's own name in Excel is its button's, Refresh All
    expect(w.find('[data-tb="refresh-all-menu"]').attributes('aria-label')).toBe('Refresh All')
    // its ScreenTip has the description under the name, as Office's
    button.element.dispatchEvent(new PointerEvent('pointerover', { bubbles: true, pointerType: 'mouse' }))
    await until(() => w.find('[data-tb-tip]').exists())
    expect(w.findAll('[data-tb-tip] > *').map((x) => x.text())).toEqual(['Refresh All (Ctrl+Alt+F5)', 'Get the latest data by refreshing all sources in the workbook.'])
    await button.trigger('click')
    await until(() => (source.refreshConnections as ReturnType<typeof vi.fn>).mock.calls.length === 1)
    w.unmount()
    const plain = memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] })
    plain.format = vi.fn(async () => ({ ok: true as const, changes: [] }))
    const none = mount(SheetToolbar, { props: { source: plain, state: null, sheet: null } })
    expect(none.find('[data-tb="refresh-all"]').exists()).toBe(false)
    none.unmount()
  })
})

