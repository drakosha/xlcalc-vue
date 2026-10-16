// The sheet's own Find and Replace and Format Cells: Control+H opens the bar with its replace row, Control+1 and the
// API's formatCells() open Format Cells at the active cell's look, OK formats the selection in one call of the source.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { memorySource, SheetToolbar, SheetView, type MemoryBook, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/source'
import { goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const book = (): MemoryBook => ({ styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: 'apple', B2: { value: 5, style: 0 } } }] })
/** A memory book with a source's `format` that takes the patches and changes nothing. */
const formatting = () => {
  const format = vi.fn(async (_sheet: string, _patches: FormatPatch[]): Promise<EditOutcome> => ({ ok: true, changes: [] }))
  return { source: { ...memorySource(book()), format } as SheetSource, format }
}

describe('the sheet’s dialogs', () => {
  it('Control+H opens the search bar with its replace row, the Russian layout too', async () => {
    const w = await mountSheet(memorySource(book()))
    await key(w, 'h', { ctrlKey: true })
    await settle()
    expect(w.find('[data-xs-replace-input]').exists()).toBe(true)
    await key(w, 'Escape')
    await settle()
    expect(w.find('.xs-find').exists()).toBe(false)
    await key(w, 'р', { ctrlKey: true, code: 'KeyH' })
    await settle()
    expect(w.find('[data-xs-replace-input]').exists()).toBe(true)
  })

  it('Control+1 opens Format Cells at the active cell’s look; OK formats the selection in one call', async () => {
    const { source, format } = formatting()
    const w = await mountSheet(source)
    await goTo(w, 'B2')
    await key(w, '1', { ctrlKey: true, code: 'Digit1' })
    await settle()
    expect(w.find('[data-fmt]').exists()).toBe(true)
    await w.find('[data-fmt-tab="font"]').trigger('click')
    expect((w.find('[data-fmt-font-style]').element as HTMLSelectElement).value).toBe('bold')
    await w.find('[data-fmt-font-style]').setValue('italic')
    await w.find('[data-fmt-ok]').trigger('submit')
    await settle()
    expect(format).toHaveBeenCalledTimes(1)
    expect(format.mock.calls[0]).toEqual(['S', [{ range: 'B2', style: { font: { bold: null, italic: true } } }]])
    expect(w.find('[data-fmt]').exists()).toBe(false)
  })

  it('a read-only sheet has no Format Cells: Control+1 does nothing and formatCells() is false', async () => {
    const w = await mountSheet(memorySource(book()), { readOnly: true })
    await key(w, '1', { ctrlKey: true, code: 'Digit1' })
    await settle()
    expect(w.find('[data-fmt]').exists()).toBe(false)
    expect((w.vm as unknown as SheetViewApi).formatCells()).toBe(false)
  })

  it('the toolbar’s Format Cells… opens the sheet’s dialog when it has the sheet', async () => {
    const { source } = formatting()
    const Host = defineComponent(() => {
      const sheet = ref<SheetViewApi | null>(null)
      const state = ref<SheetViewState | null>(null)
      return () => [
        h(SheetToolbar, { source, sheet: sheet.value, state: state.value }),
        h(SheetView, { ref: sheet, source, style: 'height: 400px', onState: (s: SheetViewState | null) => (state.value = s) }),
      ]
    })
    const w: VueWrapper = mount(Host, { attachTo: document.body })
    await settle(10)
    await w.find('[data-tb="format-cells"]').trigger('click')
    await settle()
    expect(w.find('[data-fmt]').exists()).toBe(true)
  })
})
