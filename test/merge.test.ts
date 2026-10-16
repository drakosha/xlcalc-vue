// Excel's Merge & Center on the toolbar: the button merges and centres the selection, its list merges across, merges
// or unmerges; it is pressed on a merged cell and then unmerges. Excel's warning comes before cells holding more than one
// value are merged. Through the local source a merged area is drawn as one cell, and Control+Z takes it back.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { memorySource, ruMessages, SheetToolbar, SheetView, type MemoryBook, type SheetViewApi, type SheetViewState } from '../src'
import { localSource } from '../src/local'
import type { EditOutcome, SheetOp, SheetSource } from '../src/source'
import { goTo, key, settle } from './helpers'

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

const OK: EditOutcome = { ok: true, changes: [] }

/** A memory book with the source's `structure` as a mock answering `answer`, and a `format` for the toolbar to show. */
function withStructure(cells: MemoryBook['sheets'][number]['cells'] = {}, more: Partial<MemoryBook['sheets'][number]> = {}, answer: (op: SheetOp) => EditOutcome = () => OK) {
  const source = memorySource({ sheets: [{ name: 'S', cells, ...more }] })
  const structure = vi.fn(async (op: SheetOp) => answer(op))
  source.structure = structure
  source.format = vi.fn(async () => OK)
  return { source: source as SheetSource, structure }
}

/** The toolbar over the sheet, as a host puts them together. */
async function mountBoth(source: SheetSource, props: Record<string, unknown> = {}): Promise<VueWrapper> {
  const Host = defineComponent(() => {
    const sheet = ref<SheetViewApi | null>(null)
    const state = ref<SheetViewState | null>(null)
    return () => [
      h(SheetToolbar, { source, sheet: sheet.value, state: state.value, ...props }),
      h(SheetView, { ref: sheet, source, style: 'height: 400px', ...props, onState: (s: SheetViewState | null) => (state.value = s) }),
    ]
  })
  const w = mount(Host, { attachTo: document.body })
  await settle(10)
  return w
}
const ops = (structure: ReturnType<typeof withStructure>['structure']) => structure.mock.calls.map((c) => c[0])
async function pick(w: VueWrapper, item: string): Promise<void> {
  await w.find('[data-tb="merge-menu"]').trigger('click')
  await w.find(`[data-tb="${item}"]`).trigger('click')
  await settle()
}
const alertOf = (w: VueWrapper) => w.find('[data-xs-alert="merge"]')

describe('Merge & Center on the toolbar', () => {
  it('each command calls the source’s structure over the selection: Merge & Center, Merge Across, Merge Cells, Unmerge Cells', async () => {
    const { source, structure } = withStructure()
    const w = await mountBoth(source)
    await goTo(w, 'B2:C3')
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    for (const item of ['merge-center', 'merge-across', 'merge-cells', 'unmerge-cells']) await pick(w, item)
    expect(ops(structure)).toEqual([
      { op: 'mergeCells', sheet: 'S', range: 'B2:C3', center: true },
      { op: 'mergeCells', sheet: 'S', range: 'B2:C3', center: true },
      { op: 'mergeCells', sheet: 'S', range: 'B2:C3', across: true },
      { op: 'mergeCells', sheet: 'S', range: 'B2:C3' },
      { op: 'unmergeCells', sheet: 'S', range: 'B2:C3' },
    ])
  })

  it('is pressed on a merged cell, and pressed then it unmerges, as Excel’s', async () => {
    const { source, structure } = withStructure({ A1: 'x' }, { merges: ['A1:B2'] })
    const w = await mountBoth(source)
    await goTo(w, 'D4')
    expect(w.find('[data-tb="merge"]').attributes('aria-pressed')).toBe('false')
    await goTo(w, 'A1')
    expect(w.find('[data-tb="merge"]').attributes('aria-pressed')).toBe('true')
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    expect(ops(structure)).toEqual([{ op: 'unmergeCells', sheet: 'S', range: 'A1' }])
  })

  it('does nothing on one cell outside a merged area, as Excel', async () => {
    const { source, structure } = withStructure({ C5: 1 })
    const w = await mountBoth(source)
    await goTo(w, 'C5')
    await w.find('[data-tb="merge"]').trigger('click')
    for (const item of ['merge-across', 'merge-cells', 'unmerge-cells']) await pick(w, item)
    await settle()
    expect(structure).not.toHaveBeenCalled()
  })

  it('warns before merging two cells that hold values: Cancel and Esc leave them, OK merges', async () => {
    const { source, structure } = withStructure({ A1: 1, B1: { value: 2, formula: '1+1' } })
    const w = await mountBoth(source, { messages: ruMessages })
    await goTo(w, 'A1:B1')
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    expect(alertOf(w).exists()).toBe(true)
    expect(w.find('[data-xs-alert-message]').text()).toBe(ruMessages.mergeWarning)
    expect(structure).not.toHaveBeenCalled()
    await w.find('[data-xs-alert-button="cancel"]').trigger('click')
    await settle()
    expect(alertOf(w).exists()).toBe(false)
    expect(structure).not.toHaveBeenCalled()
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    await alertOf(w).trigger('keydown', { key: 'Escape' })
    await settle()
    expect(alertOf(w).exists()).toBe(false)
    expect(structure).not.toHaveBeenCalled()
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    await w.find('[data-xs-alert-button="ok"]').trigger('click')
    await settle()
    expect(ops(structure)).toEqual([{ op: 'mergeCells', sheet: 'S', range: 'A1:B1', center: true }])
  })

  it('merges without a warning when one cell holds a value, and Merge Across when each row holds one', async () => {
    const { source, structure } = withStructure({ A1: 1, A2: 2 })
    const w = await mountBoth(source)
    await goTo(w, 'A1:B1')
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    await goTo(w, 'A1:B2')
    await pick(w, 'merge-across')
    expect(alertOf(w).exists()).toBe(false)
    expect(ops(structure)).toEqual([
      { op: 'mergeCells', sheet: 'S', range: 'A1:B1', center: true },
      { op: 'mergeCells', sheet: 'S', range: 'A1:B2', across: true },
    ])
    await pick(w, 'merge-cells')
    expect(alertOf(w).exists()).toBe(true)
  })

  it('counts the cells past the loaded windows through the source’s windows', async () => {
    const { source, structure } = withStructure({ A1: 1, A4000: 2 })
    const w = await mountBoth(source)
    await goTo(w, 'A1:A5000')
    await pick(w, 'merge-cells')
    expect(alertOf(w).exists()).toBe(true)
    await w.find('[data-xs-alert-button="ok"]').trigger('click')
    await settle()
    expect(ops(structure)).toEqual([{ op: 'mergeCells', sheet: 'S', range: 'A1:A5000' }])
  })

  it('shows the source’s refusal in the sheet’s message line', async () => {
    const { source } = withStructure({}, {}, () => ({ ok: false, reason: 'refused', message: 'A table’s cells are not merged' }))
    const w = await mountBoth(source)
    await goTo(w, 'A1:B1')
    await w.find('[data-tb="merge"]').trigger('click')
    await settle()
    expect(w.find('.xs-message').text()).toBe('A table’s cells are not merged')
  })

  it('is not there without the source’s structure, and off on a read-only sheet', async () => {
    const format = vi.fn(async (): Promise<EditOutcome> => OK)
    const plain = await mountBoth({ ...memorySource({ sheets: [{ name: 'S', cells: { A1: 1 } }] }), format })
    expect(plain.find('[data-tb="bold"]').exists()).toBe(true)
    expect(plain.find('[data-tb="merge"]').exists()).toBe(false)
    document.body.innerHTML = ''
    const { source, structure } = withStructure({ A1: 1 })
    const w = await mountBoth(source, { readOnly: true })
    await goTo(w, 'A1:B2')
    expect(w.find('[data-tb="merge"]').attributes('disabled')).toBeDefined()
    await expect((w.findComponent(SheetView).vm as unknown as SheetViewApi).merge('mergeCells')).resolves.toBeNull()
    await settle()
    expect(structure).not.toHaveBeenCalled()
  })
})

describe('Merge Cells through the local source', () => {
  it('A1:B2 merged is drawn as one cell at once, and Control+Z splits it again', async () => {
    const source = localSource({ locale: 'en-US', sheets: [{ name: 'S', cells: { A1: { value: 1 } } }] }, { calc: 'inline' })
    const w = await mountBoth(source)
    await settle(12)
    await goTo(w, 'A1:B2')
    await pick(w, 'merge-cells')
    await settle(12)
    expect(alertOf(w).exists()).toBe(false)
    const a1 = w.find('[data-cell="1:1"]')
    expect([a1.attributes('colspan'), a1.attributes('rowspan')]).toEqual(['2', '2'])
    expect(w.find('[data-cell="1:2"]').exists()).toBe(false)
    expect(w.find('[data-tb="merge"]').attributes('aria-pressed')).toBe('true')
    await key(w, 'z', { ctrlKey: true })
    await settle(12)
    expect(w.find('[data-cell="1:1"]').attributes('colspan')).toBeUndefined()
    expect(w.find('[data-cell="1:2"]').exists()).toBe(true)
    await source.close()
  })
})
