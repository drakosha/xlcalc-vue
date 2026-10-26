// Excel's AutoSum (Alt+=): the source's guess opens the active cell's edit with its SUM (or another function of its list)
// and the range selected; `./local` guesses through xlcalc's getAutoSum, with the formula in the workbook's locale. On the
// toolbar, Σ and its list go through the sheet.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { memorySource, SheetToolbar, SheetView, type SheetViewApi, type SheetViewState } from '../src'
import type { AutoSumFunction, SheetSource } from '../src/contract/source'
import { localSource } from '../src/local'
import { editorKey, editorOpen, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const numbers = { sheets: [{ name: 'S', cells: { B2: 1, B3: 2, B4: 3 } }] }
const field = (w: Awaited<ReturnType<typeof mountSheet>>) => w.find('.xs-editor').element as HTMLTextAreaElement
const altEquals = (w: Awaited<ReturnType<typeof mountSheet>>) => key(w, '≠', { code: 'Equal', altKey: true })

describe('AutoSum in ./local', () => {
  it("guesses the run of numbers above the cell, in the workbook's locale, and SUM() without numbers", async () => {
    const en = localSource({ sheets: [{ name: 'S', cells: { B2: { value: 1 }, B3: { value: 2 }, B4: { value: 3 } } }] }, { calc: 'inline' })
    expect(await en.autoSum('S', 'B5')).toEqual({ range: 'B2:B4', formula: 'SUM(B2:B4)' })
    expect(await en.autoSum('S', 'D1')).toEqual({ range: null, formula: 'SUM()' })
    await en.close()
    const ru = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 1 }, A2: { value: 2 } } }] }, { calc: 'inline' })
    expect(await ru.autoSum('Лист1', 'A3')).toEqual({ range: 'A1:A2', formula: 'СУММ(A1:A2)' })
    await ru.close()
  })

  it("builds the list's other functions over the same range, in the workbook's locale, subtotals too", async () => {
    const en = localSource({ sheets: [{ name: 'S', cells: { B2: { value: 1 }, B3: { value: 2 }, B4: { formula: 'SUM(B2:B3)' }, B5: { value: 4 }, B6: { formula: 'SUM(B5)' } } }] }, { calc: 'inline' })
    expect(en.autoSumFunctions).toEqual(['average', 'count', 'max', 'min'])
    expect(await en.autoSum('S', 'B7', undefined, 'average')).toEqual({ range: 'B6,B4', formula: 'AVERAGE(B6,B4)' })
    expect(await en.autoSum('S', 'B7', undefined, 'count')).toEqual({ range: 'B6,B4', formula: 'COUNT(B6,B4)' })
    expect(await en.autoSum('S', 'D1', undefined, 'max')).toEqual({ range: null, formula: 'MAX()' })
    await en.close()
    const ru = localSource({ locale: 'ru-RU', sheets: [{ name: 'Лист1', cells: { A1: { value: 1 }, A2: { formula: 'SUM(A1)' }, A3: { value: 2 }, A4: { formula: 'SUM(A3)' } } }] }, { calc: 'inline' })
    const names = await Promise.all((['average', 'count', 'max', 'min'] as AutoSumFunction[]).map(async (fn) => (await ru.autoSum('Лист1', 'A5', undefined, fn))?.formula))
    expect(names).toEqual(['СРЗНАЧ(A4;A2)', 'СЧЁТ(A4;A2)', 'МАКС(A4;A2)', 'МИН(A4;A2)'])
    await ru.close()
  })
})

describe('AutoSum on the sheet', () => {
  it('Alt+= opens the edit with the SUM and its range selected; Enter enters it', async () => {
    const source = memorySource(numbers)
    source.autoSum = vi.fn(async () => ({ range: 'B2:B4', formula: 'SUM(B2:B4)' }))
    const edit = vi.spyOn(source, 'edit')
    const w = await mountSheet(source)
    await goTo(w, 'B5')
    await altEquals(w)
    await settle()
    expect(source.autoSum).toHaveBeenCalledWith('S', 'B5')
    expect(editorOpen(w)).toBe(true)
    expect(field(w).value).toBe('=SUM(B2:B4)')
    expect([field(w).selectionStart, field(w).selectionEnd]).toEqual([5, 10])
    await editorKey(w, 'Enter')
    await settle()
    expect(edit.mock.calls[0]![0]).toEqual([{ sheet: 'S', cell: 'B5', text: '=SUM(B2:B4)' }])
  })

  it('without a range the caret is between the parentheses; the method does the same', async () => {
    const source = memorySource(numbers)
    source.autoSum = async () => ({ range: null, formula: 'SUM()' })
    const w = await mountSheet(source)
    await goTo(w, 'D1')
    expect(await (w.vm as unknown as SheetViewApi).autoSum()).toBe(true)
    await settle()
    expect(field(w).value).toBe('=SUM()')
    expect([field(w).selectionStart, field(w).selectionEnd]).toEqual([5, 5])
  })

  it('with a function the source lists, the edit opens with that formula; one it does not list sends nothing', async () => {
    const source = memorySource(numbers)
    source.autoSum = vi.fn(async (_s: string, _c: string, _signal?: AbortSignal, fn?: AutoSumFunction) => ({ range: 'B2:B4', formula: `${(fn ?? 'sum').toUpperCase()}(B2:B4)` }))
    const w = await mountSheet(source)
    await goTo(w, 'B5')
    const api = w.vm as unknown as SheetViewApi
    expect(await api.autoSum('max')).toBe(false)
    expect(source.autoSum).not.toHaveBeenCalled()
    Object.assign(source, { autoSumFunctions: ['max'] })
    expect(await api.autoSum('max')).toBe(true)
    await settle()
    expect(source.autoSum).toHaveBeenCalledWith('S', 'B5', undefined, 'max')
    expect(field(w).value).toBe('=MAX(B2:B4)')
    expect([field(w).selectionStart, field(w).selectionEnd]).toEqual([5, 10])
  })

  it('without the source’s autoSum, or on a read-only sheet, there is none', async () => {
    const w = await mountSheet(memorySource(numbers))
    expect(await (w.vm as unknown as SheetViewApi).autoSum()).toBe(false)
    const source = memorySource(numbers)
    source.autoSum = vi.fn(async () => ({ range: null, formula: 'SUM()' }))
    const ro = await mountSheet(source, { readOnly: true })
    await altEquals(ro)
    await settle()
    expect(source.autoSum).not.toHaveBeenCalled()
    expect(ro.find('.xs-editor').exists()).toBe(false)
  })
})

/** The toolbar over the sheet, as a host puts them together. */
async function mountBoth(source: SheetSource, props: Record<string, unknown> = {}): Promise<VueWrapper> {
  const Host = defineComponent(() => {
    const sheet = ref<SheetViewApi | null>(null)
    const state = ref<SheetViewState | null>(null)
    return () => [
      h(SheetToolbar, { source, sheet: sheet.value, state: state.value }),
      h(SheetView, { ref: sheet, source, style: 'height: 400px', ...props, onState: (s: SheetViewState | null) => (state.value = s) }),
    ]
  })
  const w = mount(Host, { attachTo: document.body })
  await settle(10)
  return w
}
const tb = (w: VueWrapper, hook: string) => w.find(`[data-tb="${hook}"]`)

describe('AutoSum on the toolbar', () => {
  function withAutoSum(functions?: AutoSumFunction[]) {
    const source = memorySource(numbers)
    source.format = vi.fn(async () => ({ ok: true as const, changes: [] }))
    source.autoSum = vi.fn(async (_s: string, _c: string, _signal?: AbortSignal, fn?: AutoSumFunction) => ({ range: 'B2:B4', formula: `${(fn ?? 'sum').toUpperCase()}(B2:B4)` }))
    if (functions) Object.assign(source, { autoSumFunctions: functions })
    return source
  }

  it('Σ enters the SUM in the active cell; its list has the functions the source builds and More Functions…', async () => {
    const source = withAutoSum(['average', 'count', 'max', 'min'])
    Object.assign(source, { functionNames: async () => ['SUM', 'AVERAGE'] })
    const w = await mountBoth(source)
    await goTo(w, 'B5')
    await tb(w, 'auto-sum').trigger('click')
    await settle()
    expect(field(w).value).toBe('=SUM(B2:B4)')
    await editorKey(w, 'Escape')
    await tb(w, 'auto-sum-menu').trigger('click')
    expect(w.findAll('[data-tb^="auto-sum-"]').map((b) => b.attributes('data-tb'))).toEqual(['auto-sum-menu', 'auto-sum-sum', 'auto-sum-average', 'auto-sum-count', 'auto-sum-max', 'auto-sum-min', 'auto-sum-more'])
    expect(w.find('[role="menu"]').text()).toContain('Count Numbers')
    await tb(w, 'auto-sum-count').trigger('click')
    await settle()
    expect(source.autoSum).toHaveBeenLastCalledWith('S', 'B5', undefined, 'count')
    expect(field(w).value).toBe('=COUNT(B2:B4)')
    expect(tb(w, 'auto-sum-count').exists()).toBe(false)
    await editorKey(w, 'Escape')
    await tb(w, 'auto-sum-menu').trigger('click')
    await tb(w, 'auto-sum-more').trigger('click')
    await settle()
    expect(w.find('[data-xs-insert-function]').exists()).toBe(true)
  })

  it('without autoSumFunctions the list is Sum alone; without functionNames there is no More Functions…', async () => {
    const w = await mountBoth(withAutoSum())
    await goTo(w, 'B5')
    await tb(w, 'auto-sum-menu').trigger('click')
    expect(w.findAll('[role="menu"] [data-tb]').map((b) => b.attributes('data-tb'))).toEqual(['auto-sum-sum'])
  })

  it("is disabled without the source's autoSum and on a read-only sheet", async () => {
    const plain = memorySource(numbers)
    plain.format = vi.fn(async () => ({ ok: true as const, changes: [] }))
    const w = await mountBoth(plain)
    expect(tb(w, 'auto-sum').attributes('disabled')).toBeDefined()
    expect(tb(w, 'auto-sum-menu').attributes('disabled')).toBeDefined()
    w.unmount()
    const source = withAutoSum()
    const ro = await mountBoth(source, { readOnly: true })
    expect(tb(ro, 'auto-sum').attributes('disabled')).toBeDefined()
    await tb(ro, 'auto-sum').trigger('click')
    expect(source.autoSum).not.toHaveBeenCalled()
  })
})
