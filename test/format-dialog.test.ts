// The Format Cells dialog: it opens with the active cell's formatting, OK emits patches for what changed and closes, Cancel
// and Esc close; inputs beyond Excel's limits block OK; the toolbar's Format Cells button asks the host to open it.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { FormatDialog, SheetToolbar, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'
import { settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const at = (ref: string, more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'S', ref, ...more })
async function open(state: SheetViewState, locale?: string): Promise<VueWrapper> {
  const w = mount(FormatDialog, { props: { open: true, state, locale }, attachTo: document.body })
  await vi.dynamicImportSettled()
  await settle()
  return w
}
const tab = (w: VueWrapper, t: string) => w.find(`[data-fmt-tab="${t}"]`).trigger('click')
const okButton = (w: VueWrapper) => w.find('[data-fmt-ok]')
const value = (w: VueWrapper, sel: string) => (w.find(sel).element as HTMLInputElement | HTMLSelectElement).value

describe('FormatDialog', () => {
  it("opens with the active cell's formatting: its number format, font, alignment and fill", async () => {
    const w = await open(at('B2', { numberFormat: '0.00%', style: { font: { bold: true, size: 14 }, alignment: { horizontal: 'right', indent: 2 }, fill: { color: '#ffff00' } } }))
    expect((w.find('[data-fmt-number="percentage"]').element as HTMLInputElement).checked).toBe(true)
    expect(value(w, '[data-fmt-code]')).toBe('0.00%')
    expect(w.find('[data-fmt-category]').text()).toBe('Category: Percentage')
    expect([value(w, '[data-fmt-font-style]'), value(w, '[data-fmt-font-size]')]).toEqual(['bold', '14'])
    expect([value(w, '[data-fmt-horizontal]'), value(w, '[data-fmt-indent]'), value(w, '[data-fmt-vertical]')]).toEqual(['right', '2', 'bottom'])
    expect(w.find('[data-fmt-colour-of="fill"] [aria-pressed="true"]').attributes('data-xs-colour')).toBe('#FFFF00')
    expect(w.find('[data-fmt-colour-of="font"] [data-fmt-colour-none]').attributes('aria-pressed')).toBe('true')
  })

  it('OK emits what changed over the selection and closes; with nothing changed it only closes', async () => {
    const w = await open(at('B2:C3', { style: { font: { bold: true } } }))
    await tab(w, 'font')
    await w.find('[data-fmt-font-style]').setValue('italic')
    await w.find('[data-fmt-font-size]').setValue('12')
    await w.find('[data-fmt-colour-of="font"] [data-xs-colour="#C00000"]').trigger('click')
    await tab(w, 'fill')
    await w.find('[data-fmt-colour-of="fill"] [data-fmt-colour-code]').setValue('92d050')
    await okButton(w).trigger('submit')
    expect(w.emitted('apply')).toEqual([[[{ style: { font: { size: 12, bold: null, italic: true, color: '#C00000' }, fill: { color: '#92D050' } } }]]])
    expect(w.emitted('close')).toHaveLength(1)
    const same = await open(at('A1'))
    await same.find('form').trigger('submit')
    expect(same.emitted('apply')).toBeUndefined()
    expect(same.emitted('close')).toHaveLength(1)
  })

  it('Cancel, Esc and the close button close without changes', async () => {
    const w = await open(at('A1'))
    await tab(w, 'font')
    await w.find('[data-fmt-strike]').setValue(true)
    await w.find('[data-fmt-cancel]').trigger('click')
    await w.find('form').trigger('keydown', { key: 'Escape' })
    await w.find('.xs-fmt-x').trigger('click')
    expect(w.emitted('apply')).toBeUndefined()
    expect(w.emitted('close')).toHaveLength(3)
  })

  it("draws the Border tab's line on the edges: a red medium outline of a range in one call", async () => {
    const w = await open(at('B2:C3'))
    await tab(w, 'border')
    await w.find('[data-fmt-line="medium"]').trigger('click')
    await w.find('[data-fmt-colour-of="line"] [data-fmt-colour-code]').setValue('#c00000')
    await w.find('[data-fmt-preset="outline"]').trigger('click')
    expect(w.find('[data-fmt-edge="top"]').attributes('aria-pressed')).toBe('true')
    expect(w.find('[data-fmt-edge="insideH"]').attributes('aria-pressed')).toBe('false')
    await w.find('form').trigger('submit')
    const RED = { style: 'medium', color: '#C00000' }
    // plain data, as a source in a worker thread clones it
    expect(() => structuredClone(w.emitted('apply')![0]![0])).not.toThrow()
    expect(w.emitted('apply')![0]![0]).toEqual([
      { range: 'B2:C2', style: { border: { top: RED } } },
      { range: 'B3:C3', style: { border: { bottom: RED } } },
      { range: 'B2:B3', style: { border: { left: RED } } },
      { range: 'C2:C3', style: { border: { right: RED } } },
    ])
  })

  it('a single cell has no inside lines, and its sides open as the cell has them', async () => {
    const w = await open(at('B2', { style: { border: { bottom: { style: 'thin' } } } }))
    await tab(w, 'border')
    expect(w.find('[data-fmt-edge="insideV"]').attributes('disabled')).toBeDefined()
    expect(w.find('[data-fmt-preset="inside"]').attributes('disabled')).toBeDefined()
    expect(w.find('[data-fmt-edge="bottom"]').attributes('aria-pressed')).toBe('true')
    await w.find('[data-fmt-edge="bottom"]').trigger('click')
    await w.find('form').trigger('submit')
    expect(w.emitted('apply')![0]![0]).toEqual([{ range: 'B2', style: { border: { bottom: null } } }])
  })

  it("the Number tab is Excel's list: an item puts its code in the input, and a typed code is named by its category", async () => {
    const w = await open(at('A1'))
    expect(w.find('[data-fmt-category]').text()).toBe('Category: General')
    await w.find('[data-fmt-number="scientific"]').setValue(true)
    expect(value(w, '[data-fmt-code]')).toBe('0.00E+00')
    await w.find('[data-fmt-code]').setValue('0.000')
    expect(w.find('[data-fmt-category]').text()).toBe('Category: Number')
    expect(w.findAll('[data-fmt-number]').filter((r) => (r.element as HTMLInputElement).checked)).toHaveLength(0)
    await w.find('form').trigger('submit')
    expect(w.emitted('apply')).toEqual([[[{ numberFormat: '0.000' }]]])
  })

  it("names the list's items in the workbook's locale", async () => {
    const w = await open(at('A1', { numberFormat: '0%' }), 'ru-RU')
    expect(w.findAll('.xs-fmt-choice').slice(0, 3).map((c) => c.text())).toEqual(['Общий', 'Числовой', 'Денежный'])
    expect(w.find('[data-fmt-category]').attributes('data-fmt-category-key')).toBe('percentage')
  })

  it("blocks OK while an input is beyond Excel's limits or a color code is invalid", async () => {
    const w = await open(at('A1'))
    await tab(w, 'font')
    await w.find('[data-fmt-font-size]').setValue('500')
    expect(w.find('.xs-fmt-error').text()).toBe('The size must be between 1 and 409')
    expect(okButton(w).attributes('disabled')).toBeDefined()
    await w.find('form').trigger('submit')
    expect(w.emitted('close')).toBeUndefined()
    await w.find('[data-fmt-font-size]').setValue('11')
    await w.find('[data-fmt-colour-of="font"] [data-fmt-colour-code]').setValue('#12345')
    expect(okButton(w).attributes('disabled')).toBeDefined()
    await w.find('[data-fmt-colour-of="font"] [data-fmt-colour-code]').setValue('#123456')
    expect(okButton(w).attributes('disabled')).toBeUndefined()
  })

  it('shows nothing while closed or without a selection', async () => {
    expect(mount(FormatDialog, { props: { open: false, state: at('A1') } }).find('[data-fmt]').exists()).toBe(false)
    expect(mount(FormatDialog, { props: { open: true, state: null } }).find('[data-fmt]').exists()).toBe(false)
  })
})

describe('the toolbar’s Format Cells button', () => {
  const source = { sheets: async () => [], window: async () => ({ pending: { elapsedMs: 0 } }), format: vi.fn(async (_s: string, _p: FormatPatch[]): Promise<EditOutcome> => ({ ok: true, changes: [] })) } as SheetSource
  it('is present when the host listens to `formatCells`, which it emits', async () => {
    const onFormatCells = vi.fn()
    const w = mount(SheetToolbar, { props: { source, state: at('A1'), onFormatCells }, attachTo: document.body })
    await settle()
    await w.find('[data-tb="format-cells"]').trigger('click')
    expect(onFormatCells).toHaveBeenCalledTimes(1)
    const plain = mount(SheetToolbar, { props: { source, state: at('A1') }, attachTo: document.body })
    expect(plain.find('[data-tb="format-cells"]').exists()).toBe(false)
  })

  it("opens the host's dialog once the host listens, even when it started listening after the toolbar was created", async () => {
    const sheet = { formatCells: vi.fn() } as unknown as SheetViewApi
    const w = mount(SheetToolbar, { props: { source, state: at('A1'), sheet }, attachTo: document.body })
    await settle()
    const onFormatCells = vi.fn()
    await w.setProps({ onFormatCells } as never)
    await w.find('[data-tb="format-cells"]').trigger('click')
    expect([onFormatCells.mock.calls.length, vi.mocked(sheet.formatCells).mock.calls.length]).toEqual([1, 0])
  })
})
