// The toolbar's Font and Font Size boxes, the size steps, vertical alignment, indent, Percent Style (with its key on the
// grid), Accounting, Comma Style and Increase and Decrease Decimal: each a patch over the selection through the source's `format`, by Excel's rules, from the active cell's formatting.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { SheetToolbar, type SheetViewState } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'
import { indentStep, stepFontSize } from '../src/toolbar'
import { fakeSource, goTo, key, mountSheet, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const OK: EditOutcome = { ok: true, changes: [] }

function source() {
  return {
    sheets: async () => [{ name: 'S', maxRow: 10, maxCol: 5 }],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    format: vi.fn(async (_sheet: string, _patches: FormatPatch[]) => OK),
  } as SheetSource & { format: ReturnType<typeof vi.fn> }
}
const at = (ref: string, more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'S', ref, ...more })
function mountBar(state: SheetViewState | null, props: Record<string, unknown> = {}) {
  const src = source()
  const w = mount(SheetToolbar, { props: { source: src, state, ...props }, attachTo: document.body })
  return { w, src, styles: () => src.format.mock.calls.map((c) => (c[1] as FormatPatch[])[0]) }
}
const field = (w: VueWrapper, hook: string) => w.find<HTMLInputElement>(`[data-tb="${hook}"]`)
async function typeInto(w: VueWrapper, hook: string, text: string, k = 'Enter'): Promise<void> {
  const f = field(w, hook)
  await f.trigger('focus')
  await f.setValue(text)
  await f.trigger('keydown', { key: k })
  await settle()
}

describe("Excel's size steps", () => {
  it('go through the list, beyond 72 by tens up to 409, below 8 by a point down to 1', () => {
    const up = (s: number) => stepFontSize(s, true)
    const down = (s: number) => stepFontSize(s, false)
    expect([8, 10.5, 11, 28, 48, 72, 75, 400, 409].map(up)).toEqual([9, 11, 12, 36, 72, 80, 80, 409, 409])
    expect([72, 75, 80, 409, 9, 8, 7.5, 2, 1].map(down)).toEqual([48, 72, 72, 400, 8, 7, 7, 1, 1])
  })
})

describe('Increase and Decrease Indent', () => {
  it('keep left, right and distributed, turn General and Center into Left, and do nothing at 0 or without an indent', () => {
    const style = (horizontal?: string, indent?: number) => ({ alignment: { horizontal, indent } }) as never
    expect(indentStep(style(), true)).toEqual({ alignment: { horizontal: 'left', indent: 1 } })
    expect(indentStep(style('center'), true)).toEqual({ alignment: { horizontal: 'left', indent: 1 } })
    expect(indentStep(style('right', 2), true)).toEqual({ alignment: { indent: 3 } })
    expect(indentStep(style('distributed', 1), false)).toEqual({ alignment: { indent: null } })
    expect(indentStep(style('left', 250), true)).toBeNull()
    expect(indentStep(style('left'), false)).toBeNull()
    expect(indentStep(style('center'), false)).toBeNull()
  })

  it('apply over the selection from the toolbar', async () => {
    const { w, styles } = mountBar(at('A1:B3', { style: { alignment: { horizontal: 'left', indent: 2 } } }))
    await w.find('[data-tb="indent-up"]').trigger('click')
    await w.find('[data-tb="indent-down"]').trigger('click')
    await w.setProps({ state: at('A1:B3') })
    await w.find('[data-tb="indent-down"]').trigger('click')
    await settle()
    expect(styles()).toEqual([
      { range: 'A1:B3', style: { alignment: { indent: 3 } } },
      { range: 'A1:B3', style: { alignment: { indent: 1 } } },
    ])
  })
})

describe('the Font and Font Size boxes', () => {
  it("show the active cell's own font, and nothing where it has none", async () => {
    const { w } = mountBar(at('A1', { style: { font: { name: 'Georgia', size: 10.5 } } }))
    expect([field(w, 'font-name').element.value, field(w, 'font-size').element.value]).toEqual(['Georgia', '10.5'])
    await w.setProps({ state: at('A2') })
    expect([field(w, 'font-name').element.value, field(w, 'font-size').element.value]).toEqual(['', ''])
  })

  it("show the workbook's default font for a cell without its own, and step its size from it", async () => {
    const { w, styles } = mountBar(at('A1', { font: { name: 'Calibri', size: 11 }, style: { font: { bold: true } } }))
    expect([field(w, 'font-name').element.value, field(w, 'font-size').element.value]).toEqual(['Calibri', '11'])
    await w.setProps({ state: at('A1', { font: { name: 'Arial', size: 10 } }) })
    await w.find('[data-tb="font-up"]').trigger('click')
    await settle()
    expect(styles().map((p) => p?.style)).toEqual([{ font: { size: 11 } }])
  })

  it('apply a typed font on Enter, an item picked from the list, and restore on Esc without a call', async () => {
    const { w, styles } = mountBar(at('B2:C3', { style: { font: { name: 'Arial' } } }))
    await typeInto(w, 'font-name', 'times new roman')
    await w.find('[data-tb="font-name-list"]').trigger('click')
    expect(w.find('[data-tb-item="Georgia"]').attributes('style')).toContain('Georgia')
    expect(w.find('[aria-selected="true"]').attributes('data-tb-item')).toBe('Arial')
    await w.find('[data-tb-item="Georgia"]').trigger('click')
    await typeInto(w, 'font-name', 'Comic Sans MS', 'Escape')
    await settle()
    expect(styles()).toEqual([
      { range: 'B2:C3', style: { font: { name: 'Times New Roman' } } },
      { range: 'B2:C3', style: { font: { name: 'Georgia' } } },
    ])
    expect(field(w, 'font-name').element.value).toBe('Arial')
    expect(w.find('[role="listbox"]').exists()).toBe(false)
    // Esc returns to the grid, as each call does
    expect(w.emitted('done')).toHaveLength(3)
  })

  it('complete a typed name with the first font it begins, from the list the host gives', async () => {
    const { w, styles } = mountBar(at('A1'), { fonts: ['PT Sans', 'PT Serif', 'Roboto'] })
    const f = field(w, 'font-name')
    await f.trigger('focus')
    f.element.value = 'pt s'
    f.element.dispatchEvent(new InputEvent('input', { inputType: 'insertText', data: 's' }))
    expect([f.element.value, f.element.selectionStart, f.element.selectionEnd]).toEqual(['PT Sans', 4, 7])
    await f.trigger('keydown', { key: 'Enter' })
    await settle()
    expect(styles()).toEqual([{ range: 'A1', style: { font: { name: 'PT Sans' } } }])
    await w.find('[data-tb="font-name-list"]').trigger('click')
    expect(w.findAll('[role="option"]').map((o) => o.text())).toEqual(['PT Sans', 'PT Serif', 'Roboto'])
  })

  it("take a size of 1 to 409 in half points, with a comma too, keep any other marked invalid, and pick from Excel's list", async () => {
    const { w, styles } = mountBar(at('A1'), { messages: { fmtSizeRange: 'Out of range' } })
    await typeInto(w, 'font-size', '10,3')
    await typeInto(w, 'font-size', '500')
    const f = field(w, 'font-size')
    expect([f.attributes('aria-invalid'), f.attributes('data-tip'), f.element.value]).toEqual(['true', 'Out of range', '500'])
    await f.trigger('keydown', { key: 'Escape' })
    expect([f.attributes('aria-invalid'), f.element.value]).toEqual(['false', ''])
    await w.find('[data-tb="font-size-list"]').trigger('click')
    expect(w.findAll('[role="option"]').map((o) => o.text()).slice(-3)).toEqual(['36', '48', '72'])
    await w.find('[data-tb-item="14"]').trigger('click')
    await settle()
    expect(styles()).toEqual([
      { range: 'A1', style: { font: { size: 10.5 } } },
      { range: 'A1', style: { font: { size: 14 } } },
    ])
  })

  it('go through the list with the arrows, putting each item into the box', async () => {
    const { w, styles } = mountBar(at('A1', { style: { font: { size: 12 } } }))
    const f = field(w, 'font-size')
    await f.trigger('focus')
    await f.trigger('keydown', { key: 'ArrowDown' })
    await f.trigger('keydown', { key: 'ArrowDown' })
    await f.trigger('keydown', { key: 'ArrowDown' })
    expect(f.element.value).toBe('16')
    await f.trigger('keydown', { key: 'Enter' })
    await settle()
    expect(styles()).toEqual([{ range: 'A1', style: { font: { size: 16 } } }])
  })
})

describe('Increase and Decrease Font Size', () => {
  it("step from the active cell's size, from 11 pt for a cell without one, and send nothing at a limit", async () => {
    const { w, styles } = mountBar(at('A1:A2'))
    await w.find('[data-tb="font-up"]').trigger('click')
    await w.setProps({ state: at('A1:A2', { style: { font: { size: 72 } } }) })
    await w.find('[data-tb="font-up"]').trigger('click')
    await w.setProps({ state: at('A1:A2', { style: { font: { size: 8 } } }) })
    await w.find('[data-tb="font-down"]').trigger('click')
    await w.setProps({ state: at('A1:A2', { style: { font: { size: 409 } } }) })
    await w.find('[data-tb="font-up"]').trigger('click')
    await settle()
    expect(styles().map((p) => p?.style)).toEqual([{ font: { size: 12 } }, { font: { size: 80 } }, { font: { size: 7 } }])
  })
})

describe('Top, Middle and Bottom Align', () => {
  it('show the active cell pressed, Bottom without an alignment of its own, and toggle back to Bottom', async () => {
    const { w, styles } = mountBar(at('C3'))
    const pressed = () => ['align-top', 'align-middle', 'align-bottom'].map((b) => w.find(`[data-tb="${b}"]`).attributes('aria-pressed'))
    expect(pressed()).toEqual(['false', 'false', 'true'])
    await w.find('[data-tb="align-middle"]').trigger('click')
    await w.setProps({ state: at('C3', { style: { alignment: { vertical: 'top' } } }) })
    expect(pressed()).toEqual(['true', 'false', 'false'])
    await w.find('[data-tb="align-top"]').trigger('click')
    await w.find('[data-tb="align-bottom"]').trigger('click')
    await settle()
    expect(styles().map((p) => p?.style)).toEqual([{ alignment: { vertical: 'center' } }, { alignment: { vertical: null } }, { alignment: { vertical: null } }])
  })
})

describe('Percent Style', () => {
  it("sets Excel's 0% (built-in format 9) over the selection from the button", async () => {
    const { w, src } = mountBar(at('B2:B4'))
    await w.find('[data-tb="percent"]').trigger('click')
    await settle()
    expect(src.format).toHaveBeenCalledWith('S', [{ range: 'B2:B4', numberFormat: 9 }])
  })

  it('and from Ctrl+Shift+% on the grid, Control on a Mac too, by the key of 5 on any layout', async () => {
    const f = fakeSource([{ name: 'S', maxRow: 10, maxCol: 5 }])
    const format = vi.fn(async () => OK)
    Object.assign(f.source, { format })
    const w = await mountSheet(f.source)
    await key(w, '%', { ctrlKey: true, shiftKey: true, code: 'Digit5' })
    await goTo(w, 'B2:C3')
    await key(w, '5', { ctrlKey: true, shiftKey: true, code: 'Digit5' })
    // Cmd+Shift+5 is the Mac's screenshot key, not Excel's
    await key(w, '%', { metaKey: true, shiftKey: true, code: 'Digit5' })
    await settle()
    expect(format.mock.calls).toEqual([
      ['S', [{ range: 'A1', numberFormat: 9 }]],
      ['S', [{ range: 'B2:C3', numberFormat: 9 }]],
    ])
  })
})

describe('Accounting, Comma Style, Increase and Decrease Decimal', () => {
  it("set Excel's built-in Accounting (44) and Comma Style (43) over the selection", async () => {
    const { w, styles } = mountBar(at('A1:B3'))
    await w.find('[data-tb="accounting"]').trigger('click')
    await w.find('[data-tb="comma"]').trigger('click')
    await settle()
    expect(styles()).toEqual([
      { range: 'A1:B3', numberFormat: 44 },
      { range: 'A1:B3', numberFormat: 43 },
    ])
  })

  it("step the active cell's decimals by xlcalc's rule and put the result on the selection; nothing when the step changes nothing", async () => {
    const { w, styles } = mountBar(at('A1:A4', { numberFormat: '0.00', cellValue: 1.5 }))
    await w.find('[data-tb="decimal-up"]').trigger('click')
    await w.find('[data-tb="decimal-down"]').trigger('click')
    await w.setProps({ state: at('A1:A4', { cellValue: 1.5 }) })
    await w.find('[data-tb="decimal-up"]').trigger('click')
    // General without a number, and a format without a decimal to take away, change nothing
    await w.setProps({ state: at('A1', { cellValue: 'text' }) })
    await w.find('[data-tb="decimal-up"]').trigger('click')
    await w.setProps({ state: at('A1', { numberFormat: '0', cellValue: 2 }) })
    await w.find('[data-tb="decimal-down"]').trigger('click')
    await settle()
    expect(styles()).toEqual([
      { range: 'A1:A4', numberFormat: '0.000' },
      { range: 'A1:A4', numberFormat: '0.0' },
      { range: 'A1:A4', numberFormat: '0.00' },
    ])
  })
})
