// The formatting toolbar: its buttons call the source's `format`, `undo` and `redo` with Excel's patches over the
// selection, and show what the active cell shows (the `state` event of the sheet).
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { SheetToolbar, SheetView, ruMessages, type SheetViewState, type ToolbarNumberFormat } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/source'
import { fakeSource, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const OK: EditOutcome = { ok: true, changes: [] }
const THIN = { style: 'thin' }

/** A source with the methods the toolbar calls, each a mock answering ok. */
function source(extra: Record<string, unknown> = {}) {
  return {
    sheets: async () => [{ name: 'Лист1', maxRow: 10, maxCol: 5 }],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    format: vi.fn(async (_sheet: string, _patches: FormatPatch[]) => OK),
    undo: vi.fn(async () => OK),
    redo: vi.fn(async () => OK),
    ...extra,
  } as SheetSource & { format: ReturnType<typeof vi.fn>; undo: ReturnType<typeof vi.fn>; redo: ReturnType<typeof vi.fn> }
}

/** The ru-RU number formats as the example builds them with the library: built-in formats by number, money as typed. */
const FORMATS: ToolbarNumberFormat[] = [
  { kind: 'general', format: 0, code: 'General' },
  { kind: 'number', format: 2, code: '0.00' },
  { kind: 'currency', format: '#,##0.00\\ "₽";[Red]\\-#,##0.00\\ "₽"', code: '#,##0.00\\ "₽";[Red]\\-#,##0.00\\ "₽"' },
  { kind: 'percent', format: 9, code: '0%' },
  { kind: 'date', format: 14, code: 'dd.mm.yyyy' },
  { kind: 'time', format: '[$-F400]h:mm:ss AM/PM', code: '[$-F400]h:mm:ss AM/PM' },
  { kind: 'text', format: 49, code: '@' },
]

function mountBar(src: SheetSource, state: SheetViewState | null, props: Record<string, unknown> = {}) {
  return mount(SheetToolbar, { props: { source: src, state, messages: ruMessages, numberFormats: FORMATS, ...props }, attachTo: document.body })
}
const at = (ref: string, more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'Лист1', ref, ...more })

describe('SheetToolbar: formatting through the source', () => {
  it('draws outside borders as Excel does: four edges of the selection in one call', async () => {
    const src = source()
    const w = mountBar(src, at('B2:D5'))
    await w.find('[data-tb="borders"]').trigger('click')
    await w.find('[data-tb="border-outside"]').trigger('click')
    await settle()
    expect(src.format).toHaveBeenCalledTimes(1)
    expect(src.format).toHaveBeenCalledWith('Лист1', [
      { range: 'B2:D2', style: { border: { top: THIN } } },
      { range: 'B5:D5', style: { border: { bottom: THIN } } },
      { range: 'B2:B5', style: { border: { left: THIN } } },
      { range: 'D2:D5', style: { border: { right: THIN } } },
    ])
  })

  it('gives all borders, the bottom border and none over the selection', async () => {
    const src = source()
    const w = mountBar(src, at('B2:C3'))
    for (const kind of ['all', 'bottom', 'none']) {
      await w.find('[data-tb="borders"]').trigger('click')
      await w.find(`[data-tb="border-${kind}"]`).trigger('click')
    }
    await settle()
    expect(src.format.mock.calls.map((c) => c[1])).toEqual([
      [{ range: 'B2:C3', style: { border: { top: THIN, right: THIN, bottom: THIN, left: THIN } } }],
      [{ range: 'B3:C3', style: { border: { bottom: THIN } } }],
      [{ range: 'B2:C3', style: { border: null } }],
    ])
  })

  it('makes the selection bold, or not bold when the active cell is', async () => {
    const src = source()
    const w = mountBar(src, at('A1:B2'))
    await w.find('[data-tb="bold"]').trigger('click')
    await w.setProps({ state: at('A1:B2', { style: { font: { bold: true } } }) })
    await w.find('[data-tb="bold"]').trigger('click')
    await settle()
    expect(src.format.mock.calls).toEqual([
      ['Лист1', [{ range: 'A1:B2', style: { font: { bold: true } } }]],
      ['Лист1', [{ range: 'A1:B2', style: { font: { bold: null } } }]],
    ])
  })

  it('sets italic, underline, a horizontal alignment and wrapping as toggles of the active cell', async () => {
    const src = source()
    const w = mountBar(src, at('C3', { style: { font: { underline: 'double' }, alignment: { horizontal: 'center', wrap: true } } }))
    for (const b of ['italic', 'underline', 'align-center', 'align-right', 'wrap']) await w.find(`[data-tb="${b}"]`).trigger('click')
    await settle()
    expect(src.format.mock.calls.map((c) => c[1][0].style)).toEqual([
      { font: { italic: true } },
      { font: { underline: null } },
      { alignment: { horizontal: null } },
      { alignment: { horizontal: 'right' } },
      { alignment: { wrap: null } },
    ])
  })

  it('colours the text and the fill from the palette, the automatic colour and no fill remove them', async () => {
    const src = source()
    const w = mountBar(src, at('A1'))
    await w.find('[data-tb="font-palette"]').trigger('click')
    await w.find('[data-tb-colour="#C00000"]').trigger('click')
    await w.find('[data-tb="font-colour"]').trigger('click')
    await w.find('[data-tb="font-palette"]').trigger('click')
    await w.find('[data-tb="colour-auto"]').trigger('click')
    await w.find('[data-tb="fill-palette"]').trigger('click')
    await w.find('[data-tb-colour="#FFFF00"]').trigger('click')
    await w.find('[data-tb="fill-palette"]').trigger('click')
    await w.find('[data-tb="fill-none"]').trigger('click')
    await settle()
    expect(src.format.mock.calls.map((c) => c[1][0].style)).toEqual([
      { font: { color: '#C00000' } },
      { font: { color: '#C00000' } },
      { font: { color: null } },
      { fill: { color: '#FFFF00' } },
      { fill: null },
    ])
  })

  it('sets a number format by what it offers: a built-in format by its number', async () => {
    const src = source()
    const w = mountBar(src, at('A1:A3'))
    await w.find('[data-tb="number"]').setValue('percent')
    await w.find('[data-tb="number"]').setValue('date')
    await settle()
    expect(src.format.mock.calls).toEqual([
      ['Лист1', [{ range: 'A1:A3', numberFormat: 9 }]],
      ['Лист1', [{ range: 'A1:A3', numberFormat: 14 }]],
    ])
  })
})

describe('SheetToolbar: what the active cell shows', () => {
  it('presses bold, italic, underline, the alignment and wrapping the active cell has', async () => {
    const w = mountBar(source(), at('A1', { style: { font: { bold: true, underline: 'single' }, alignment: { horizontal: 'right', wrap: true } } }))
    await settle()
    const pressed = (b: string) => w.find(`[data-tb="${b}"]`).attributes('aria-pressed')
    expect(['bold', 'italic', 'underline', 'align-left', 'align-center', 'align-right', 'wrap'].map(pressed)).toEqual(['true', 'false', 'true', 'false', 'false', 'true', 'true'])
  })

  it('selects the number format the active cell shows by: General without a code, the offered one by its code, other formats as custom', async () => {
    const w = mountBar(source(), at('A1', { numberFormat: '0%' }))
    await settle()
    const select = () => w.find<HTMLSelectElement>('[data-tb="number"]').element
    expect(select().value).toBe('percent')
    await w.setProps({ state: at('A2', { numberFormat: 'dd.mm.yyyy' }) })
    expect(select().value).toBe('date')
    await w.setProps({ state: at('A3') })
    expect(select().value).toBe('general')
    await w.setProps({ state: at('A4', { numberFormat: 'General' }) })
    expect(select().value).toBe('general')
    await w.setProps({ state: at('A5', { numberFormat: '0.000' }) })
    expect(select().value).toBe('')
    expect(select().selectedOptions[0]!.textContent).toBe(ruMessages.formatOther)
  })

  it('marks the colours the active cell has in the palettes, the automatic colour and no fill where it has none', async () => {
    const w = mountBar(source(), at('A1', { style: { font: { color: '#ff0000' }, fill: { color: '#92D050' } } }))
    await w.find('[data-tb="font-palette"]').trigger('click')
    const marked = () => w.findAll('[aria-pressed="true"][data-tb-colour], [aria-pressed="true"][data-tb^="colour-"], [aria-pressed="true"][data-tb^="fill-"]').map((b) => b.attributes('data-tb-colour') ?? b.attributes('data-tb'))
    expect(marked()).toEqual(['#FF0000'])
    await w.find('[data-tb="fill-palette"]').trigger('click')
    expect(marked()).toEqual(['#92D050'])
    await w.setProps({ state: at('A2') })
    expect(marked()).toEqual(['fill-none'])
  })

  it('has nothing to press without a selection', async () => {
    const w = mountBar(source(), null)
    await settle()
    expect(w.find('[data-tb="bold"]').attributes('disabled')).toBeDefined()
    expect(w.find('[data-tb="number"]').attributes('disabled')).toBeDefined()
  })
})

describe('SheetToolbar: when it shows', () => {
  it('is not there without the source’s format', async () => {
    const src = source()
    const { format: _, ...rest } = src
    const w = mountBar(rest as SheetSource, at('A1'))
    await settle()
    expect(w.find('.xs-toolbar').exists()).toBe(false)
    expect(w.find('[data-tb="bold"]').exists()).toBe(false)
  })

  it('has no number formats without the list of them, and no undo without the source’s undo', async () => {
    const { undo: _u, ...rest } = source()
    const w = mountBar(rest as SheetSource, at('A1'), { numberFormats: undefined })
    await settle()
    expect(w.find('.xs-toolbar').exists()).toBe(true)
    expect(w.find('[data-tb="number"]').exists()).toBe(false)
    expect(w.find('[data-tb="undo"]').exists()).toBe(false)
    expect(w.find('[data-tb="redo"]').exists()).toBe(true)
  })
})

describe('SheetToolbar: undo and redo', () => {
  it('calls the source’s undo and redo, and passes on where the selection goes', async () => {
    const src = source({ undo: vi.fn(async () => ({ ok: true, changes: [], select: { sheet: 'Лист1', range: 'B2:D5' } })) })
    const w = mountBar(src, at('A1'))
    await w.find('[data-tb="undo"]').trigger('click')
    await settle()
    await w.find('[data-tb="redo"]').trigger('click')
    await settle()
    expect(src.undo).toHaveBeenCalledTimes(1)
    expect(src.redo).toHaveBeenCalledTimes(1)
    expect(w.emitted('select')).toEqual([[{ sheet: 'Лист1', range: 'B2:D5' }]])
  })

  it('is off where the source’s history has no step, on where it has one or gives no history', async () => {
    const history = { undo: 0, redo: 1 }
    // the step the format makes is in the history once the call is done, as the local source's
    const src = source({ format: vi.fn(async () => ((history.undo = 1), OK)) })
    Object.defineProperty(src, 'history', { get: () => history })
    const w = mountBar(src, at('A1'))
    await settle()
    expect(w.find('[data-tb="undo"]').attributes('disabled')).toBeDefined()
    expect(w.find('[data-tb="redo"]').attributes('disabled')).toBeUndefined()
    await w.find('[data-tb="bold"]').trigger('click')
    await settle()
    expect(w.find('[data-tb="undo"]').attributes('disabled')).toBeUndefined()
    const plain = mountBar(source(), at('A1'))
    await settle()
    expect(plain.find('[data-tb="undo"]').attributes('disabled')).toBeUndefined()
  })

  it('reports a refusal of the source', async () => {
    const src = source({ format: vi.fn(async () => ({ ok: false, reason: 'refused', message: 'Лист защищён' })) })
    const w = mountBar(src, at('A1'))
    await w.find('[data-tb="italic"]').trigger('click')
    await settle()
    expect(w.emitted('error')).toEqual([['Лист защищён']])
  })
})

describe('SheetToolbar through the sheet', () => {
  it('formats, undoes and redoes through the sheet it is given: the sheet sets the selection and reads itself again', async () => {
    const f = fakeSource([{ name: 'Лист1', maxRow: 10, maxCol: 5 }])
    delete (f.source as { subscribe?: unknown }).subscribe
    const src = Object.assign(f.source, source({ undo: vi.fn(async (): Promise<EditOutcome> => ({ ok: true, changes: [], select: { sheet: 'Лист1', range: 'C3' } })) }), { sheets: f.source.sheets, window: f.source.window })
    const states: SheetViewState[] = []
    const sheet = mount(SheetView, { props: { source: src, onState: (s: SheetViewState | null) => s && states.push(s) }, attachTo: document.body })
    await settle()
    await sheet.findAll('.xs-head')[1]!.trigger('click')
    const bar = mountBar(src, states.at(-1)!, { sheet: sheet.vm })
    const reads = f.sheetsCalls
    await bar.find('[data-tb="bold"]').trigger('click')
    await settle()
    expect(src.format).toHaveBeenCalledWith('Лист1', [{ range: 'B:B', style: { font: { bold: true } } }])
    expect(f.sheetsCalls).toBe(reads + 1)
    await bar.find('[data-tb="undo"]').trigger('click')
    await settle()
    expect(src.undo).toHaveBeenCalledTimes(1)
    expect(states.at(-1)!.ref).toBe('C3')
    expect(sheet.findAll('.xs-flash').length).toBe(0)
    expect(bar.emitted('select')).toBeUndefined()
  })

  it('leaves a refusal to the sheet’s message line', async () => {
    const src = source({ sheets: async () => [{ name: 'Лист1', maxRow: 3, maxCol: 3 }], window: fakeSource([{ name: 'Лист1', maxRow: 3, maxCol: 3 }]).source.window, edit: async () => OK })
    src.format = vi.fn(async (): Promise<EditOutcome> => ({ ok: false, reason: 'refused', message: 'Лист защищён' }))
    const sheet = mount(SheetView, { props: { source: src }, attachTo: document.body })
    await settle()
    const bar = mountBar(src, at('A1'), { sheet: sheet.vm })
    await bar.find('[data-tb="italic"]').trigger('click')
    await settle()
    expect(sheet.find('.xs-message').text()).toBe('Лист защищён')
    expect(bar.emitted('error')).toBeUndefined()
  })
})
