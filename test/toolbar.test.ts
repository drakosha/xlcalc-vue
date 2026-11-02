// The formatting toolbar: its buttons call the source's `format`, `undo` and `redo` with Excel's patches over the
// selection, and show the active cell's formatting (from the sheet's `state` event).
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { SheetToolbar, SheetView, ruMessages, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'
import { fakeSource, numberShown, pickNumber, settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const OK: EditOutcome = { ok: true, changes: [] }
const THIN = { style: 'thin' }

/** A source with the methods the toolbar calls, each a mock returning ok. */
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

function mountBar(src: SheetSource, state: SheetViewState | null, props: Record<string, unknown> = {}) {
  return mount(SheetToolbar, { props: { source: src, state, messages: ruMessages, ...props }, attachTo: document.body })
}
const at = (ref: string, more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'Лист1', ref, ...more })

describe('SheetToolbar: formatting through the source', () => {
  it('formats every range of a selection of several (state.areas), styles and borders alike, in one call each', async () => {
    const src = source()
    const w = mountBar(src, at('C3', { areas: ['A1', 'C3', 'E5:E6'] }))
    await w.find('[data-tb="bold"]').trigger('click')
    await w.find('[data-tb="borders"]').trigger('click')
    await w.find('[data-tb="border-bottom"]').trigger('click')
    await settle()
    expect(src.format.mock.calls).toEqual([
      ['Лист1', ['A1', 'C3', 'E5:E6'].map((range) => ({ range, style: { font: { bold: true } } }))],
      ['Лист1', [{ range: 'A1', style: { border: { bottom: THIN } } }, { range: 'C3', style: { border: { bottom: THIN } } }, { range: 'E6', style: { border: { bottom: THIN } } }]],
    ])
  })

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

  it('applies All Borders, Bottom Border and No Border over the selection', async () => {
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

  it('makes the selection bold, or removes bold when the active cell is bold', async () => {
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

  it('sets italic, underline, a horizontal alignment and wrapping as toggles based on the active cell (Underline over a double one puts a single one, as Excel)', async () => {
    const src = source()
    const w = mountBar(src, at('C3', { style: { font: { underline: 'double' }, alignment: { horizontal: 'center', wrap: true } } }))
    for (const b of ['italic', 'underline', 'align-center', 'align-right', 'wrap']) await w.find(`[data-tb="${b}"]`).trigger('click')
    await settle()
    expect(src.format.mock.calls.map((c) => c[1][0].style)).toEqual([
      { font: { italic: true } },
      { font: { underline: 'single' } },
      { alignment: { horizontal: null } },
      { alignment: { horizontal: 'right' } },
      { alignment: { wrap: null } },
    ])
  })

  it('colors the text and the fill from the palette; Automatic and No Fill remove them', async () => {
    const src = source()
    const w = mountBar(src, at('A1'))
    await w.find('[data-tb="font-palette"]').trigger('click')
    await w.find('[data-xs-colour="#C00000"]').trigger('click')
    await w.find('[data-tb="font-colour"]').trigger('click')
    await w.find('[data-tb="font-palette"]').trigger('click')
    await w.find('[data-tb="colour-auto"]').trigger('click')
    await w.find('[data-tb="fill-palette"]').trigger('click')
    await w.find('[data-xs-colour="#FFFF00"]').trigger('click')
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

  it("sets a number format from Excel's list, as xlcalc returns it: a built-in format by its number, a code as written", async () => {
    const src = source()
    const w = mountBar(src, at('A1:A3'))
    await settle()
    for (const item of ['percentage', 'shortDate', 'time']) await pickNumber(w, item)
    await settle()
    expect(src.format.mock.calls).toEqual([
      ['Лист1', [{ range: 'A1:A3', numberFormat: 10 }]],
      ['Лист1', [{ range: 'A1:A3', numberFormat: 14 }]],
      ['Лист1', [{ range: 'A1:A3', numberFormat: '[$-F400]h:mm:ss\\ AM/PM' }]],
    ])
    // the box names the category again, as Excel's does
    expect([numberShown(w), w.find('[role="listbox"]').exists()]).toEqual(['General', false])
  })
})

describe("SheetToolbar: Excel's Number Format box", () => {
  const items = ['general', 'number', 'currency', 'accounting', 'shortDate', 'longDate', 'time', 'percentage', 'fraction', 'scientific', 'text']
  const listed = (w: ReturnType<typeof mountBar>) => w.findAll('[role="option"]').map((o) => o.attributes('data-tb-number'))
  const selected = (w: ReturnType<typeof mountBar>) => w.find('[role="option"][aria-selected="true"]').attributes('data-tb-number')

  it('is drawn as the Font box: a field that names the category, joined to its own arrow, taking no typing', async () => {
    const w = mountBar(source(), at('A1', { numberFormat: '0%' }), { messages: undefined })
    await settle()
    const field = w.find<HTMLInputElement>('[data-tb="number"]')
    expect(field.element.tagName).toBe('INPUT')
    expect([field.element.value, field.attributes('readonly'), field.attributes('data-tip'), field.attributes('data-tb-category')]).toEqual(['Percentage', '', 'Number Format', 'percentage'])
    const box = field.element.closest('.xs-tb-box')!
    expect(box).not.toBeNull()
    expect(box.querySelector('.xs-tb-caret[data-tb="number-list"] svg')).not.toBeNull()
    expect(w.find('.xs-toolbar select').exists()).toBe(false)
  })

  it("opens its list from the arrow and from the field, the active cell's item marked, and a click on an item applies it", async () => {
    const src = source()
    const w = mountBar(src, at('B2:C3', { numberFormat: '0.00%' }))
    await settle()
    await w.find('[data-tb="number-list"]').trigger('click')
    expect([listed(w), selected(w)]).toEqual([items, 'percentage'])
    await w.find('[data-tb="number-list"]').trigger('click')
    expect(w.find('[role="listbox"]').exists()).toBe(false)
    await w.find('[data-tb="number"]').trigger('click')
    expect(listed(w)).toEqual(items)
    await w.find('[data-tb-number="scientific"]').trigger('click')
    await settle()
    expect(src.format.mock.calls).toEqual([['Лист1', [{ range: 'B2:C3', numberFormat: 11 }]]])
    expect([w.find('[role="listbox"]').exists(), numberShown(w)]).toEqual([false, 'Percentage'])
  })

  it('opens on Alt+Down, goes through the list with the arrows, picks on Enter, and closes on Esc back to the grid', async () => {
    const src = source()
    const w = mountBar(src, at('A1'), { messages: undefined })
    await settle()
    const field = w.find('[data-tb="number"]')
    await field.trigger('focus')
    await field.trigger('keydown', { key: 'ArrowDown', altKey: true })
    expect(selected(w)).toBe('general')
    await field.trigger('keydown', { key: 'ArrowDown' })
    await field.trigger('keydown', { key: 'ArrowDown' })
    await field.trigger('keydown', { key: 'ArrowUp' })
    // the field keeps naming the active cell's category while the arrows go through the list
    expect([selected(w), numberShown(w)]).toEqual(['number', 'General'])
    await field.trigger('keydown', { key: 'Enter' })
    await settle()
    expect(src.format.mock.calls).toEqual([['Лист1', [{ range: 'A1', numberFormat: 2 }]]])
    expect(w.find('[role="listbox"]').exists()).toBe(false)
    await field.trigger('focus')
    await field.trigger('keydown', { key: 'ArrowDown', altKey: true })
    await field.trigger('keydown', { key: 'ArrowDown' })
    await field.trigger('keydown', { key: 'Escape' })
    await settle()
    expect([src.format.mock.calls.length, w.find('[role="listbox"]').exists(), numberShown(w)]).toEqual([1, false, 'General'])
    // the pick and the Esc each return to the grid
    expect(w.emitted('done')).toHaveLength(2)
  })

  it("marks the item whose format the active cell has, a date's and a time's too; a Custom code marks none, and Enter on none leaves for the grid", async () => {
    const src = source()
    const w = mountBar(src, at('A1', { numberFormat: 'm/d/yyyy' }), { messages: undefined })
    await settle()
    const opened = async (numberFormat: string) => {
      await w.setProps({ state: at('A1', { numberFormat }) })
      await w.find('[data-tb="number-list"]').trigger('click')
      const item = w.find('[role="option"][aria-selected="true"]')
      const out = [numberShown(w), item.exists() ? item.attributes('data-tb-number') : null]
      await w.find('[data-tb="number-list"]').trigger('click')
      return out
    }
    expect(await opened('m/d/yyyy')).toEqual(['Date', 'shortDate'])
    expect(await opened('[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy')).toEqual(['Date', 'longDate'])
    expect(await opened('[$-F400]h:mm:ss\\ AM/PM')).toEqual(['Time', 'time'])
    expect(await opened('m/d/yy')).toEqual(['Date', null])
    expect(await opened('0.00;-0.00')).toEqual(['Custom', null])
    const field = w.find('[data-tb="number"]')
    ;(field.element as HTMLElement).focus()
    await field.trigger('keydown', { key: 'ArrowDown', altKey: true })
    expect(w.find('[role="listbox"]').exists()).toBe(true)
    await field.trigger('keydown', { key: 'Enter' })
    await settle()
    expect([w.find('[role="listbox"]').exists(), document.activeElement === field.element, src.format.mock.calls.length, w.emitted('done')?.length]).toEqual([false, false, 0, 1])
  })

  it('names the highlighted item to a screen reader: the field points at it while the list is open, the line is not an item', async () => {
    const w = mountBar(source(), at('A1', { numberFormat: '0.00%', style: { font: { size: 12 } } }), { sheet: { formatCells: vi.fn() } as unknown as SheetViewApi })
    await settle()
    for (const hook of ['number', 'font-size']) {
      const field = w.find(`[data-tb="${hook}"]`)
      expect(field.attributes('aria-activedescendant')).toBeUndefined()
      await field.trigger('focus')
      await field.trigger('keydown', { key: 'ArrowDown', altKey: true })
      const pointed = () => document.getElementById(field.attributes('aria-activedescendant') ?? '')
      expect(pointed()?.getAttribute('aria-selected')).toBe('true')
      const before = pointed()?.textContent?.trim()
      await field.trigger('keydown', { key: 'ArrowDown' })
      expect([pointed()?.getAttribute('aria-selected'), pointed()?.textContent?.trim() !== before]).toEqual(['true', true])
      if (hook === 'number') expect(w.findAll('[role="listbox"] > :not([role="option"])').map((x) => x.attributes('role'))).toEqual(['presentation'])
      await field.trigger('keydown', { key: 'Escape' })
      expect(field.attributes('aria-activedescendant')).toBeUndefined()
    }
  })

  it('keeps an open list within the window as it is resized or scrolled', async () => {
    const w = mountBar(source(), at('A1'))
    await settle()
    const height = window.innerHeight
    await w.find('[data-tb="number-list"]').trigger('click')
    await settle()
    const room = () => (w.find('[role="listbox"]').element as HTMLElement).style.getPropertyValue('--xs-tb-room')
    const first = room()
    try {
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: height + 200 })
      window.dispatchEvent(new Event('resize'))
      await settle()
      expect(parseFloat(room()) - parseFloat(first)).toBe(200)
    } finally {
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: height })
    }
  })

  it("ends the list, after a line, with More Number Formats..., which opens Format Cells on the Number tab at the active cell's format", async () => {
    const sheet = { formatCells: vi.fn(), format: vi.fn() } as unknown as SheetViewApi
    const w = mountBar(source(), at('A1', { numberFormat: '0.000' }), { sheet })
    await settle()
    await w.find('[data-tb="number-list"]').trigger('click')
    const list = w.find('[role="listbox"]')
    expect(list.findAll(':scope > *').slice(-2).map((x) => x.attributes('role') + ':' + (x.attributes('data-tb-number') ?? '') + ':' + x.text())).toEqual(['presentation::', 'option:more:Другие числовые форматы...'])
    await w.find('[data-tb-number="more"]').trigger('click')
    expect(sheet.formatCells).toHaveBeenCalledWith({ tab: 'number' })
    expect(w.find('[role="listbox"]').exists()).toBe(false)
    // the arrows reach it after the eleven items, and Enter opens the dialog
    const field = w.find('[data-tb="number"]')
    await field.trigger('focus')
    await field.trigger('keydown', { key: 'ArrowDown', altKey: true })
    for (let i = 0; i < 12; i++) await field.trigger('keydown', { key: 'ArrowDown' })
    expect(w.find('[role="option"][aria-selected="true"]').attributes('data-tb-number')).toBe('more')
    await field.trigger('keydown', { key: 'Enter' })
    expect(sheet.formatCells).toHaveBeenCalledTimes(2)
    expect(sheet.format).not.toHaveBeenCalled()
    // a host that listens to formatCells gets where; without a dialog to open the item is not there
    const onFormatCells = vi.fn()
    const host = mountBar(source(), at('A1'), { onFormatCells })
    await settle()
    await pickNumber(host, 'more')
    expect(onFormatCells).toHaveBeenCalledWith({ tab: 'number' })
    const none = mountBar(source(), at('A1'))
    await settle()
    await none.find('[data-tb="number-list"]').trigger('click')
    expect([none.findAll('[role="option"]').length, none.find('[data-tb-number="more"]').exists(), none.find('[role="separator"]').exists()]).toEqual([11, false, false])
  })
})

describe("SheetToolbar: the active cell's formatting", () => {
  it('shows as pressed the bold, italic, underline, alignment and wrapping the active cell has', async () => {
    const w = mountBar(source(), at('A1', { style: { font: { bold: true, underline: 'single' }, alignment: { horizontal: 'right', wrap: true } } }))
    await settle()
    const pressed = (b: string) => w.find(`[data-tb="${b}"]`).attributes('aria-pressed')
    expect(['bold', 'italic', 'underline', 'align-left', 'align-center', 'align-right', 'wrap'].map(pressed)).toEqual(['true', 'false', 'true', 'false', 'false', 'true', 'true'])
  })

  it("names the category of the active cell's format as Excel's box does, by xlcalc's rules", async () => {
    const w = mountBar(source(), at('A1', { numberFormat: '0.00%' }), { messages: undefined })
    await settle()
    const shown = () => [w.find('[data-tb="number"]').attributes('data-tb-category'), numberShown(w)]
    expect(shown()).toEqual(['percentage', 'Percentage'])
    await w.setProps({ state: at('A2', { numberFormat: 'm/d/yyyy' }) })
    expect(shown()).toEqual(['date', 'Date'])
    await w.setProps({ state: at('A3') })
    expect(shown()).toEqual(['general', 'General'])
    await w.setProps({ state: at('A4', { numberFormat: '0.000' }) })
    expect(shown()).toEqual(['number', 'Number'])
    await w.setProps({ state: at('A5', { numberFormat: '0.00;-0.00' }) })
    expect(shown()).toEqual(['custom', 'Custom'])
    await w.find('[data-tb="number-list"]').trigger('click')
    expect(w.findAll('[role="option"]').map((o) => o.attributes('data-tb-number'))).toEqual(['general', 'number', 'currency', 'accounting', 'shortDate', 'longDate', 'time', 'percentage', 'fraction', 'scientific', 'text'])
  })

  it("names the list's items and the categories in the workbook's locale", async () => {
    const w = mountBar(source(), at('A1', { numberFormat: '0%' }), { locale: 'ru-RU' })
    // the locale's module arrives as a separate chunk
    await vi.dynamicImportSettled()
    await settle()
    expect(numberShown(w)).toBe('Процентный')
    await w.find('[data-tb="number-list"]').trigger('click')
    expect(w.findAll('[role="option"]').slice(0, 3).map((o) => o.text())).toEqual(['Общий', 'Числовой', 'Денежный'])
    await w.setProps({ state: at('A2', { numberFormat: '@' }) })
    expect(numberShown(w)).toBe('Текстовый')
  })

  it("marks the active cell's colors in the palettes, and Automatic and No Fill where it has none", async () => {
    const w = mountBar(source(), at('A1', { style: { font: { color: '#ff0000' }, fill: { color: '#92D050' } } }))
    await w.find('[data-tb="font-palette"]').trigger('click')
    const marked = () => w.findAll('[aria-pressed="true"][data-xs-colour], [aria-pressed="true"][data-tb^="colour-"], [aria-pressed="true"][data-tb^="fill-"]').map((b) => b.attributes('data-xs-colour') ?? b.attributes('data-tb'))
    expect(marked()).toEqual(['#FF0000'])
    await w.find('[data-tb="fill-palette"]').trigger('click')
    expect(marked()).toEqual(['#92D050'])
    await w.setProps({ state: at('A2') })
    expect(marked()).toEqual(['fill-none'])
  })

  it('has nothing enabled without a selection', async () => {
    const w = mountBar(source(), null)
    await settle()
    expect(w.find('[data-tb="bold"]').attributes('disabled')).toBeDefined()
    expect(w.find('[data-tb="number"]').attributes('disabled')).toBeDefined()
  })
})

describe('SheetToolbar: when it is shown', () => {
  it("is absent without the source's format", async () => {
    const src = source()
    const { format: _, ...rest } = src
    const w = mountBar(rest as SheetSource, at('A1'))
    await settle()
    expect(w.find('.xs-toolbar').exists()).toBe(false)
    expect(w.find('[data-tb="bold"]').exists()).toBe(false)
  })

  it("has no Undo without the source's undo", async () => {
    const { undo: _u, ...rest } = source()
    const w = mountBar(rest as SheetSource, at('A1'))
    await settle()
    expect(w.find('.xs-toolbar').exists()).toBe(true)
    expect(w.find('[data-tb="undo"]').exists()).toBe(false)
    expect(w.find('[data-tb="redo"]').exists()).toBe(true)
  })
})

describe('SheetToolbar: undo and redo', () => {
  it("calls the source's undo and redo, and passes on where the selection goes", async () => {
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

  it("is disabled where the source's history has no steps, enabled where it has one or returns no history", async () => {
    const history = { undo: 0, redo: 1 }
    // the format's step is in the history once the call completes, as with the local source
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

  it("reports the source's rejection", async () => {
    const src = source({ format: vi.fn(async () => ({ ok: false, reason: 'refused', message: 'Лист защищён' })) })
    const w = mountBar(src, at('A1'))
    await w.find('[data-tb="italic"]').trigger('click')
    await settle()
    expect(w.emitted('error')).toEqual([['Лист защищён']])
  })
})

describe('SheetToolbar through the sheet', () => {
  it('formats, undoes and redoes through the given sheet: the sheet sets the selection and reloads', async () => {
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

  it("leaves a rejection to the sheet's message line", async () => {
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
