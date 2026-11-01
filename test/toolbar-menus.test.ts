// The toolbar's Underline, Accounting Number Format and Orientation menus as Excel for Windows' ribbon has them, measured in
// Excel 16.0 in en-US and ru-RU (each item applied to a cell and read back): Underline's part applies the underline picked last
// and an item picked again removes it; the currencies come with the workbook's locale's first and Excel's code each; an
// Orientation item turns the text as Excel stores it, shows checked on a cell turned so, and picked again levels the text.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defineComponent, h, ref } from 'vue'
import { FormatDialog, SheetToolbar, SheetView, enMessages, ruMessages, type SheetViewApi, type SheetViewState } from '../src'
import { localSource, type LocalSource } from '../src/local'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'
import { goTo, settle, until } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
const source = () =>
  ({
    sheets: async () => [{ name: 'S', maxRow: 10, maxCol: 5 }],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    format: vi.fn(async (_sheet: string, _patches: FormatPatch[]) => OK),
  }) as SheetSource & { format: ReturnType<typeof vi.fn> }
const at = (more: Partial<SheetViewState> = {}): SheetViewState => ({ sheet: 'S', ref: 'B2', ...more })
const bar = (props: Record<string, unknown>) => mount(SheetToolbar, { props: { source: source(), state: at(), ...props }, attachTo: document.body })
/** The patches' styles and formats sent to the source, one per call. */
const sent = (w: VueWrapper) => (w.props() as { source: ReturnType<typeof source> }).source.format.mock.calls.map((c) => (c[1] as FormatPatch[]).map(({ range: _, ...p }) => p))
const items = (w: VueWrapper) => w.findAll('[role="menu"] [data-tb]').map((e) => [e.attributes('data-tb'), e.text(), e.attributes('aria-checked') ?? null])
async function pick(w: VueWrapper, arrow: string, item: string): Promise<void> {
  await w.find(`[data-tb="${arrow}"]`).trigger('click')
  await w.find(`[data-tb="${item}"]`).trigger('click')
}

describe("Underline, Excel's split button", () => {
  it('applies Underline at first and the underline picked last after a pick, named after it, pressed while the active cell has it', async () => {
    const w = bar({ state: at({ style: { font: { underline: 'double' } } }) })
    const part = () => w.find('[data-tb="underline"]')
    expect([part().attributes('data-tip'), part().attributes('aria-pressed')]).toEqual(['Underline (Ctrl+U)', 'false'])
    // Underline over a double underline puts a single one in its place, as Excel's part does
    await part().trigger('click')
    await pick(w, 'underline-menu', 'underline-double')
    expect(w.find('[role="menu"]').exists()).toBe(false)
    expect([part().attributes('data-tip'), part().attributes('aria-pressed')]).toEqual(['Double Underline', 'true'])
    await w.setProps({ state: at() })
    expect(part().attributes('aria-pressed')).toBe('false')
    await part().trigger('click')
    await w.setProps({ state: at({ style: { font: { underline: 'single' } } }) })
    await part().trigger('click')
    await settle()
    expect(sent(w)).toEqual([
      [{ style: { font: { underline: 'single' } } }],
      [{ style: { font: { underline: null } } }],
      [{ style: { font: { underline: 'double' } } }],
      [{ style: { font: { underline: 'double' } } }],
    ])
  })

  it("checks the active cell's underline in the menu, and an item picked again removes it", async () => {
    const w = bar({ state: at({ style: { font: { underline: 'single' } } }) })
    await w.find('[data-tb="underline-menu"]').trigger('click')
    expect(items(w)).toEqual([
      ['underline-single', 'Underline', 'true'],
      ['underline-double', 'Double Underline', 'false'],
    ])
    await w.find('[data-tb="underline-single"]').trigger('click')
    await settle()
    expect(sent(w)).toEqual([[{ style: { font: { underline: null } } }]])
  })

  it('names its arrow and items as Excel for Windows does, in Russian', async () => {
    const w = bar({ messages: ruMessages })
    const arrow = w.find('[data-tb="underline-menu"]')
    expect([arrow.attributes('aria-label'), arrow.attributes('aria-haspopup'), arrow.attributes('aria-expanded')]).toEqual(['Подчеркнутый', 'menu', 'false'])
    await arrow.trigger('click')
    expect(w.find('[data-tb="underline-menu"]').attributes('aria-expanded')).toBe('true')
    expect(items(w).map((x) => x[1])).toEqual(['Подчеркнутый', 'Двойное подчеркивание'])
    await w.find('[data-tb="underline-double"]').trigger('click')
    expect(w.find('[data-tb="underline"]').attributes('data-tip')).toBe('Двойное подчеркивание')
  })
})

describe("Accounting Number Format, Excel's split button", () => {
  it("lists the workbook's locale's currency first, then Excel's set without it, as Excel's menu in each language", async () => {
    const sheet = { formatCells: vi.fn() } as unknown as SheetViewApi
    const en = bar({ sheet })
    await en.find('[data-tb="accounting-menu"]').trigger('click')
    expect(items(en).map((x) => [x[0], x[1]])).toEqual([
      ['accounting-usd', '$ English (United States)'],
      ['accounting-rub', '₽ Russian'],
      ['accounting-gbp', '£ English (United Kingdom)'],
      ['accounting-eur', '€ Euro (€ 123)'],
      ['accounting-cny', '¥ Chinese (Simplified, Mainland China)'],
      ['accounting-chf', 'CHF French (Switzerland)'],
      ['accounting-more', 'More Accounting Formats...'],
    ])
    const ru = bar({ sheet, locale: 'ru-RU', messages: ruMessages })
    const arrow = ru.find('[data-tb="accounting-menu"]')
    expect(arrow.attributes('aria-label')).toBe('Финансовый числовой формат')
    await arrow.trigger('click')
    expect(items(ru).map((x) => [x[0], x[1]])).toEqual([
      ['accounting-rub', '₽ русский'],
      ['accounting-usd', '$ английский (США)'],
      ['accounting-gbp', '£ английский (Соединенное Королевство)'],
      ['accounting-eur', '€ Евро (€ 123)'],
      ['accounting-cny', '¥ Китайский, упрощенное письмо (континентальный Китай)'],
      ['accounting-chf', 'CHF французский (Швейцария)'],
      ['accounting-more', 'Другие финансовые форматы...'],
    ])
  })

  it("applies Excel's code of each currency, which follows the workbook's locale for the dollar and the euro; the part applies the built-in 44", async () => {
    const codes = async (locale?: string) => {
      const w = bar({ locale })
      await w.find('[data-tb="accounting"]').trigger('click')
      for (const c of ['usd', 'rub', 'gbp', 'eur', 'cny', 'chf']) await pick(w, 'accounting-menu', `accounting-${c}`)
      await settle()
      return sent(w).map((p) => p[0]!.numberFormat)
    }
    expect(await codes()).toEqual([
      44,
      '_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "-"??_);_(@_)',
      '_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-',
      '_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
      '_([$€-2] * #,##0.00_);_([$€-2] * (#,##0.00);_([$€-2] * "-"??_);_(@_)',
      '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ',
      '_-* #,##0.00 [$CHF-100C]_-;-* #,##0.00 [$CHF-100C]_-;_-* "-"?? [$CHF-100C]_-;_-@_-',
    ])
    const ru = await codes('ru-RU')
    expect([ru[1], ru[4]]).toEqual(['_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* "-"??_ ;_-@_ ', '_-[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_-'])
  })

  it('opens Format Cells from More Accounting Formats... on the Number tab at Accounting: the sheet’s, or the host’s with where', async () => {
    const sheet = { formatCells: vi.fn() } as unknown as SheetViewApi
    const w = bar({ sheet })
    await pick(w, 'accounting-menu', 'accounting-more')
    expect(sheet.formatCells).toHaveBeenCalledWith({ tab: 'number', choice: 'accounting' })
    expect(w.find('[role="menu"]').exists()).toBe(false)
    const onFormatCells = vi.fn()
    const host = bar({ onFormatCells })
    await pick(host, 'accounting-menu', 'accounting-more')
    expect(onFormatCells).toHaveBeenCalledWith({ tab: 'number', choice: 'accounting' })
    // without a dialog to open, the item is not there
    const none = bar({})
    await none.find('[data-tb="accounting-menu"]').trigger('click')
    expect(none.find('[data-tb="accounting-more"]').exists()).toBe(false)
  })
})

describe('Orientation, in the Alignment group', () => {
  it("sits between the vertical and the horizontal alignments, a button that opens Excel's menu", async () => {
    const w = bar({ sheet: { formatCells: vi.fn() } as unknown as SheetViewApi })
    const hooks = w.findAll('[data-tb]').map((e) => e.attributes('data-tb'))
    expect(hooks.slice(hooks.indexOf('align-bottom'), hooks.indexOf('align-left') + 1)).toEqual(['align-bottom', 'orientation', 'align-left'])
    await w.find('[data-tb="orientation"]').trigger('click')
    expect(items(w)).toEqual([
      ['orientation-counterclockwise', 'Angle Counterclockwise', 'false'],
      ['orientation-clockwise', 'Angle Clockwise', 'false'],
      ['orientation-vertical', 'Vertical Text', 'false'],
      ['orientation-up', 'Rotate Text Up', 'false'],
      ['orientation-down', 'Rotate Text Down', 'false'],
      ['orientation-format', 'Format Cell Alignment', null],
    ])
    expect(w.find('[role="menu"] [role="separator"] + [data-tb="orientation-format"]').exists()).toBe(true)
  })

  it('names the button and its items as Excel for Windows does, in Russian', async () => {
    const w = bar({ messages: ruMessages, sheet: { formatCells: vi.fn() } as unknown as SheetViewApi })
    expect(w.find('[data-tb="orientation"]').attributes('aria-label')).toBe('Ориентация')
    await w.find('[data-tb="orientation"]').trigger('click')
    expect(items(w).map((x) => x[1])).toEqual(['Текст против часовой стрелки', 'Текст по часовой стрелке', 'Вертикальный текст', 'Повернуть текст вверх', 'Повернуть текст вниз', 'Формат выравнивания ячейки'])
  })

  it('turns the text as each item does in Excel, checks the one the active cell has, and the same again levels the text', async () => {
    const w = bar({})
    for (const kind of ['counterclockwise', 'clockwise', 'vertical', 'up', 'down']) await pick(w, 'orientation', `orientation-${kind}`)
    await w.setProps({ state: at({ style: { alignment: { rotation: -45 } } }) })
    await w.find('[data-tb="orientation"]').trigger('click')
    expect(items(w).map((x) => x[2])).toEqual(['false', 'true', 'false', 'false', 'false'])
    await w.find('[data-tb="orientation-clockwise"]').trigger('click')
    // another angle checks none
    await w.setProps({ state: at({ style: { alignment: { rotation: 30 } } }) })
    await w.find('[data-tb="orientation"]').trigger('click')
    expect(items(w).map((x) => x[2])).toEqual(['false', 'false', 'false', 'false', 'false'])
    await settle()
    expect(sent(w).map((p) => p[0]!.style)).toEqual([
      { alignment: { rotation: 45 } },
      { alignment: { rotation: -45 } },
      { alignment: { rotation: 'vertical' } },
      { alignment: { rotation: 90 } },
      { alignment: { rotation: -90 } },
      { alignment: { rotation: null } },
    ])
  })

  it('opens Format Cells on the Alignment tab from Format Cell Alignment', async () => {
    const sheet = { formatCells: vi.fn() } as unknown as SheetViewApi
    const w = bar({ sheet })
    await pick(w, 'orientation', 'orientation-format')
    expect(sheet.formatCells).toHaveBeenCalledWith({ tab: 'alignment' })
  })
})

describe('the menus and the keyboard', () => {
  it('Esc closes each of them', async () => {
    const w = bar({})
    for (const arrow of ['underline-menu', 'accounting-menu', 'orientation']) {
      await w.find(`[data-tb="${arrow}"]`).trigger('click')
      expect(w.find('[role="menu"]').exists()).toBe(true)
      await w.find('[role="menu"] [data-tb]').trigger('keydown', { key: 'Escape' })
      expect(w.find('[role="menu"]').exists()).toBe(false)
    }
  })

  it('Esc closes them while the focus is on the grid, as the toolbar’s buttons leave it there, and the grid does not see it', async () => {
    const w = bar({})
    const grid = document.createElement('div')
    document.body.append(grid)
    const seen = vi.fn()
    grid.addEventListener('keydown', seen)
    await w.find('[data-tb="orientation"]').trigger('click')
    grid.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    await settle()
    expect(w.find('[role="menu"]').exists()).toBe(false)
    expect(seen).not.toHaveBeenCalled()
    // with no menu open, Esc is the grid's
    grid.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    expect(seen).toHaveBeenCalledOnce()
  })

  it('are disabled where the sheet cannot be formatted', () => {
    const w = bar({ state: null })
    expect(['underline', 'underline-menu', 'accounting-menu', 'orientation'].map((x) => w.find(`[data-tb="${x}"]`).attributes('disabled'))).toEqual(['', '', '', ''])
  })
})

describe("Format Cells' place", () => {
  it('opens on the tab given, with the Number tab’s item picked, and OK applies it, as Excel’s More Accounting Formats...', async () => {
    const state = at({ numberFormat: '0%', style: {} })
    const w = mount(FormatDialog, { props: { open: false, state, at: { tab: 'number', choice: 'accounting' } }, attachTo: document.body })
    await w.setProps({ open: true })
    await settle()
    expect(w.find('[data-fmt-tab][aria-selected="true"]').attributes('data-fmt-tab')).toBe('number')
    expect((w.find('[data-fmt-number="accounting"]').element as HTMLInputElement).checked).toBe(true)
    await w.find('form').trigger('submit')
    expect(w.emitted('apply')).toEqual([[[{ numberFormat: 44 }]]])
  })

  it('opens on the Alignment tab, and without a place on the tab it was last on', async () => {
    const state = at({ style: {} })
    const w = mount(FormatDialog, { props: { open: false, state, at: { tab: 'alignment' } }, attachTo: document.body })
    await w.setProps({ open: true })
    await settle()
    expect(w.find('[data-fmt-tab][aria-selected="true"]').attributes('data-fmt-tab')).toBe('alignment')
    await w.setProps({ open: false, at: null })
    await w.setProps({ open: true })
    await settle()
    expect(w.find('[data-fmt-tab][aria-selected="true"]').attributes('data-fmt-tab')).toBe('alignment')
  })
})

describe('through the local source', () => {
  /** The toolbar over the sheet over a workbook of `locale` with 1234.5 in A1:A9, as a host puts them together. */
  async function mountLocal(locale: string, messages = enMessages): Promise<{ w: VueWrapper; s: LocalSource }> {
    const cells = Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`A${i + 1}`, { value: 1234.5 }]))
    const s = localSource({ locale, sheets: [{ name: 'S', cells }] }, { calc: 'inline' })
    const Host = defineComponent(() => {
      const sheet = ref<SheetViewApi | null>(null)
      const state = ref<SheetViewState | null>(null)
      return () => [
        h(SheetToolbar, { source: s, sheet: sheet.value, state: state.value, locale, messages }),
        h(SheetView, { ref: sheet, source: s, locale, messages, style: 'height: 400px', onState: (x: SheetViewState | null) => (state.value = x) }),
      ]
    })
    const w = mount(Host, { attachTo: document.body })
    await settle(10)
    return { w, s }
  }
  /** A cell's style and number format as the workbook holds them. */
  async function cell(s: LocalSource, addr: string) {
    const book = await s.book()
    const c = book.sheets[0]!.cells![addr]
    const style = c?.style === undefined ? undefined : typeof c.style === 'number' ? book.styles?.[c.style] : c.style
    return { ...style, numberFormat: c?.numberFormat }
  }
  const close: LocalSource[] = []
  afterEach(async () => {
    await Promise.all(close.splice(0).map((s) => s.close()))
  })

  it('sets the double underline, the currencies and the rotations on the cells, each one step', async () => {
    const { w, s } = await mountLocal('en-US')
    close.push(s)
    const run = async (place: string, arrow: string, item: string) => {
      await goTo(w, place)
      await pick(w, arrow, item)
      await settle(12)
    }
    await run('A1', 'underline-menu', 'underline-double')
    await run('A2', 'accounting-menu', 'accounting-gbp')
    await run('A3', 'accounting-menu', 'accounting-usd')
    await run('A4', 'orientation', 'orientation-vertical')
    await run('A5', 'orientation', 'orientation-clockwise')
    await until(() => w.find('[data-tb="orientation"]').exists())
    expect((await cell(s, 'A1')).font?.underline).toBe('double')
    expect((await cell(s, 'A2')).numberFormat).toBe('_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-')
    expect((await cell(s, 'A3')).numberFormat).toBe('_([$$-409]* #,##0.00_);_([$$-409]* (#,##0.00);_([$$-409]* "-"??_);_(@_)')
    expect((await cell(s, 'A4')).alignment?.rotation).toBe('vertical')
    expect((await cell(s, 'A5')).alignment?.rotation).toBe(-45)
    // the same item on A5 again levels its text
    await run('A5', 'orientation', 'orientation-clockwise')
    expect((await cell(s, 'A5')).alignment?.rotation).toBeUndefined()
  })

  it('opens the sheet’s Format Cells at Accounting from More Accounting Formats..., whose OK formats the cell in the workbook’s locale', async () => {
    const { w, s } = await mountLocal('ru-RU', ruMessages)
    close.push(s)
    await goTo(w, 'A6')
    await pick(w, 'accounting-menu', 'accounting-more')
    await until(() => !!w.find<HTMLInputElement>('[data-fmt-number="accounting"]').element?.checked)
    expect(w.find('[data-fmt-tab][aria-selected="true"]').attributes('data-fmt-tab')).toBe('number')
    await w.find('[data-fmt] form, form[data-fmt]').trigger('submit')
    await settle(12)
    expect((await cell(s, 'A6')).numberFormat).toBe(44)
    await goTo(w, 'A7')
    await pick(w, 'accounting-menu', 'accounting-rub')
    await settle(12)
    expect((await cell(s, 'A7')).numberFormat).toBe('_-* #,##0.00 [$₽-419]_-;-* #,##0.00 [$₽-419]_-;_-* "-"?? [$₽-419]_-;_-@_-')
  })
})
