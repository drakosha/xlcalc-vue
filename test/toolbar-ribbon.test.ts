// The toolbar against Excel for Windows' ribbon: its default order (Home's groups, then the commands of other tabs in their
// tabs' order) and the Borders split button with its menu in Excel's order.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { SheetToolbar, TOOLBAR_GROUPS, ruMessages, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'
import { settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
})

const OK: EditOutcome = { ok: true, changes: [] }
const THIN = { style: 'thin' }
const source = () =>
  ({
    sheets: async () => [{ name: 'S', maxRow: 10, maxCol: 5 }],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    format: vi.fn(async (_sheet: string, _patches: FormatPatch[]) => OK),
    undo: vi.fn(async () => OK),
    redo: vi.fn(async () => OK),
    structure: vi.fn(async () => OK),
    refreshConnections: vi.fn(async () => OK),
  }) as SheetSource & { format: ReturnType<typeof vi.fn> }
const state: SheetViewState = { sheet: 'S', ref: 'B2:C3' }
const hooks = (w: VueWrapper) => w.findAll('[data-tb], [data-tb-host]').map((e) => e.attributes('data-tb') ?? `host:${e.attributes('data-tb-host')}`)

describe("the toolbar's default order", () => {
  it("is Excel's Home (Clipboard, Font, Alignment, Number, Editing) after Undo and Redo, then Insert's PivotTable, Data's Refresh All and the PivotTable tabs", () => {
    expect(TOOLBAR_GROUPS).toEqual(['undo', 'formatPainter', 'font', 'fontStyle', 'borders', 'colors', 'verticalAlign', 'align', 'merge', 'number', 'formatCells', 'autoSum', 'pivotTable', 'refreshAll', 'pivotAnalyze', 'pivotDesign'])
  })

  it("draws the buttons in the order of Excel's groups: Fill Color before Font Color, Format Cells right after the Number group", async () => {
    const sheet = { formatCells: vi.fn() } as unknown as SheetViewApi
    const w = mount(SheetToolbar, { props: { source: source(), state, sheet }, attachTo: document.body })
    await vi.waitFor(() => expect(w.find('[data-tb="number"]').exists()).toBe(true))
    expect(hooks(w).filter((x) => x !== 'more')).toEqual([
      'undo', 'redo', 'format-painter',
      'font-name', 'font-name-list', 'font-size', 'font-size-list', 'font-up', 'font-down',
      'bold', 'italic', 'underline', 'border-last', 'borders', 'fill-colour', 'fill-palette', 'font-colour', 'font-palette',
      'align-top', 'align-middle', 'align-bottom', 'align-left', 'align-center', 'align-right', 'indent-down', 'indent-up', 'wrap', 'merge', 'merge-menu',
      'number', 'accounting', 'percent', 'comma', 'decimal-up', 'decimal-down', 'format-cells',
      'auto-sum', 'auto-sum-menu', 'pivot-table', 'refresh-all', 'refresh-all-menu',
    ])
  })
})

describe("the Borders split button, as Excel's", () => {
  it('applies Bottom Border at first, the last border picked after a pick, and names its part after it', async () => {
    const src = source()
    const w = mount(SheetToolbar, { props: { source: src, state }, attachTo: document.body })
    const part = () => w.find('[data-tb="border-last"]')
    expect([part().attributes('data-tip'), part().attributes('aria-label')]).toEqual(['Bottom Border', 'Bottom Border'])
    await part().trigger('click')
    expect(w.find('[role="menu"]').exists()).toBe(false)
    await w.find('[data-tb="borders"]').trigger('click')
    await w.find('[data-tb="border-all"]').trigger('click')
    expect(w.find('[role="menu"]').exists()).toBe(false)
    expect(part().attributes('data-tip')).toBe('All Borders')
    await part().trigger('click')
    await w.find('[data-tb="borders"]').trigger('click')
    await w.find('[data-tb="border-none"]').trigger('click')
    expect(part().attributes('data-tip')).toBe('No Border')
    await settle()
    expect(src.format.mock.calls.map((c) => c[1])).toEqual([
      [{ range: 'B3:C3', style: { border: { bottom: THIN } } }],
      [{ range: 'B2:C3', style: { border: { top: THIN, right: THIN, bottom: THIN, left: THIN } } }],
      [{ range: 'B2:C3', style: { border: { top: THIN, right: THIN, bottom: THIN, left: THIN } } }],
      [{ range: 'B2:C3', style: { border: null } }],
    ])
  })

  it("opens its menu from the arrow, named Borders, with Excel's items in Excel's order", async () => {
    const w = mount(SheetToolbar, { props: { source: source(), state, messages: ruMessages }, attachTo: document.body })
    const arrow = w.find('[data-tb="borders"]')
    expect([arrow.attributes('data-tip'), arrow.attributes('aria-expanded')]).toEqual(['Границы', 'false'])
    await arrow.trigger('click')
    expect(arrow.attributes('aria-expanded')).toBe('true')
    expect(w.findAll('[role="menu"] [role="menuitem"]').map((b) => [b.attributes('data-tb'), b.text()])).toEqual([
      ['border-bottom', 'Нижняя граница'],
      ['border-none', 'Нет границы'],
      ['border-all', 'Все границы'],
      ['border-outside', 'Внешние границы'],
    ])
  })

  it('is disabled, both parts, where the selection cannot be formatted', () => {
    const w = mount(SheetToolbar, { props: { source: source(), state: null }, attachTo: document.body })
    expect(['border-last', 'borders'].map((x) => w.find(`[data-tb="${x}"]`).attributes('disabled'))).toEqual(['', ''])
  })
})
