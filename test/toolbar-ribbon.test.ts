// The toolbar against Excel for Windows' ribbon: its default order (Home's groups, then the commands of other tabs in their
// tabs' order), the Borders split button with its menu in Excel's order, the labels as Excel names them, and with `rows: 2`
// the toolbar laid out as Excel's ribbon, each group two rows tall.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { nextTick } from 'vue'
import { SheetToolbar, TOOLBAR_GROUPS, enMessages, ruMessages, type SheetViewApi, type SheetViewState, type ToolbarGroup } from '../src'
import { ribbonGroups } from '../src/toolbar'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'
import { settle } from './helpers'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
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

describe("the toolbar's labels, as Excel for Windows names them", () => {
  it('in Russian', () => {
    const ru = ruMessages
    expect([ru.underline, ru.fontSize, ru.fontColor, ru.wrapText, ru.autoSumCount, ru.pivotUngroup]).toEqual([
      'Подчеркнутый (Ctrl+U)',
      'Размер',
      'Цвет шрифта',
      'Переносить текст',
      'Счетчик',
      'Разгруппировать...',
    ])
    expect([ru.alignTop, ru.alignMiddle, ru.alignBottom, ru.alignLeft, ru.alignCenter, ru.alignRight]).toEqual([
      'Выровнять по верхнему краю',
      'Выровнять по середине',
      'Выровнять по нижнему краю',
      'Выровнять по левому краю',
      'Выровнять по центру',
      'Выровнять по правому краю',
    ])
  })

  it('in English', () => {
    const en = enMessages
    expect([en.underline, en.fontSize, en.fontColor, en.fillColor, en.wrapText, en.autoSumCount, en.pivotUngroup, en.borders]).toEqual([
      'Underline (Ctrl+U)',
      'Font Size',
      'Font Color',
      'Fill Color',
      'Wrap Text',
      'Count Numbers',
      'Ungroup...',
      'Borders',
    ])
    expect([en.alignTop, en.alignMiddle, en.alignBottom, en.alignLeft, en.alignCenter, en.alignRight]).toEqual(['Top Align', 'Middle Align', 'Bottom Align', 'Align Left', 'Center', 'Align Right'])
  })
})

describe('ribbonGroups', () => {
  const blocks = (keys: readonly string[]) => keys.map((key) => ({ key, group: key as ToolbarGroup }))
  const shape = (keys: readonly string[]) =>
    ribbonGroups(blocks(keys), true).map((g) => [g.key, g.kind, !!g.large, g.stacks.map((s) => s.lines.map((l) => l.map((p) => `${p.block.key}${p.slot ? `:${p.slot}` : ''}`)))])

  it("joins Font, Alignment and Number into Excel's two-row groups and stacks the others", () => {
    expect(shape(TOOLBAR_GROUPS)).toEqual([
      ['undo', 'cols', false, [[['undo']]]],
      ['formatPainter', 'cols', false, [[['formatPainter']]]],
      ['font', 'rows', false, [[['font:a1'], ['fontStyle:a2', 'borders:a2', 'colors:a2']]]],
      ['verticalAlign', 'rows', false, [[['verticalAlign:a1'], ['align:a2']], [['align:b1'], ['merge:b2']]]],
      ['number', 'rows', false, [[['number:a1'], ['number:a2']], [[], ['formatCells:b2']]]],
      ['autoSum', 'cols', false, [[['autoSum']]]],
      ['pivotTable', 'cols', true, [[['pivotTable']]]],
      ['refreshAll', 'cols', true, [[['refreshAll']]]],
      ['pivotAnalyze', 'cols', false, [[['pivotAnalyze']]]],
      ['pivotDesign', 'cols', true, [[['pivotDesign']]]],
    ])
  })

  it('keeps apart blocks of one Excel group that do not follow each other, and stacks a run that would leave a row empty', () => {
    expect(shape(['fontStyle', 'colors', 'undo', 'font']).map((g) => [g[0], g[1]])).toEqual([
      ['fontStyle', 'cols'],
      ['undo', 'cols'],
      ['font', 'cols'],
    ])
    expect(shape(['merge', 'number'])[0]![1]).toBe('cols')
  })

  it('without ribbon draws each block alone on one line', () => {
    expect(ribbonGroups(blocks(['font', 'fontStyle']), false).map((g) => [g.key, g.kind])).toEqual([
      ['font', 'flat'],
      ['fontStyle', 'flat'],
    ])
  })
})

describe('rows: 2, as Excel’s ribbon', () => {
  /** Each group's lines as the hooks they hold, stack by stack. */
  const lines = (w: VueWrapper, key: string) =>
    w.findAll(`.xs-tb-row > [data-tb-block="${key}"] > .xs-tb-stack`).map((st) => st.findAll(':scope > .xs-tb-line').map((ln) => ln.findAll('[data-tb]').map((e) => e.attributes('data-tb'))))
  const wide = () => vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1440)

  it("draws Font, Alignment and Number in Excel's rows, and the others stacked two to a column, PivotTable and Refresh All large", async () => {
    wide()
    const sheet = { formatCells: vi.fn() } as unknown as SheetViewApi
    const w = mount(SheetToolbar, { props: { source: source(), state, sheet, rows: 2 }, attachTo: document.body })
    await vi.waitFor(() => expect(w.find('[data-tb="number"]').exists()).toBe(true))
    expect(w.find('.xs-tb-row').classes()).toEqual(expect.arrayContaining(['xs-tb-ribbon']))
    expect(w.find('.xs-tb-row').classes()).not.toContain('xs-tb-rows')
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((g) => g.attributes('data-tb-block'))).toEqual(['undo', 'formatPainter', 'font', 'verticalAlign', 'number', 'autoSum', 'pivotTable', 'refreshAll'])
    expect(lines(w, 'font')).toEqual([
      [
        ['font-name', 'font-name-list', 'font-size', 'font-size-list', 'font-up', 'font-down'],
        ['bold', 'italic', 'underline', 'border-last', 'borders', 'fill-colour', 'fill-palette', 'font-colour', 'font-palette'],
      ],
    ])
    expect(lines(w, 'verticalAlign')).toEqual([
      [
        ['align-top', 'align-middle', 'align-bottom'],
        ['align-left', 'align-center', 'align-right', 'indent-down', 'indent-up'],
      ],
      [['wrap'], ['merge', 'merge-menu']],
    ])
    expect(lines(w, 'number')).toEqual([
      [['number'], ['accounting', 'percent', 'comma', 'decimal-up', 'decimal-down']],
      [[], ['format-cells']],
    ])
    expect(w.find('[data-tb-block="undo"]').classes()).toContain('xs-tb-g-cols')
    expect(['pivotTable', 'refreshAll', 'autoSum'].map((k) => w.find(`[data-tb-block="${k}"]`).classes().includes('xs-tb-large'))).toEqual([true, true, false])
  })

  it("makes a host's run of buttons a group of its own, stacked two to a column", async () => {
    wide()
    const run = vi.fn()
    const w = mount(SheetToolbar, { props: { source: source(), state, rows: 2, items: ['undo', { id: 'a', label: 'A', run }, { id: 'b', label: 'B', run }, { id: 'c', label: 'C', run }] }, attachTo: document.body })
    await nextTick()
    const host = w.find('.xs-tb-row > [data-tb-block="host-1"]')
    expect(host.classes()).toContain('xs-tb-g-cols')
    expect(host.findAll('[data-tb-host]').map((b) => b.attributes('data-tb-host'))).toEqual(['a', 'b', 'c'])
  })

  it('moves whole groups into More where they do not fit, the menu drawing their blocks on one line', async () => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1440)
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-tb-row') ? 700 : 0
    })
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      const width = this.dataset.tbBlock ? 300 : this.dataset.tb === 'more' ? 44 : 0
      return { width, height: 64, top: 0, left: 0, right: width, bottom: 64, x: 0, y: 0, toJSON: () => ({}) } as DOMRect
    })
    const w = mount(SheetToolbar, { props: { source: source(), state, rows: 2, items: ['undo', 'font', 'fontStyle', 'verticalAlign', 'align', 'number'] }, attachTo: document.body })
    await vi.waitFor(() => expect(w.find('[data-tb-menu]').exists()).toBe(true))
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['undo', 'font'])
    expect(w.findAll('.xs-tb-row > [data-tb-line]')).toHaveLength(0)
    expect(w.findAll('[data-tb-menu] > [data-tb-block]').map((b) => [b.attributes('data-tb-block'), b.classes().includes('xs-tb-g-flat')])).toEqual([
      ['verticalAlign', true],
      ['align', true],
      ['number', true],
    ])
    expect(w.find('[data-tb-menu] [data-tb="wrap"]').exists()).toBe(true)
  })

  it('moves a group too wide for what is left into More and keeps the later groups that fit, in order', async () => {
    const WIDTH: Record<string, number> = { undo: 150, font: 600 }
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1440)
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-tb-row') ? 700 : 0
    })
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      const key = this.dataset.tbBlock
      const width = key ? (WIDTH[key] ?? 150) : this.dataset.tb === 'more' ? 44 : 0
      return { width, height: 64, top: 0, left: 0, right: width, bottom: 64, x: 0, y: 0, toJSON: () => ({}) } as DOMRect
    })
    const w = mount(SheetToolbar, { props: { source: source(), state, rows: 2, items: ['undo', 'font', 'fontStyle', 'verticalAlign', 'align', 'number'] }, attachTo: document.body })
    await vi.waitFor(() => expect(w.find('[data-tb-menu]').exists()).toBe(true))
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['undo', 'verticalAlign', 'number'])
    expect(w.findAll('[data-tb-menu] > [data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['font', 'fontStyle'])
  })

  it('below 600 px keeps one row, as with rows: 1', async () => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(390)
    const w = mount(SheetToolbar, { props: { source: source(), state, rows: 2, items: ['font', 'fontStyle'] }, attachTo: document.body })
    await nextTick()
    expect(w.find('.xs-tb-row').classes()).not.toContain('xs-tb-ribbon')
    expect(w.findAll('[data-tb-block]').map((b) => b.attributes('data-tb-block'))).toEqual(['font', 'fontStyle'])
  })
})

describe('Format Cells without the sheet', () => {
  it("shows once the host listens to formatCells, at the toolbar's next render", async () => {
    const w = mount(SheetToolbar, { props: { source: source(), state }, attachTo: document.body })
    expect(w.find('[data-tb="format-cells"]').exists()).toBe(false)
    const onFormatCells = vi.fn()
    await w.setProps({ onFormatCells } as never)
    await w.setProps({ state: { ...state, ref: 'A1' } })
    await w.find('[data-tb="format-cells"]').trigger('click')
    expect(onFormatCells).toHaveBeenCalledOnce()
  })
})
