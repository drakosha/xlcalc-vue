// The toolbar's large buttons with rows: 2 as Excel's ribbon draws them: the icon at every button's size, the label under it on
// one or two lines broken as Excel breaks it, a menu's arrow after the label; on one line and in More no label. The row is
// measured with the labels, so a group that the labels widen goes into More.
import { mount, type VueWrapper } from '@vue/test-utils'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { nextTick } from 'vue'
import { SheetToolbar, ruMessages, type SheetViewApi, type SheetViewState, type ToolbarItem } from '../src'
import type { EditOutcome, SheetSource } from '../src/contract/source'
import { largeLabel } from '../src/toolbar'

afterEach(() => {
  document.body.innerHTML = ''
  vi.restoreAllMocks()
})

const OK: EditOutcome = { ok: true, changes: [] }
const source = () =>
  ({
    sheets: async () => [{ name: 'S', maxRow: 10, maxCol: 5 }],
    window: async () => ({ pending: { elapsedMs: 0 } }),
    format: vi.fn(async () => OK),
    undo: vi.fn(async () => OK),
    redo: vi.fn(async () => OK),
    structure: vi.fn(async () => OK),
    refreshConnections: vi.fn(async () => OK),
    pivotCells: vi.fn(async () => null),
  }) as unknown as SheetSource
/** The active cell in a PivotTable, so that PivotTable Analyze and Design show. */
const state: SheetViewState = { sheet: 'S', ref: 'A3', pivot: 'P', canEditPivot: true, canInsertPivot: true }
const sheet = {} as SheetViewApi
const run = () => {}
const ICON = '<svg viewBox="0 0 24 24"><path d="M4 4h16v16H4z"/></svg>'
/** Excel's Formulas tab's large buttons as a host's: a plain one, a menu (Calculation Options) and a split one. */
const HOST: ToolbarItem[] = [
  { id: 'insertFunction', label: 'Insert Function (Shift+F3)', icon: ICON, run, size: 'large', group: 'f' },
  { id: 'calculation', label: 'Calculation Options', icon: ICON, menu: [{ id: 'auto', label: 'Automatic', run }], size: 'large', group: 'f' },
  { id: 'errorChecking', label: 'Error Checking', icon: ICON, run, menu: [{ id: 'trace', label: 'Trace Error', run }], size: 'large', group: 'f' },
]
const ITEMS: ToolbarItem[] = ['undo', 'pivotTable', 'refreshAll', 'pivotAnalyze', 'pivotDesign', ...HOST]
const wide = () => vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1440)
const mountBar = (rows: number, items = ITEMS, messages?: typeof ruMessages) => mount(SheetToolbar, { props: { source: source(), state, sheet, items, rows, ...(messages ? { messages } : {}) }, attachTo: document.body })
/** A button's label as its lines, the arrow drawn as `˅`; null without a label. */
function labelOf(w: VueWrapper, selector: string): string[] | null {
  const label = w.find(`.xs-tb-row ${selector} > .xs-tb-label`)
  if (!label.exists()) return null
  return label.findAll('.xs-tb-label-line').map((l) => `${l.text()}${l.find('.xs-tb-label-down').exists() ? '˅' : ''}`)
}

describe('largeLabel', () => {
  it("breaks a name as Excel's ribbon does (measured at 1920 px): one word on one line, more on two with the longer line shortest", () => {
    expect(largeLabel('PivotTable', false)).toEqual(['PivotTable'])
    expect(largeLabel('Сводная таблица', false)).toEqual(['Сводная', 'таблица'])
    expect(largeLabel('Промежуточные итоги', true)).toEqual(['Промежуточные', 'итоги'])
    expect(largeLabel('Рекомендуемые сводные таблицы', false)).toEqual(['Рекомендуемые', 'сводные таблицы'])
    expect(largeLabel('Text to Columns', false)).toEqual(['Text to', 'Columns'])
    expect(largeLabel('Lookup & Reference', true)).toEqual(['Lookup &', 'Reference'])
    expect(largeLabel('Date & Time', true)).toEqual(['Date &', 'Time'])
    expect(largeLabel('Анализ "что если"', true)).toEqual(['Анализ "что', 'если"'])
  })

  it("puts a one-word menu's arrow on a second line of its own, as Excel's Subtotals", () => {
    expect(largeLabel('Subtotals', true)).toEqual(['Subtotals', ''])
  })

  it('leaves out the key that a ScreenTip shows after the name', () => {
    expect(largeLabel('Refresh All (Ctrl+Alt+F5)', true)).toEqual(['Refresh', 'All'])
    expect(largeLabel('Обновить все (Ctrl+Alt+F5)', true)).toEqual(['Обновить', 'все'])
    expect(largeLabel('Insert Function (Shift+F3)', false)).toEqual(['Insert', 'Function'])
    expect(largeLabel('Show Formulas (Ctrl+`)', false)).toEqual(['Show', 'Formulas'])
    expect(largeLabel('Repeat (F4)', false)).toEqual(['Repeat'])
    expect(largeLabel('Area (sq m)', false)).toEqual(['Area', '(sq m)'])
  })
})

describe("a large button's icon", () => {
  /** The stylesheet's rules as [selector, declarations]. */
  const rules = (file: string) => [...readFileSync(resolve(import.meta.dirname, '..', file), 'utf8').replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/([^{}]+)\{([^{}]*)\}/g)].map((m) => [m[1]!.trim(), m[2]!] as const)

  it("is sized only by the rule every toolbar icon takes (--xs-tb-icon), none of a large button's own but its label's arrow", () => {
    const own = rules('src/toolbar.css').filter(([sel, decl]) => sel.includes('xs-tb-large') && sel.includes('.xs-icon') && !sel.includes('xs-tb-label') && /(^|[\s;])(width|height)\s*:/.test(decl))
    expect(own).toEqual([])
    const all = rules('src/toolbar.css').find(([sel]) => sel === '.xs-toolbar:not(.xs-touch) .xs-tool .xs-icon')
    expect(all?.[1].replace(/\s+/g, ' ').trim()).toBe('width: var(--xs-tb-icon); height: var(--xs-tb-icon);')
  })

  it('is the same element as a small button\'s, the first child of its button, with the label after it', async () => {
    wide()
    const w = mountBar(2)
    await nextTick()
    const first = (s: string) => w.find(`.xs-tb-row ${s}`).element.firstElementChild!
    for (const s of ['[data-tb="pivot-table"]', '[data-tb="refresh-all"]', '[data-tb="pivot-subtotals"]', '[data-tb-host="insertFunction"]', '[data-tb-host-menu="calculation"]', '[data-tb-host="errorChecking"]']) {
      expect([s, first(s).getAttribute('class')]).toEqual([s, expect.stringMatching(/^xs-icon( xs-tb-svg)?$/)])
    }
    expect(first('[data-tb="pivot-table"]').nextElementSibling?.classList.contains('xs-tb-label')).toBe(true)
  })
})

describe("a large button's label", () => {
  it("in rows: 2 is under the icon, from the button's name without its key; a menu's arrow ends it, a split button's arrow carries it", async () => {
    wide()
    const w = mountBar(2)
    await nextTick()
    expect(labelOf(w, '[data-tb="pivot-table"]')).toEqual(['PivotTable'])
    expect(labelOf(w, '[data-tb="refresh-all"]')).toBeNull()
    expect(labelOf(w, '[data-tb="refresh-all-menu"]')).toEqual(['Refresh', 'All˅'])
    expect(['subtotals', 'grandTotals', 'layout', 'blankRows'].map((id) => labelOf(w, `[data-tb="pivot-${id}"]`))).toEqual([
      ['Subtotals', '˅'],
      ['Grand', 'Totals˅'],
      ['Report', 'Layout˅'],
      ['Blank', 'Rows˅'],
    ])
    expect(labelOf(w, '[data-tb-host="insertFunction"]')).toEqual(['Insert', 'Function'])
    expect(labelOf(w, '[data-tb-host-menu="calculation"]')).toEqual(['Calculation', 'Options˅'])
    expect(labelOf(w, '[data-tb-host="errorChecking"]')).toBeNull()
    expect(labelOf(w, '[data-tb-host-menu="errorChecking"]')).toEqual(['Error', 'Checking˅'])
    // The label is shown, not read: the button's name stays its aria-label
    expect(w.find('[data-tb="pivot-table"] .xs-tb-label').attributes('aria-hidden')).toBe('true')
    expect(w.find('[data-tb="pivot-table"]').attributes('aria-label')).toBe('PivotTable')
    // Small buttons have none
    expect(w.findAll('[data-tb-block="pivotAnalyze"] .xs-tb-label')).toHaveLength(0)
  })

  it('takes the texts of the interface language', async () => {
    wide()
    const w = mountBar(2, ITEMS, ruMessages)
    await nextTick()
    expect(labelOf(w, '[data-tb="pivot-table"]')).toEqual(['Сводная', 'таблица'])
    expect(labelOf(w, '[data-tb="refresh-all-menu"]')).toEqual(['Обновить', 'все˅'])
    expect(labelOf(w, '[data-tb="pivot-subtotals"]')).toEqual(['Промежуточные', 'итоги˅'])
  })

  it('in rows: 1 is not drawn: the buttons keep their icon and arrow', async () => {
    wide()
    const w = mountBar(1)
    await nextTick()
    expect(w.findAll('.xs-tb-label')).toHaveLength(0)
    expect(w.findAll('.xs-tb-large')).toHaveLength(0)
    expect(w.find('[data-tb="refresh-all-menu"] .xs-icon').exists()).toBe(true)
    expect(w.find('[data-tb="pivot-subtotals"]').findAll('.xs-icon')).toHaveLength(2)
    expect(w.find('[data-tb-host-menu="calculation"]').findAll('.xs-icon')).toHaveLength(2)
  })
})

describe('the ribbon measured with the labels', () => {
  /** The row 500 px wide, a group 100 px plus 50 px for each label it draws, More 44 px. */
  function layout(): void {
    wide()
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
      return this.classList.contains('xs-tb-row') ? 500 : 0
    })
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
      const width = this.dataset.tbBlock ? 100 + 50 * this.querySelectorAll('.xs-tb-label').length : this.dataset.tb === 'more' ? 44 : 0
      return { width, height: 64, top: 0, left: 0, right: width, bottom: 64, x: 0, y: 0, toJSON: () => ({}) } as DOMRect
    })
  }

  it('moves into More the group that its labels make too wide, and draws it there without them', async () => {
    layout()
    // undo 100, pivotTable 150, pivotDesign 300: 550 > 500, so Design goes; without the labels all three would fit in 300
    const w = mountBar(2, ['undo', 'pivotTable', 'pivotDesign'])
    await vi.waitFor(() => expect(w.find('[data-tb-menu]').exists()).toBe(true))
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((g) => g.attributes('data-tb-block'))).toEqual(['undo', 'pivotTable'])
    expect(w.findAll('[data-tb-menu] > [data-tb-block]').map((g) => g.attributes('data-tb-block'))).toEqual(['pivotDesign'])
    expect(w.findAll('[data-tb-menu] .xs-tb-label')).toHaveLength(0)
    expect(labelOf(w, '[data-tb="pivot-table"]')).toEqual(['PivotTable'])
  })

  it('keeps them all when the labels fit', async () => {
    layout()
    const w = mountBar(2, ['undo', 'pivotTable', 'refreshAll'])
    await nextTick()
    await nextTick()
    expect(w.find('[data-tb-menu]').exists()).toBe(false)
    expect(w.find('[data-tb="more"]').exists()).toBe(false)
  })
})
