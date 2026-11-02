// A host's buttons in ribbon groups (`ToolbarButton.group`) and large (`size: 'large'`), as Excel's Formulas tab draws its
// Function Library, Defined Names, Formula Auditing and Calculation groups; hosts that set neither keep their one group.
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { nextTick } from 'vue'
import { SheetToolbar, type SheetViewApi, type SheetViewState, type ToolbarButton, type ToolbarItem } from '../src'
import type { EditOutcome, SheetSource } from '../src/contract/source'
import { ribbonGroups } from '../src/toolbar'

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
  }) as SheetSource
const state: SheetViewState = { sheet: 'S', ref: 'A1', canAutoSum: true }
const sheet = {} as SheetViewApi
const run = () => {}
const button = (id: string, more: Partial<ToolbarButton> = {}): ToolbarButton => ({ id, label: id, run, ...more })

/** Excel's Formulas tab as a host draws it: Insert Function joins AutoSum (Function Library), then Defined Names, Formula Auditing and Calculation. */
const FORMULAS: ToolbarItem[] = [
  'undo',
  button('insertFunction', { group: 'autoSum', size: 'large' }),
  'autoSum',
  button('defineName', { group: 'names' }),
  button('nameManager', { group: 'names', size: 'large' }),
  button('createNames', { group: 'names' }),
  ...['tracePrecedents', 'traceDependents', 'removeArrows', 'showFormulas'].map((id) => button(id, { group: 'auditing' })),
  button('watchWindow', { group: 'auditing', size: 'large' }),
  { id: 'calculation', label: 'Calculation Options', group: 'calculation', size: 'large', menu: [{ id: 'automatic', label: 'Automatic', run }] },
]
/** The same buttons as a host draws them without `group` and `size`. */
const PLAIN: ToolbarItem[] = FORMULAS.map((x) => (typeof x === 'string' ? x : { ...x, group: undefined, size: undefined }))

const mountBar = (items: readonly ToolbarItem[], rows = 2) => mount(SheetToolbar, { props: { source: source(), state, sheet, items, rows }, attachTo: document.body })
/** Each group in the row as the hooks of its buttons, in the order drawn (the order Tab takes them). */
const groupsOf = (w: VueWrapper, within = '.xs-tb-row') =>
  w.findAll(`${within} > [data-tb-block]`).map((g) => g.findAll('[data-tb], [data-tb-host], [data-tb-host-menu]').map((e) => e.attributes('data-tb') ?? e.attributes('data-tb-host') ?? `${e.attributes('data-tb-host-menu')}-menu`))
const wide = () => vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1920)
/** The row `room` px wide, each group as wide as `widths` says (150 px otherwise), More 44 px. */
function layout(room: number, widths: Record<string, number>, viewport = 1920): void {
  vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(viewport)
  vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (this: HTMLElement) {
    return this.classList.contains('xs-tb-row') ? room : 0
  })
  vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (this: HTMLElement) {
    const key = this.dataset.tbBlock
    const width = key ? (widths[key] ?? 150) : this.dataset.tb === 'more' ? 44 : 0
    return { width, height: 64, top: 0, left: 0, right: width, bottom: 64, x: 0, y: 0, toJSON: () => ({}) } as DOMRect
  })
}

describe("a host's ribbon groups", () => {
  it("draws Excel's Formulas tab in rows: 2 as its four groups after Undo, each stacked two to a column in the order of items", async () => {
    wide()
    const w = mountBar(FORMULAS)
    await nextTick()
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((g) => [g.attributes('data-tb-block'), g.classes().includes('xs-tb-g-cols')])).toEqual([
      ['undo', true],
      ['host-1', true],
      ['host-3', true],
      ['host-4', true],
      ['host-5', true],
    ])
    expect(groupsOf(w)).toEqual([
      ['undo', 'redo'],
      ['insertFunction', 'auto-sum', 'auto-sum-menu'],
      ['defineName', 'nameManager', 'createNames'],
      ['tracePrecedents', 'traceDependents', 'removeArrows', 'showFormulas', 'watchWindow'],
      ['calculation-menu'],
    ])
  })

  it('without group and size keeps the layout of before: adjacent host buttons one group, none large', async () => {
    wide()
    const w = mountBar(PLAIN)
    await nextTick()
    expect(groupsOf(w)).toEqual([
      ['undo', 'redo'],
      ['insertFunction'],
      ['auto-sum', 'auto-sum-menu'],
      ['defineName', 'nameManager', 'createNames', 'tracePrecedents', 'traceDependents', 'removeArrows', 'showFormulas', 'watchWindow', 'calculation-menu'],
    ])
    expect(w.findAll('.xs-tb-large')).toHaveLength(0)
  })

  it("draws a large host button as tall as both rows in rows: 2 (Excel's PivotTable), and as the others on one line", async () => {
    wide()
    const w = mountBar(FORMULAS)
    await nextTick()
    const large = w.findAll('.xs-tb-row .xs-tb-large').map((e) => e.attributes('data-tb-host') ?? `${e.find('[data-tb-host-menu]').attributes('data-tb-host-menu')}-split`)
    expect(large).toEqual(['insertFunction', 'nameManager', 'watchWindow', 'calculation-split'])
    const one = mountBar(FORMULAS, 1)
    await nextTick()
    expect(one.findAll('.xs-tb-large')).toHaveLength(0)
  })

  it('in rows: 1 sends a host group to More whole, with the toolbar group it joins, the menu keeping the order of items', async () => {
    layout(700, { undo: 100, 'host-1': 600, 'host-3': 200, 'host-4': 250, 'host-5': 100 })
    const w = mountBar(FORMULAS, 1)
    await vi.waitFor(() => expect(w.find('[data-tb-menu]').exists()).toBe(true))
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((g) => g.attributes('data-tb-block'))).toEqual(['undo', 'host-3', 'host-4', 'host-5'])
    expect(groupsOf(w, '[data-tb-menu]')).toEqual([['insertFunction', 'auto-sum', 'auto-sum-menu']])
  })

  it('in rows: 2 sends a host group to More whole and keeps the later groups that fit', async () => {
    layout(700, { undo: 100, 'host-1': 150, 'host-3': 150, 'host-4': 400, 'host-5': 100 })
    const w = mountBar(FORMULAS)
    await vi.waitFor(() => expect(w.find('[data-tb-menu]').exists()).toBe(true))
    expect(w.findAll('.xs-tb-row > [data-tb-block]').map((g) => g.attributes('data-tb-block'))).toEqual(['undo', 'host-1', 'host-3', 'host-5'])
    expect(groupsOf(w, '[data-tb-menu]')).toEqual([['tracePrecedents', 'traceDependents', 'removeArrows', 'showFormulas', 'watchWindow']])
  })
})

describe('ribbonGroups with host blocks', () => {
  const host = (key: string, group?: string) => ({ key, buttons: [{ group }] })
  const shape = (blocks: Parameters<typeof ribbonGroups>[0], layout: 'blocks' | 'groups' | 'ribbon') => ribbonGroups(blocks, layout).map((g) => g.stacks.flatMap((s) => s.lines.flat().map((p) => p.block.key)))

  it("joins a host's group to an adjacent stacked group of the toolbar's named so, not to Font, Alignment or Number, and not apart", () => {
    const blocks = [host('a', 'autoSum'), { key: 'autoSum', group: 'autoSum' as const }, host('b', 'autoSum'), host('c', 'font'), { key: 'font', group: 'font' as const }, host('d', 'autoSum')]
    expect(shape(blocks, 'ribbon')).toEqual([['a', 'autoSum', 'b'], ['c'], ['font'], ['d']])
    expect(shape(blocks, 'groups')).toEqual(shape(blocks, 'ribbon'))
  })

  it('below 600 px draws each block alone, as before', () => {
    expect(shape([host('a', 'autoSum'), { key: 'autoSum', group: 'autoSum' as const }, host('b', 'x')], 'blocks')).toEqual([['a'], ['autoSum'], ['b']])
  })

  it("marks Excel's large toolbar groups large only where they are stacked", () => {
    const blocks = [host('a', 'refreshAll'), { key: 'refreshAll', group: 'refreshAll' as const }]
    expect(ribbonGroups(blocks, 'ribbon')[0]!.stacks[0]!.lines[0]!.map((p) => [p.block.key, !!p.large])).toEqual([
      ['a', false],
      ['refreshAll', true],
    ])
    expect(ribbonGroups(blocks, 'groups')[0]!.stacks[0]!.lines[0]!.some((p) => p.large)).toBe(false)
  })
})
