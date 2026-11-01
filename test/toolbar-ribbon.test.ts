// The toolbar against Excel for Windows' ribbon: its default order (Home's groups, then the commands of other tabs in their tabs' order).
import { mount, type VueWrapper } from '@vue/test-utils'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { SheetToolbar, TOOLBAR_GROUPS, type SheetViewApi, type SheetViewState } from '../src'
import type { EditOutcome, FormatPatch, SheetSource } from '../src/contract/source'

afterEach(() => {
  document.body.innerHTML = ''
})

const OK: EditOutcome = { ok: true, changes: [] }
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
      'bold', 'italic', 'underline', 'borders', 'fill-colour', 'fill-palette', 'font-colour', 'font-palette',
      'align-top', 'align-middle', 'align-bottom', 'align-left', 'align-center', 'align-right', 'indent-down', 'indent-up', 'wrap', 'merge', 'merge-menu',
      'number', 'accounting', 'percent', 'comma', 'decimal-up', 'decimal-down', 'format-cells',
      'auto-sum', 'auto-sum-menu', 'pivot-table', 'refresh-all', 'refresh-all-menu',
    ])
  })
})
