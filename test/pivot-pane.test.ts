// The PivotTable Fields pane's state over a source whose answers the test holds back: changes made while others are on their
// way build on them, a change made during Update waits for the next one, a draft back where it started leaves nothing behind,
// and a change from elsewhere drops a draft built on the old layout.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { defineComponent, h, type PropType } from 'vue'
import type { PivotCacheField, PivotTable, PivotTablePatch } from 'xlcalc/contract'
import type { SheetPivotTable, SheetSource, SourceEvent } from '../src/contract/source'
import type { SheetViewApi, SheetViewState } from '../src/contract/view'
import { usePivotPane, type PivotPane } from '../src/components/pivot-pane'
import { check } from '../src/core/pivot-layout'
import { settle } from './helpers'

const FIELDS: PivotCacheField[] = ['Region', 'Kind', 'City', 'Team'].map((name) => ({ name, facts: { types: ['string'] } }))

/** A PivotTable the source holds, the sheet's `structure` calls held until `answer`, and the source's events. */
function fixture() {
  let table: Omit<PivotTable, 'rowLines' | 'columnLines'> = { name: 'PivotTable1', cache: 1, location: { range: 'A3:C20', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, fields: FIELDS.map(() => ({})) }
  const patches: PivotTablePatch[] = []
  const held: (() => void)[] = []
  let listener: ((e: SourceEvent) => void) | null = null
  const source = {
    sheets: async () => [],
    window: async () => ({ row: 1, col: 1, text: [], values: [] }),
    structure: async () => ({ ok: true, changes: [] }),
    pivotTable: async (): Promise<SheetPivotTable> => ({ table: structuredClone(table), cache: { id: 1, source: { type: 'range', sheet: 'Data', range: 'A1:D9' }, fields: FIELDS } }),
    subscribe: (l: (e: SourceEvent) => void) => ((listener = l), () => (listener = null)),
  } as unknown as SheetSource
  const sheet = {
    structure: (op: { set: PivotTablePatch }) =>
      new Promise((done) => {
        patches.push(op.set)
        held.push(() => {
          table = { ...table, ...Object.fromEntries(Object.entries(op.set).map(([k, v]) => [k, v ?? undefined])) }
          done({ ok: true, changes: [] })
        })
      }),
    pivotDialog: () => true,
  } as unknown as SheetViewApi
  const state: SheetViewState = { sheet: 'Report', ref: 'A3', pivot: 'PivotTable1', fieldList: true, canInsertPivot: true }
  let pane!: PivotPane
  const Harness = defineComponent({
    props: { source: { type: Object as PropType<SheetSource>, required: true } },
    setup() {
      pane = usePivotPane({ source, sheet, state })
      return () => h('div')
    },
  })
  const w = mount(Harness, { props: { source } })
  return {
    w,
    pane: () => pane,
    patches,
    /** Answers the oldest call held. */
    answer: async () => {
      held.shift()?.()
      await settle()
    },
    /** A change from elsewhere: the table set, then `reload`. */
    outside: async (set: Partial<PivotTable>) => {
      table = { ...table, ...set }
      listener?.({ kind: 'reload' })
      await settle()
    },
    rows: () => pane.areas.value?.rows,
  }
}
let mounted: ReturnType<typeof fixture>[] = []
afterEach(() => {
  for (const f of mounted) f.w.unmount()
  mounted = []
})
async function start() {
  const f = fixture()
  mounted.push(f)
  await settle()
  return f
}
const tick = (f: Awaited<ReturnType<typeof start>>, field: number, on = true) => f.pane().apply(check(f.pane().areas.value!, field, on, FIELDS))

describe('the PivotTable Fields pane while changes are on their way', () => {
  it('builds each change on the ones before it, though the source has not answered them', async () => {
    const f = await start()
    tick(f, 0)
    tick(f, 1)
    tick(f, 2)
    await settle()
    expect(f.rows()).toEqual([0, 1, 2])
    for (let i = 0; i < 3; i++) await f.answer()
    expect(f.patches.map((p) => p.rows)).toEqual([[0], [0, 1], [0, 1, 2]])
    expect(f.rows()).toEqual([0, 1, 2])
  })

  it('keeps a field checked during Update, with Defer Layout Update on, for the next Update', async () => {
    const f = await start()
    f.pane().setDefer(true)
    tick(f, 0)
    void f.pane().update()
    await settle()
    tick(f, 1)
    await f.answer()
    expect(f.rows()).toEqual([0, 1])
    expect(f.pane().waiting.value).toBe(true)
    void f.pane().update()
    await f.answer()
    expect(f.patches.map((p) => p.rows)).toEqual([[0], [0, 1]])
  })

  it('leaves nothing behind for a draft back where it started, and shows a change from elsewhere', async () => {
    const f = await start()
    f.pane().setDefer(true)
    tick(f, 0)
    tick(f, 0, false)
    f.pane().setDefer(false)
    await settle()
    expect(f.patches).toEqual([])
    await f.outside({ rows: [3] })
    expect(f.rows()).toEqual([3])
    tick(f, 1)
    await settle()
    expect(f.patches.at(-1)?.rows).toEqual([3, 1])
  })

  it('drops a draft when the PivotTable changes from elsewhere', async () => {
    const f = await start()
    f.pane().setDefer(true)
    tick(f, 0)
    await f.outside({ rows: [2] })
    expect(f.rows()).toEqual([2])
    expect(f.pane().waiting.value).toBe(false)
  })
})
