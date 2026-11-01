// The PivotTable Fields pane's state over a source whose answers the test holds back: changes made while others are on their
// way build on them, a change made during Update waits for the next one, a draft back where it started leaves nothing behind,
// and a change from elsewhere drops a draft built on the old layout.
import { mount } from '@vue/test-utils'
import { afterEach, describe, expect, it } from 'vitest'
import { defineComponent, h, reactive, type PropType } from 'vue'
import type { PivotCache, PivotTable, PivotTablePatch } from 'xlcalc/contract'
import type { SheetPivotTable, SheetSource, SourceEvent } from '../src/contract/source'
import type { SheetViewApi, SheetViewState } from '../src/contract/view'
import { usePivotPane, type PivotPane } from '../src/components/pivot-pane'
import { check } from '../src/core/pivot-layout'
import { settle, until } from './helpers'

const CACHE: PivotCache = { id: 1, source: { type: 'range', sheet: 'Data', range: 'A1:D9' }, fields: ['Region', 'Kind', 'City', 'Team'].map((name) => ({ name, facts: { types: ['string'] } })) }
const FIELDS = CACHE.fields

/**
 * PivotTables the source holds (PivotTable1, and PivotTable2 for `two`), the sheet's `structure` calls held until `answer`,
 * and the source's events. `answered` is what the source does to a table besides the patch (xlcalc putting fields on the
 * axes itself); with `late` a read takes 10 ms and the source's `reload` comes right after the answer, as a remote
 * source's event can come after its HTTP answer.
 */
function fixture(opts: { answered?: (t: Table) => Partial<Table>; late?: boolean } = {}) {
  const empty = (name: string): Table => ({ name, cache: 1, location: { range: 'A3:C20', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, fields: FIELDS.map(() => ({})) })
  const tables: Record<string, Table> = { PivotTable1: empty('PivotTable1'), PivotTable2: empty('PivotTable2') }
  const patches: PivotTablePatch[] = []
  const pivots: string[] = []
  const held: (() => void)[] = []
  let listener: ((e: SourceEvent) => void) | null = null
  const source = {
    sheets: async () => [],
    window: async () => ({ row: 1, col: 1, text: [], values: [] }),
    structure: async () => ({ ok: true, changes: [] }),
    pivotTable: async (_sheet: string, name: string): Promise<SheetPivotTable> => {
      if (opts.late) await new Promise((r) => setTimeout(r, 10))
      return { table: structuredClone(tables[name]!), cache: CACHE }
    },
    subscribe: (l: (e: SourceEvent) => void) => ((listener = l), () => (listener = null)),
  } as unknown as SheetSource
  const sheet = {
    structure: (op: { pivot: string; set: PivotTablePatch }) =>
      new Promise((done) => {
        patches.push(op.set)
        pivots.push(op.pivot)
        held.push(() => {
          const t = { ...tables[op.pivot]!, ...Object.fromEntries(Object.entries(op.set).map(([k, v]) => [k, v ?? undefined])) }
          tables[op.pivot] = { ...t, ...opts.answered?.(t) }
          done({ ok: true, changes: [] })
          if (opts.late) setTimeout(() => listener?.({ kind: 'reload' }))
        })
      }),
    pivotDialog: () => true,
  } as unknown as SheetViewApi
  const state = reactive<SheetViewState>({ sheet: 'Report', ref: 'A3', pivot: 'PivotTable1', fieldList: true, canInsertPivot: true })
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
    pivots,
    state,
    /** Answers the oldest call held. */
    answer: async () => {
      held.shift()?.()
      await settle()
    },
    /** A change from elsewhere: the table set, then `reload`. */
    outside: async (set: Partial<PivotTable>) => {
      tables.PivotTable1 = { ...tables.PivotTable1!, ...set }
      listener?.({ kind: 'reload' })
      await settle()
    },
    rows: () => pane.areas.value?.rows,
  }
}
type Table = Omit<PivotTable, 'rowLines' | 'columnLines'>
let mounted: ReturnType<typeof fixture>[] = []
afterEach(() => {
  for (const f of mounted) f.w.unmount()
  mounted = []
})
async function start(opts: Parameters<typeof fixture>[0] = {}) {
  const f = fixture(opts)
  mounted.push(f)
  await settle()
  await until(() => !!f.pane().definition.value)
  return f
}
const tick = (f: Awaited<ReturnType<typeof start>>, field: number, on = true) => f.pane().apply((a) => check(a, field, on, CACHE))

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

  it("builds a change waiting in the queue over the answer before it, though the source's reload comes after the answer", async () => {
    // the source puts field 3 first on the rows with the answer, as xlcalc puts a date's periods before it
    const f = await start({ late: true, answered: (t) => (t.rows?.length && t.rows[0] !== 3 ? { rows: [3, ...t.rows] } : {}) })
    tick(f, 0)
    tick(f, 1)
    await settle()
    await f.answer()
    await until(() => f.patches.length === 2)
    expect(f.patches.map((p) => p.rows)).toEqual([[0], [3, 0, 1]])
  })

  it('sends a change waiting in the queue to the PivotTable it was made on, though the active cell moved to another', async () => {
    const f = await start()
    tick(f, 0)
    tick(f, 1)
    await settle()
    f.state.pivot = 'PivotTable2'
    await settle()
    await f.answer()
    await f.answer()
    expect(f.pivots).toEqual(['PivotTable1', 'PivotTable1'])
    expect(f.patches.map((p) => p.rows)).toEqual([[0], [0, 1]])
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
