// @vitest-environment node
// localSource and xlcalc's PivotTable engine: loaded at once for a workbook with PivotTables, and for one without them only
// when an operation on PivotTables needs it. A file of its own, since the engine, once registered, serves every session of
// the module's realm; the subpath is watched to tell when it is imported.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { closeAll, source, win } from './local-fixture'
import { pivotBook } from './pivot-fixture'

const imported = vi.hoisted(() => ({ count: 0 }))
vi.mock('xlcalc/pivot', async (original) => {
  imported.count++
  return original()
})

afterEach(closeAll)

describe('localSource: the PivotTable engine', () => {
  it('is not loaded for a workbook without PivotTables; the first PivotTable added loads it and is laid out, renamed and refreshed', async () => {
    const s = source({ sheets: [pivotBook().sheets[0]!, { name: 'Report', cells: {} }] })
    await s.ready
    expect(imported.count).toBe(0)
    const out = await s.structure({ op: 'addPivotTable', sheet: 'Report', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:D4' }, definition: { rows: [0], values: [{ field: 1 }] } })
    expect(out).toMatchObject({ ok: true })
    expect(imported.count).toBe(1)
    expect((await s.book()).pivotCaches?.[0]?.stale).toBeUndefined()
    expect((await win(s, 'Report', 4, 1, 1, 2)).values).toEqual([['North', 15]])
    expect((await s.sheets())[1]!.pivots).toMatchObject([{ name: 'PivotTable1', range: 'A3:B6', headers: 'A3:B3', rowLabels: 'A4:A6', values: 'B4:B6' }])
    // a label typed over renames it, and Refresh All refreshes it from the source
    expect(await s.edit([{ sheet: 'Report', cell: 'A4', text: 'Nord' }])).toMatchObject({ ok: true })
    expect(await s.edit([{ sheet: 'Data', cell: 'B2', text: '100' }])).toMatchObject({ ok: true })
    expect(await s.refreshConnections()).toMatchObject({ ok: true })
    expect((await win(s, 'Report', 4, 1, 1, 2)).values).toEqual([['Nord', 105]])
  })

  it('is loaded before the session opens for a workbook with PivotTables: Refresh reads the source again', async () => {
    const s = source(pivotBook())
    expect(await s.edit([{ sheet: 'Data', cell: 'B2', text: '100' }])).toMatchObject({ ok: true })
    // the PivotTable keeps its values until a refresh, as in Excel
    expect((await win(s, 'Report', 4, 1, 2, 1)).values).toEqual([[15]])
    expect(await s.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] })).toMatchObject({ ok: true })
    expect((await win(s, 'Report', 4, 1, 2, 1)).values).toEqual([[105]])
  })
})
