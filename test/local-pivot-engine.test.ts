// @vitest-environment node
// localSource and xlcalc's PivotTable engine: loaded for a workbook with PivotTables, not for one without. A file of its own,
// since the engine, once registered, serves every session of the module's realm.
import { afterEach, describe, expect, it } from 'vitest'
import { closeAll, source, win } from './local-fixture'
import { pivotBook } from './pivot-fixture'

afterEach(closeAll)

describe('localSource: the PivotTable engine', () => {
  it('is not loaded for a workbook without PivotTables: a PivotTable added there keeps its cells, its cache stale', async () => {
    const s = source({ sheets: [pivotBook().sheets[0]!, { name: 'Report', cells: {} }] })
    const out = await s.structure({ op: 'addPivotTable', sheet: 'Report', cell: 'A3', source: { type: 'range', sheet: 'Data', range: 'A1:D4' }, definition: { rows: [0], values: [{ field: 1 }] } })
    expect(out).toMatchObject({ ok: true })
    expect((await s.book()).pivotCaches?.[0]?.stale).toBe(true)
    expect((await win(s, 'Report', 3, 1, 1, 2)).text).toEqual([['', '']])
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
