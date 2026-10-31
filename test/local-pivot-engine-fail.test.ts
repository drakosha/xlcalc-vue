// @vitest-environment node
// localSource when xlcalc's PivotTable engine cannot be loaded (its chunk fails): the workbook opens with its PivotTables'
// values, and Refresh and the operations on PivotTables refuse with a message for the user, never the library's.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { closeAll, source, win } from './local-fixture'
import { pivotBook } from './pivot-fixture'

vi.mock('xlcalc/pivot', () => {
  throw new Error('Failed to fetch dynamically imported module')
})

afterEach(closeAll)

const ENGINE = "PivotTables can't be refreshed or changed now: this page could not load what they need. Reload the page and try again."

describe('localSource without the PivotTable engine', () => {
  it('opens the workbook with its values and GETPIVOTDATA', async () => {
    const s = source(pivotBook({ extra: { E1: { formula: 'GETPIVOTDATA("Amount",A3,"Region","South")' } } }))
    await expect(s.ready).resolves.toBeUndefined()
    expect((await win(s, 'Report', 4, 1, 2, 1)).values).toEqual([[15]])
    expect((await win(s, 'Report', 1, 1, 5, 1)).values).toEqual([[20]])
  })

  it('refuses Refresh, Refresh All and an operation on a PivotTable with a message for the user', async () => {
    const s = source(pivotBook())
    const engine = { ok: false, reason: 'refused', code: 'pivot', message: ENGINE }
    expect(await s.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] })).toEqual(engine)
    expect(await s.refreshConnections()).toEqual(engine)
    expect(await s.structure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { grandTotals: { rows: false } } })).toEqual(engine)
    expect((await win(s, 'Report', 6, 1, 2, 1)).values).toEqual([[35]])
  })
})
