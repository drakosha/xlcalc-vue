// @vitest-environment node
// localSource over a workbook with a PivotTable: its areas among the sheet's information, Excel's alerts for what its cells
// do not take, its labels renamed by typing, Refresh and Refresh All, and a protected sheet's Use PivotTable and PivotChart.
import { afterEach, describe, expect, it } from 'vitest'
import type { EditOutcome } from '../src/contract/source'
import { closeAll, source, win } from './local-fixture'
import { filteredPivotBook, pivotBook } from './pivot-fixture'

afterEach(closeAll)

// Excel 16.0 for Windows' answers to xlcalc's set `edit`
const PART = "We can't change this part of the PivotTable."
const AFFECT = "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again."
const FORMULA = 'Cannot enter a formula for an item or field name in a PivotTable report.'
const PROTECTED = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet. You might be requested to enter a password."
const refusedWith = (message: string) => ({ ok: false, reason: 'refused', code: 'pivot', message })
const ok = (out: EditOutcome) => expect(out).toMatchObject({ ok: true })
const valueAt = async (s: ReturnType<typeof source>, cell: string) => {
  const m = /^([A-Z])(\d+)$/.exec(cell)!
  return (await win(s, 'Report', Number(m[2]), 1, m[1]!.charCodeAt(0) - 64, 1)).values[0]![0]
}

describe('localSource: PivotTables', () => {
  it("lists each sheet's PivotTables with their areas", async () => {
    expect((await source(pivotBook()).sheets()).map((x) => x.pivots)).toEqual([undefined, [{ name: 'PivotTable1', range: 'A3:B6', headers: 'A3:B3', rowLabels: 'A4:A6', values: 'B4:B6' }]])
    expect((await source(filteredPivotBook()).sheets())[1]!.pivots).toEqual([{ name: 'PivotTable1', range: 'A3:D7', filters: 'A1:B1', headers: 'A3:D4', rowLabels: 'A5:A7', values: 'B5:D7' }])
  })

  it("refuses a value or a formula typed, pasted or filled over its cells with Excel's alert; a formula over a label has its own", async () => {
    const s = source(pivotBook())
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '100' }])).toEqual(refusedWith(PART))
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '=1+1' }])).toEqual(refusedWith(PART))
    expect(await s.edit([{ sheet: 'Report', cell: 'A4', text: '="East"' }])).toEqual(refusedWith(FORMULA))
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '1' }, { sheet: 'Report', cell: 'C4', text: '2' }])).toEqual(refusedWith(PART))
    expect(await s.paste('Report', 'B5', '7\t8')).toEqual(refusedWith(PART))
    expect(await s.fill('Report', 'B4:C5', '9', 'B4')).toEqual(refusedWith(PART))
    expect(await valueAt(s, 'B4')).toBe(15)
    // the cells beside it take anything
    ok(await s.edit([{ sheet: 'Report', cell: 'C4', text: '2' }]))
  })

  it("refuses clearing its cells and an operation that would cut it with Excel's alert for a change that affects it", async () => {
    const s = source(pivotBook())
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: null }])).toEqual(refusedWith(AFFECT))
    expect(await s.clear('Report', ['A4:B5'], 'all')).toEqual(refusedWith(AFFECT))
    expect(await s.structure({ op: 'insertRows', sheet: 'Report', row: 5, count: 1 })).toEqual(refusedWith(AFFECT))
    expect(await s.structure({ op: 'deleteColumns', sheet: 'Report', column: 'B', count: 1 })).toEqual(refusedWith(AFFECT))
    expect(await valueAt(s, 'B6')).toBe(35)
    // an operation on the PivotTable keeps the library's reason
    expect(await s.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'Nope' }] })).toMatchObject({ ok: false, code: 'pivot', message: expect.stringMatching(/Nope/) })
  })

  it("words the alerts in the workbook's language", async () => {
    const s = source(pivotBook({ locale: 'ru-RU' }))
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: '100' }])).toMatchObject({ code: 'pivot', message: 'Нельзя изменить эту часть сводной таблицы.' })
    expect(await s.edit([{ sheet: 'Report', cell: 'B4', text: null }])).toMatchObject({ code: 'pivot', message: expect.stringMatching(/^Нельзя внести это изменение/) })
  })

  it('renames a label typed over, as Excel does, in one undo step', async () => {
    const s = source(pivotBook())
    ok(await s.edit([{ sheet: 'Report', cell: 'A4', text: 'Nord' }]))
    expect(await valueAt(s, 'A4')).toBe('Nord')
    ok(await s.edit([{ sheet: 'Report', cell: 'B3', text: 'Total amount' }]))
    expect(await valueAt(s, 'B3')).toBe('Total amount')
    ok(await s.undo())
    ok(await s.undo())
    expect([await valueAt(s, 'A4'), await valueAt(s, 'B3')]).toEqual(['North', 'Sum of Amount'])
  })

  it('refreshes the PivotTable from its source as one undo step, and Refresh All refreshes every PivotTable', async () => {
    const s = source(pivotBook())
    ok(await s.edit([{ sheet: 'Data', cell: 'B3', text: '200' }]))
    ok(await s.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] }))
    expect(await valueAt(s, 'B5')).toBe(200)
    ok(await s.undo())
    expect(await valueAt(s, 'B5')).toBe(20)
    // a workbook with PivotTables has Refresh All without calls outside it to ask
    expect(s.refreshConnections).toBeTypeOf('function')
    ok(await s.refreshConnections!())
    expect(await valueAt(s, 'B6')).toBe(215)
    // Refresh All of some functions' calls leaves the PivotTables
    ok(await s.edit([{ sheet: 'Data', cell: 'B4', text: '0' }]))
    ok(await s.refreshConnections!(['WEBSERVICE']))
    expect(await valueAt(s, 'B6')).toBe(215)
  })

  it('has no Refresh All for a workbook without PivotTables or calls outside it', () => {
    expect(source({ sheets: [{ name: 'S', cells: {} }] }).refreshConnections).toBeUndefined()
  })

  it('refreshes and changes a PivotTable on a protected sheet only with Use PivotTable and PivotChart, as Microsoft documents the box', async () => {
    const locked = source(pivotBook({ protection: {} }))
    const refused = { ok: false, reason: 'refused', code: 'protected', message: PROTECTED }
    expect(await locked.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] })).toEqual(refused)
    expect(await locked.refreshConnections!()).toEqual(refused)
    expect(await locked.structure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { grandTotals: { rows: false } } })).toEqual(refused)
    expect(await locked.structure({ op: 'addPivotTable', sheet: 'Report', cell: 'F3', source: { type: 'range', sheet: 'Data', range: 'A1:D4' } })).toEqual(refused)
    const allowed = source(pivotBook({ protection: { pivotTables: true } }))
    ok(await allowed.structure({ op: 'refreshPivots', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] }))
    ok(await allowed.refreshConnections!())
    ok(await allowed.structure({ op: 'editPivotTable', sheet: 'Report', pivot: 'PivotTable1', set: { grandTotals: { rows: false } } }))
    expect(await valueAt(allowed, 'A6')).toBeNull()
  })
})
