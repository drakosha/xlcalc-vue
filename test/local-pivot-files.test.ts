// @vitest-environment node
// A workbook with a PivotTable through openXlsx and saveXlsx: the PivotTable read back, what a file loses of it grouped as
// the other issues, and xlcalc's KeptRefusal `slicers` when a slicer added cannot go over the original file.
import { afterEach, describe, expect, it } from 'vitest'
import { groupIssues, openXlsx, saveXlsx } from '../src/local'
import { closeAll, source } from './local-fixture'
import { pivotBook } from './pivot-fixture'
import { unzipText } from './zip'

afterEach(closeAll)

describe('PivotTables in files', () => {
  it('writes the PivotTable and reads it back with its areas', async () => {
    const { data } = await saveXlsx(source(pivotBook()))
    const { book } = await openXlsx(data)
    expect((await source(book).sheets())[1]!.pivots).toMatchObject([{ name: 'PivotTable1', range: 'A3:B6', headers: 'A3:B3', rowLabels: 'A4:A6', values: 'B4:B6' }])
  })

  it("marks the filter buttons' cells (pivotButton) in the file it writes, so Excel draws them before a refresh", async () => {
    const files = unzipText((await saveXlsx(source(pivotBook()))).data)
    expect(files['xl/styles.xml']).toMatch(/pivotButton="1"/)
  })

  it("groups a PivotTable's issues as the others, under xlcalc's code pivotTable: a cache's source on a deleted sheet", async () => {
    const s = source(pivotBook())
    expect(await s.structure({ op: 'deleteSheet', sheet: 'Data' })).toMatchObject({ ok: true })
    const { issues } = await saveXlsx(s)
    expect(groupIssues(issues)).toEqual([{ code: 'pivotTable', count: 1, sheets: expect.any(Array), message: expect.stringMatching(/Data/) }])
  })

  it("rejects a slicer added over the original file with xlcalc's KeptRefusal code slicers; the workbook saves anew", async () => {
    const file = (await saveXlsx(source(pivotBook()))).data
    const { book, cached } = await openXlsx(file, { keep: true })
    const s = source(book, { session: { cached } })
    expect(await s.structure({ op: 'addSlicer', sheet: 'Report', slicer: {}, cache: { field: 'Region', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } })).toMatchObject({ ok: true })
    expect(await saveXlsx(s, { over: file }).catch((e: unknown) => e)).toMatchObject({ name: 'KeptRefusal', code: 'slicers' })
    expect((await saveXlsx(s)).data.byteLength).toBeGreaterThan(0)
  })
})
