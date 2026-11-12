// @vitest-environment node
// A workbook with a PivotTable through openXlsx and saveXlsx: the PivotTable read back, what a file loses of it grouped as
// the other issues, and added and deleted slicers saved over the original file.
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

  it('saves a slicer added over the original file and removes its parts when deleted', async () => {
    const file = (await saveXlsx(source(pivotBook()))).data
    const { book, cached } = await openXlsx(file, { keep: true })
    const s = source(book, { session: { cached } })
    expect(await s.structure({ op: 'addSlicer', sheet: 'Report', slicer: {}, cache: { field: 'Region', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } })).toMatchObject({ ok: true })
    const saved = await saveXlsx(s, { over: file })
    expect(saved.issues).toEqual([])
    const back = await openXlsx(saved.data, { keep: true })
    const t = source(back.book, { session: { cached: back.cached } })
    expect((await t.sheets())[1]!.slicers).toMatchObject([{ name: 'Region', field: 'Region', items: [{ label: 'North' }, { label: 'South' }] }])
    expect(await t.structure({ op: 'deleteSlicer', slicer: 'Region' })).toMatchObject({ ok: true })
    const removed = await saveXlsx(t, { over: saved.data })
    expect(removed.issues).toEqual([])
    expect((await openXlsx(removed.data)).book.sheets[1]!.slicers ?? []).toEqual([])
    expect(Object.keys(unzipText(removed.data)).filter(p => /^xl\/(slicers|slicerCaches)\//.test(p))).toEqual([])
  })
})
