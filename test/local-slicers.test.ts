// @vitest-environment node
// localSource over a workbook with a slicer and a timeline: each sheet's slicers and timelines with their items, labels,
// choice and dates, the PivotTables' fields for Insert Slicer, and the choice filtering the PivotTable in one undo step.
import { afterEach, describe, expect, it } from 'vitest'
import { closeAll, source, win } from './local-fixture'
import { slicerBook } from './slicer-fixture'

afterEach(closeAll)

const sumAt = async (s: ReturnType<typeof source>, row: number) => (await win(s, 'Report', row, 1, 2, 1)).values[0]![0]

describe('localSource: slicers and timelines', () => {
  it("lists a sheet's slicers and timelines with their caches, and the PivotTable's fields with the dates marked", async () => {
    const [data, report] = await source(slicerBook()).sheets()
    expect(data!.slicers).toBeUndefined()
    expect(report!.slicers).toEqual([
      {
        name: 'Region',
        cache: 'Slicer_Region',
        caption: 'Region',
        place: { from: 'D3', to: 'F12' },
        field: 'Region',
        pivots: [{ sheet: 'Report', name: 'PivotTable1' }],
        pivotCache: 1,
        items: [
          { item: 0, label: 'North' },
          { item: 1, label: 'South' },
        ],
      },
    ])
    expect(report!.timelines).toEqual([
      {
        name: 'Date',
        cache: 'NativeTimeline_Date',
        caption: 'Date',
        place: { from: 'H3', to: 'N9' },
        level: 'quarters',
        field: 'Date',
        pivots: [{ sheet: 'Report', name: 'PivotTable1' }],
        pivotCache: 1,
        bounds: { start: '2024-01-01T00:00:00', end: '2026-01-01T00:00:00' },
      },
    ])
    expect(report!.pivots![0]!.fields).toEqual([{ name: 'Region' }, { name: 'Amount' }, { name: 'Kind' }, { name: 'Year' }, { name: 'Date', dates: true }])
  })

  it('a choice filters the PivotTable and comes back with the choice of each item; undo takes it back', async () => {
    const s = source(slicerBook())
    expect(await s.structure({ op: 'editSlicer', slicer: 'Region', cache: { items: [{ item: 0, selected: false }, { item: 1, selected: true }] } })).toMatchObject({ ok: true })
    const report = (await s.sheets())[1]!
    expect(report.slicers![0]!.items).toEqual([
      { item: 0, label: 'North', selected: false },
      { item: 1, label: 'South' },
    ])
    expect([await sumAt(s, 4), await sumAt(s, 5)]).toEqual([20, 20])
    await s.undo()
    expect((await s.sheets())[1]!.slicers![0]!.items.every((x) => x.selected === undefined)).toBe(true)
    expect(await sumAt(s, 6)).toBe(35)
  })

  it("a timeline's dates filter the PivotTable by whole days and come back with the level they span", async () => {
    const s = source(slicerBook())
    expect(await s.structure({ op: 'editTimeline', timeline: 'Date', cache: { selection: { start: '2024-01-01T00:00:00', end: '2024-12-31T00:00:00' } } })).toMatchObject({ ok: true })
    const t = (await s.sheets())[1]!.timelines![0]!
    expect(t.selection).toEqual({ start: '2024-01-01T00:00:00', end: '2024-12-31T00:00:00' })
    expect(t.selectionLevel).toBe('years')
    // 2024 holds North's 10 and South's 20, not North's 5 of 2025
    expect([await sumAt(s, 4), await sumAt(s, 5), await sumAt(s, 6)]).toEqual([10, 20, 30])
  })

  it('labels a slicer of dates in the short date, as the PivotTable labels its items', async () => {
    const s = source(slicerBook())
    expect(await s.structure({ op: 'addSlicer', sheet: 'Report', slicer: {}, cache: { field: 'Date', pivots: [{ sheet: 'Report', name: 'PivotTable1' }] } })).toMatchObject({ ok: true })
    const added = (await s.sheets())[1]!.slicers!.find((x) => x.field === 'Date')!
    expect(added.items.map((x) => x.label)).toEqual(['1/15/2024', '6/10/2024', '3/10/2025'])
    expect(added.place).toBeDefined()
  })
})
