// A synthetic workbook with a PivotTable, a slicer over its Region and a timeline over its Date, as a file read gives them:
// Sum of Amount by Region at Report!A3:B6 over Data!A1:E4, the slicer at D3:F12, the timeline at H3:N9 by quarters.
import type { BookInput, SheetInput } from 'xlcalc'
import { pivotBook } from './pivot-fixture.ts'

/** Data!E: a date in each row (2024-01-15, 2024-06-10, 2025-03-10), shown as dates. */
const DATES = { E1: { value: 'Date' }, E2: { value: 45306, numberFormat: 14 }, E3: { value: 45453, numberFormat: 14 }, E4: { value: 45726, numberFormat: 14 } }
const REF = { sheet: 'Report', name: 'PivotTable1' }

export function slicerBook(opts: { protection?: SheetInput['protection']; locked?: boolean; freeze?: boolean } = {}): BookInput {
  const book = pivotBook(opts.protection ? { protection: opts.protection } : {})
  const data = book.sheets[0]!
  data.cells = { ...data.cells, ...DATES }
  const cache = book.pivotCaches![0]!
  cache.source = { type: 'range', sheet: 'Data', range: 'A1:E4' }
  cache.fields.push({ name: 'Date', facts: { types: ['date'], min: { date: '2024-01-15T00:00:00' }, max: { date: '2025-03-10T00:00:00' } } })
  const dates = ['2024-01-15T00:00:00', '2024-06-10T00:00:00', '2025-03-10T00:00:00']
  cache.records = cache.records!.map((r, i) => [...r, { date: dates[i]! }])
  const report = book.sheets[1]!
  report.pivotTables![0]!.fields.push({})
  report.slicers = [{ name: 'Region', cache: 'Slicer_Region', caption: 'Region', place: { from: 'D3', to: 'F12' }, ...(opts.locked === false ? { locked: false } : {}) }]
  report.timelines = [{ name: 'Date', cache: 'NativeTimeline_Date', caption: 'Date', place: { from: 'H3', to: 'N9' }, level: 'quarters' }]
  if (opts.freeze) report.frozen = { rows: 4 }
  book.slicerCaches = [{ name: 'Slicer_Region', field: 'Region', cache: 1, pivots: [REF], items: [{ item: 0 }, { item: 1 }] }]
  book.timelineCaches = [{ name: 'NativeTimeline_Date', field: 'Date', cache: 1, pivots: [REF], bounds: { start: '2024-01-01T00:00:00', end: '2026-01-01T00:00:00' } }]
  return book
}
