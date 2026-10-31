// Synthetic workbooks with a PivotTable as a file read gives them: the cache with its records, the PivotTable laid out by
// xlcalc's engine (addPivotTable over Data!A1:D4) and the values of its cells. Built without the engine, so a test file can
// tell whether the source loads it.
import type { BookInput, PivotCache, PivotTable, PivotTablePatch, SheetInput } from 'xlcalc'

/** Region, Amount, Kind and Year over three rows. */
const DATA: SheetInput = {
  name: 'Data',
  cells: {
    A1: { value: 'Region' }, B1: { value: 'Amount' }, C1: { value: 'Kind' }, D1: { value: 'Year' },
    A2: { value: 'North' }, B2: { value: 10 }, C2: { value: 'a' }, D2: { value: 2024 },
    A3: { value: 'South' }, B3: { value: 20 }, C3: { value: 'b' }, D3: { value: 2025 },
    A4: { value: 'North' }, B4: { value: 5 }, C4: { value: 'a' }, D4: { value: 2025 },
  },
}

const cache = (): PivotCache => ({
  id: 1,
  source: { type: 'range', sheet: 'Data', range: 'A1:D4' },
  fields: [
    { name: 'Region', facts: { types: ['string'] }, items: ['North', 'South'] },
    { name: 'Amount', facts: { types: ['number'], integer: true, min: 5, max: 20 } },
    { name: 'Kind', facts: { types: ['string'] }, items: ['a', 'b'] },
    { name: 'Year', facts: { types: ['number'], integer: true, min: 2024, max: 2025 }, items: [2024, 2025] },
  ],
  records: [[0, 10, 0, 0], [1, 20, 1, 1], [0, 5, 0, 1]],
})

const items = () => ({ items: [{ item: 0 }, { item: 1 }] })
const STYLE: PivotTable['style'] = { name: 'PivotStyleLight16', rowHeaders: true, columnHeaders: true, lastColumn: true }

/** Sum of Amount by Region at Report!A3:B6 (PivotTable1), with `extra` cells on Report, and `table` over its definition. */
export function pivotBook(opts: { locale?: string; protection?: SheetInput['protection']; extra?: SheetInput['cells']; styles?: BookInput['styles']; table?: Partial<PivotTable> } = {}): BookInput {
  const table: PivotTable = {
    name: 'PivotTable1',
    cache: 1,
    location: { range: 'A3:B6', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 },
    fields: [items(), {}, {}, {}],
    indent: 0,
    style: STYLE,
    rows: [0],
    values: [{ field: 1, name: 'Sum of Amount' }],
    rowLines: [{ items: [0] }, { items: [1] }, { items: [0], type: 'grand' }],
    columnLines: [{ items: [] }],
    ...opts.table,
  }
  const report: SheetInput = {
    name: 'Report',
    cells: {
      A3: { value: 'Row Labels' }, B3: { value: 'Sum of Amount' },
      A4: { value: 'North' }, B4: { value: 15 },
      A5: { value: 'South' }, B5: { value: 20 },
      A6: { value: 'Grand Total' }, B6: { value: 35 },
      ...opts.extra,
    },
    pivotTables: [table],
    ...(opts.protection ? { protection: opts.protection } : {}),
  }
  return { ...(opts.locale ? { locale: opts.locale } : {}), ...(opts.styles ? { styles: opts.styles } : {}), sheets: [DATA, report], pivotCaches: [cache()] }
}

/** Sum of Amount by Region down and Kind across at Report!A3:D7, Year in its Filters area at A1:B1 (PivotTable1). */
export function filteredPivotBook(): BookInput {
  const table: PivotTable = {
    name: 'PivotTable1',
    cache: 1,
    location: { range: 'A3:D7', firstHeaderRow: 1, firstDataRow: 2, firstDataColumn: 1, pageRows: 1, pageColumns: 1 },
    fields: [items(), {}, items(), items()],
    indent: 0,
    style: STYLE,
    rows: [0],
    columns: [2],
    pages: [{ field: 3 }],
    values: [{ field: 1, name: 'Sum of Amount' }],
    rowLines: [{ items: [0] }, { items: [1] }, { items: [0], type: 'grand' }],
    columnLines: [{ items: [0] }, { items: [1] }, { items: [0], type: 'grand' }],
  }
  const report: SheetInput = {
    name: 'Report',
    cells: {
      A1: { value: 'Year' }, B1: { value: '(All)' },
      A3: { value: 'Sum of Amount' }, B3: { value: 'Column Labels' },
      A4: { value: 'Row Labels' }, B4: { value: 'a' }, C4: { value: 'b' }, D4: { value: 'Grand Total' },
      A5: { value: 'North' }, B5: { value: 15 }, D5: { value: 15 },
      A6: { value: 'South' }, C6: { value: 20 }, D6: { value: 20 },
      A7: { value: 'Grand Total' }, B7: { value: 15 }, C7: { value: 20 }, D7: { value: 35 },
    },
    pivotTables: [table],
  }
  return { sheets: [DATA, report], pivotCaches: [cache()] }
}

/** Region, Kind, Year, Amount and Date over four rows on Data, and an empty sheet R for a PivotTable added at R!A3. */
export function salesBook(locale?: string): BookInput {
  const rows: [string, string, number, number, number][] = [['North', 'a', 2024, 10, 45383], ['South', 'b', 2025, 20, 45425], ['North', 'b', 2025, 5, 45717], ['East', 'a', 2024, 7, 45725]]
  const cells: SheetInput['cells'] = { A1: { value: 'Region' }, B1: { value: 'Kind' }, C1: { value: 'Year' }, D1: { value: 'Amount' }, E1: { value: 'Date' } }
  rows.forEach((r, i) => {
    const n = i + 2
    Object.assign(cells, { [`A${n}`]: { value: r[0] }, [`B${n}`]: { value: r[1] }, [`C${n}`]: { value: r[2] }, [`D${n}`]: { value: r[3] }, [`E${n}`]: { value: r[4], numberFormat: 'm/d/yyyy' } })
  })
  return { ...(locale ? { locale } : {}), sheets: [{ name: 'Data', cells }, { name: 'R', cells: {} }] }
}

/** The operation adding PivotTable1 at R!A3 over salesBook's data with `definition`. */
export const addSales = (definition: PivotTablePatch) => ({ op: 'addPivotTable' as const, sheet: 'R', cell: 'A3', source: { type: 'range' as const, sheet: 'Data', range: 'A1:E5' }, definition })
