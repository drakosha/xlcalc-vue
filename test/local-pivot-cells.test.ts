// @vitest-environment node
// localSource's pivotCells: the labels of a PivotTable's items, its filter buttons, and what a cell is, in Excel's three
// forms, with value fields across the columns or down the rows, and with a Filters area.
import { afterEach, describe, expect, it } from 'vitest'
import { closeAll, source } from './local-fixture'
import { addSales, salesBook } from './pivot-fixture'

afterEach(closeAll)

async function laid(definition: Parameters<typeof addSales>[0], locale?: string) {
  const s = source(salesBook(locale))
  expect(await s.structure(addSales(definition))).toMatchObject({ ok: true })
  return s
}

describe('localSource: pivotCells', () => {
  it('compact form: Row Labels and Column Labels have the buttons; items, values, totals and headers by cell', async () => {
    const s = await laid({ rows: [0, 1], columns: [2], values: [{ field: 3 }] })
    const got = await s.pivotCells('R', 'PivotTable1', ['A3', 'B3', 'A4', 'B4', 'D4', 'A5', 'A6', 'B7', 'A12', 'C12', 'D3', 'F20'])
    expect(got?.labels.slice(0, 3)).toEqual([['East', 'North', 'South'], ['a', 'b'], ['2024', '2025']])
    expect(got?.labels[3]).toBeNull()
    expect(got?.buttons).toEqual([{ cell: 'A4', axis: 'rows', field: 0 }, { cell: 'B3', axis: 'columns', field: 2 }])
    expect(got?.cells).toEqual([
      { kind: 'valueHeader', value: 0 },
      { kind: 'header', axis: 'columns', field: 2 },
      { kind: 'header', axis: 'rows', field: 0 },
      { kind: 'item', axis: 'columns', field: 2, item: 0 },
      { kind: 'grand', axis: 'columns' },
      { kind: 'item', axis: 'rows', field: 0, item: 0 },
      { kind: 'item', axis: 'rows', field: 1, item: 0 },
      { kind: 'value', value: 0 },
      { kind: 'grand', axis: 'rows' },
      { kind: 'value', value: 0 },
      { kind: 'blank' },
      null,
    ])
  })

  it("outline and tabular forms: each row field's header has a button; tabular subtotals by their field", async () => {
    const outline = await laid({ rows: [0, 1], columns: [2], values: [{ field: 3 }], layout: 'outline' })
    expect((await outline.pivotCells('R', 'PivotTable1'))?.buttons).toEqual([
      { cell: 'A4', axis: 'rows', field: 0 },
      { cell: 'B4', axis: 'rows', field: 1 },
      { cell: 'C3', axis: 'columns', field: 2 },
    ])
    expect((await outline.pivotCells('R', 'PivotTable1', ['B6', 'A6', 'B4']))?.cells).toEqual([
      { kind: 'item', axis: 'rows', field: 1, item: 0 },
      { kind: 'item', axis: 'rows', field: 0, item: 0 },
      { kind: 'header', axis: 'rows', field: 1 },
    ])
    const tabular = await laid({ rows: [0, 1], values: [{ field: 3 }], layout: 'tabular' })
    expect((await tabular.pivotCells('R', 'PivotTable1', ['A4', 'B4', 'A5']))?.cells).toEqual([
      { kind: 'item', axis: 'rows', field: 0, item: 0 },
      { kind: 'item', axis: 'rows', field: 1, item: 0 },
      { kind: 'subtotal', axis: 'rows', field: 0, item: 0 },
    ])
  })

  it('value fields down the rows, and the Filters area with its buttons', async () => {
    const s = await laid({ rows: [0, 'values'], pages: [{ field: 1 }, { field: 2 }], values: [{ field: 3 }, { field: 3, function: 'count' }] })
    const got = await s.pivotCells('R', 'PivotTable1', ['A1', 'B1', 'B2', 'A6', 'B6', 'A7'])
    expect(got?.buttons).toEqual([{ cell: 'A4', axis: 'rows', field: 0 }, { cell: 'B1', axis: 'pages', field: 1 }, { cell: 'B2', axis: 'pages', field: 2 }])
    expect(got?.cells).toEqual([
      { kind: 'pageField', axis: 'pages', field: 1 },
      { kind: 'pageItem', axis: 'pages', field: 1 },
      { kind: 'pageItem', axis: 'pages', field: 2 },
      { kind: 'valueHeader', value: 0 },
      { kind: 'value', value: 0 },
      { kind: 'valueHeader', value: 1 },
    ])
  })

  it('labels dates as short dates, the blank item in the workbook language, and no buttons without field captions', async () => {
    const s = await laid({ rows: [4], values: [{ field: 3 }], fieldHeaders: false })
    const got = await s.pivotCells('R', 'PivotTable1')
    expect(got?.labels[4]).toEqual(['4/1/2024', '5/13/2024', '3/1/2025', '3/9/2025'])
    expect(got?.buttons).toEqual([])
    expect(await s.pivotCells('R', 'PivotTable9')).toBeNull()
  })
})
