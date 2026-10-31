// @vitest-environment node
// localSource's pivotCells: the labels of a PivotTable's items, its filter buttons, and what a cell is as xlcalc's pivotCells
// tells it, in Excel's three forms, with value fields across the columns or down the rows, and with a Filters area.
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
  // what each cell is comes from xlcalc's pivotCells (Excel's Range.PivotCell); the source passes it through by address
  it("compact form: Row Labels and Column Labels name no field and have the buttons of their axes, as does the corner value field's name", async () => {
    const s = await laid({ rows: [0, 1], columns: [2], values: [{ field: 3 }] })
    const got = await s.pivotCells('R', 'PivotTable1', ['A3', 'B3', 'A4', 'B4', 'D4', 'A5', 'A6', 'B7', 'A12', 'C12', 'D3', 'F20'])
    expect(got?.labels.slice(0, 3)).toEqual([['East', 'North', 'South'], ['a', 'b'], ['2024', '2025']])
    expect(got?.labels[3]).toBeNull()
    // Row Labels and Column Labels filter their axis's first field; the corner value field's name filters none
    expect(got?.buttons).toEqual([{ cell: 'B3', axis: 'columns', field: 2 }, { cell: 'A4', axis: 'rows', field: 0 }, { cell: 'A3', axis: 'rows' }])
    expect(got?.cells).toMatchObject([
      { kind: 'valueName', value: 0, button: true },
      { kind: 'blank', button: true },
      { kind: 'blank', button: true },
      { kind: 'item', axis: 'columns', field: 2, item: 0 },
      { kind: 'grand', axis: 'columns', value: 0 },
      { kind: 'item', axis: 'rows', field: 0, item: 0 },
      { kind: 'item', axis: 'rows', field: 1, item: 0 },
      { kind: 'value', value: 0 },
      { kind: 'grand', axis: 'rows', value: 0 },
      { kind: 'value', value: 0 },
      { kind: 'blank' },
      null,
    ])
    // a value's cell carries the items of its row and column
    expect(got?.cells?.[7]).toEqual({ pivot: 'PivotTable1', kind: 'value', value: 0, rowItems: [{ field: 0, item: 1 }], columnItems: [{ field: 2, item: 0 }] })
  })

  it("outline and tabular forms: each field's name has its button; every cell of a total's line is the total's", async () => {
    const outline = await laid({ rows: [0, 1], columns: [2], values: [{ field: 3 }], layout: 'outline' })
    expect((await outline.pivotCells('R', 'PivotTable1'))?.buttons).toEqual([
      { cell: 'C3', axis: 'columns', field: 2 },
      { cell: 'A4', axis: 'rows', field: 0 },
      { cell: 'B4', axis: 'rows', field: 1 },
      { cell: 'A3', axis: 'rows' },
    ])
    expect((await outline.pivotCells('R', 'PivotTable1', ['B6', 'A6', 'B4']))?.cells).toMatchObject([
      { kind: 'item', axis: 'rows', field: 1, item: 0 },
      // the outer field's column beside its inner item is its item's
      { kind: 'item', axis: 'rows', field: 0, item: 0 },
      { kind: 'field', axis: 'rows', field: 1, button: true },
    ])
    // A4 East a, A5 East Total, A6 North a, A7 (North) b
    const tabular = await laid({ rows: [0, 1], values: [{ field: 3 }], layout: 'tabular' })
    expect((await tabular.pivotCells('R', 'PivotTable1', ['A4', 'B4', 'A5', 'B5', 'A7', 'B7']))?.cells).toMatchObject([
      { kind: 'item', axis: 'rows', field: 0, item: 0 },
      { kind: 'item', axis: 'rows', field: 1, item: 0 },
      { kind: 'subtotal', axis: 'rows', field: 0, item: 0 },
      { kind: 'subtotal', axis: 'rows', field: 0, item: 0 },
      { kind: 'item', axis: 'rows', field: 0, item: 1 },
      { kind: 'item', axis: 'rows', field: 1, item: 1 },
    ])
  })

  it('the value fields down the rows, the value field without any row field, and the Filters area with its buttons on the names', async () => {
    const values = await laid({ rows: ['values', 0], values: [{ field: 3 }, { field: 3, function: 'count' }] })
    const down = await values.pivotCells('R', 'PivotTable1', ['A3', 'A4'])
    expect(down?.cells).toMatchObject([{ kind: 'blank', button: true }, { kind: 'valueName', axis: 'rows', value: 0 }])
    // Row Labels over Σ Values and Region: Region, the first field of the rows that is no Σ Values
    expect(down?.buttons).toEqual([{ cell: 'A3', axis: 'rows', field: 0 }])
    // Row Labels over Σ Values alone filters no field, nor Region on the columns
    const alone = await laid({ columns: [0], rows: ['values'], values: [{ field: 3 }, { field: 2 }] })
    expect((await alone.pivotCells('R', 'PivotTable1'))?.buttons).toEqual([{ cell: 'B3', axis: 'columns', field: 0 }, { cell: 'A4', axis: 'rows' }])
    const columns = await laid({ columns: [1], values: [{ field: 3 }] })
    const got = await columns.pivotCells('R', 'PivotTable1', ['A3', 'B3', 'A5'])
    expect(got?.cells).toMatchObject([{ kind: 'valueName', value: 0 }, { kind: 'blank', button: true }, { kind: 'valueName', value: 0 }])
    expect(got?.buttons).toEqual([{ cell: 'B3', axis: 'columns', field: 1 }])
    const s = await laid({ rows: [0, 'values'], pages: [{ field: 1 }, { field: 2 }], values: [{ field: 3 }, { field: 3, function: 'count' }] })
    const pages = await s.pivotCells('R', 'PivotTable1', ['A1', 'B1', 'B2', 'A6', 'B6', 'A7', 'C1'])
    expect(pages?.buttons).toEqual([{ cell: 'A1', axis: 'pages', field: 1 }, { cell: 'A2', axis: 'pages', field: 2 }, { cell: 'A4', axis: 'rows', field: 0 }])
    expect(pages?.cells).toMatchObject([
      { kind: 'field', axis: 'pages', field: 1, button: true },
      { kind: 'pageItem', axis: 'pages', field: 1 },
      { kind: 'pageItem', axis: 'pages', field: 2 },
      { kind: 'valueName', axis: 'rows', value: 0 },
      { kind: 'value', value: 0 },
      { kind: 'valueName', axis: 'rows', value: 1 },
      // beside the filters: no cell of the PivotTable
      null,
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
