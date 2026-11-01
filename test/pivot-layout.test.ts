// @vitest-environment node
// The PivotTable Fields pane's model: where a checked field goes and its function by xlcalc's suggestPivotField, moves
// between and within the areas as Excel's field list makes them, Σ Values moved by the user, and the editPivotTable patch
// of a change, in which xlcalc keeps Σ Values as the value fields come and go.
import { describe, expect, it } from 'vitest'
import type { PivotCache } from 'xlcalc/contract'
import { areasOf, check, move, patchOf, remove, usedFields } from '../src/core/pivot-layout'

const CACHE: PivotCache = {
  id: 1,
  source: { type: 'range', sheet: 'Data', range: 'A1:D9' },
  fields: [
    { name: 'Region', facts: { types: ['string'] } },
    { name: 'Amount', facts: { types: ['number'] } },
    { name: 'Date', facts: { types: ['date'] } },
    { name: 'Qty', facts: { types: ['number', 'blank'] } },
  ],
}
const FIELDS = CACHE

describe('the PivotTable Fields pane: where a field goes', () => {
  it("checking puts a field where xlcalc's suggestPivotField says, Sum or Count; unchecking removes it from every area", () => {
    let a = areasOf({})
    a = check(a, 1, true, CACHE)
    a = check(a, 0, true, CACHE)
    a = check(a, 2, true, CACHE)
    a = check(a, 3, true, CACHE)
    expect(a).toEqual({ filters: [], columns: [], rows: [0, 2], values: [{ field: 1, function: 'sum' }, { field: 3, function: 'sum' }] })
    a = move(a, { area: 'list', field: 0 }, 'values', undefined, CACHE)!
    expect(a.rows).toEqual([0, 2])
    expect(a.values.at(-1)).toEqual({ field: 0, function: 'count' })
    // Σ Values is xlcalc's to place: the pane's change carries the value fields alone
    expect(a.columns).toEqual([])
    expect([...usedFields(a)].sort()).toEqual([0, 1, 2, 3])
    a = check(a, 0, false, CACHE)
    expect(a).toEqual({ filters: [], columns: [], rows: [2], values: [{ field: 1, function: 'sum' }, { field: 3, function: 'sum' }] })
  })
})

describe('the PivotTable Fields pane: moves', () => {
  // two value fields as xlcalc holds them: Σ Values at the end of the columns
  const two = () => areasOf({ rows: [0], columns: ['values'], values: [{ field: 1, name: 'Sum of Amount' }, { field: 3, name: 'Sum of Qty' }] })

  it('shows Σ Values where xlcalc holds it', () => {
    expect(two().columns).toEqual(['values'])
    expect(areasOf({ rows: [0, 'values', 2], values: [{ field: 1 }, { field: 3 }] }).rows).toEqual([0, 'values', 2])
  })

  it('moves within an area by place, Move Up and Move Down included', () => {
    const a = areasOf({ rows: [0, 2, 3] })
    expect(move(a, { area: 'rows', index: 2 }, 'rows', 0, FIELDS)!.rows).toEqual([3, 0, 2])
    expect(move(a, { area: 'rows', index: 0 }, 'rows', 2, FIELDS)!.rows).toEqual([2, 0, 3])
    expect(move(a, { area: 'rows', index: 0 }, 'rows', 3, FIELDS)!.rows).toEqual([2, 3, 0])
    expect(move(a, { area: 'rows', index: 1 }, 'rows', 1, FIELDS)).toBeNull()
  })

  it('a field lies on one axis; a value field dragged to an axis leaves Values, an axis field dragged to Values leaves its axis', () => {
    let a = areasOf({ rows: [0], columns: [2], values: [{ field: 1, name: 'Sum of Amount' }] })
    a = move(a, { area: 'list', field: 2 }, 'rows', 0, FIELDS)!
    expect(a).toMatchObject({ rows: [2, 0], columns: [] })
    a = move(a, { area: 'values', index: 0 }, 'filters', undefined, FIELDS)!
    expect(a).toMatchObject({ filters: [{ field: 1 }], values: [] })
    a = move(a, { area: 'rows', index: 1 }, 'values', undefined, FIELDS)!
    expect(a).toMatchObject({ rows: [2], values: [{ field: 0, function: 'count' }] })
  })

  it('a field dropped onto the axis it lies on already lands where it was dropped, counted without it', async () => {
    const a = areasOf({ rows: [0, 2, 3], values: [{ field: 2, function: 'count' }] })
    // from the list: Region dropped before Kind's place lands between Date and Qty
    expect(move(a, { area: 'list', field: 0 }, 'rows', 2, FIELDS)!.rows).toEqual([2, 0, 3])
    // from Values onto Rows, where Date lies first: dropped at the end
    expect(move(a, { area: 'values', index: 0 }, 'rows', 3, FIELDS)).toMatchObject({ rows: [0, 3, 2], values: [] })
    expect(move(a, { area: 'list', field: 3 }, 'rows', 0, FIELDS)!.rows).toEqual([3, 0, 2])
  })

  it('Σ Values goes between Columns and Rows only; removing it removes every value field', () => {
    const a = two()
    expect(move(a, { area: 'columns', index: 0 }, 'rows', 0, FIELDS)).toMatchObject({ rows: ['values', 0], columns: [] })
    expect(move(a, { area: 'columns', index: 0 }, 'filters', 0, FIELDS)).toBeNull()
    expect(move(a, { area: 'columns', index: 0 }, 'values', 0, FIELDS)).toBeNull()
    expect(remove(a, { area: 'columns', index: 0 })).toEqual({ filters: [], columns: [], rows: [0], values: [] })
    // the second to last value field removed: xlcalc takes Σ Values away with it
    expect(remove(a, { area: 'values', index: 1 })).toMatchObject({ columns: ['values'], values: [{ field: 1 }] })
  })
})

describe('the PivotTable Fields pane: the patch', () => {
  it('sets the areas changed whole, an empty one as null, Σ Values where the user moved it; none for no change', () => {
    const t = { rows: [0], columns: ['values' as const], values: [{ field: 1, name: 'Sum of Amount' }, { field: 3, name: 'Sum of Qty' }] }
    expect(patchOf(t, areasOf(t))).toBeNull()
    const onRows = move(areasOf(t), { area: 'columns', index: 0 }, 'rows', undefined, FIELDS)!
    expect(patchOf(t, onRows)).toEqual({ rows: [0, 'values'], columns: null })
    expect(patchOf(t, check(areasOf(t), 2, true, FIELDS))).toEqual({ rows: [0, 2] })
    expect(patchOf(t, remove(areasOf(t), { area: 'values', index: 1 }))).toEqual({ values: [{ field: 1, name: 'Sum of Amount' }] })
  })
})
