// @vitest-environment node
// The PivotTable Fields pane's model: Σ Values among the axes' fields, where a checked field goes, moves between and within
// the areas as Excel's field list makes them, and the editPivotTable patch of a change.
import { describe, expect, it } from 'vitest'
import type { PivotCacheField } from 'xlcalc/contract'
import { areasOf, check, fieldKind, move, patchOf, remove, usedFields, valueFunction } from '../src/core/pivot-layout'

const FIELDS: PivotCacheField[] = [
  { name: 'Region', facts: { types: ['string'] } },
  { name: 'Amount', facts: { types: ['number'] } },
  { name: 'Date', facts: { types: ['date'] } },
  { name: 'Qty', facts: { types: ['number', 'blank'] } },
]

describe('the PivotTable Fields pane: where a field goes', () => {
  it('a field of numbers alone, blanks allowed, is a number field, Sum; dates a date field; anything else Count', () => {
    expect(FIELDS.map(fieldKind)).toEqual(['text', 'number', 'date', 'number'])
    expect(FIELDS.map(valueFunction)).toEqual(['count', 'sum', 'count', 'sum'])
    const kinds = (types: NonNullable<PivotCacheField['facts']>['types']) => fieldKind({ name: 'F', facts: { types } })
    expect([kinds(['blank']), kinds(['number', 'string']), kinds(['number', 'boolean']), kinds(['number', 'date'])]).toEqual(['text', 'text', 'text', 'text'])
  })

  it('checking puts a number field on Values and others on Rows; unchecking removes it from every area', () => {
    let a = areasOf({})
    a = check(a, 1, true, FIELDS)
    a = check(a, 0, true, FIELDS)
    a = check(a, 2, true, FIELDS)
    expect(a).toEqual({ filters: [], columns: [], rows: [0, 2], values: [{ field: 1, function: 'sum' }] })
    a = move(a, { area: 'list', field: 0 }, 'values', undefined, FIELDS)!
    expect(a.rows).toEqual([0, 2])
    expect(a.columns).toEqual(['values'])
    expect([...usedFields(a)].sort()).toEqual([0, 1, 2])
    a = check(a, 0, false, FIELDS)
    expect(a).toEqual({ filters: [], columns: [], rows: [2], values: [{ field: 1, function: 'sum' }] })
  })
})

describe('the PivotTable Fields pane: moves', () => {
  const two = () => areasOf({ rows: [0], values: [{ field: 1, name: 'Sum of Amount' }, { field: 3, name: 'Sum of Qty' }] })

  it('shows Σ Values on Columns for two value fields kept without it, at valuesPosition or on Rows with valuesOnRows', () => {
    expect(two().columns).toEqual(['values'])
    expect(areasOf({ rows: [0, 2], valuesOnRows: true, valuesPosition: 1, values: [{ field: 1 }, { field: 3 }] }).rows).toEqual([0, 'values', 2])
    expect(areasOf({ columns: ['values'], values: [{ field: 1 }] }).columns).toEqual([])
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
    expect(remove(a, { area: 'values', index: 1 })).toMatchObject({ columns: [], values: [{ field: 1 }] })
  })
})

describe('the PivotTable Fields pane: the patch', () => {
  it('sets the areas changed whole, an empty one as null, Σ Values with valuesOnRows; none for no change', () => {
    const t = { rows: [0], values: [{ field: 1, name: 'Sum of Amount' }, { field: 3, name: 'Sum of Qty' }] }
    expect(patchOf(t, areasOf(t))).toBeNull()
    const onRows = move(areasOf(t), { area: 'columns', index: 0 }, 'rows', undefined, FIELDS)!
    expect(patchOf(t, onRows)).toEqual({ rows: [0, 'values'], columns: null, valuesOnRows: true })
    expect(patchOf({ ...t, valuesPosition: 0 }, onRows)).toEqual({ rows: [0, 'values'], columns: null, valuesOnRows: true, valuesPosition: null })
    expect(patchOf(t, check(areasOf(t), 2, true, FIELDS))).toEqual({ rows: [0, 2] })
  })
})
