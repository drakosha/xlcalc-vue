// The PivotTable Fields pane's model: a PivotTable's four areas (Filters, Columns, Rows, Values) with Σ Values among the
// axes' fields, the changes the pane makes to them as Excel's field list makes them (where a field goes and its function
// by xlcalc's `suggestPivotField`), and the `editPivotTable` patch that carries a change in one step of undo, in which
// xlcalc keeps Σ Values. Pure, over xlcalc's PivotTable model.
import type { PivotAxisField, PivotCache, PivotDataField, PivotPageField, PivotTable, PivotTablePatch } from 'xlcalc/contract'
import { suggestPivotField } from 'xlcalc/text'

/** The pane's areas, in Excel's order. */
export type PivotArea = 'filters' | 'columns' | 'rows' | 'values'

/** The areas as the pane shows them: Σ Values (`values`) is among the Columns' or Rows' fields when there are several value fields. */
export interface PivotAreas {
  filters: PivotPageField[]
  columns: PivotAxisField[]
  rows: PivotAxisField[]
  values: PivotDataField[]
}

/** A field in an area by its place, or a field of the list by its cache index (`list`). */
export type PivotSpot = { area: PivotArea; index: number } | { area: 'list'; field: number }

const insertAt = <T>(list: T[], at: number, item: T): T[] => [...list.slice(0, at), item, ...list.slice(at)]

/** A PivotTable's areas as the pane shows them: Σ Values where xlcalc holds it, on an axis while there are two value fields or more. */
export function areasOf(t: Pick<PivotTable, 'rows' | 'columns' | 'pages' | 'values'>): PivotAreas {
  return { filters: [...(t.pages ?? [])], columns: [...(t.columns ?? [])], rows: [...(t.rows ?? [])], values: [...(t.values ?? [])] }
}

// ── the changes ──

/** The field a spot holds: a cache field's index, or `values` for Σ Values. */
export function fieldAt(a: PivotAreas, spot: PivotSpot): number | 'values' | null {
  if (spot.area === 'list') return spot.field
  if (spot.area === 'filters') return a.filters[spot.index]?.field ?? null
  if (spot.area === 'values') return a.values[spot.index]?.field ?? null
  return a[spot.area][spot.index] ?? null
}

/** The fields on any area, which the list shows checked. */
export function usedFields(a: PivotAreas): Set<number> {
  const out = new Set<number>()
  for (const f of [...a.rows, ...a.columns]) if (f !== 'values') out.add(f)
  for (const p of a.filters) out.add(p.field)
  for (const v of a.values) out.add(v.field)
  return out
}

/** The areas without the field on an axis (a field lies on one axis); its value fields stay. */
function offAxes(a: PivotAreas, field: number | 'values'): PivotAreas {
  return { ...a, rows: a.rows.filter((f) => f !== field), columns: a.columns.filter((f) => f !== field), filters: a.filters.filter((p) => p.field !== field) }
}

/** The areas without what the spot holds; a field of the list leaves every area. */
function without(a: PivotAreas, spot: PivotSpot): PivotAreas {
  if (spot.area === 'list') return { ...offAxes(a, spot.field), values: a.values.filter((v) => v.field !== spot.field) }
  if (spot.area === 'values') return { ...a, values: a.values.filter((_, i) => i !== spot.index) }
  if (spot.area === 'filters') return { ...a, filters: a.filters.filter((_, i) => i !== spot.index) }
  return { ...a, [spot.area]: a[spot.area].filter((_, i) => i !== spot.index) }
}

/**
 * A field moved from a spot to an area at `index` (its end without it), as a drag in Excel's field list moves it: within an
 * area it changes places; onto Values it becomes a value field (leaving its axis when it came from one, staying there when it
 * came from the list), summarized as Excel summarizes a new one; onto an axis it leaves the axis it was on, and a value field
 * dragged there leaves Values. Σ Values goes only between Columns and Rows. Null when nothing changes.
 */
export function move(a: PivotAreas, from: PivotSpot, to: PivotArea, index: number | undefined, cache: PivotCache): PivotAreas | null {
  const field = fieldAt(a, from)
  if (field === null) return null
  if (field === 'values' && (to === 'filters' || to === 'values')) return null
  if (from.area === to) {
    const list = a[to] as unknown[]
    const at = Math.min(index ?? list.length, list.length)
    // dropped where it lies, or right after itself
    if (at === from.index || at === from.index + 1) return null
    const item = list[from.index]
    const rest = list.filter((_, i) => i !== from.index)
    return { ...a, [to]: insertAt(rest, at > from.index ? at - 1 : at, item) }
  }
  if (field === 'values') {
    const rest = offAxes(a, 'values')
    const axis = rest[to as 'rows' | 'columns']
    return { ...rest, [to]: insertAt(axis, Math.min(index ?? axis.length, axis.length), 'values') }
  }
  let rest = from.area === 'list' ? a : without(a, from)
  if (to === 'values') {
    if (from.area !== 'list') rest = offAxes(rest, field)
    const value: PivotDataField = { field, function: suggestPivotField(cache, field).function }
    return { ...rest, values: insertAt(rest.values, Math.min(index ?? rest.values.length, rest.values.length), value) }
  }
  // where the field lies on the axis it goes to already (from the list, or from Values): places after it count without it
  const was = to === 'filters' ? rest.filters.findIndex((p) => p.field === field) : rest[to].indexOf(field)
  const at = index !== undefined && was >= 0 && index > was ? index - 1 : index
  rest = offAxes(rest, field)
  if (to === 'filters') {
    const page: PivotPageField = { field }
    return { ...rest, filters: insertAt(rest.filters, Math.min(at ?? rest.filters.length, rest.filters.length), page) }
  }
  const axis = rest[to]
  return { ...rest, [to]: insertAt(axis, Math.min(at ?? axis.length, axis.length), field) }
}

/** A field's chip removed from its area (Remove Field, or dragged out of the areas); Σ Values removes every value field. */
export function remove(a: PivotAreas, spot: PivotSpot): PivotAreas {
  if (fieldAt(a, spot) === 'values') return { ...offAxes(a, 'values'), values: [] }
  return without(a, spot)
}

/** A field checked in the list: where Excel's field list puts it (xlcalc's `suggestPivotField`); unchecked, it leaves every area. */
export function check(a: PivotAreas, field: number, on: boolean, cache: PivotCache): PivotAreas {
  if (!on) return without(a, { area: 'list', field })
  return move(a, { area: 'list', field }, suggestPivotField(cache, field).axis, undefined, cache) ?? a
}

const same = (x: unknown, y: unknown) => JSON.stringify(x) === JSON.stringify(y)

/**
 * The `editPivotTable` patch from the PivotTable's areas to `next`: each area changed set whole (null for an empty one), Σ
 * Values among the axis's fields where the user put it; xlcalc puts it there and takes it away as the value fields come and
 * go. Null for no change.
 */
export function patchOf(t: Pick<PivotTable, 'rows' | 'columns' | 'pages' | 'values'>, next: PivotAreas): PivotTablePatch | null {
  const now = areasOf(t)
  const set: PivotTablePatch = {}
  const list = <T>(x: T[]) => (x.length ? x : null)
  if (!same(now.rows, next.rows)) set.rows = list(next.rows)
  if (!same(now.columns, next.columns)) set.columns = list(next.columns)
  if (!same(now.filters, next.filters)) set.pages = list(next.filters)
  if (!same(now.values, next.values)) set.values = list(next.values)
  return Object.keys(set).length ? set : null
}
