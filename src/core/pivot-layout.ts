// The PivotTable Fields pane's model: a PivotTable's four areas (Filters, Columns, Rows, Values) with Σ Values among the
// axes' fields, the changes the pane makes to them as Excel's field list makes them, and the `editPivotTable` patch that
// carries a change in one step of undo. Pure, over xlcalc's PivotTable model.
import type { PivotAxisField, PivotCacheField, PivotDataField, PivotFunction, PivotPageField, PivotTable, PivotTablePatch } from 'xlcalc/contract'

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

// ── where a field goes and how it is summarized: one place, for xlcalc's coming suggestPivotField to take over ──

/** Where Excel puts a field checked in the list, and how a value field summarizes it. */
export type FieldKind = 'number' | 'date' | 'text'

const insertAt = <T>(list: T[], at: number, item: T): T[] => [...list.slice(0, at), item, ...list.slice(at)]

/**
 * The kind of a cache field by the values it holds (`facts`), as xlcalc measured Excel for Windows' field list check: numbers
 * alone (blanks allowed) are a number field, which the check puts on Values with Sum; dates alone (blanks allowed) a date
 * field; anything else, blanks alone and numbers mixed with text, logicals, errors or dates included, text, put on Rows.
 */
export function fieldKind(f: PivotCacheField): FieldKind {
  const types = f.facts?.types ?? []
  if (types.includes('number') && types.every((t) => t === 'number' || t === 'blank')) return 'number'
  if (types.includes('date') && types.every((t) => t === 'date' || t === 'blank')) return 'date'
  return 'text'
}

/** Excel's function for a new value field: Sum over a number field, Count over any other dropped onto Values. */
export const valueFunction = (f: PivotCacheField | undefined): PivotFunction => (f && fieldKind(f) === 'number' ? 'sum' : 'count')

/** The area a field checked in the list goes to: a number field onto Values, a date or text field onto Rows. */
const checkedArea = (f: PivotCacheField | undefined): PivotArea => (f && fieldKind(f) === 'number' ? 'values' : 'rows')

// ── Σ Values upkeep: one place, for xlcalc's coming editPivotTable that keeps the marker itself to take over ──

const hasValuesField = (axis: PivotAxisField[]) => axis.includes('values')

/**
 * Σ Values follows the value fields, as Excel's field list does: gone below two; with two or more and none on an axis, on
 * Rows (`onRows`) or Columns at `at`, at the end without it.
 */
function settle(a: PivotAreas, place: { onRows?: boolean; at?: number } = {}): PivotAreas {
  if (a.values.length < 2) return { ...a, rows: a.rows.filter((f) => f !== 'values'), columns: a.columns.filter((f) => f !== 'values') }
  if (hasValuesField(a.rows) || hasValuesField(a.columns)) return a
  const key = place.onRows ? 'rows' : 'columns'
  return { ...a, [key]: insertAt(a[key], Math.min(place.at ?? a[key].length, a[key].length), 'values') }
}

/**
 * A PivotTable's areas as the pane shows them: Σ Values where xlcalc holds it on an axis, or, for several value fields
 * without it, on Rows (`valuesOnRows`) or Columns at `valuesPosition`, at the end without it.
 */
export function areasOf(t: Pick<PivotTable, 'rows' | 'columns' | 'pages' | 'values' | 'valuesOnRows' | 'valuesPosition'>): PivotAreas {
  const areas: PivotAreas = { filters: [...(t.pages ?? [])], columns: [...(t.columns ?? [])], rows: [...(t.rows ?? [])], values: [...(t.values ?? [])] }
  return settle(areas, { onRows: t.valuesOnRows, at: t.valuesPosition })
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
export function move(a: PivotAreas, from: PivotSpot, to: PivotArea, index: number | undefined, fields: PivotCacheField[]): PivotAreas | null {
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
    const value: PivotDataField = { field, function: valueFunction(fields[field]) }
    return settle({ ...rest, values: insertAt(rest.values, Math.min(index ?? rest.values.length, rest.values.length), value) })
  }
  // where the field lies on the axis it goes to already (from the list, or from Values): places after it count without it
  const was = to === 'filters' ? rest.filters.findIndex((p) => p.field === field) : rest[to].indexOf(field)
  const at = index !== undefined && was >= 0 && index > was ? index - 1 : index
  rest = offAxes(rest, field)
  if (to === 'filters') {
    const page: PivotPageField = { field }
    return settle({ ...rest, filters: insertAt(rest.filters, Math.min(at ?? rest.filters.length, rest.filters.length), page) })
  }
  const axis = rest[to]
  return settle({ ...rest, [to]: insertAt(axis, Math.min(at ?? axis.length, axis.length), field) })
}

/** A field's chip removed from its area (Remove Field, or dragged out of the areas); Σ Values removes every value field. */
export function remove(a: PivotAreas, spot: PivotSpot): PivotAreas {
  if (fieldAt(a, spot) === 'values') return settle({ ...offAxes(a, 'values'), values: [] })
  return settle(without(a, spot))
}

/**
 * A field checked in the list: a number field onto Values, a date or text field onto Rows, as Excel puts it; unchecked, it
 * leaves every area.
 */
export function check(a: PivotAreas, field: number, on: boolean, fields: PivotCacheField[]): PivotAreas {
  if (!on) return settle(without(a, { area: 'list', field }))
  return move(a, { area: 'list', field }, checkedArea(fields[field]), undefined, fields) ?? a
}

const same = (x: unknown, y: unknown) => JSON.stringify(x) === JSON.stringify(y)

/** The patch that gives the PivotTable these value fields, Σ Values kept on an axis as Excel keeps it (`settle`); null for no change. */
export function valuesPatch(t: Pick<PivotTable, 'rows' | 'columns' | 'pages' | 'values' | 'valuesOnRows' | 'valuesPosition'>, values: PivotDataField[]): PivotTablePatch | null {
  return patchOf(t, settle({ ...areasOf(t), values }))
}

/**
 * The `editPivotTable` patch from the PivotTable's areas to `next`: each area changed set whole (null for an empty one), and
 * with Σ Values moved, `valuesOnRows` where it lies (`valuesPosition` dropped, its place now on the axis); null for no change.
 */
export function patchOf(t: Pick<PivotTable, 'rows' | 'columns' | 'pages' | 'values' | 'valuesOnRows' | 'valuesPosition'>, next: PivotAreas): PivotTablePatch | null {
  const now = areasOf(t)
  const set: PivotTablePatch = {}
  const list = <T>(x: T[]) => (x.length ? x : null)
  if (!same(now.rows, next.rows)) set.rows = list(next.rows)
  if (!same(now.columns, next.columns)) set.columns = list(next.columns)
  if (!same(now.filters, next.filters)) set.pages = list(next.filters)
  if (!same(now.values, next.values)) set.values = list(next.values)
  if (set.rows !== undefined || set.columns !== undefined) {
    const onRows = hasValuesField(next.rows)
    if (!!t.valuesOnRows !== onRows) set.valuesOnRows = onRows || null
    if (t.valuesPosition !== undefined) set.valuesPosition = null
  }
  return Object.keys(set).length ? set : null
}
