// A PivotTable's definition as the dialogs over it read and change it, without Vue: names of fields, the kinds of value a
// field holds, the Design tab's commands and their checked items, filters by comparison, and dates as xlcalc's ISO text.
import type { PivotDataField, PivotFieldPatch, PivotFilter, PivotFilterType, PivotFunction, PivotLayout, PivotTablePatch } from 'xlcalc/contract'
import type { PivotCompare } from './contract/messages'
import type { SheetPivotCells, SheetPivotTable } from './contract/source'
import type { PivotDesignChange } from './contract/view'

/** A PivotTable the dialogs act on: its sheet and name, its definition and what its cells show. */
export interface PivotTarget {
  sheet: string
  pivot: string
  def: SheetPivotTable
  cells: SheetPivotCells
}

/** Excel's 11 functions in the order of Summarize Values By's list. */
export const PIVOT_FUNCTIONS: readonly PivotFunction[] = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp']

/** A field's name on the PivotTable: its own, or its cache field's. */
export const fieldName = (t: PivotTarget, field: number): string => t.def.table.fields[field]?.name ?? t.def.cache.fields[field]?.name ?? ''

/** A value field's name as the PivotTable shows it. */
export const valueName = (t: PivotTarget, index: number): string => t.def.table.values?.[index]?.name ?? fieldName(t, t.def.table.values?.[index]?.field ?? -1)

/** The cache fields the PivotTable shows on its rows, columns or Filters area, in that order. */
export function axisFields(t: PivotTarget, axis: 'rows' | 'columns' | 'pages'): number[] {
  const tb = t.def.table
  const list = axis === 'pages' ? (tb.pages ?? []).map((p) => p.field) : (tb[axis] ?? [])
  return list.filter((f): f is number => f !== 'values')
}

/** The kind of value a field's items are, which decides its filters (Date Filters) and its Grouping dialog. */
export function fieldKind(t: PivotTarget, field: number): 'dates' | 'numbers' | 'text' {
  const f = t.def.cache.fields[field]
  const by = f?.group?.range?.by
  if (by) return by === 'numbers' ? 'numbers' : 'dates'
  const types = (f?.facts?.types ?? []).filter((x) => x !== 'blank')
  if (types.length && types.every((x) => x === 'date')) return 'dates'
  if (types.length && types.every((x) => x === 'number')) return 'numbers'
  return 'text'
}

/** The field a Grouping dialog or Ungroup acts on: a date's period field (`Months (Date)`) stands for its date field. */
export function groupedField(t: PivotTarget, field: number): number {
  const g = t.def.cache.fields[field]?.group
  return g?.base !== undefined && g.range ? g.base : field
}

/** Whether a field has filters or hidden items, for Clear Filter From. */
export function filtered(t: PivotTarget, field: number): boolean {
  return !!t.def.table.filters?.some((x) => x.field === field) || !!t.def.table.fields[field]?.items?.some((x) => x.hidden)
}

/** The PivotTable field's items without what hides them: every item shown. */
export const shownItems = (t: PivotTarget, field: number) => (t.def.table.fields[field]?.items ?? []).map(({ hidden: _hidden, ...rest }) => rest)

// ── filters ──
const LABEL: Partial<Record<PivotCompare, PivotFilterType>> = {
  equal: 'captionEqual',
  notEqual: 'captionNotEqual',
  beginsWith: 'captionBeginsWith',
  notBeginsWith: 'captionNotBeginsWith',
  endsWith: 'captionEndsWith',
  notEndsWith: 'captionNotEndsWith',
  contains: 'captionContains',
  notContains: 'captionNotContains',
  greaterThan: 'captionGreaterThan',
  greaterThanOrEqual: 'captionGreaterThanOrEqual',
  lessThan: 'captionLessThan',
  lessThanOrEqual: 'captionLessThanOrEqual',
  between: 'captionBetween',
  notBetween: 'captionNotBetween',
}
const VALUE: Partial<Record<PivotCompare, PivotFilterType>> = {
  equal: 'valueEqual',
  notEqual: 'valueNotEqual',
  greaterThan: 'valueGreaterThan',
  greaterThanOrEqual: 'valueGreaterThanOrEqual',
  lessThan: 'valueLessThan',
  lessThanOrEqual: 'valueLessThanOrEqual',
  between: 'valueBetween',
  notBetween: 'valueNotBetween',
}
const DATE: Partial<Record<PivotCompare, PivotFilterType>> = {
  equal: 'dateEqual',
  notEqual: 'dateNotEqual',
  before: 'dateOlderThan',
  beforeOrEqual: 'dateOlderThanOrEqual',
  after: 'dateNewerThan',
  afterOrEqual: 'dateNewerThanOrEqual',
  between: 'dateBetween',
  notBetween: 'dateNotBetween',
}
/** The kinds of filter: by label, by a value field's values, by date. */
export type FilterKind = 'label' | 'value' | 'date'
const BY_KIND: Record<FilterKind, Partial<Record<PivotCompare, PivotFilterType>>> = { label: LABEL, value: VALUE, date: DATE }

/** The comparisons each kind of filter's submenu and dialog list, in Excel's order. */
export const COMPARES: Record<FilterKind, readonly PivotCompare[]> = {
  label: Object.keys(LABEL) as PivotCompare[],
  value: Object.keys(VALUE) as PivotCompare[],
  date: Object.keys(DATE) as PivotCompare[],
}
/** The submenu of Date Filters lists these comparisons; its dialog lists every one. */
export const DATE_MENU: readonly PivotCompare[] = ['equal', 'before', 'after', 'between']

export const filterType = (kind: FilterKind, compare: PivotCompare): PivotFilterType => BY_KIND[kind][compare]!
/** The comparison and kind of a filter's type; null for Top 10 and the dates' fixed periods. */
export function compareOf(type: PivotFilterType): { kind: FilterKind; compare: PivotCompare } | null {
  for (const kind of ['label', 'value', 'date'] as const) {
    const hit = (Object.entries(BY_KIND[kind]) as [PivotCompare, PivotFilterType][]).find(([, t]) => t === type)
    if (hit) return { kind, compare: hit[0] }
  }
  return null
}
export const isBetween = (compare: PivotCompare) => compare === 'between' || compare === 'notBetween'
const isValueFilter = (x: PivotFilter) => x.type.startsWith('value') || x.type === 'count' || x.type === 'percent' || x.type === 'sum'

/** The Date Filters periods Excel lists before All Dates in the Period, in its order. */
export const DATE_PERIODS: readonly PivotFilterType[] = ['tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate']

/**
 * The edit that puts a filter on a field: a filter by label or date takes the place of the field's other filters and of
 * its items unchecked, a value filter or Top 10 that of its other value filter; with Allow multiple filters per field the
 * other kind stays, as Excel keeps it, and without it the field keeps one filter.
 */
export function withFilter(t: PivotTarget, filter: PivotFilter): PivotTablePatch {
  const tb = t.def.table
  const value = isValueFilter(filter)
  const kept = (tb.filters ?? []).filter((x) => x.field !== filter.field || (tb.multipleFieldFilters && isValueFilter(x) !== value))
  return { filters: [...kept, filter], ...(value ? {} : { fields: [{ field: filter.field, items: shownItems(t, filter.field) }] }) }
}

/** The edit of Clear Filter From: the field's filters gone and every item shown. */
export function withoutFilters(t: PivotTarget, field: number): PivotTablePatch {
  const left = (t.def.table.filters ?? []).filter((x) => x.field !== field)
  return { filters: left.length ? left : null, fields: [{ field, items: shownItems(t, field) }] }
}

/** The edit of the item list's OK: the items unchecked hidden; filters by label or date on the field give way to it. */
export function withHidden(t: PivotTarget, field: number, hidden: boolean[]): PivotTablePatch {
  const items = (t.def.table.fields[field]?.items ?? []).map(({ hidden: _was, ...rest }, i) => (hidden[i] ? { ...rest, hidden: true } : rest))
  const left = (t.def.table.filters ?? []).filter((x) => x.field !== field || isValueFilter(x))
  return { fields: [{ field, items }], ...(left.length !== (t.def.table.filters?.length ?? 0) ? { filters: left.length ? left : null } : {}) }
}

// ── Design tab ──
/** What the Design tab's lists show checked: the subtotals of every field on the rows and columns, the grand totals, the form. */
export function designOf(t: PivotTarget) {
  const tb = t.def.table
  const fields = [...axisFields(t, 'rows'), ...axisFields(t, 'columns')].map((f) => tb.fields[f] ?? {})
  const subtotals = fields.length && fields.every((f) => f.subtotals?.length === 0) ? 'none' : fields.length && fields.every((f) => f.subtotalTop === false) ? 'bottom' : 'top'
  const rows = tb.grandTotals?.rows !== false
  const columns = tb.grandTotals?.columns !== false
  const grandTotals = rows && columns ? 'on' : rows ? 'rows' : columns ? 'columns' : 'off'
  return { subtotals, grandTotals, layout: tb.layout ?? 'compact', repeatLabels: !!tb.repeatLabels, blankLine: axisFields(t, 'rows').some((f) => tb.fields[f]?.blankLine) } as const
}

/** The edit a Design command makes, over every field on the rows and columns where it sets their fields. */
export function designPatch(t: PivotTarget, change: PivotDesignChange): PivotTablePatch {
  const onAxes = [...axisFields(t, 'rows'), ...axisFields(t, 'columns')]
  const each = (patch: Omit<PivotFieldPatch, 'field'>, list = onAxes) => list.map((field) => ({ field, ...patch }))
  if ('subtotals' in change) {
    if (change.subtotals === 'none') return { fields: each({ subtotals: [] }) }
    return { fields: each({ subtotals: null, subtotalTop: change.subtotals === 'top' ? null : false }) }
  }
  if ('grandTotals' in change) {
    const g = change.grandTotals
    return { grandTotals: g === 'on' ? null : { rows: g === 'rows', columns: g === 'columns' } }
  }
  if ('layout' in change) return { layout: change.layout === 'compact' ? null : (change.layout as PivotLayout), fields: t.def.table.fields.map((_, field) => ({ field, layout: null })) }
  if ('repeatLabels' in change) return { repeatLabels: change.repeatLabels || null, fields: t.def.table.fields.map((_, field) => ({ field, repeatLabels: null })) }
  return { fields: each({ blankLine: change.blankLine || null }, axisFields(t, 'rows')) }
}

// ── value fields ──
/** The value fields with the one at `index` replaced. */
export const withValue = (t: PivotTarget, index: number, next: PivotDataField): PivotTablePatch => ({ values: (t.def.table.values ?? []).map((v, i) => (i === index ? next : v)) })

/**
 * A value field summarized by another function: its name goes with the function, xlcalc naming it as Excel does, unless
 * `name` is given. TODO(xlcalc): the library does not tell whether a name is its own default, which Excel keeps custom.
 */
export function withFunction(v: PivotDataField, fn: PivotFunction, name?: string): PivotDataField {
  const { name: _old, function: _fn, ...rest } = v
  return { ...rest, ...(fn !== 'sum' ? { function: fn } : {}), ...(name ? { name } : {}) }
}

// ── dates ──
const DAY_MS = 86_400_000
/** The 1900 date system's day 0 as Excel counts it from March 1900 on (its phantom 29 February before). */
const EPOCH = Date.UTC(1899, 11, 30)

/** A serial date as xlcalc's ISO text without a zone (`2024-01-31T00:00:00`). */
export function isoOfSerial(serial: number): string {
  return new Date(EPOCH + Math.round(serial * DAY_MS)).toISOString().slice(0, 19)
}
/** xlcalc's ISO text as a serial date. */
export function serialOfIso(iso: string): number {
  const ms = Date.parse(iso.length <= 10 ? `${iso}T00:00:00Z` : `${iso}Z`)
  return (ms - EPOCH) / DAY_MS
}
