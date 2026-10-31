// A PivotTable's definition as the dialogs over it read and change it, without Vue: names of fields, the kinds of value a
// field holds, the Design tab's commands and their checked items, filters by comparison, and dates as xlcalc's ISO text.
import type { PivotDataField, PivotFieldPatch, PivotFilter, PivotFilterType, PivotFunction, PivotLayout, PivotTablePatch } from 'xlcalc/contract'
import type { PivotCompare } from './contract/messages'
import type { PivotAxis, SheetPivotCell, SheetPivotCells, SheetPivotTable } from './contract/source'
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
/** A cell's cache field, none for Σ Values (`values`) or a cell of no field. */
export const cellField = (c: SheetPivotCell | null | undefined): number | undefined => (typeof c?.field === 'number' ? c.field : undefined)
/**
 * The field a cell's menu and commands are about: the cell's own, or on Row Labels and Column Labels, which name no field
 * (xlcalc's `blank`), the field of the filter button on it as the source chose it (`SheetPivotButton`); none for others.
 */
export function cellTarget(t: PivotTarget, cell: string, c: SheetPivotCell | null | undefined): { axis?: PivotAxis; field?: number } {
  const own = cellField(c)
  if (own !== undefined) return { ...(c?.axis ? { axis: c.axis } : {}), field: own }
  const b = c?.kind === 'blank' ? t.cells.buttons.find((x) => x.cell === cell && x.field !== undefined) : undefined
  return b ? { axis: b.axis, field: b.field } : {}
}
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

/**
 * The field Ungroup takes away for an item of `field`, as xlcalc ungroups: a field grouped by hand (a group field), a field of
 * numbers grouped in place, or a date's period field standing for its date field; null for any other, a field grouped by
 * hand's base among them, where Excel does not offer Ungroup.
 */
export function ungroupedField(t: PivotTarget, field: number): number | null {
  const g = t.def.cache.fields[field]?.group
  if (g?.discrete) return field
  if (g?.range) return g.base ?? field
  return null
}

/** Whether a field has filters or hidden items, or a Filters area's field one item chosen, for its funnel and Clear Filter From. */
export function filtered(t: PivotTarget, field: number): boolean {
  const tb = t.def.table
  return !!tb.filters?.some((x) => x.field === field) || !!tb.fields[field]?.items?.some((x) => x.hidden) || !!tb.pages?.some((p) => p.field === field && p.item !== undefined)
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

// ── grand totals ──
/**
 * Excel's Show grand totals for rows and for columns (RowGrand, ColumnGrand), as xlcalc holds them: RowGrand is xlsx
 * `colGrandTotals`, the grand total column at the right, xlcalc's `grandTotals.columns`, and ColumnGrand the grand total row
 * at the bottom, `grandTotals.rows` (xlcalc's docs/development.md, PivotTables: the case `no-row-grand`, where Excel for
 * Windows' RowGrand false wrote `colGrandTotals="0"` and the column at the right went).
 */
export function grandOf(t: PivotTarget): { rows: boolean; columns: boolean } {
  const g = t.def.table.grandTotals
  return { rows: g?.columns !== false, columns: g?.rows !== false }
}
/** The patch of Excel's two boxes, null for both on (xlcalc's default). */
export const grandPatch = (rows: boolean, columns: boolean): PivotTablePatch['grandTotals'] => (rows && columns ? null : { rows: columns, columns: rows })

// ── Design tab ──
/** What the Design tab's lists show checked: the subtotals of every field on the rows and columns, the grand totals, the form. */
export function designOf(t: PivotTarget) {
  const tb = t.def.table
  const fields = [...axisFields(t, 'rows'), ...axisFields(t, 'columns')].map((f) => tb.fields[f] ?? {})
  const subtotals = fields.length && fields.every((f) => f.subtotals?.length === 0) ? 'none' : fields.length && fields.every((f) => f.subtotalTop === false) ? 'bottom' : 'top'
  const { rows, columns } = grandOf(t)
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
    return { grandTotals: grandPatch(g === 'on' || g === 'rows', g === 'on' || g === 'columns') }
  }
  if ('layout' in change) return { layout: change.layout === 'compact' ? null : (change.layout as PivotLayout), fields: t.def.table.fields.map((_, field) => ({ field, layout: null })) }
  if ('repeatLabels' in change) return { repeatLabels: change.repeatLabels || null, fields: t.def.table.fields.map((_, field) => ({ field, repeatLabels: null })) }
  return { fields: each({ blankLine: change.blankLine || null }, axisFields(t, 'rows')) }
}

// ── value fields ──
/** The value fields with the one at `index` replaced. */
export const withValue = (t: PivotTarget, index: number, next: PivotDataField): PivotTablePatch => ({ values: (t.def.table.values ?? []).map((v, i) => (i === index ? next : v)) })

/** A value field summarized by another function, with the name given; without one xlcalc names it as Excel does (its `renamedValueField`). */
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
