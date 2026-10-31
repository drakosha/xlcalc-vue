// Excel's filter menu of a PivotTable's field (its button on Row Labels, Column Labels, a field's header or a Filters
// area's item, a field's arrow in the PivotTable Fields pane): sorting, Clear Filter From, Label, Value and Date Filters
// with their dialogs, Top 10, the search box and the items' list; and More Sort Options' Sort dialog. Each OK is one
// `editPivotTable` through the source's `structure`.
import { markRaw, ref } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { PivotFilter, PivotFilterType } from 'xlcalc/contract'
import type { PivotCompare } from '../contract/messages'
import type { PivotAxis } from '../contract/source'
import { loadEntry } from '../entry'
import { axisFields, compareOf, fieldKind, filterType, isBetween, isoOfSerial, serialOfIso, shownItems, withFilter, withHidden, withoutFilters, type FilterKind, type PivotTarget } from '../pivot-model'
import type { PivotDialogs } from './pivot-dialogs'
import type { SheetState } from './state'

/** The open filter menu: its field among those of its axis, where it opens, the search and the items left unchecked. */
export interface PivotFilterMenu {
  t: PivotTarget
  axis: PivotAxis
  fields: number[]
  field: number
  at: { left: number; top: number }
  search: string
  hidden: boolean[]
  /** A Filters area's field: Select Multiple Items, and the one item shown without it (-1: All). */
  multiple: boolean
  page: number
}

export interface CompareDraft {
  compare: PivotCompare
  first: string
  second: string
  value: number
  wholeDays: boolean
}
export interface TopDraft {
  top: boolean
  count: string
  type: 'count' | 'percent' | 'sum'
  value: number
}
export interface SortDraft {
  mode: 'manual' | 'ascending' | 'descending'
  /** -1 sorts by the field's labels, otherwise by that value field's values. */
  ascBy: number
  descBy: number
}

/** A dialog of the filter menu: Label, Value or Date Filter, Top 10 Filter, or Sort. */
export type PivotFilterDialog =
  | { kind: 'compare'; t: PivotTarget; field: number; filter: FilterKind; draft: CompareDraft; error: string }
  | { kind: 'top10'; t: PivotTarget; field: number; draft: TopDraft; error: string }
  | { kind: 'sort'; t: PivotTarget; field: number; draft: SortDraft }

export function usePivotFilter(s: SheetState, dialogs: PivotDialogs) {
  const menu = ref<PivotFilterMenu | null>(null)
  const dialog = ref<PivotFilterDialog | null>(null)
  const m = () => s.messages.value

  /** Opens the menu of `field` on `axis` at a point of the viewport; a field of the axis may be picked in it (Select field). */
  function openMenu(t: PivotTarget, axis: PivotAxis, field: number, at: { left: number; top: number }): void {
    const fields = axis === 'pages' ? [field] : axisFields(t, axis)
    const page = (t.def.table.pages ?? []).find((p) => p.field === field)
    menu.value = { t: markRaw(t), axis, fields: fields.includes(field) ? fields : [field], field, at, search: '', hidden: [], multiple: !!t.def.table.fields[field]?.multipleItems, page: page?.item ?? -1 }
    pick(field)
  }
  /** Select field: the menu over another field of its axis, its items checked as the PivotTable shows them. */
  function pick(field: number): void {
    const x = menu.value
    if (!x) return
    x.field = field
    x.search = ''
    x.hidden = (x.t.def.table.fields[field]?.items ?? []).map((i) => !!i.hidden)
  }
  const closeMenu = () => {
    menu.value = null
  }
  const done = () => {
    closeMenu()
    s.focusGrid()
  }

  /** The items' labels the list shows, with their indexes, those the search leaves. */
  function listed(): { i: number; label: string }[] {
    const x = menu.value
    if (!x) return []
    const q = x.search.trim().toLowerCase()
    return (x.t.cells.labels[x.field] ?? []).map((label, i) => ({ i, label })).filter((y) => !q || y.label.toLowerCase().includes(q))
  }
  /** OK of the list: the items unchecked hidden (with a search, every item it leaves out too); one item stays shown at least. */
  function hiddenOf(x: PivotFilterMenu): boolean[] {
    if (!x.search.trim()) return x.hidden
    const shown = new Set(listed().map((y) => y.i))
    return x.hidden.map((h, i) => h || !shown.has(i))
  }
  const canApply = () => {
    const x = menu.value
    if (!x) return false
    if (x.axis === 'pages' && !x.multiple) return true
    return hiddenOf(x).some((h) => !h)
  }
  async function applyList(): Promise<void> {
    const x = menu.value
    if (!x || !canApply()) return
    // what the search leaves is read before the menu closes
    const hidden = hiddenOf(x)
    done()
    if (x.axis !== 'pages') return void (await dialogs.edit(x.t, withHidden(x.t, x.field, hidden)))
    // a Filters area's field: the one item shown, or with Select Multiple Items those left checked
    const pages = (x.t.def.table.pages ?? []).map((p) => (p.field !== x.field ? p : x.multiple || x.page < 0 ? { field: p.field } : { field: p.field, item: x.page }))
    const items = x.multiple ? withHidden(x.t, x.field, hidden).fields! : [{ field: x.field, items: shownItems(x.t, x.field) }]
    await dialogs.edit(x.t, { pages, fields: [{ ...items[0]!, multipleItems: x.multiple || null }] })
  }

  async function sort(order: 'ascending' | 'descending', t = menu.value?.t, field = menu.value?.field): Promise<void> {
    if (!t || field === undefined) return
    done()
    await dialogs.edit(t, { fields: [{ field, sort: { order } }] })
  }
  async function clear(t = menu.value?.t, field = menu.value?.field): Promise<void> {
    if (!t || field === undefined) return
    done()
    await dialogs.edit(t, withoutFilters(t, field))
  }
  /** A date filter of a fixed period (Today, Quarter 1, January…), applied at once. */
  async function period(type: PivotFilterType): Promise<void> {
    const x = menu.value
    if (!x) return
    done()
    await dialogs.edit(x.t, withFilter(x.t, { field: x.field, type }))
  }

  // ── the dialogs ──
  /** A comparison's dialog, filled from the field's filter of that kind where it has one. */
  async function openCompare(kind: FilterKind, compare: PivotCompare, t = menu.value?.t, field = menu.value?.field): Promise<void> {
    if (!t || field === undefined) return
    closeMenu()
    const had = (t.def.table.filters ?? []).find((f) => f.field === field && compareOf(f.type)?.kind === kind)
    const entry = await loadEntry()
    const locale = dialogs.locale()
    const text = (v: string | number | undefined) => (v === undefined ? '' : kind === 'date' ? entry.entryText(serialOfIso(String(v)), { locale, numberFormat: 14 }) : typeof v === 'number' ? entry.entryText(v, { locale }) : v)
    const same = had && compareOf(had.type)?.compare === compare
    dialog.value = {
      kind: 'compare',
      t: markRaw(t),
      field,
      filter: kind,
      draft: { compare, first: same ? text(had.args?.[0]) : '', second: same ? text(had.args?.[1]) : '', value: had?.value ?? 0, wholeDays: had?.wholeDays !== false },
      error: '',
    }
  }
  function openTop10(t = menu.value?.t, field = menu.value?.field): void {
    if (!t || field === undefined) return
    closeMenu()
    const had = (t.def.table.filters ?? []).find((f) => f.field === field && (f.type === 'count' || f.type === 'percent' || f.type === 'sum'))
    dialog.value = { kind: 'top10', t: markRaw(t), field, draft: { top: !had?.bottom, count: String(had?.args?.[0] ?? 10), type: (had?.type as TopDraft['type']) ?? 'count', value: had?.value ?? 0 }, error: '' }
  }
  function openSort(t = menu.value?.t, field = menu.value?.field): void {
    if (!t || field === undefined) return
    closeMenu()
    const sort = t.def.table.fields[field]?.sort
    const by = sort?.by?.value ?? -1
    dialog.value = { kind: 'sort', t: markRaw(t), field, draft: { mode: sort ? sort.order : 'manual', ascBy: by, descBy: by } }
  }
  const closeDialog = () => {
    dialog.value = null
    s.focusGrid()
  }

  /** OK of a comparison: its boxes read in the workbook's locale (a number, a date), Excel's alert for one that does not read. */
  async function applyCompare(): Promise<void> {
    const d = dialog.value
    if (d?.kind !== 'compare') return
    const x = d.draft
    const entry = await loadEntry()
    const locale = dialogs.locale()
    const boxes = isBetween(x.compare) ? [x.first, x.second] : [x.first]
    const args: (string | number)[] = []
    for (const box of boxes) {
      if (d.filter === 'label') {
        args.push(box)
        continue
      }
      const got = entry.parseEntry(box, { locale })
      const n = got.ok && typeof got.input?.value === 'number' ? got.input.value : null
      if (n === null) return void (d.error = d.filter === 'date' ? m().pivotBadDate : m().pivotBadNumber)
      args.push(d.filter === 'date' ? isoOfSerial(n) : n)
    }
    const filter: PivotFilter = { field: d.field, type: filterType(d.filter, x.compare), args, ...(d.filter === 'value' ? { value: x.value } : {}), ...(d.filter === 'date' && !x.wholeDays ? { wholeDays: false } : {}) }
    closeDialog()
    await dialogs.edit(d.t, withFilter(d.t, filter))
  }
  async function applyTop10(): Promise<void> {
    const d = dialog.value
    if (d?.kind !== 'top10') return
    const entry = await loadEntry()
    const got = entry.parseEntry(d.draft.count, { locale: dialogs.locale() })
    const n = got.ok && typeof got.input?.value === 'number' ? got.input.value : null
    if (n === null || n <= 0) return void (d.error = m().pivotBadNumber)
    const x = d.draft
    closeDialog()
    await dialogs.edit(d.t, withFilter(d.t, { field: d.field, type: x.type, args: [n], value: x.value, ...(x.top ? {} : { bottom: true }) }))
  }
  async function applySort(): Promise<void> {
    const d = dialog.value
    if (d?.kind !== 'sort') return
    const x = d.draft
    const by = x.mode === 'ascending' ? x.ascBy : x.descBy
    closeDialog()
    await dialogs.edit(d.t, { fields: [{ field: d.field, sort: x.mode === 'manual' ? null : { order: x.mode, ...(by >= 0 ? { by: { value: by } } : {}) } }] })
  }

  /** Alt+Down on a cell with a PivotTable's filter button opens its menu, as Excel's; true when handled. */
  function key(e: KeyboardEvent): boolean {
    const a = s.selected.value
    if (e.key !== 'ArrowDown' || !e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || !a || !dialogs.can.value) return false
    const cell = cellAddress(a.row, a.col)
    for (const t of dialogs.views.value) {
      const b = t.cells.buttons.find((x) => x.cell === cell)
      if (!b) continue
      e.preventDefault()
      const r = s.scroller.value?.querySelector<HTMLElement>(`[data-pivot-button="${cell}"]`)?.getBoundingClientRect()
      openMenu(t, b.axis, b.field ?? axisFields(t, b.axis)[0] ?? 0, r ? { left: r.left, top: r.bottom } : { left: 80, top: 80 })
      return true
    }
    return false
  }

  /** The kind of filter a field takes by its values: Date Filters for a field of dates, Label Filters otherwise. */
  const labelKind = (t: PivotTarget, field: number): FilterKind => (fieldKind(t, field) === 'dates' ? 'date' : 'label')

  return { menu, dialog, key, openMenu, pick, closeMenu, listed, canApply, applyList, sort, clear, period, openCompare, openTop10, openSort, closeDialog, applyCompare, applyTop10, applySort, labelKind }
}

export type PivotFilterMenus = ReturnType<typeof usePivotFilter>
