// The cell menu of a PivotTable's cell, as Excel for Windows lists it after Refresh: Number Format… on a value, Sort,
// Filter, Subtotal "Region", Group…, Ungroup…, Remove, Summarize Values By and Show Values As, Value Field Settings…,
// Field Settings… and PivotTable Options…; and the PivotTable Fields pane's requests for a field's dialogs and menu.
import type { PivotDialogRequest } from '../contract/view'
import { axisFields, fieldKind, fieldName, filtered, PIVOT_FUNCTIONS, valueName, withHidden } from '../pivot-model'
import type { MenuItem } from './menu'
import { SHOW_AS, type PivotDialogs } from './pivot-dialogs'
import type { PivotFilterMenus } from './pivot-filter'
import type { SheetState } from './state'

/** Where a menu opens without a control to open it at: near the top left of the sheet's grid. */
const AT_GRID = { left: 80, top: 80 }

export function usePivotMenu(s: SheetState, dialogs: PivotDialogs, filter: PivotFilterMenus) {
  const m = () => s.messages.value

  /** The menu's items over the cell under the active cell, before and after Refresh; none while it is read or without the commands. */
  function items(refresh: MenuItem[]): MenuItem[] {
    const here = dialogs.here.value
    if (!here || !dialogs.can.value) return refresh
    const { t, part } = here
    const x = m()
    const out: MenuItem[] = []
    const value = part.value
    const isValue = (part.kind === 'value' || part.kind === 'valueHeader') && value !== undefined && !!t.def.table.values?.[value]
    if (isValue) out.push({ id: 'pivotNumberFormat', label: x.pivotMenuNumberFormat, run: () => dialogs.openNumberFormat(t, value), group: true })
    out.push(...refresh)
    if (refresh.length) out.push({ id: 'pivotDelete', label: x.pivotMenuDelete, run: () => void dialogs.send({ op: 'deletePivotTable', sheet: t.sheet, pivot: t.pivot }) })
    const field = part.field
    const onAxis = field !== undefined && (part.axis === 'rows' || part.axis === 'columns')
    if (onAxis) {
      const dates = fieldKind(t, field) === 'dates'
      out.push({
        id: 'pivotSort',
        label: x.pivotMenuSort,
        group: true,
        run: () => {},
        sub: [
          { id: 'pivotSortAsc', label: dates ? x.pivotSortOldest : x.pivotSortAsc, run: () => void filter.sort('ascending', t, field) },
          { id: 'pivotSortDesc', label: dates ? x.pivotSortNewest : x.pivotSortDesc, run: () => void filter.sort('descending', t, field) },
          { id: 'pivotMoreSort', label: x.pivotMoreSort, run: () => filter.openSort(t, field), group: true },
        ],
      })
      out.push({
        id: 'pivotFilter',
        label: x.pivotMenuFilter,
        run: () => {},
        sub: [
          { id: 'pivotClear', label: x.pivotClearFrom(fieldName(t, field)), disabled: !filtered(t, field), run: () => void filter.clear(t, field) },
          { id: 'pivotKeep', label: x.pivotKeepSelected, disabled: part.kind !== 'item', run: () => void keep(true), group: true },
          { id: 'pivotHide', label: x.pivotHideSelected, disabled: part.kind !== 'item', run: () => void keep(false) },
          { id: 'pivotTop10', label: x.pivotTop10Menu, disabled: !t.def.table.values?.length, run: () => filter.openTop10(t, field), group: true },
          { id: 'pivotLabelFilters', label: x.pivotLabelFiltersMenu, run: () => void filter.openCompare(filter.labelKind(t, field), 'equal', t, field) },
          { id: 'pivotValueFilters', label: x.pivotValueFiltersMenu, disabled: !t.def.table.values?.length, run: () => void filter.openCompare('value', 'equal', t, field) },
        ],
      })
      const subs = t.def.table.fields[field]?.subtotals
      out.push({ id: 'pivotSubtotal', label: x.pivotMenuSubtotal(fieldName(t, field)), checked: subs?.length !== 0, run: () => void dialogs.edit(t, { fields: [{ field, subtotals: subs?.length === 0 ? null : [] }] }), group: true })
      const kind = fieldKind(t, field)
      out.push({ id: 'pivotGroup', label: x.pivotMenuGroup, run: () => void dialogs.group(), group: true })
      const g = t.def.cache.fields[field]?.group
      const grouped = !!g || t.def.cache.fields.some((f) => f.group?.base === field)
      out.push({ id: 'pivotUngroup', label: x.pivotMenuUngroup, disabled: !grouped && kind === 'text', run: () => void dialogs.ungroup() })
    }
    if (field !== undefined && part.axis) {
      out.push({ id: 'pivotRemove', label: x.pivotMenuRemove(fieldName(t, field)), run: () => void remove(field), group: true })
    }
    if (isValue) {
      const v = t.def.table.values![value]!
      out.push({ id: 'pivotRemoveValue', label: x.pivotMenuRemove(valueName(t, value)), run: () => void dialogs.edit(t, { values: t.def.table.values!.filter((_, i) => i !== value) }), group: true })
      out.push({
        id: 'pivotSummarize',
        label: x.pivotMenuSummarize,
        run: () => {},
        group: true,
        sub: [
          ...PIVOT_FUNCTIONS.slice(0, 6).map((fn) => ({ id: `pivotFn-${fn}`, label: x.pivotFunction(fn), checked: (v.function ?? 'sum') === fn, run: () => void dialogs.setFunction(t, value, fn) })),
          { id: 'pivotSummarizeMore', label: x.pivotMenuSummarizeMore, run: () => void dialogs.openValue(t, value, 'summarize'), group: true },
        ],
      })
      out.push({
        id: 'pivotShowAs',
        label: x.pivotMenuShowAs,
        run: () => {},
        sub: [
          ...SHOW_AS.map((calc) => ({ id: `pivotShowAs-${calc}`, label: x.pivotShowAsMenu(calc), checked: (v.showAs ?? 'normal') === calc, run: () => void dialogs.setShowAs(t, value, calc) })),
          { id: 'pivotShowAsMore', label: x.pivotMenuShowAsMore, run: () => void dialogs.openValue(t, value, 'showAs'), group: true },
        ],
      })
      out.push({ id: 'pivotValueSettings', label: x.valueFieldSettings, run: () => void dialogs.openValue(t, value), group: true })
    } else if (field !== undefined && part.axis) {
      out.push({ id: 'pivotFieldSettings', label: x.fieldSettings, run: () => void dialogs.openField(t, field), group: true })
    }
    out.push({ id: 'pivotOptions', label: x.pivotMenuOptions, run: () => dialogs.openOptions(t), group: !isValue && !(field !== undefined && part.axis) })
    return out
  }

  /** Remove "Region": the field off its axis, as Excel's item does. */
  async function remove(field: number): Promise<void> {
    const here = dialogs.here.value
    if (!here) return
    const tb = here.t.def.table
    const off = (list: typeof tb.rows) => list?.filter((f) => f !== field)
    await dialogs.edit(here.t, {
      ...(tb.rows?.includes(field) ? { rows: off(tb.rows)!.length ? off(tb.rows)! : null } : {}),
      ...(tb.columns?.includes(field) ? { columns: off(tb.columns)!.length ? off(tb.columns)! : null } : {}),
      ...(tb.pages?.some((p) => p.field === field) ? { pages: tb.pages.filter((p) => p.field !== field).length ? tb.pages.filter((p) => p.field !== field) : null } : {}),
    })
  }

  /** Keep Only Selected Items or Hide Selected Items: the items of the active cell's field among the cells selected. */
  async function keep(only: boolean): Promise<void> {
    const here = dialogs.here.value
    const got = await dialogs.picked()
    const field = here?.part.field
    if (!got || field === undefined) return
    const chosen = new Set(got.parts.filter((p) => p.kind === 'item' && p.field === field && p.item !== undefined).map((p) => p.item!))
    const items = got.t.def.table.fields[field]?.items ?? []
    const hidden = items.map((it, i) => (only ? !chosen.has(i) : !!it.hidden || chosen.has(i)))
    if (hidden.every(Boolean)) return
    await dialogs.edit(got.t, withHidden(got.t, field, hidden))
  }

  /** The PivotTable Fields pane's request: Field Settings, Value Field Settings, or the field's filter menu at its arrow. */
  async function request(r: PivotDialogRequest): Promise<boolean> {
    const got = await dialogs.target(r.sheet, r.pivot)
    if (!got) return false
    const t = got.t
    if (r.dialog === 'valueFieldSettings') return dialogs.openValue(t, r.value ?? (t.def.table.values ?? []).findIndex((v) => v.field === r.field))
    if (r.dialog === 'fieldSettings') return dialogs.openField(t, r.field)
    const axis = axisFields(t, 'rows').includes(r.field) ? 'rows' : axisFields(t, 'columns').includes(r.field) ? 'columns' : axisFields(t, 'pages').includes(r.field) ? 'pages' : 'rows'
    filter.openMenu(t, axis, r.field, r.at ? { left: r.at.left, top: r.at.top + r.at.height } : AT_GRID)
    return true
  }
  /** Whether the sheet opens a request itself; one it does not know goes on to the host. */
  const handles = (r: PivotDialogRequest) => r.dialog === 'fieldSettings' || r.dialog === 'valueFieldSettings' || r.dialog === 'fieldFilter'

  return { items, request, handles }
}

export type PivotMenu = ReturnType<typeof usePivotMenu>
