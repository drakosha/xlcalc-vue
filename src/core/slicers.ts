// Excel's slicers and timelines over PivotTables on the sheet: the one selected, a slicer's choice by click, Ctrl+click,
// Shift+click and Multi-Select, Clear Filter, a timeline's periods and level, moving, resizing and removing one, its menu,
// and the dialogs Insert Slicers, Insert Timelines, Slicer Settings and Report Connections. Each change is one operation
// of the source's `structure` (xlcalc's `editSlicer`, `editTimeline` and the others), one step of undo.
import { computed, ref, shallowRef, watch } from 'vue'
import type { DrawingPlace, PivotRef, TimelineLevel } from 'xlcalc/contract'
import { DEFAULT_SLICER_SIZE, DEFAULT_TIMELINE_SIZE } from 'xlcalc/text'
import type { EditOutcome, SheetOp, SheetSlicer, SheetTimeline } from '../contract/source'
import type { SheetObjectRef } from '../contract/view'
import type { Commands } from './commands'
import { EMU_PX, placeOfBox } from './drawing'
import type { Geometry } from './geometry'
import type { MenuItem } from './menu'
import type { Pivots } from './pivots'
import { allows, objectLocks, objectNeeds, type ObjectChange } from '../protection'
import type { SheetState } from './state'
import type { AlertAnswer, SheetAlert } from './validation'

/** The open dialog: Insert Slicers or Insert Timelines over a PivotTable, Slicer Settings, Report Connections. */
export type SlicerDialog =
  | { kind: 'insert'; timeline: boolean; pivot: PivotRef; fields: { name: string; checked: boolean }[] }
  | {
      kind: 'settings'
      slicer: string
      /** Name, which the source refuses empty or taken by another slicer or timeline. */
      name: string
      field: string
      cache: string
      caption: string
      showCaption: boolean
      descending: boolean
      customListSort: boolean
      hide: boolean
      mark: boolean
      last: boolean
      showMissing: boolean
    }
  | { kind: 'connections'; object: SheetObjectRef; pivots: (PivotRef & { checked: boolean })[] }

export interface SlicersDeps {
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
  geo: Geometry
  /** The middle of the cells the active pane shows, in the offsets' space of the grid. */
  middle: () => { x: number; y: number }
}

/** Excel's size of a new slicer and timeline, xlcalc's in EMU, in px. */
const px = (size: { width: number; height: number }) => ({ width: size.width / EMU_PX, height: size.height / EMU_PX })
const NEW_SIZE = { slicer: px(DEFAULT_SLICER_SIZE), timeline: px(DEFAULT_TIMELINE_SIZE) }
/** How far each next slicer or timeline of one insert lies down and to the right of the one before (not measured in Excel). */
const CASCADE_PX = 20

/** A slicer's or timeline's operation by its kind. */
const opOf = (o: SheetObjectRef, edit: { set?: Record<string, unknown>; cache?: Record<string, unknown> }): SheetOp =>
  (o.kind === 'slicer' ? { op: 'editSlicer', slicer: o.name, ...edit } : { op: 'editTimeline', timeline: o.name, ...edit }) as SheetOp

export function useSlicers(s: SheetState, cmd: Commands, pivots: Pivots, deps: SlicersDeps) {
  const m = () => s.messages.value
  const slicers = computed(() => s.sheet.value?.slicers ?? [])
  const timelines = computed(() => s.sheet.value?.timelines ?? [])
  const find = (o: SheetObjectRef | null): SheetSlicer | SheetTimeline | null =>
    !o ? null : ((o.kind === 'slicer' ? slicers.value : timelines.value) as (SheetSlicer | SheetTimeline)[]).find((x) => x.name === o.name) ?? null

  /** The slicer or timeline selected, while the displayed sheet still has it. */
  const picked = shallowRef<SheetObjectRef | null>(null)
  const selected = computed(() => (find(picked.value) ? picked.value : null))
  const select = (o: SheetObjectRef | null) => (picked.value = o)
  watch(() => s.sheetName.value, () => (picked.value = null))
  // a cell made active by any means (a click, the Name Box, Go To, the keys) leaves the object, as in Excel
  watch(
    () => [s.selected.value?.row, s.selected.value?.col, s.extent.value?.row, s.extent.value?.col].join(),
    () => (picked.value = null),
  )
  /** Asks the selected object's box for the keys (after an insert), which the view gives it once drawn. */
  const focusRequest = ref(0)

  /** Multi-Select on, by slicer: a toggle of the slicer's header, which Excel does not keep in the file. */
  const multi = ref<Record<string, boolean>>({})
  /** The item a Shift+click extends from, by slicer: the last one clicked. */
  const anchors = new Map<string, number>()

  // ── what the protection allows: `objectNeeds`, the rule the local source checks, over the sheets' protection ──
  /** Whether every protected sheet a change touches allows it, the sheets of the PivotTables it filters among them. */
  const sheetsAllow = (needs: ReturnType<typeof objectNeeds>) => needs.every((n) => allows(s.sheets.value.find((x) => x.name === n.sheet)?.protection, n.allow))
  function allowed(change: ObjectChange, o: SheetSlicer | SheetTimeline): boolean {
    return sheetsAllow(objectNeeds(change, { sheet: s.sheetName.value, locked: objectLocks(o).sheet }, o.pivots))
  }
  /** A choice, a timeline's level, the slicer's settings and Report Connections: with `structure` and the PivotTables' sheets allowing them; a table's slicer takes none. */
  const canChoose = (o: SheetSlicer | SheetTimeline | null): boolean => !!o && cmd.can('structure') && !('items' in o && o.pivotCache === undefined) && allowed('choose', o)
  /**
   * Moving and resizing: as a choice, not with Disable resizing and moving on, on a protected sheet with Edit objects
   * unless it is unlocked; not while the display filter shows rows in another order than the sheet's.
   */
  const canEdit = (o: SheetSlicer | SheetTimeline | null): boolean => canChoose(o) && !objectLocks(o!).position && !s.visible.value && allowed('place', o!)
  /** Removing one: Edit objects and Use PivotTable and PivotChart. */
  const canRemove = (o: SheetSlicer | SheetTimeline | null): boolean => !!o && cmd.can('structure') && allowed('remove', o)

  const send = (op: SheetOp): Promise<EditOutcome | null> => cmd.structure(op)

  // ── a slicer's choice ──
  /** The items shown: all, or without those with no data where the slicer hides them. */
  const shownItems = (sl: SheetSlicer) => (sl.noData?.hide ? sl.items.filter((x) => !x.noData) : sl.items)
  const filtered = (sl: SheetSlicer) => sl.items.some((x) => x.selected === false)
  /** The choice as `editSlicer` takes it, every item named; one that leaves nothing chosen is Clear Filter. */
  function choice(sl: SheetSlicer, chosen: Set<number>): Promise<EditOutcome | null> {
    const all = chosen.size === 0
    return send(opOf({ kind: 'slicer', name: sl.name }, { cache: { items: sl.items.map((x) => ({ item: x.item, selected: all || chosen.has(x.item) })) } }))
  }
  /**
   * A click on an item: that item alone; with Ctrl or Multi-Select it is added or taken away, with Shift the items from the
   * last one clicked; leaving none chosen clears the filter (not measured in Excel).
   */
  async function choose(sl: SheetSlicer, item: number, mods: { ctrl?: boolean; shift?: boolean } = {}): Promise<EditOutcome | null> {
    if (!canChoose(sl)) return null
    const shown = shownItems(sl)
    const now = new Set(sl.items.filter((x) => x.selected !== false).map((x) => x.item))
    const toggles = !!mods.ctrl || !!multi.value[sl.name]
    let next: Set<number>
    const from = anchors.get(sl.name)
    const a = shown.findIndex((x) => x.item === from)
    const b = shown.findIndex((x) => x.item === item)
    if (mods.shift && a >= 0 && b >= 0) {
      const run = shown.slice(Math.min(a, b), Math.max(a, b) + 1).map((x) => x.item)
      next = toggles && filtered(sl) ? new Set([...now, ...run]) : new Set(run)
    } else {
      anchors.set(sl.name, item)
      if (!toggles) next = new Set([item])
      else {
        next = new Set(now)
        if (next.has(item)) next.delete(item)
        else next.add(item)
      }
    }
    if (next.size === now.size && [...next].every((x) => now.has(x))) return null
    return choice(sl, next)
  }
  /** Clear Filter (Alt+C): every item chosen. */
  async function clear(o: SheetObjectRef): Promise<EditOutcome | null> {
    const x = find(o)
    if (!x || !canChoose(x)) return null
    if ('items' in x) return filtered(x) ? choice(x, new Set()) : null
    return x.selection ? send(opOf(o, { cache: { selection: null } })) : null
  }
  /** Multi-Select (Alt+S): Excel's toggle in the slicer's header. */
  function toggleMulti(name: string): void {
    multi.value = { ...multi.value, [name]: !multi.value[name] }
  }

  // ── a timeline ──
  /** The dates of a run of periods chosen, from the first day of the first to the last day of the last. */
  async function period(t: SheetTimeline, start: string, end: string): Promise<EditOutcome | null> {
    if (!canChoose(t) || (t.selection?.start === start && t.selection.end === end)) return null
    return send(opOf({ kind: 'timeline', name: t.name }, { cache: { selection: { start, end } } }))
  }
  /** The level the timeline shows, which does not change its dates. */
  async function level(t: SheetTimeline, to: TimelineLevel): Promise<EditOutcome | null> {
    if (!canChoose(t) || (t.level ?? 'months') === to) return null
    return send(opOf({ kind: 'timeline', name: t.name }, { set: { level: to } }))
  }

  // ── the object ──
  /** A slicer or a timeline moved or resized to a new place. */
  async function place(o: SheetObjectRef, to: DrawingPlace): Promise<EditOutcome | null> {
    return canEdit(find(o)) ? send(opOf(o, { set: { place: to } })) : null
  }
  /** Removes it (Delete, the menu's Remove), as one step of undo. */
  async function remove(o: SheetObjectRef): Promise<EditOutcome | null> {
    if (!canRemove(find(o))) return null
    const out = await send(o.kind === 'slicer' ? { op: 'deleteSlicer', slicer: o.name } : { op: 'deleteTimeline', timeline: o.name })
    if (out?.ok) {
      picked.value = null
      s.focusGrid()
    }
    return out
  }
  /** The menu's Refresh: the PivotTables it filters, refreshed from their source. */
  async function refresh(o: SheetObjectRef): Promise<EditOutcome | null> {
    const x = find(o)
    return x && x.pivots.length && canChoose(x) ? send({ op: 'refreshPivots', pivots: x.pivots }) : null
  }
  /** The menu's Sort A to Z and Sort Z to A. */
  async function sort(sl: SheetSlicer, dir: 'ascending' | 'descending'): Promise<EditOutcome | null> {
    return canChoose(sl) && (sl.sort ?? 'ascending') !== dir ? send(opOf({ kind: 'slicer', name: sl.name }, { cache: { sort: dir } })) : null
  }

  /** The menu of the selected slicer or timeline, with Excel's items the sheet has. */
  /**
   * The items of Excel for Windows' menu of a slicer the sheet has, in its order and groups (measured): Refresh; Sort A to Z,
   * Sort Z to A; Clear Filter from, Multi-Select, Report Connections…, Remove; Slicer Settings…. A timeline's has those of
   * them it can have, in the same order (not measured). Remove names the slicer by its name, as Microsoft's page says.
   */
  function items(): MenuItem[] {
    const o = selected.value
    const x = find(o)
    if (!o || !x) return []
    const caption = x.caption ?? x.name
    const slicer = 'items' in x ? x : null
    const out: MenuItem[] = [{ id: 'objectRefresh', label: m().menuRefresh, disabled: !x.pivots.length || !canChoose(x), run: () => void refresh(o) }]
    if (slicer)
      out.push(
        { id: 'slicerSortAZ', label: m().menuSortAZ, disabled: !canChoose(slicer), run: () => void sort(slicer, 'ascending'), group: true },
        { id: 'slicerSortZA', label: m().menuSortZA, disabled: !canChoose(slicer), run: () => void sort(slicer, 'descending') },
      )
    out.push({ id: 'objectClear', label: m().menuClearFilterFrom(caption), disabled: !(slicer ? filtered(slicer) : 'selection' in x && x.selection) || !canChoose(x), run: () => void clear(o), group: true })
    if (slicer) out.push({ id: 'slicerMulti', label: m().menuMultiSelect(x.name), run: () => toggleMulti(x.name) })
    out.push(
      { id: 'reportConnections', label: m().menuReportConnections, disabled: !canChoose(x), run: () => void openConnections() },
      { id: 'objectRemove', label: m().menuRemoveObject(x.name), disabled: !canRemove(x), run: () => void remove(o) },
    )
    if (slicer) out.push({ id: 'slicerSettings', label: m().menuSlicerSettings, disabled: !canChoose(slicer), run: () => void openSettings(), group: true })
    return out
  }

  // ── the dialogs ──
  const dialog = ref<SlicerDialog | null>(null)
  /** Insert Slicer and Insert Timeline: the active cell in a PivotTable, with `structure`, on a sheet that allows objects. */
  const canInsert = computed(() => {
    const p = pivots.active.value
    const at = { sheet: s.sheetName.value, locked: true }
    return !!p && cmd.can('structure') && sheetsAllow(objectNeeds('add', at, [at]))
  })
  /** Opens Insert Slicers or Insert Timelines over the PivotTable under the active cell; without a field of dates, Excel's alert. */
  function openInsert(timeline: boolean): boolean {
    const p = pivots.active.value
    if (!p || !canInsert.value) return false
    const fields = (p.fields ?? []).filter((f) => !timeline || f.dates)
    if (timeline && !fields.length) {
      void deps.ask({ severity: 'notice', message: m().timelineNoDates }).then(() => s.focusGrid())
      return true
    }
    dialog.value = { kind: 'insert', timeline, pivot: { sheet: s.sheetName.value, name: p.name }, fields: fields.map((f) => ({ name: f.name, checked: false })) }
    return true
  }
  /**
   * An insert under way: the names there were before it and the field of the last slicer or timeline it adds, which is
   * selected once the sheets are read again with it, and its box focused, as Excel leaves the last one inserted.
   */
  let inserting: { kind: SheetObjectRef['kind']; before: Set<string>; field: string } | null = null
  const namesOf = (kind: SheetObjectRef['kind']) => new Set((kind === 'timeline' ? timelines.value : slicers.value).map((x) => x.name))
  function pickInserted(): void {
    const i = inserting
    const added = i && (i.kind === 'timeline' ? timelines.value : slicers.value).filter((x) => !i.before.has(x.name) && x.field === i.field).at(-1)
    if (!i || !added) return
    inserting = null
    picked.value = { kind: i.kind, name: added.name }
    focusRequest.value++
  }
  watch([slicers, timelines], pickInserted)
  /**
   * Where Excel puts a new slicer or timeline: at the middle of the cells the active pane shows, in Excel's size, and each
   * next one of the same insert `CASCADE_PX` lower and to the right. The library cannot know the window, so the sheet sends it.
   */
  function placeOf(kind: SheetObjectRef['kind'], i: number) {
    const size = NEW_SIZE[kind]
    const mid = deps.middle()
    const box = { x: Math.max(0, mid.x - size.width / 2) + i * CASCADE_PX, y: Math.max(0, mid.y - size.height / 2) + i * CASCADE_PX, ...size }
    return placeOfBox(box, deps.geo, { rowOf: s.rowOf, indexOf: s.indexOf })
  }
  /** OK in Insert Slicers or Insert Timelines: a slicer or timeline of each field checked, the last one selected. */
  async function insert(): Promise<void> {
    const d = dialog.value
    if (d?.kind !== 'insert') return
    dialog.value = null
    const sheet = s.sheetName.value
    const kind = d.timeline ? 'timeline' : 'slicer'
    const fields = d.fields.filter((x) => x.checked).map((x) => x.name)
    if (!fields.length) return s.focusGrid()
    inserting = { kind, before: namesOf(kind), field: fields.at(-1)! }
    let last: string | null = null
    for (const [i, field] of fields.entries()) {
      const place = placeOf(kind, i)
      const out = await send(d.timeline ? { op: 'addTimeline', sheet, timeline: { place }, cache: { field, pivots: [d.pivot] } } : { op: 'addSlicer', sheet, slicer: { place }, cache: { field, pivots: [d.pivot] } })
      if (!out?.ok) break
      last = field
    }
    // a refused insert selects the last one added, or none
    if (inserting && last !== inserting.field) inserting = last === null ? null : { ...inserting, field: last }
    if (inserting) pickInserted()
    else if (last === null) s.focusGrid()
  }
  /** Opens Slicer Settings for the selected slicer. */
  function openSettings(): boolean {
    const o = selected.value
    const x = find(o)
    if (!x || !('items' in x) || !canChoose(x)) return false
    dialog.value = {
      kind: 'settings',
      slicer: x.name,
      name: x.name,
      field: x.field,
      cache: x.cache,
      caption: x.caption ?? x.name,
      showCaption: x.showCaption !== false,
      descending: x.sort === 'descending',
      customListSort: x.customListSort !== false,
      hide: !!x.noData?.hide,
      mark: x.noData?.mark !== false,
      last: x.noData?.last !== false,
      showMissing: x.showMissing !== false,
    }
    return true
  }
  /** OK in Slicer Settings: the name, the caption and header, the sorting and the items with no data, in one operation; a new name keeps it selected. */
  async function applySettings(): Promise<void> {
    const d = dialog.value
    // an empty name is never sent: OK waits for one, as Excel takes no slicer without a name (its alert not measured)
    if (d?.kind !== 'settings' || !d.name.trim()) return
    dialog.value = null
    const o = { kind: 'slicer' as const, name: d.slicer }
    const name = d.name !== d.slicer ? { name: d.name } : {}
    const out = await send(opOf(o, { set: { ...name, caption: d.caption, showCaption: d.showCaption }, cache: { sort: d.descending ? 'descending' : 'ascending', customListSort: d.customListSort, noData: { hide: d.hide, mark: d.mark, last: d.last }, showMissing: d.showMissing } }))
    if (out?.ok && 'name' in name && picked.value?.kind === 'slicer' && picked.value.name === d.slicer) picked.value = { kind: 'slicer', name: d.name }
  }
  /** Opens Report Connections: every PivotTable of the workbook over the same pivot cache, those it filters checked. */
  function openConnections(): boolean {
    const o = selected.value
    const x = find(o)
    if (!o || !x || x.pivotCache === undefined || !canChoose(x)) return false
    const on = (p: PivotRef) => x.pivots.some((q) => q.sheet === p.sheet && q.name === p.name)
    const list = s.sheets.value.flatMap((sh) => (sh.pivots ?? []).filter((p) => p.cache === x.pivotCache).map((p) => ({ sheet: sh.name, name: p.name })))
    dialog.value = { kind: 'connections', object: o, pivots: list.map((p) => ({ ...p, checked: on(p) })) }
    return true
  }
  /** OK in Report Connections: the PivotTables checked, unchanged when they are the same. */
  async function applyConnections(): Promise<void> {
    const d = dialog.value
    if (d?.kind !== 'connections') return
    dialog.value = null
    const x = find(d.object)
    const next = d.pivots.filter((p) => p.checked).map(({ sheet, name }) => ({ sheet, name }))
    if (!x || (next.length === x.pivots.length && next.every((p) => x.pivots.some((q) => q.sheet === p.sheet && q.name === p.name)))) return
    await send(opOf(d.object, { cache: { pivots: next } }))
  }
  function close(): void {
    dialog.value = null
    s.focusGrid()
  }

  return { slicers, timelines, find, selected, select, focusRequest, multi, toggleMulti, canEdit, canChoose, canRemove, shownItems, filtered, choose, clear, period, level, place, remove, refresh, sort, items, dialog, canInsert, openInsert, insert, openSettings, applySettings, openConnections, applyConnections, close }
}

export type Slicers = ReturnType<typeof useSlicers>
