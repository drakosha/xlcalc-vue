// PivotTables in the local source: xlcalc's PivotTable engine (`xlcalc/pivot`), loaded as its own chunk for a workbook that
// has PivotTables, since without them nothing needs it; their areas as xlcalc gives them; Excel's alerts for a change
// the library refuses over their cells; and the PivotTables an operation touches, which a protected sheet gates.
import { registerPivot, type BookInput, type BookLayout, type CellRect, type PivotAreas, type PivotCache, type PivotRef, type PivotRefusalReason, type PivotTable, type SheetLayout, type StructureOp } from 'xlcalc'
import { parseRange, rectsMeet, sameName } from 'xlcalc/a1'
import type { SheetPivot, SheetPivotTable } from '../contract/source'
import { pivotFields } from './slicers'
import { objectLocks, objectNeeds } from '../protection'

/** Whether the workbook holds PivotTables or their caches, which the engine refreshes and lays out. */
export const hasPivots = (book: BookInput | BookLayout): boolean => !!book.pivotCaches?.length || book.sheets.some((s) => !!s.pivotTables?.length)

let loading: Promise<void> | null = null
/** Loads xlcalc's PivotTable engine once for the page; a failed load is tried again by the next call. */
export function loadPivot(): Promise<void> {
  loading ??= import('xlcalc/pivot').then(
    (m) => registerPivot(m.pivotEngine),
    (e: unknown) => {
      loading = null
      throw e
    },
  )
  return loading
}

/** A PivotTable's areas as xlcalc gives them with its place (`PivotLocation.areas`, on every PivotTable a layout holds). */
const areasOf = (t: PivotTable): PivotAreas => t.location.areas ?? { table: t.location.range, data: t.location.range }

/** A sheet's PivotTables as `SheetInfo.pivots` lists them, with their cache's id and fields; none for a sheet without them. */
export function pivotsOf(sheet: SheetLayout, caches: PivotCache[] = []): SheetPivot[] {
  return (sheet.pivotTables ?? []).map((t) => {
    const a = areasOf(t)
    return {
      name: t.name,
      range: a.table,
      ...(a.page ? { filters: a.page } : {}),
      ...(a.columnHeaders ? { headers: a.columnHeaders } : {}),
      ...(a.rowLabels ? { rowLabels: a.rowLabels } : {}),
      values: a.data,
      cache: t.cache,
      fields: pivotFields(caches.find((c) => c.id === t.cache)),
    }
  })
}

/** The PivotTable `name` on `sheet` without its saved lines, with its cache without its rows (`SheetSource.pivotTable`); null when there is none. */
export function pivotDefinition(layout: BookLayout, sheet: string, name: string): SheetPivotTable | null {
  const found = layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables?.find((t) => sameName(t.name, name))
  const cache = found && layout.pivotCaches?.find((c) => c.id === found.cache)
  if (!found || !cache) return null
  const { rowLines: _rows, columnLines: _columns, ...table } = found
  const { records: _records, ...fields } = cache
  return { table, cache: fields }
}

/** The PivotTable whose table or Filters area meets `rect` on `sheet`; null when none does. */
export function pivotMeeting(layout: BookLayout, sheet: string, rect: CellRect): PivotTable | null {
  for (const t of layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables ?? []) {
    const a = areasOf(t)
    const meets = (range: string | undefined) => {
      const x = range ? parseRange(range) : null
      return !!x && rectsMeet(x, rect)
    }
    if (meets(a.table) || meets(a.page)) return t
  }
  return null
}

/**
 * Excel for Windows' alerts over a PivotTable's cells, as xlcalc's refusal `reason` names them: a value typed (`part`), a
 * formula over its cells (`formula`), a cleared cell, a paste or an operation that would insert or delete cells in it
 * (`affect`), in Excel 16.0 for Windows' words in en-US and ru-RU (Russian `переместить` is Excel's). `engine` is the
 * source's own: the PivotTable engine could not be loaded, so nothing needing it is done; so is `moveLocked`: a slicer or
 * a timeline with Disable resizing and moving on, which Excel's interface does not move.
 */
const ALERTS = {
  en: {
    part: "We can't change this part of the PivotTable.",
    affect: "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.",
    formula: 'Cannot enter a formula for an item or field name in a PivotTable report.',
    engine: "PivotTables can't be refreshed or changed now: this page could not load what they need. Reload the page and try again.",
    moveLocked: "This slicer or timeline can't be moved or resized: Disable resizing and moving is on.",
  },
  ru: {
    part: 'Мы не можем изменить эту часть сводной таблицы.',
    affect: 'Мы не можем изменить выделенные ячейки, так как это повлияет на сводную таблицу. Для изменения отчета используйте список полей. Если вы пытаетесь вставить или удалить ячейки, переместить сводную таблицу и повторите попытку.',
    formula: 'Нельзя задавать формулу для имени поля или элемента в отчете сводной таблицы.',
    engine: 'Сводные таблицы сейчас нельзя обновить или изменить: страница не смогла загрузить то, что для этого нужно. Обновите страницу и повторите попытку.',
    moveLocked: 'Этот срез или эту временную шкалу нельзя переместить или изменить в размере: для них отключено изменение размеров и перемещение.',
  },
}
export type PivotAlert = keyof (typeof ALERTS)['en']

/** The outcome code of a change a PivotTable's cells do not take, as xlcalc's refusals name it. */
export const PIVOT = 'pivot'

/**
 * xlcalc's refusal over PivotTables, or its `PivotInputError`, as the outcome the sheet shows in an alert: Excel's alert for
 * its `reason`, else the library's own message (a name taken, an item a field does not have).
 */
export function pivotRefusals(locale: string) {
  const words = locale.toLowerCase().startsWith('ru') ? ALERTS.ru : ALERTS.en
  const alert = (kind: PivotAlert) => ({ ok: false as const, reason: 'refused' as const, code: PIVOT, message: words[kind] })
  const refusal = (r: { message: string; reason?: PivotRefusalReason }) => (r.reason ? alert(r.reason) : { ok: false as const, reason: 'refused' as const, code: PIVOT, message: r.message })
  return { alert, refusal }
}

/** The operations of xlcalc on PivotTables themselves, whose refusals keep the library's reason; any other refused with `pivot` would cut one. */
const PIVOT_OPS = new Set<StructureOp['op']>(['addPivotTable', 'editPivotTable', 'renamePivotTable', 'movePivotTable', 'deletePivotTable', 'refreshPivots', 'groupPivotField', 'ungroupPivotField', 'renamePivotGroup', 'setPivotCalculatedField', 'setPivotCalculatedItem', 'addSlicer', 'editSlicer', 'deleteSlicer', 'addTimeline', 'editTimeline', 'deleteTimeline'])
export const isPivotOp = (op: StructureOp): boolean => PIVOT_OPS.has(op.op)

/** What a protected sheet must allow for an operation on PivotTables: a box of Protect Sheet on a sheet, or `refresh`, which none allows. */
export interface PivotNeed {
  sheet: string
  allow: 'pivotTables' | 'objects' | 'refresh'
}

/**
 * What an operation on PivotTables needs of protected sheets. A refresh is refused on a protected sheet with a PivotTable
 * over the cache (every sheet with one for Refresh All), whatever its boxes allow, as Excel 16.0 for Windows refused Refresh
 * All there with and without Use PivotTable and PivotChart (`refresh`). The rest by Microsoft's Protect a worksheet (not
 * measured in Excel): Use PivotTable reports ("Format, change the layout, refresh, or otherwise modify PivotTable reports, or
 * create new reports") on each sheet whose PivotTables it lays out again, every sheet with a PivotTable over the cache for a
 * change of the cache; a slicer's or a timeline's choice refilters the PivotTables it names; adding or deleting a slicer or
 * a timeline also needs Edit objects on its sheet, as slicers and timelines are objects locked when added.
 */
export function pivotNeeds(layout: BookLayout, op: StructureOp): PivotNeed[] {
  const cacheOf = (sheet: string, name: string) => layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables?.find((t) => sameName(t.name, name))?.cache
  const over = (caches: (number | undefined)[]) => layout.sheets.filter((x) => x.pivotTables?.some((t) => caches.includes(t.cache))).map((x) => x.name)
  const pivots = (sheets: string[]): PivotNeed[] => sheets.map((sheet) => ({ sheet, allow: 'pivotTables' }))
  const filtered = (refs: readonly PivotRef[] | undefined) => (refs ?? []).map((p) => p.sheet)
  /** A slicer or a timeline: its sheet, whether it is locked, and the PivotTables its cache filters. */
  const objectOf = (kind: 'slicers' | 'timelines', name: string) => {
    const sheet = layout.sheets.find((x) => x[kind]?.some((o) => sameName(o.name, name)))
    const o = sheet?.[kind]?.find((x) => sameName(x.name, name)) as { cache: string; locked?: boolean } | undefined
    const caches: readonly { name: string; pivots: PivotRef[] }[] = (kind === 'slicers' ? layout.slicerCaches : layout.timelineCaches) ?? []
    const at = { sheet: sheet?.name ?? '', locked: objectLocks(o ?? {}).sheet }
    return { at, found: !!sheet, pivots: caches.find((c) => sameName(c.name, o?.cache ?? ''))?.pivots ?? [] }
  }
  const cachePivots = (kind: 'slicers' | 'timelines', cache: string | { pivots: PivotRef[] }) =>
    (typeof cache === 'string' ? ((kind === 'slicers' ? layout.slicerCaches : layout.timelineCaches) ?? []).find((c) => sameName(c.name, cache))?.pivots : cache.pivots) ?? []
  switch (op.op) {
    case 'refreshPivots':
      return (op.pivots ? over(op.pivots.map((p) => cacheOf(p.sheet, p.name))) : layout.sheets.filter((x) => x.pivotTables?.length).map((x) => x.name)).map((sheet) => ({ sheet, allow: 'refresh' }))
    case 'groupPivotField':
    case 'ungroupPivotField':
    case 'renamePivotGroup':
    case 'setPivotCalculatedField':
    case 'setPivotCalculatedItem':
      return pivots([op.sheet, ...over([cacheOf(op.sheet, op.pivot)])])
    case 'movePivotTable':
      return pivots([op.sheet, op.to.sheet])
    case 'addPivotTable':
      // a new sheet is not protected
      return op.newSheet ? [] : pivots([op.sheet])
    case 'editPivotTable':
    case 'renamePivotTable':
    case 'deletePivotTable':
      return pivots([op.sheet])
    // a slicer's or a timeline's change: the rule the sheet asks too (`objectNeeds`)
    case 'addSlicer':
    case 'addTimeline': {
      return objectNeeds('add', { sheet: op.sheet, locked: true }, cachePivots(op.op === 'addSlicer' ? 'slicers' : 'timelines', op.cache))
    }
    case 'editSlicer':
    case 'editTimeline': {
      const o = op.op === 'editSlicer' ? objectOf('slicers', op.slicer) : objectOf('timelines', op.timeline)
      if (!o.found) return []
      // Report Connections may name more PivotTables, which take its choice
      const pivots = [...o.pivots, ...((op.cache as { pivots?: PivotRef[] } | undefined)?.pivots ?? [])]
      return objectNeeds(op.set?.place ? 'place' : 'choose', o.at, pivots)
    }
    case 'deleteSlicer':
    case 'deleteTimeline': {
      const o = op.op === 'deleteSlicer' ? objectOf('slicers', op.slicer) : objectOf('timelines', op.timeline)
      return o.found ? objectNeeds('remove', o.at, o.pivots) : []
    }
    default:
      return []
  }
}

/** Whether an operation moves or resizes a slicer or a timeline whose position is locked (Disable resizing and moving), which nothing moves. */
export function movesLocked(layout: BookLayout, op: StructureOp): boolean {
  if ((op.op !== 'editSlicer' && op.op !== 'editTimeline') || !op.set?.place) return false
  const kind = op.op === 'editSlicer' ? 'slicers' : 'timelines'
  const name = op.op === 'editSlicer' ? op.slicer : op.timeline
  for (const sheet of layout.sheets) {
    const o = (sheet[kind] as { name: string; locked?: boolean }[] | undefined)?.find((x) => sameName(x.name, name))
    if (o) return objectLocks(o).position
  }
  return false
}
