// PivotTables in the local source: xlcalc's PivotTable engine (`xlcalc/pivot`), loaded as its own chunk for a workbook that
// has PivotTables, since without them nothing needs it; their areas as the sheet lists them; Excel's alerts for a change
// the library refuses over their cells; and the PivotTables an operation touches, which a protected sheet gates.
import { LOCALES, registerPivot, type BookInput, type BookLayout, type CellRect, type PivotCache, type PivotRef, type PivotTable, type SheetLayout, type StructureOp } from 'xlcalc'
import { parseRange, rangeAddress, rectsMeet, sameName } from 'xlcalc/a1'
import type { SheetPivot, SheetPivotTable } from '../contract/source'
import { pivotFields } from './slicers'

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

/**
 * A PivotTable's areas from xlcalc's `PivotLocation`, the one place they are worked out (xlcalc is to give them ready-made):
 * the filters take `pageRows` rows above the table with a blank row between, each column of filters a field's name and its
 * item and a gap (`pageColumns`), as xlcalc lays them out; the table's headers end above `firstDataRow`, its values start at
 * `firstDataColumn`. A PivotTable not laid out (no saved lines: added without the engine) or shrunk to its top-left cell
 * (`#SPILL!`, no header row) has its table alone.
 */
function areasOf(t: PivotTable): { table: CellRect; filters?: CellRect; headers?: CellRect; rowLabels?: CellRect; values?: CellRect } | null {
  const x = parseRange(t.location.range)
  if (!x) return null
  const { firstDataRow: dataRow, firstDataColumn: dataCol, pageRows = 0, pageColumns } = t.location
  const top = x.top + dataRow
  const left = x.left + dataCol
  const laid = !!t.rowLines && dataRow > 0 && top <= x.bottom
  return {
    table: x,
    ...(pageRows > 0 && x.top - pageRows - 1 >= 1 ? { filters: { top: x.top - pageRows - 1, bottom: x.top - 2, left: x.left, right: x.left + 3 * Math.max(1, pageColumns ?? 1) - 2 } } : {}),
    ...(laid ? { headers: { ...x, bottom: top - 1 } } : {}),
    ...(laid && dataCol > 0 ? { rowLabels: { ...x, top, right: Math.min(x.right, left - 1) } } : {}),
    ...(laid && t.values?.length && left <= x.right ? { values: { ...x, top, left } } : {}),
  }
}

/** A sheet's PivotTables as `SheetInfo.pivots` lists them, with their cache's id and fields; none for a sheet without them. */
export function pivotsOf(sheet: SheetLayout, caches: PivotCache[] = []): SheetPivot[] {
  const out: SheetPivot[] = []
  for (const t of sheet.pivotTables ?? []) {
    const a = areasOf(t)
    if (!a) continue
    const range = (r: CellRect | undefined) => r && rangeAddress(r)
    out.push({
      name: t.name,
      range: rangeAddress(a.table),
      ...(a.filters ? { filters: range(a.filters) } : {}),
      ...(a.headers ? { headers: range(a.headers) } : {}),
      ...(a.rowLabels ? { rowLabels: range(a.rowLabels) } : {}),
      ...(a.values ? { values: range(a.values) } : {}),
      cache: t.cache,
      fields: pivotFields(caches.find((c) => c.id === t.cache)),
    })
  }
  return out
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
    if (a && (rectsMeet(a.table, rect) || (a.filters && rectsMeet(a.filters, rect)))) return t
  }
  return null
}

/**
 * Excel for Windows' alerts over a PivotTable's cells, as xlcalc's refusal `reason` names them: a value typed or pasted
 * (`part`), a formula over a label (`formula`), a cleared cell or an operation that would insert or delete cells in it
 * (`affect`); the English ones are those Excel 16.0 for Windows answered xlcalc's set `edit` with, the Russian ones were
 * not measured. `engine` is the source's own: the PivotTable engine could not be loaded, so nothing needing it is done.
 */
const ALERTS = {
  en: {
    part: "We can't change this part of the PivotTable.",
    affect: "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.",
    formula: 'Cannot enter a formula for an item or field name in a PivotTable report.',
    engine: "PivotTables can't be refreshed or changed now: this page could not load what they need. Reload the page and try again.",
  },
  ru: {
    part: 'Нельзя изменить эту часть сводной таблицы.',
    affect: 'Нельзя внести это изменение в выделенные ячейки, так как оно повлияет на сводную таблицу. Используйте список полей, чтобы изменить отчет. Если вы пытаетесь вставить или удалить ячейки, переместите сводную таблицу и повторите попытку.',
    formula: 'Нельзя ввести формулу для имени элемента или поля в отчете сводной таблицы.',
    engine: 'Сводные таблицы сейчас нельзя обновить или изменить: страница не смогла загрузить то, что для этого нужно. Обновите страницу и повторите попытку.',
  },
}
export type PivotAlert = keyof (typeof ALERTS)['en']

/** The outcome code of a change a PivotTable's cells do not take, as xlcalc's refusals name it. */
export const PIVOT = 'pivot'

// Temporary until xlcalc's refusals carry `reason`: its internal message keys and its en-US words for them, after the place.
const KEYS = { pivotPart: 'part', pivotFormula: 'formula', pivotClear: 'affect', pivotNotLoaded: 'engine' } as const
const EN_WORDS: Record<keyof typeof KEYS, string> = {
  pivotPart: "we can't change this part of the PivotTable",
  pivotFormula: 'cannot enter a formula for an item or field name in a PivotTable report',
  pivotClear: "we can't make this change for the selected cells because it will affect a PivotTable",
  pivotNotLoaded: 'registerPivot(pivotEngine) loads it',
}

/** The alert of xlcalc's message by its key (temporary, see `KEYS`): the locale's own words where it has them, else en-US's. */
function keyed(message: string, locale: string): PivotAlert | null {
  const own = (LOCALES.get(locale) as { messages?: Record<string, unknown> } | undefined)?.messages
  for (const key of Object.keys(KEYS) as (keyof typeof KEYS)[]) {
    const say = own?.[key]
    const said = typeof say === 'function' ? String((say as (p: object) => string)({ sheet: 'S', cell: 'A1', pivot: 'P' })) : null
    const words = said === null ? EN_WORDS[key] : said.slice(said.indexOf(': ') + 2)
    if (message === words || message.endsWith(`: ${words}`)) return KEYS[key]
  }
  return null
}

/**
 * xlcalc's refusal over PivotTables as the outcome the sheet shows in an alert: Excel's alert for its `reason`, or for the
 * kind `by` the operation makes it (Delete, and an operation that cuts one, `affect`), else the library's own message.
 */
export function pivotRefusals(locale: string) {
  const words = locale.toLowerCase().startsWith('ru') ? ALERTS.ru : ALERTS.en
  const alert = (kind: PivotAlert) => ({ ok: false as const, reason: 'refused' as const, code: PIVOT, message: words[kind] })
  const refusal = (r: { message: string; reason?: unknown }, by?: PivotAlert) => {
    const reason = typeof r.reason === 'string' && r.reason in words ? (r.reason as PivotAlert) : null
    const kind = reason ?? keyed(r.message, locale) ?? by
    return kind ? alert(kind) : { ok: false as const, reason: 'refused' as const, code: PIVOT, message: r.message }
  }
  return { alert, refusal }
}

/** The operations of xlcalc on PivotTables themselves, whose refusals keep the library's reason; any other refused with `pivot` would cut one. */
const PIVOT_OPS = new Set<StructureOp['op']>(['addPivotTable', 'editPivotTable', 'renamePivotTable', 'movePivotTable', 'deletePivotTable', 'refreshPivots', 'groupPivotField', 'ungroupPivotField', 'renamePivotGroup', 'setPivotCalculatedField', 'setPivotCalculatedItem', 'addSlicer', 'editSlicer', 'deleteSlicer', 'addTimeline', 'editTimeline', 'deleteTimeline'])
export const isPivotOp = (op: StructureOp): boolean => PIVOT_OPS.has(op.op)

/** What a protected sheet must allow for an operation on PivotTables: a box of Protect Sheet on a sheet. */
export interface PivotNeed {
  sheet: string
  allow: 'pivotTables' | 'objects'
}

/**
 * What an operation on PivotTables needs of protected sheets, by Microsoft's Protect a worksheet (not measured in Excel): Use
 * PivotTable reports ("Format, change the layout, refresh, or otherwise modify PivotTable reports, or create new reports")
 * on each sheet whose PivotTables it lays out again, every sheet with a PivotTable over the cache for a refresh or a change
 * of the cache, and Refresh All every sheet with one; a slicer's or a timeline's choice refilters the PivotTables it names;
 * adding or deleting a slicer or a timeline also needs Edit objects on its sheet, as slicers and timelines are objects
 * locked when added.
 */
export function pivotNeeds(layout: BookLayout, op: StructureOp): PivotNeed[] {
  const cacheOf = (sheet: string, name: string) => layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables?.find((t) => sameName(t.name, name))?.cache
  const over = (caches: (number | undefined)[]) => layout.sheets.filter((x) => x.pivotTables?.some((t) => caches.includes(t.cache))).map((x) => x.name)
  const pivots = (sheets: string[]): PivotNeed[] => sheets.map((sheet) => ({ sheet, allow: 'pivotTables' }))
  const filtered = (refs: readonly PivotRef[] | undefined) => (refs ?? []).map((p) => p.sheet)
  /** The sheet of a slicer or a timeline, and the PivotTables its cache filters. */
  const objectOf = (kind: 'slicers' | 'timelines', name: string) => {
    const sheet = layout.sheets.find((x) => x[kind]?.some((o) => sameName(o.name, name)))
    const cache = sheet?.[kind]?.find((o) => sameName(o.name, name))?.cache
    const caches: readonly { name: string; pivots: PivotRef[] }[] = (kind === 'slicers' ? layout.slicerCaches : layout.timelineCaches) ?? []
    return { sheet: sheet?.name, pivots: caches.find((c) => sameName(c.name, cache ?? ''))?.pivots }
  }
  const cachePivots = (kind: 'slicers' | 'timelines', cache: string | { pivots: PivotRef[] }) =>
    typeof cache === 'string' ? ((kind === 'slicers' ? layout.slicerCaches : layout.timelineCaches) ?? []).find((c) => sameName(c.name, cache))?.pivots : cache.pivots
  switch (op.op) {
    case 'refreshPivots':
      return pivots(op.pivots ? over(op.pivots.map((p) => cacheOf(p.sheet, p.name))) : layout.sheets.filter((x) => x.pivotTables?.length).map((x) => x.name))
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
    case 'addSlicer':
    case 'addTimeline': {
      const kind = op.op === 'addSlicer' ? 'slicers' : 'timelines'
      return [{ sheet: op.sheet, allow: 'objects' }, ...pivots([op.sheet, ...filtered(cachePivots(kind, op.cache))])]
    }
    case 'editSlicer':
    case 'editTimeline': {
      const o = op.op === 'editSlicer' ? objectOf('slicers', op.slicer) : objectOf('timelines', op.timeline)
      // Report Connections may name more PivotTables, which take its choice
      return pivots([...filtered(o.pivots), ...filtered((op.cache as { pivots?: PivotRef[] } | undefined)?.pivots)])
    }
    case 'deleteSlicer':
    case 'deleteTimeline': {
      const o = op.op === 'deleteSlicer' ? objectOf('slicers', op.slicer) : objectOf('timelines', op.timeline)
      return o.sheet ? [{ sheet: o.sheet, allow: 'objects' }, ...pivots([o.sheet, ...filtered(o.pivots)])] : []
    }
    default:
      return []
  }
}
