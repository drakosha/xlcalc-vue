// PivotTables in the local source: xlcalc's PivotTable engine (`xlcalc/pivot`), loaded as its own chunk for a workbook that
// has PivotTables, since without them nothing needs it; their areas as the sheet lists them; Excel's alerts for a change
// the library refuses over their cells; and the PivotTables an operation touches, which a protected sheet gates.
import { registerPivot, type BookInput, type BookLayout, type CellRect, type PivotTable, type SheetLayout, type StructureOp } from 'xlcalc'
import { parseRange, rangeAddress, rectsMeet, sameName } from 'xlcalc/a1'
import type { SheetPivot } from '../contract/source'

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
 * A PivotTable's table and its Filters area as xlcalc's `PivotLocation` gives them: the filters take `pageRows` rows above
 * the table with a blank row between, each column of filters a field's name and its item and a gap (`pageColumns`), as
 * xlcalc lays them out; the table's headers end above `firstDataRow`, its values start at `firstDataColumn`.
 */
function areasOf(t: PivotTable): { table: CellRect; filters?: CellRect; headers?: CellRect; rowLabels?: CellRect; values?: CellRect } | null {
  const x = parseRange(t.location.range)
  if (!x) return null
  const { firstDataRow: dataRow, firstDataColumn: dataCol, pageRows = 0, pageColumns } = t.location
  const top = x.top + dataRow
  const left = x.left + dataCol
  return {
    table: x,
    ...(pageRows > 0 && x.top - pageRows - 1 >= 1 ? { filters: { top: x.top - pageRows - 1, bottom: x.top - 2, left: x.left, right: x.left + 3 * Math.max(1, pageColumns ?? 1) - 2 } } : {}),
    ...(dataRow > 0 ? { headers: { ...x, bottom: Math.min(x.bottom, top - 1) } } : {}),
    ...(dataCol > 0 && top <= x.bottom ? { rowLabels: { ...x, top, right: Math.min(x.right, left - 1) } } : {}),
    ...(t.values?.length && top <= x.bottom && left <= x.right ? { values: { ...x, top, left } } : {}),
  }
}

/** A sheet's PivotTables as `SheetInfo.pivots` lists them; none for a sheet without them. */
export function pivotsOf(sheet: SheetLayout): SheetPivot[] {
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
    })
  }
  return out
}

/** The PivotTable whose table or Filters area meets `rect` on `sheet`, with its values area; null when none does. */
export function pivotMeeting(layout: BookLayout, sheet: string, rect: CellRect): { table: PivotTable; values?: CellRect } | null {
  for (const t of layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables ?? []) {
    const a = areasOf(t)
    if (a && (rectsMeet(a.table, rect) || (a.filters && rectsMeet(a.filters, rect)))) return { table: t, ...(a.values ? { values: a.values } : {}) }
  }
  return null
}

/**
 * Excel for Windows' alerts over a PivotTable's cells: a value typed or pasted (`part`), a cleared cell or an operation
 * that would insert or delete cells in it (`affect`), a formula over a label (`formula`). The English ones are those Excel
 * 16.0 for Windows answered xlcalc's set `edit` with through its object model; the Russian ones were not measured.
 */
const ALERTS = {
  en: {
    part: "We can't change this part of the PivotTable.",
    affect: "We can't make this change for the selected cells because it will affect a PivotTable. Use the field list to change the report. If you are trying to insert or delete cells, move the PivotTable and try again.",
    formula: 'Cannot enter a formula for an item or field name in a PivotTable report.',
  },
  ru: {
    part: 'Нельзя изменить эту часть сводной таблицы.',
    affect: 'Нельзя внести это изменение в выделенные ячейки, так как оно повлияет на сводную таблицу. Используйте список полей, чтобы изменить отчет. Если вы пытаетесь вставить или удалить ячейки, переместите сводную таблицу и повторите попытку.',
    formula: 'Нельзя ввести формулу для имени элемента или поля в отчете сводной таблицы.',
  },
}
export type PivotAlert = keyof (typeof ALERTS)['en']

/** The outcome code of a change a PivotTable's cells do not take, as xlcalc's refusals name it. */
export const PIVOT = 'pivot'

/** Excel's alerts over PivotTables in the workbook's language. */
export function pivotAlerts(locale: string) {
  const words = locale.toLowerCase().startsWith('ru') ? ALERTS.ru : ALERTS.en
  return (kind: PivotAlert) => ({ ok: false as const, reason: 'refused' as const, code: PIVOT, message: words[kind] })
}

/** The operations of xlcalc on PivotTables themselves, whose refusals keep the library's reason; any other refused with `pivot` would cut one. */
const PIVOT_OPS = new Set<StructureOp['op']>(['addPivotTable', 'editPivotTable', 'renamePivotTable', 'movePivotTable', 'deletePivotTable', 'refreshPivots', 'groupPivotField', 'ungroupPivotField', 'renamePivotGroup', 'setPivotCalculatedField', 'setPivotCalculatedItem', 'addSlicer', 'editSlicer', 'deleteSlicer', 'addTimeline', 'editTimeline', 'deleteTimeline'])
export const isPivotOp = (op: StructureOp): boolean => PIVOT_OPS.has(op.op)

/**
 * The sheets whose PivotTables an operation on PivotTables lays out again: the PivotTable's own, the sheet it moves to, and
 * for a refresh or a change of its cache (grouping, calculated fields and items) every sheet with a PivotTable over that
 * cache; Refresh All (`refreshPivots` without `pivots`), every sheet with a PivotTable. Slicers and timelines are none.
 */
export function pivotSheets(layout: BookLayout, op: StructureOp): string[] {
  const cacheOf = (sheet: string, name: string) => layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables?.find((t) => sameName(t.name, name))?.cache
  const over = (caches: (number | undefined)[]) => layout.sheets.filter((x) => x.pivotTables?.some((t) => caches.includes(t.cache))).map((x) => x.name)
  switch (op.op) {
    case 'refreshPivots':
      return op.pivots ? over(op.pivots.map((p) => cacheOf(p.sheet, p.name))) : layout.sheets.filter((x) => x.pivotTables?.length).map((x) => x.name)
    case 'groupPivotField':
    case 'ungroupPivotField':
    case 'renamePivotGroup':
    case 'setPivotCalculatedField':
    case 'setPivotCalculatedItem':
      return [op.sheet, ...over([cacheOf(op.sheet, op.pivot)])]
    case 'movePivotTable':
      return [op.sheet, op.to.sheet]
    case 'addPivotTable':
      // a new sheet is not protected
      return op.newSheet ? [] : [op.sheet]
    case 'editPivotTable':
    case 'renamePivotTable':
    case 'deletePivotTable':
      return [op.sheet]
    default:
      return []
  }
}
