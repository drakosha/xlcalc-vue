// What a PivotTable's cells show, for the sheet's cell menu, filter buttons and dialogs: what each cell is and where Excel's
// filter buttons are, as xlcalc's `pivotCells` tells it (Excel's Range.PivotCell); the labels of its fields' items, the
// value fields' formats and the calculated fields and items, which xlcalc does not list.
import { builtinFormatCode, dateToSerial, displayValue, pivotCells, toLocal, type BookLayout, type CellValue, type PivotCacheField, type PivotCells, type PivotTable, type PivotValue } from 'xlcalc'
import { cellAddress, parseCell, parseRange, sameName } from 'xlcalc/a1'
import type { SheetPivotButton, SheetPivotCalculated, SheetPivotCells } from '../contract/source'

/** Excel's short date (built-in format 14), in which a PivotTable shows a date without a format of its own. */
const SHORT_DATE = 14
/** The label of the blank item, as xlcalc lays it out in the workbook's language. */
const BLANK = { en: '(blank)', ru: '(пусто)' }

/** A value of the cache as a cell holds it. */
function cellValueOf(v: PivotValue): CellValue {
  if (v === null) return null
  if (typeof v === 'object' && 'date' in v) return dateToSerial(v.date)
  return v as CellValue
}

/** A field's items' labels in the PivotTable's order: the caption, or the value in the field's number format. */
function labelsOf(t: PivotTable, f: PivotCacheField, index: number, locale: string): string[] {
  const own = t.fields[index]
  const values = f.group?.items ?? f.items ?? []
  const format = own?.numberFormat ?? f.numberFormat
  const blank = locale.toLowerCase().startsWith('ru') ? BLANK.ru : BLANK.en
  return (own?.items ?? []).map((x) => {
    if (x.caption !== undefined) return x.caption
    const v = x.item === undefined ? undefined : values[x.item]
    if (v === undefined) return ''
    if (v === null) return blank
    const date = typeof v === 'object' && 'date' in v
    return displayValue(cellValueOf(v), { locale, numberFormat: date ? (format ?? SHORT_DATE) : format }).text
  })
}

/**
 * The filter buttons from xlcalc's `pivotCells`, each with the field its menu is about. A field's name filters that field;
 * Σ Values' name filters none. Row Labels and Column Labels name no field: their axis is the one whose labels lie under
 * them (the rows over the row labels' columns), their field its first, as Excel's menu starts (none when the axis holds
 * Σ Values alone). A value field's name at the corner filters none. TODO(xlcalc): its next release gives Row Labels' and
 * Column Labels' axis and draws no button at the corner (Excel shows none there); this is the one place to take them.
 */
function buttonsOf(t: PivotTable, got: PivotCells): SheetPivotButton[] {
  const labels = t.location.areas?.rowLabels ? parseRange(t.location.areas.rowLabels) : null
  const first = (axis: 'rows' | 'columns') => (axis === 'rows' ? t.rows : t.columns)?.find((f): f is number => f !== 'values')
  return got.buttons.flatMap((cell): SheetPivotButton[] => {
    const c = got.cells[cell]
    const at = parseCell(cell)
    if (!c || !at) return []
    if (c.kind === 'field' && c.axis && typeof c.field === 'number') return [{ cell, axis: c.axis, field: c.field }]
    const axis = c.axis ?? (labels && at.col >= labels.left && at.col <= labels.right ? 'rows' : 'columns')
    const field = c.kind === 'blank' && axis !== 'pages' ? first(axis) : undefined
    return [{ cell, axis, ...(field !== undefined ? { field } : {}) }]
  })
}

/** A formula of the cache as Excel's dialog shows it: in the workbook's locale, with `=`; as stored where it does not translate. */
function shown(formula: string, locale: string): string {
  const got = toLocal(formula, locale)
  return `=${got.ok ? got.formula : formula}`
}
/** The cache's calculated fields and items; an item's formula without its field's name before each item (`Region[East]` as `East`). */
function calculatedOf(cache: { fields: PivotCacheField[]; calculatedItems?: { field: number; item: number; formula: string }[] }, locale: string): SheetPivotCalculated[] {
  const out: SheetPivotCalculated[] = []
  cache.fields.forEach((f, field) => f.formula !== undefined && out.push({ field, name: f.name, formula: shown(f.formula, locale) }))
  for (const c of cache.calculatedItems ?? []) {
    const name = cache.fields[c.field]?.name ?? ''
    const plain = c.formula.split(`${name}[`).join('').replace(/\]/g, '')
    out.push({ field: c.field, item: c.item, name: String(cache.fields[c.field]?.items?.[c.item] ?? ''), formula: shown(plain, locale) })
  }
  return out
}

/** `SheetSource.pivotCells` over the session's layout; null when the sheet has no PivotTable `name`. */
export function pivotCellsOf(layout: BookLayout, sheet: string, name: string, cells: string[] | undefined, locale: string): SheetPivotCells | null {
  const t = layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables?.find((p) => sameName(p.name, name))
  const cache = t && layout.pivotCaches?.find((c) => c.id === t.cache)
  if (!t || !cache) return null
  const onAxis = new Set<number>([...(t.rows ?? []), ...(t.columns ?? []), ...(t.pages ?? []).map((p) => p.field)].filter((f): f is number => f !== 'values'))
  const labels = cache.fields.map((f, i) => (onAxis.has(i) || t.fields[i]?.items ? labelsOf(t, f, i, locale) : null))
  const got = pivotCells(layout, sheet, name)
  if (!got) return null
  const at = (cell: string) => {
    const a = parseCell(cell)
    return (a && got.cells[cellAddress(a.row, a.col)]) || null
  }
  const formats = (t.values ?? []).map((v) => (v.numberFormat === undefined ? null : typeof v.numberFormat === 'string' ? v.numberFormat : (builtinFormatCode(v.numberFormat, locale) ?? null)))
  return { labels, formats, calculated: calculatedOf(cache, locale), buttons: buttonsOf(t, got), ...(cells ? { cells: cells.map(at) } : {}) }
}
