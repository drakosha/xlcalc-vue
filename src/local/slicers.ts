// Slicers and timelines in the local source: each sheet's, as the sheet draws them, from the workbook's layout (xlcalc keeps
// their caches with the items chosen and with data after each calculation), and the fields of a pivot cache Insert Slicer
// and Insert Timeline offer. What a protected sheet allows of them is `pivotNeeds`'s (./pivots).
import { dateToSerial, displayValue, type BookLayout, type PivotCache, type PivotValue, type SheetLayout } from 'xlcalc'
import { sameName } from 'xlcalc/a1'
import type { SheetPivotField, SheetSlicer, SheetTimeline } from '../contract/source'

/**
 * A blank item's label, as the PivotTable engine words it (`(blank)`, `(пусто)` in ru-RU).
 * TODO(xlcalc): the library gives no slicer item's caption; this function is the one place that makes it, to be replaced
 * by the library's once it gives them.
 */
const BLANK = { en: '(blank)', ru: '(пусто)' }

/**
 * A date of a pivot cache as a serial: a time alone, which Excel writes on its day 0 of 1899-12-30 (the library on
 * 1900-01-00), as its fraction of a day; a date the library cannot take, null.
 */
function serialOf(date: string): number | null {
  const time = /^(?:1899-12-30|1900-01-00)T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?/.exec(date)
  if (time) return (Number(time[1]) * 3600 + Number(time[2]) * 60 + Number(time[3] ?? 0)) / 86_400
  try {
    return dateToSerial(date)
  } catch {
    return null
  }
}

/**
 * An item as a slicer shows it: its value in its field's format, a date in the short date (with the time when it has one),
 * as the PivotTable engine formats an item's label (xlcalc's `itemFormat`); a blank in the engine's words. It never fails:
 * a value the library cannot display shows as the cache holds it.
 */
function itemLabel(v: PivotValue | undefined, format: PivotCache['fields'][number]['numberFormat'], locale: string): string {
  if (v === null || v === undefined) return locale.toLowerCase().startsWith('ru') ? BLANK.ru : BLANK.en
  if (typeof v === 'string') return v
  try {
    if (typeof v === 'object' && 'date' in v) {
      const serial = serialOf(v.date)
      if (serial === null) return v.date
      const numberFormat = format ?? (v.date.endsWith('T00:00:00') || !v.date.includes('T') ? 14 : 22)
      return displayValue(serial, { numberFormat, locale }).text
    }
    return displayValue(v, { ...(format !== undefined ? { numberFormat: format } : {}), locale }).text
  } catch {
    return typeof v === 'object' ? ('date' in v ? v.date : `#${v.error}`) : String(v)
  }
}

/** The fields of a pivot cache Insert Slicer offers (not the calculated ones), those holding dates marked for Insert Timeline. */
export function pivotFields(cache: PivotCache | undefined): SheetPivotField[] {
  return (cache?.fields ?? []).flatMap((f) => (f.formula !== undefined ? [] : [{ name: f.name, ...(f.facts?.types.includes('date') ? { dates: true as const } : {}) }]))
}

/** A sheet's slicers with their caches' fields, PivotTables and items, as `SheetInfo.slicers` lists them; none without them. */
export function slicersOf(layout: BookLayout, sheet: SheetLayout, locale: string): SheetSlicer[] {
  const out: SheetSlicer[] = []
  for (const s of sheet.slicers ?? []) {
    const sc = layout.slicerCaches?.find((c) => c.name === s.cache)
    if (!sc) continue
    const pc = sc.cache === undefined ? undefined : layout.pivotCaches?.find((c) => c.id === sc.cache)
    const field = pc?.fields.find((f) => sameName(f.name, sc.field))
    const items = (sc.items ?? []).map((x) => ({
      item: x.item,
      label: itemLabel(field?.items?.[x.item], field?.numberFormat, locale),
      ...(x.selected === false ? { selected: false as const } : {}),
      ...(x.noData ? { noData: true as const } : {}),
    }))
    out.push({
      ...s,
      field: sc.field,
      pivots: sc.pivots,
      ...(sc.cache !== undefined ? { pivotCache: sc.cache } : {}),
      ...(sc.sort ? { sort: sc.sort } : {}),
      ...(sc.customListSort !== undefined ? { customListSort: sc.customListSort } : {}),
      ...(sc.noData ? { noData: sc.noData } : {}),
      ...(sc.showMissing !== undefined ? { showMissing: sc.showMissing } : {}),
      items,
    })
  }
  return out
}

/** A sheet's timelines with their caches' fields, PivotTables, dates chosen and bounds, as `SheetInfo.timelines` lists them. */
export function timelinesOf(layout: BookLayout, sheet: SheetLayout): SheetTimeline[] {
  return (sheet.timelines ?? []).flatMap((t) => {
    const tc = layout.timelineCaches?.find((c) => c.name === t.cache)
    if (!tc) return []
    return [{ ...t, field: tc.field, pivots: tc.pivots, pivotCache: tc.cache, ...(tc.selection ? { selection: tc.selection } : {}), ...(tc.bounds ? { bounds: tc.bounds } : {}) }]
  })
}
