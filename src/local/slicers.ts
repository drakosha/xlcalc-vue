// Slicers and timelines in the local source: each sheet's, as the sheet draws them, from the workbook's layout (xlcalc keeps
// their caches with the items chosen and with data after each calculation, and labels the items), and the fields of a
// pivot cache Insert Slicer and Insert Timeline offer. What a protected sheet allows of them is `pivotNeeds`'s (./pivots).
import type { BookLayout, PivotCache, SheetLayout } from 'xlcalc'
import type { SheetPivotField, SheetSlicer, SheetTimeline } from '../contract/source'

/** The fields of a pivot cache Insert Slicer offers (not the calculated ones), those holding dates marked for Insert Timeline. */
export function pivotFields(cache: PivotCache | undefined): SheetPivotField[] {
  return (cache?.fields ?? []).flatMap((f) => (f.formula !== undefined ? [] : [{ name: f.name, ...(f.facts?.types.includes('date') ? { dates: true as const } : {}) }]))
}

/** A sheet's slicers with their caches' fields, PivotTables and items, as `SheetInfo.slicers` lists them; none without them. */
export function slicersOf(layout: BookLayout, sheet: SheetLayout): SheetSlicer[] {
  const out: SheetSlicer[] = []
  for (const s of sheet.slicers ?? []) {
    const sc = layout.slicerCaches?.find((c) => c.name === s.cache)
    if (!sc) continue
    const items = (sc.items ?? []).map((x) => ({
      item: x.item,
      label: x.label,
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
