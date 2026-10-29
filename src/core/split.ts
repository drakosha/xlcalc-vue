// Excel's View > Split: the split window placed as Excel places it from the active cell, kept through the source's `split`
// (a view setting, not an undo step). The sheet keeps and reports the split; it draws a single pane.
import { computed } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import { errorText } from '../failure'
import type { SheetSplit } from '../contract/source'
import type { Geometry } from './geometry'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** Points in a CSS pixel: the file's split lines are points on the screen. */
const PT = 0.75

/** Where a window's split goes: lines after `rows` and `cols` of the window's first row and column (0: no line on that axis). */
export interface SplitPlace {
  rows: number
  cols: number
}

/**
 * Excel's place for a split (Excel for Mac's object model, `Split = True`, measured from C5, A5, C1 and A1): above and to
 * the left of the active cell, counted from the window's first row and column, so a cell in the first row gives a vertical
 * line alone and one in the first column a horizontal one; the window's first cell active splits it at its middle (`middle`).
 */
export function splitPlace(active: { d: number; col: number }, first: { d: number; col: number }, middle: () => SplitPlace): SplitPlace {
  const rows = active.d - first.d
  const cols = active.col - first.col
  if (rows <= 0 && cols <= 0) return middle()
  return { rows: Math.max(0, rows), cols: Math.max(0, cols) }
}

export function useSplit(s: SheetState, geo: Geometry, v: Viewport, notify: (text: string) => void) {
  const split = computed(() => s.sheet.value?.split ?? null)
  /** Split is offered: the source keeps it, the sheet is editable, has no frozen panes (which take its place) and no display filter. */
  const available = computed(() => !!s.sheet.value && s.editable.value && !!s.sourceMethod('split') && !v.bookFreeze.value && !s.visible.value)

  function put(sheet: string, next: SheetSplit | null): void {
    s.sheets.value = s.sheets.value.map((x) => {
      if (x.name !== sheet) return x
      const { split: _, ...rest } = x
      return next ? { ...rest, split: next } : rest
    })
  }

  /** The split for the active cell: lines in the file's points (headings counted in, at the zoom) and the pane after them. */
  function placed(): SheetSplit | null {
    const a = s.selected.value
    if (!a) return null
    const d = s.indexOf(a.row)
    if (d === undefined) return null
    const first = { d: geo.rowAt(s.scrollTop.value), col: geo.colAt(s.scrollLeft.value) }
    const middle = (): SplitPlace => ({
      rows: geo.rowAt(s.scrollTop.value + Math.max(0, (s.viewH.value - v.headH.value) / 2)) - first.d,
      cols: geo.colAt(s.scrollLeft.value + Math.max(0, (s.viewW.value - v.leadW.value) / 2)) - first.col,
    })
    const at = splitPlace({ d, col: a.col }, first, middle)
    if (at.rows <= 0 && at.cols <= 0) return null
    const pt = (px: number) => Math.round(px * s.zoom.value * PT * 20) / 20
    const rows = geo.rowOffsets.value
    const cols = geo.offsets.value
    const row = first.d + at.rows
    const col = first.col + at.cols
    const out: SheetSplit = { topLeft: cellAddress(s.rowOf(row), col) }
    if (at.cols > 0) out.x = pt(v.leadW.value + (cols[col - 1] ?? 0) - (cols[first.col - 1] ?? 0))
    if (at.rows > 0) out.y = pt(v.headH.value + (rows[row - 1] ?? 0) - (rows[first.d - 1] ?? 0))
    // the active cell stays in the pane past the lines, unless the window's first cell split it at its middle
    if (d >= row && a.col >= col) out.activePane = at.rows > 0 && at.cols > 0 ? 'bottomRight' : at.rows > 0 ? 'bottomLeft' : 'topRight'
    return out
  }

  /** Excel's View > Split, toggled without `on`: set at once, then kept by the source; whether the sheet is split afterwards. */
  function toggle(on?: boolean): boolean {
    const was = split.value
    if (!available.value) return !!was
    const next = on ?? !was
    if (next === !!was) return next
    const value = next ? placed() : null
    if (next && !value) return false
    const sheet = s.sheetName.value
    put(sheet, value)
    const outdated = s.outdated()
    // Not an undo step, as in Excel: a refusal puts the split back
    s.sourceMethod('split')!(sheet, value).then(
      (out) => out.ok || refused(out.message),
      (e: unknown) => refused(errorText(e)),
    )
    function refused(text: string): void {
      if (outdated()) return
      put(sheet, was)
      notify(text || s.messages.value.editFailed)
    }
    return next
  }

  return { split, available, toggle }
}

export type Split = ReturnType<typeof useSplit>
