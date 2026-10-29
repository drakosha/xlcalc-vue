// Excel's View > Split: the split window placed as Excel places it from the active cell, moved and removed with its bars, kept
// through the source's `split` (a view setting, not an undo step). The panes themselves are drawn and scrolled by `panes`.
import { computed } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import { errorText } from '../failure'
import type { SheetSplit } from '../contract/source'
import type { Geometry } from './geometry'
import { colGroup, PANE_NAMES, PT, rowGroup, type Panes } from './panes'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

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
/** Screen pixels to the file's points, to a twentieth of a point as xlsx keeps them. */
const points = (px: number): number => Math.round(px * PT * 20) / 20

export function splitPlace(active: { d: number; col: number }, first: { d: number; col: number }, middle: () => SplitPlace): SplitPlace {
  const rows = active.d - first.d
  const cols = active.col - first.col
  if (rows <= 0 && cols <= 0) return middle()
  return { rows: Math.max(0, rows), cols: Math.max(0, cols) }
}

export function useSplit(s: SheetState, geo: Geometry, v: Viewport, panes: Panes, notify: (text: string) => void) {
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
    const rows = geo.rowOffsets.value
    const cols = geo.offsets.value
    const row = first.d + at.rows
    const col = first.col + at.cols
    const out: SheetSplit = { topLeft: cellAddress(s.rowOf(row), col) }
    if (at.cols > 0) out.x = points((v.leadW.value + (cols[col - 1] ?? 0) - (cols[first.col - 1] ?? 0)) * s.zoom.value)
    if (at.rows > 0) out.y = points((v.headH.value + (rows[row - 1] ?? 0) - (rows[first.d - 1] ?? 0)) * s.zoom.value)
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
    store(value)
    return next
  }

  /**
   * The lines where a bar was dragged to (screen pixels from the window's edges, null for a line removed): the split there,
   * with the second panes' scroll as its `topLeft` and the active pane; none once both lines are gone. A read-only sheet
   * moves its lines on screen only, as the source keeps nothing for it.
   */
  function move(lines: { x: number | null; y: number | null }): void {
    const was = split.value
    if (!was) return
    let next: SheetSplit | null = null
    if (lines.x !== null || lines.y !== null) {
      const row = s.rowOf(geo.rowAt(panes.topOf(lines.y !== null ? 1 : 0)))
      const col = geo.colAt(panes.leftOf(lines.x !== null ? 1 : 0))
      const id = panes.active.value
      const pane = PANE_NAMES[(lines.y !== null ? rowGroup(id) : 0) * 2 + (lines.x !== null ? colGroup(id) : 0)]!
      next = { ...(lines.x !== null ? { x: points(lines.x) } : {}), ...(lines.y !== null ? { y: points(lines.y) } : {}), topLeft: cellAddress(row, col), ...(pane !== 'topLeft' ? { activePane: pane } : {}) }
    }
    if (available.value) store(next)
    else put(s.sheetName.value, next)
  }

  /** Sets the split at once and keeps it through the source; a refusal puts it back. */
  function store(value: SheetSplit | null): void {
    const was = split.value
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
  }

  return { split, available, toggle, move }
}

export type Split = ReturnType<typeof useSplit>
