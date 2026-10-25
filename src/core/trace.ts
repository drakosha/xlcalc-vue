// Excel's Trace Dependents and Remove Arrows over the source's `dependents`: arrows from a cell to the cells whose formulas
// refer to it, one level per command, kept per sheet until Remove Arrows or a change of cells, and their place over the grid.
import { computed, shallowRef, watch } from 'vue'
import { cellAddress, parseCell, sameName } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import { useSheetReads, type SheetState } from './state'
import type { Viewport } from './viewport'
import type { Windows } from './windows'

/** The most cells one command traces, so a sheet of long chains does not send thousands of queries at once. */
const MAX_TRACED = 200

/** An arrow from a cell to its dependent on the sheet, or to the worksheet icon for dependents on other sheets (`to` null). */
export interface TraceArrow {
  from: CellPosition
  to: CellPosition | null
}

/** A sheet's arrows and the cells whose dependents they already show. */
interface SheetTrace {
  arrows: TraceArrow[]
  traced: Set<string>
}

/** An arrow as drawn, in the scroller's content coordinates: blue, red from a cell with an error, dashed to the worksheet icon; `frozen` when an end is in frozen panes. */
export interface DrawnArrow {
  key: string
  x1: number
  y1: number
  x2: number
  y2: number
  kind: 'cell' | 'error' | 'sheet'
  frozen: boolean
}

/** Where the worksheet icon is drawn from its cell's center, beside the cell as Excel draws it. */
const ICON_DX = 40
const ICON_DY = -24

const keyOf = (p: CellPosition) => cellAddress(p.row, p.col)

export function useTrace(s: SheetState, v: Viewport, w: Windows, notify: (text: string) => void) {
  const read = useSheetReads(s)
  const available = computed(() => typeof s.source().dependents === 'function')
  /** Arrows by sheet name, as Excel keeps them on each sheet. */
  const sheets = shallowRef(new Map<string, SheetTrace>())
  const traceOf = (sheet: string) => [...sheets.value].find(([name]) => sameName(name, sheet))?.[1] ?? null

  /** The sheet's cells reached by arrows from `origin`, `origin` included. */
  function reached(t: SheetTrace, origin: string): string[] {
    const out = [origin]
    const seen = new Set(out)
    for (let i = 0; i < out.length; i++)
      for (const a of t.arrows) {
        const k = a.to && keyOf(a.from) === out[i] ? keyOf(a.to) : null
        if (k === null || seen.has(k)) continue
        seen.add(k)
        out.push(k)
      }
    return out
  }

  let busy = false
  /**
   * Excel's Trace Dependents from the active cell: its direct dependents the first time, then each time the next level, the
   * dependents of the cells the arrows reach that have none traced yet. False when nothing was added: no `dependents`, a call
   * already running, no further level, or no formula referring to the cell (with Excel's message).
   */
  async function traceDependents(): Promise<boolean> {
    const call = s.sourceMethod('dependents')
    const a = s.selected.value
    if (!call || !a || !s.sheet.value || busy) return false
    const sheet = s.sheetName.value
    const origin = keyOf(a)
    const before = traceOf(sheet)
    const t: SheetTrace = before ? { arrows: [...before.arrows], traced: new Set(before.traced) } : { arrows: [], traced: new Set() }
    const frontier = reached(t, origin).filter((k) => !t.traced.has(k)).slice(0, MAX_TRACED)
    if (!frontier.length) return false
    const signal = read()
    const outdated = s.outdated()
    const count = t.arrows.length
    busy = true
    try {
      for (const cell of frontier) {
        const got = await call(sheet, cell, signal)
        if (signal.aborted || outdated()) return false
        const from = parseCell(cell)!
        const has = new Set(t.arrows.filter((x) => keyOf(x.from) === cell).map((x) => (x.to ? keyOf(x.to) : '')))
        for (const d of got.cells) {
          const to = sameName(d.sheet, sheet) ? parseCell(d.cell) : null
          const k = to ? keyOf(to) : ''
          if (has.has(k) || (to && keyOf(to) === cell)) continue
          has.add(k)
          t.arrows.push({ from, to })
        }
        t.traced.add(cell)
      }
    } catch {
      if (!signal.aborted && !outdated()) notify(s.messages.value.traceFailed)
      return false
    } finally {
      busy = false
    }
    const next = new Map(sheets.value)
    next.set(sheet, t)
    sheets.value = next
    if (t.arrows.length > count) return true
    if (!before?.traced.has(origin)) notify(s.messages.value.traceNone)
    return false
  }

  /** Excel's Remove Arrows: every arrow on every sheet; false when there were none. */
  function removeArrows(): boolean {
    if (!sheets.value.size) return false
    sheets.value = new Map()
    return true
  }
  // Excel removes the arrows when cells change (an edit, a paste, rows inserted or deleted), and a new workbook has none.
  const changing = s.hooks.changing
  s.hooks.changing = () => {
    changing()
    removeArrows()
  }
  watch(
    () => s.source(),
    () => void removeArrows(),
  )

  /** A cell's center on the grid, null when it is not drawn (a hidden row or column, a row the display filter hides). */
  function centerOf(p: CellPosition): { x: number; y: number; frozen: boolean } | null {
    const d = s.indexOf(p.row)
    if (d === undefined) return null
    const b = v.boxOf(d, p.col, d, p.col)
    if (b.width <= 0 || b.height <= 0) return null
    return { x: b.left + b.width / 2, y: b.top + b.height / 2, frozen: b.frozen }
  }
  /** The arrows of the displayed sheet as drawn: red from a cell whose value is an error, as Excel draws them. */
  const drawn = computed<DrawnArrow[]>(() => {
    const t = traceOf(s.sheetName.value)
    if (!t) return []
    const out: DrawnArrow[] = []
    for (const a of t.arrows) {
      const p = centerOf(a.from)
      // the icon above and to the right of its cell, below it for a cell in the first row
      const q = a.to ? centerOf(a.to) : p && { x: p.x + ICON_DX, y: p.y + (a.from.row === 1 ? -ICON_DY : ICON_DY), frozen: p.frozen }
      if (!p || !q) continue
      const value = w.valueOfRow(a.from.row, a.from.col)
      const error = !!value && typeof value === 'object' && 'error' in value
      out.push({ key: `${keyOf(a.from)}>${a.to ? keyOf(a.to) : '*'}`, x1: p.x, y1: p.y, x2: q.x, y2: q.y, kind: a.to ? (error ? 'error' : 'cell') : 'sheet', frozen: p.frozen || q.frozen })
    }
    return out
  })

  return { available, drawn, traceDependents, removeArrows }
}

export type Trace = ReturnType<typeof useTrace>
