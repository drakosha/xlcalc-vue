// Excel's Watch Window: a list of cells whose values and formulas stay in view whatever sheet is displayed. Add Watch adds
// each cell of the selection, Delete Watch removes the chosen ones, a double click goes to a cell; the values are read
// through the source's windows and read again when cells change, the sheets are read again, or the list changes.
import { computed, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { cellAddress, parseCell, quoteSheetIfNeeded, sameName } from 'xlcalc/a1'
import type { SheetWatch } from '../contract/view'
import { DEFAULT_MAX_CELLS } from '../window-size'
import type { Navigation } from './navigation'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** At most this many cells are watched; Add Watch over a larger selection adds its first cells. */
export const WATCH_MAX = 1000
/** A workbook being recalculated answers later: the values are requested again after this long. */
const PENDING_RETRY_MS = 1000

/** What a watched cell shows: its displayed value and its formula with `=`; null while it is read, or when its sheet is gone. */
export interface WatchShown {
  value: string
  formula: string
}

interface WatchDeps {
  /** The host's list (`SheetView`'s `watches`). */
  given: () => SheetWatch[] | undefined
  /** The user changed the list: the `watches` event. */
  changed: (list: SheetWatch[]) => void
  notify: (text: string) => void
}

const keyOf = (w: SheetWatch) => `${w.sheet.toLowerCase()}\u0000${w.cell.toUpperCase()}`

export function useWatch(s: SheetState, sel: Selection, nav: Navigation, deps: WatchDeps) {
  const list = shallowRef<SheetWatch[]>(deps.given() ?? [])
  const open = ref(false)
  /** The rows chosen in the list (indexes), for Delete Watch. */
  const chosen = shallowRef<Set<number>>(new Set())
  /** The row a Shift+click extends from. */
  let anchor = -1
  const shown = shallowRef(new Map<string, WatchShown>())
  const valueOf = (w: SheetWatch): WatchShown | null => shown.value.get(keyOf(w)) ?? null

  watch(deps.given, (next) => {
    list.value = next ?? []
    chosen.value = new Set()
  })
  function change(next: SheetWatch[]): void {
    list.value = next
    chosen.value = new Set()
    deps.changed(next)
  }

  /** Excel's Add Watch: every cell of the selection on the active sheet that is not watched yet, by rows. */
  function add(): void {
    const r = sel.selRect.value
    const sheet = s.sheetName.value
    if (!r || !s.sheet.value) return
    const have = new Set(list.value.map(keyOf))
    const next = [...list.value]
    for (let row = r.top; row <= r.bottom; row++) {
      for (let col = r.left; col <= r.right; col++) {
        if (next.length >= WATCH_MAX) {
          deps.notify(s.messages.value.watchTooMany(WATCH_MAX))
          return change(next)
        }
        const w = { sheet, cell: cellAddress(row, col) }
        if (have.has(keyOf(w))) continue
        have.add(keyOf(w))
        next.push(w)
      }
    }
    if (next.length !== list.value.length) change(next)
  }
  /** Excel's Delete Watch: removes the chosen rows. */
  function remove(): void {
    if (!chosen.value.size) return
    change(list.value.filter((_, i) => !chosen.value.has(i)))
    anchor = -1
  }
  /** A click on a row: chooses it; Ctrl (Cmd on a Mac) adds it or takes it out, Shift chooses the rows from the last one clicked. */
  function choose(i: number, how: { add?: boolean; range?: boolean } = {}): void {
    const next = new Set(how.add || how.range ? chosen.value : [])
    if (how.range && anchor >= 0) for (let j = Math.min(anchor, i); j <= Math.max(anchor, i); j++) next.add(j)
    else if (how.add && next.has(i)) next.delete(i)
    else next.add(i)
    if (!how.range) anchor = i
    chosen.value = next
  }
  /** A double click: goes to the watched cell, on its sheet. */
  async function go(i: number): Promise<void> {
    const w = list.value[i]
    if (w && (await nav.goPlace(`${quoteSheetIfNeeded(w.sheet)}!${w.cell}`))) s.focusGrid()
  }

  // ── values: one window over each sheet's watched cells when it fits in the source's limit, otherwise a window per cell ──
  let ctrl: AbortController | null = null
  let timer: ReturnType<typeof setTimeout> | null = null
  let queued = false
  /** Reads the values again once the current burst of changes ends (after this task), or after `ms`. */
  function schedule(ms = 0): void {
    if (!open.value || !list.value.length) return
    if (ms > 0) {
      if (timer) clearTimeout(timer)
      timer = setTimeout(() => {
        timer = null
        void read()
      }, ms)
      return
    }
    if (queued) return
    queued = true
    queueMicrotask(() => {
      queued = false
      void read()
    })
  }
  async function read(): Promise<void> {
    ctrl?.abort()
    const mine = (ctrl = new AbortController())
    const outdated = s.outdated()
    const maxCells = s.source().maxCells ?? DEFAULT_MAX_CELLS
    const bySheet = new Map<string, { name: string; cells: { row: number; col: number; key: string }[] }>()
    for (const w of list.value) {
      const info = s.sheets.value.find((x) => sameName(x.name, w.sheet))
      const at = parseCell(w.cell)
      if (!info || !at) continue
      const group = bySheet.get(info.name) ?? { name: info.name, cells: [] }
      group.cells.push({ ...at, key: keyOf(w) })
      bySheet.set(info.name, group)
    }
    const next = new Map<string, WatchShown>()
    let pending = false
    const reads: Promise<void>[] = []
    for (const { name, cells } of bySheet.values()) {
      const top = Math.min(...cells.map((c) => c.row))
      const left = Math.min(...cells.map((c) => c.col))
      const rows = Math.max(...cells.map((c) => c.row)) - top + 1
      const cols = Math.max(...cells.map((c) => c.col)) - left + 1
      const boxes = rows * cols <= maxCells ? [{ row: top, col: left, rows, cols }] : cells.map((c) => ({ row: c.row, col: c.col, rows: 1, cols: 1 }))
      for (const box of boxes) {
        reads.push(
          s
            .source()
            .window({ sheet: name, ...box }, mine.signal)
            .then((got) => {
              if ('pending' in got) return void (pending = true)
              for (const c of cells) {
                const i = c.row - got.row
                const j = c.col - got.col
                const text = got.text[i]?.[j]
                if (text === undefined) continue
                const value = got.values[i]?.[j]
                const formula = got.formulas[cellAddress(c.row, c.col)]
                // a narrow column's ##### is the grid's: the Watch Window shows the value
                next.set(c.key, { value: /^#+$/.test(text) && typeof value === 'number' ? String(value) : text, formula: formula ? `=${formula}` : '' })
              }
            }),
        )
      }
    }
    // a failed read leaves the cells it covers empty, as the sheet's own cells stay unloaded
    await Promise.allSettled(reads)
    if (mine !== ctrl || outdated()) return
    ctrl = null
    shown.value = next
    if (pending) schedule(PENDING_RETRY_MS)
  }
  s.hooks.changed = () => schedule()
  watch([open, list, s.sheets], () => schedule())
  onScopeDispose(() => {
    ctrl?.abort()
    if (timer) clearTimeout(timer)
  })

  function setOpen(on = true): boolean {
    open.value = on
    if (!on) chosen.value = new Set()
    return on
  }
  const canAdd = computed(() => !!s.sheet.value && !!sel.selRect.value)

  return { list, open, chosen, valueOf, add, remove, choose, go, setOpen, canAdd }
}

export type Watch = ReturnType<typeof useWatch>
