// The selection's totals for the status line: the source's `stats` when it has them, else the loaded windows.
import { computed, onScopeDispose, shallowRef, watch } from 'vue'
import type { CellValue } from 'xlcalc/contract'
import type { RangeStats } from '../source'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** A pause before asking the source for totals: a drag or Shift+arrows changes the selection many times a second. */
const STATS_DELAY_MS = 150

/** Display rows of a selection: a run of them without a filter (a whole column is a million), else the filter's list. */
export type ShownRows = { from: number; to: number } | number[]
const countOf = (ds: ShownRows): number => (Array.isArray(ds) ? ds.length : Math.max(0, ds.to - ds.from + 1))
/** The rows as a list, for what needs each one (a copy's text). */
export const listOf = (ds: ShownRows): number[] => (Array.isArray(ds) ? ds : Array.from({ length: countOf(ds) }, (_, i) => ds.from + i))

export function useTotals(s: SheetState, w: Windows, sel: Selection) {
  /** Display rows of the selection: under a filter only the shown rows between its corners, in display order. */
  function shownRows(r1: number, r2: number): ShownRows {
    const list = s.visible.value?.rows
    if (!list) return { from: r1, to: r2 }
    const out: number[] = []
    list.forEach((r, i) => r >= r1 && r <= r2 && out.push(i + 1))
    return out
  }
  /** Whether the loaded windows hold these display rows across the columns; a run is walked window by window. */
  function inWindows(ds: ShownRows, c1: number, c2: number): boolean {
    const holds = (d: number) => {
      const slot = w.slotAt(d, c1)
      return slot && c2 < slot.win!.col + slot.cols ? slot : null
    }
    if (Array.isArray(ds)) return ds.every((d) => !!holds(d))
    for (let d = ds.from; d <= ds.to; ) {
      const slot = holds(d)
      if (!slot) return false
      d = slot.win!.row + slot.rows
    }
    return true
  }
  /** Number formats by locale: made once, a drag changes the totals many times a second. */
  const formats = new Map<string, Intl.NumberFormat>()
  const numberFormat = (locale: string, digits?: number): Intl.NumberFormat => {
    const key = `${locale}\u0000${digits ?? ''}`
    let fmt = formats.get(key)
    if (!fmt) formats.set(key, (fmt = new Intl.NumberFormat(locale, digits === undefined ? undefined : { maximumFractionDigits: digits })))
    return fmt
  }
  const format = (st: RangeStats): string => {
    const m = s.messages.value
    const fmt = numberFormat(s.locale.value, 2)
    const parts = [m.statsCount(fmt.format(st.cells))]
    if (st.numbers) {
      parts.push(m.statsSum(fmt.format(st.sum)), m.statsAverage(fmt.format(st.sum / st.numbers)))
      if (st.numbers > 1) parts.push(m.statsMin(fmt.format(st.min ?? 0)), m.statsMax(fmt.format(st.max ?? 0)))
      if (st.numbers !== st.cells) parts.push(m.statsNumbers(fmt.format(st.numbers)))
    }
    return parts.join(' · ')
  }

  // ── totals from the source ──
  const fromSource = shallowRef<{ key: string; stats: RangeStats } | null>(null)
  const statsKey = () => `${s.sheetName.value}\u0000${sel.gridRef.value}\u0000${s.visible.value?.key ?? ''}`
  let statsTimer: ReturnType<typeof setTimeout> | null = null
  let statsCtrl: AbortController | null = null
  watch(
    () => (sel.multi.value ? statsKey() : ''),
    (key) => {
      if (statsTimer) clearTimeout(statsTimer)
      statsCtrl?.abort()
      const stats = s.sourceMethod('stats')
      if (!key || !stats) return
      const sheet = s.sheetName.value
      const range = sel.gridRef.value
      const filter = s.visible.value?.spec
      statsTimer = setTimeout(() => {
        const ctrl = (statsCtrl = new AbortController())
        stats(sheet, range, filter, ctrl.signal)
          .then((got) => {
            if (!ctrl.signal.aborted && statsKey() === key) fromSource.value = { key, stats: got }
          })
          .catch(() => {
            // Totals are an ornament: without them the loaded windows count.
          })
      }, STATS_DELAY_MS)
    },
  )

  onScopeDispose(() => {
    if (statsTimer) clearTimeout(statsTimer)
    statsCtrl?.abort()
  })

  /** Totals as Excel's status line: count of non-empty cells, and over numbers sum, average, min and max. */
  const statsText = computed(() => {
    const r = sel.selRect.value
    if (!r || !sel.multi.value) return ''
    const got = fromSource.value
    if (got && got.key === statsKey()) return format(got.stats)
    const r2 = Math.min(r.bottom, s.dataRows.value)
    const c2 = Math.min(r.right, s.dataCols.value)
    const ds = shownRows(r.top, r2)
    if (r2 >= r.top && c2 >= r.left && !inWindows(ds, r.left, c2)) {
      const rows = s.visible.value ? countOf(ds) : r.bottom - r.top + 1
      return s.messages.value.statsCells(numberFormat(s.locale.value).format(rows * (r.right - r.left + 1)))
    }
    const st: RangeStats = { cells: 0, numbers: 0, sum: 0, min: null, max: null }
    // Each row's cells straight from the window holding its first one, which holds them all (inWindows).
    const add = (d: number) => {
      const win = w.slotAt(d, r.left)!.win!
      const text = win.text[d - win.row]
      const values = win.values[d - win.row]
      for (let col = r.left; col <= c2; col++) {
        const value: CellValue | undefined = values?.[col - win.col]
        if ((text?.[col - win.col] ?? '') === '' && (value === null || value === undefined || value === '')) continue
        st.cells++
        if (typeof value !== 'number') continue
        st.numbers++
        st.sum += value
        st.min = st.min === null ? value : Math.min(st.min, value)
        st.max = st.max === null ? value : Math.max(st.max, value)
      }
    }
    if (c2 >= r.left) {
      if (Array.isArray(ds)) ds.forEach(add)
      else for (let d = ds.from; d <= ds.to; d++) add(d)
    }
    return format(st)
  })

  return { statsText, shownRows, inWindows }
}

export type Totals = ReturnType<typeof useTotals>
