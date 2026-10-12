// The selection's totals for the status line: the source's `stats` when it has them, else the loaded windows.
import { computed, shallowRef, watch } from 'vue'
import type { CellValue } from 'xlcalc/contract'
import type { RangeStats } from '../source'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** A pause before asking the source for totals: a drag or Shift+arrows changes the selection many times a second. */
const STATS_DELAY_MS = 150

export function useTotals(s: SheetState, w: Windows, sel: Selection) {
  /** Display rows of the selection: under a filter only the shown rows between its corners, in display order. */
  function shownRows(r1: number, r2: number): number[] {
    const out: number[] = []
    const list = s.visible.value?.rows
    if (list) list.forEach((r, i) => r >= r1 && r <= r2 && out.push(i + 1))
    else for (let r = r1; r <= r2; r++) out.push(r)
    return out
  }
  /** Whether the loaded windows hold these display rows across the columns. */
  function inWindows(ds: number[], c1: number, c2: number): boolean {
    return ds.every((d) => {
      const slot = w.slotAt(d, c1)
      return !!slot && c2 < slot.win!.col + slot.cols
    })
  }
  const format = (st: RangeStats): string => {
    const m = s.messages.value
    const fmt = new Intl.NumberFormat(s.locale.value, { maximumFractionDigits: 2 })
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
      const stats = s.source().stats
      if (!key || !stats) return
      const sheet = s.sheetName.value
      const range = sel.gridRef.value
      const filter = s.visible.value?.spec
      statsTimer = setTimeout(() => {
        const ctrl = (statsCtrl = new AbortController())
        stats
          .call(s.source(), sheet, range, filter, ctrl.signal)
          .then((got) => {
            if (!ctrl.signal.aborted && statsKey() === key) fromSource.value = { key, stats: got }
          })
          .catch(() => {
            // Totals are an ornament: without them the loaded windows count.
          })
      }, STATS_DELAY_MS)
    },
  )

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
      const fmt = new Intl.NumberFormat(s.locale.value)
      const rows = s.visible.value ? ds.length : r.bottom - r.top + 1
      return s.messages.value.statsCells(fmt.format(rows * (r.right - r.left + 1)))
    }
    const st: RangeStats = { cells: 0, numbers: 0, sum: 0, min: null, max: null }
    for (const d of ds) {
      for (let col = r.left; col <= c2; col++) {
        const value: CellValue | undefined = w.valueAt(d, col)
        if (w.textAt(d, col) === '' && (value === null || value === undefined || value === '')) continue
        st.cells++
        if (typeof value !== 'number') continue
        st.numbers++
        st.sum += value
        st.min = st.min === null ? value : Math.min(st.min, value)
        st.max = st.max === null ? value : Math.max(st.max, value)
      }
    }
    return format(st)
  })

  return { statsText, shownRows, inWindows }
}

export type Totals = ReturnType<typeof useTotals>
