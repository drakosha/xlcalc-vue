// Copying the selection as TSV and its totals for the status line: the source's `stats` when it has them, else the loaded windows.
import { computed, shallowRef, watch } from 'vue'
import type { CellValue } from 'xlcalc/contract'
import type { RangeStats } from '../source'
import { toTsv } from './clipboard'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** A pause before asking the source for totals: a drag or Shift+arrows changes the selection many times a second. */
const STATS_DELAY_MS = 150

export function useTotals(s: SheetState, w: Windows, sel: Selection, notify: (text: string) => void) {
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
  const statsKey = () => `${s.sheetName.value}\u0000${sel.selectionRef.value}\u0000${s.visible.value?.key ?? ''}`
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
      const range = sel.selectionRef.value
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
    const r2 = Math.min(r.r2, s.dataRows.value)
    const c2 = Math.min(r.c2, s.dataCols.value)
    const ds = shownRows(r.r1, r2)
    if (r2 >= r.r1 && c2 >= r.c1 && !inWindows(ds, r.c1, c2)) {
      const fmt = new Intl.NumberFormat(s.locale.value)
      const rows = s.visible.value ? ds.length : r.r2 - r.r1 + 1
      return s.messages.value.statsCells(fmt.format(rows * (r.c2 - r.c1 + 1)))
    }
    const st: RangeStats = { cells: 0, numbers: 0, sum: 0, min: null, max: null }
    for (const d of ds) {
      for (let col = r.c1; col <= c2; col++) {
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

  /** Copy the selection as TSV: from the loaded windows, else one window asked for it (cut by the cell limit). */
  async function onCopy(e: ClipboardEvent): Promise<void> {
    const r = sel.selRect.value
    if (!r || e.target !== s.scroller.value) return
    e.preventDefault()
    const ds = shownRows(r.r1, Math.min(r.r2, Math.max(r.r1, s.dataRows.value)))
    const c2 = Math.min(r.c2, Math.max(r.c1, s.dataCols.value))
    if (!ds.length) return
    const grid = (cell: (d: number, col: number) => string, rows: number[], lastCol: number) =>
      rows.map((d) => {
        const line: string[] = []
        for (let col = r.c1; col <= lastCol; col++) line.push(cell(d, col))
        return line
      })
    if (inWindows(ds, r.c1, c2)) {
      e.clipboardData?.setData('text/plain', toTsv(grid(w.textAt, ds, c2)))
      return
    }
    const d1 = ds[0]!
    const d2 = ds[ds.length - 1]!
    try {
      const filter = s.visible.value?.spec
      const got = await s.source().window({ sheet: s.sheetName.value, row: d1, rows: d2 - d1 + 1, col: r.c1, cols: c2 - r.c1 + 1, ...(filter ? { filter } : {}) })
      if ('pending' in got) throw new Error('pending')
      const lastD = got.row + got.text.length - 1
      const lastCol = got.col + (got.text[0]?.length ?? 0) - 1
      const rows = ds.filter((d) => d <= lastD)
      const text = toTsv(grid((d, col) => got.text[d - got.row]?.[col - got.col] ?? '', rows, Math.min(c2, lastCol)))
      try {
        await navigator.clipboard.writeText(text)
        if (lastD < d2 || lastCol < c2) notify(s.messages.value.copyPartial)
      } catch {
        notify(s.messages.value.clipboardDenied)
      }
    } catch {
      notify(s.messages.value.copyFailed)
    }
  }

  return { statsText, onCopy }
}
