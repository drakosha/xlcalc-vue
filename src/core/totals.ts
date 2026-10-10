// Copying the selection as TSV and its totals for the status line, both from the loaded windows where they hold it.
import { computed } from 'vue'
import type { CellValue } from 'xlcalc/contract'
import { toTsv } from './clipboard'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

export function useTotals(s: SheetState, w: Windows, sel: Selection, notify: (text: string) => void) {
  /** Whether the loaded windows hold the whole rectangle: each of its rows lies across one of them. */
  function inWindows(r1: number, c1: number, r2: number, c2: number): boolean {
    for (let row = r1; row <= r2; row++) {
      const slot = w.slotAt(row, c1)
      if (!slot || c2 >= slot.win!.col + slot.cols) return false
    }
    return true
  }

  /** Totals as Excel's status line: count of non-empty cells, and over numbers sum, average, min and max. */
  const statsText = computed(() => {
    const r = sel.selRect.value
    if (!r || !sel.multi.value) return ''
    const m = s.messages.value
    const fmt = new Intl.NumberFormat(s.locale.value, { maximumFractionDigits: 2 })
    const r2 = Math.min(r.r2, s.dataRows.value)
    const c2 = Math.min(r.c2, s.dataCols.value)
    if (r2 >= r.r1 && c2 >= r.c1 && !inWindows(r.r1, r.c1, r2, c2)) {
      return m.statsCells(fmt.format((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1)))
    }
    let cells = 0
    let numbers = 0
    let sum = 0
    let min = Infinity
    let max = -Infinity
    for (let row = r.r1; row <= r2; row++) {
      for (let col = r.c1; col <= c2; col++) {
        const value: CellValue | undefined = w.valueAt(row, col)
        if (w.textAt(row, col) === '' && (value === null || value === undefined || value === '')) continue
        cells++
        if (typeof value !== 'number') continue
        numbers++
        sum += value
        if (value < min) min = value
        if (value > max) max = value
      }
    }
    const parts = [m.statsCount(fmt.format(cells))]
    if (numbers) {
      parts.push(m.statsSum(fmt.format(sum)), m.statsAverage(fmt.format(sum / numbers)))
      if (numbers > 1) parts.push(m.statsMin(fmt.format(min)), m.statsMax(fmt.format(max)))
      if (numbers !== cells) parts.push(m.statsNumbers(fmt.format(numbers)))
    }
    return parts.join(' · ')
  })

  /** Copy the selection as TSV: from the loaded windows, else one window asked for it (cut by the cell limit). */
  async function onCopy(e: ClipboardEvent): Promise<void> {
    const r = sel.selRect.value
    if (!r || e.target !== s.scroller.value) return
    e.preventDefault()
    const r2 = Math.min(r.r2, Math.max(r.r1, s.dataRows.value))
    const c2 = Math.min(r.c2, Math.max(r.c1, s.dataCols.value))
    const grid = (cell: (row: number, col: number) => string, lastRow: number, lastCol: number) => {
      const out: string[][] = []
      for (let row = r.r1; row <= lastRow; row++) {
        const line: string[] = []
        for (let col = r.c1; col <= lastCol; col++) line.push(cell(row, col))
        out.push(line)
      }
      return out
    }
    if (inWindows(r.r1, r.c1, r2, c2)) {
      e.clipboardData?.setData('text/plain', toTsv(grid(w.textAt, r2, c2)))
      return
    }
    try {
      const got = await s.source().window({ sheet: s.sheetName.value, row: r.r1, rows: r2 - r.r1 + 1, col: r.c1, cols: c2 - r.c1 + 1 })
      if ('pending' in got) throw new Error('pending')
      const lastRow = got.row + got.text.length - 1
      const lastCol = got.col + (got.text[0]?.length ?? 0) - 1
      const text = toTsv(grid((row, col) => got.text[row - got.row]?.[col - got.col] ?? '', Math.min(r2, lastRow), Math.min(c2, lastCol)))
      try {
        await navigator.clipboard.writeText(text)
        if (lastRow < r2 || lastCol < c2) notify(s.messages.value.copyPartial)
      } catch {
        notify(s.messages.value.clipboardDenied)
      }
    } catch {
      notify(s.messages.value.copyFailed)
    }
  }

  return { statsText, onCopy }
}
