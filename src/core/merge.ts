// Excel's Merge & Center and its dropdown (Merge Across, Merge Cells, Unmerge Cells) over the selection through the source's
// `structure`, with Excel's warning before cells containing more than one value are merged: the library keeps the top-left one.
import { computed } from 'vue'
import { cellAddress, mergedArea, rangeAddress, rectContains, rectsMeet } from 'xlcalc/a1'
import type { CellRect, CellValue } from 'xlcalc/contract'
import type { EditOutcome, SheetOp } from '../contract/source'
import { cellRect } from '../rect'
import { isBlank } from '../stats'
import { DEFAULT_MAX_CELLS, rowsFor } from '../window-size'
import type { Commands } from './commands'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'
import type { AlertAnswer, SheetAlert } from './validation'
import type { Windows } from './windows'
import type { MergeCommand } from '../contract/view'

interface MergeDeps {
  /** Shows an alert and waits for the answer. */
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
}

const holds = (text: string | undefined, value: CellValue | undefined, formula: string | undefined): boolean => formula !== undefined || !isBlank(text, value)

export function useMerge(s: SheetState, w: Windows, sel: Selection, cmd: Commands, deps: MergeDeps) {
  const read = useSheetReads(s)
  /** The commands are available: the source has `structure`, the sheet is editable, and no display filter scatters its rows. */
  const available = computed(() => !!s.sheet.value && cmd.canRanges.value)
  /** The merged range containing the active cell, from the loaded windows: Excel's button shows as pressed on it. */
  const area = computed<CellRect | null>(() => {
    const a = s.selected.value
    if (!a) return null
    return w.merges.value.find((m) => rectContains(m, cellRect(a.row, a.col))) ?? null
  })

  /**
   * Whether a merge loses values, as Excel warns: a range (each row of it with `across`) where more than one cell contains a
   * value or a formula. Cells beyond the sheet's data are empty; the loaded windows answer when they hold the range,
   * otherwise the source's windows over it, a tile of `maxCells` at a time. A window that cannot be read warns, to be safe;
   * null once `signal` stopped the reads (another sheet, the sheet unmounted).
   */
  async function loses(sheet: string, r: CellRect, across: boolean, signal: AbortSignal): Promise<boolean | null> {
    const bottom = Math.min(r.bottom, s.dataRows.value)
    const right = Math.min(r.right, s.dataCols.value)
    if (bottom < r.top || right < r.left) return false
    const counts = new Map<number, number>()
    /** One more non-empty cell in the row's range: true once a range has two. */
    const add = (row: number): boolean => {
      const key = across ? row : 0
      const n = (counts.get(key) ?? 0) + 1
      counts.set(key, n)
      return n > 1
    }
    let loaded = true
    for (let row = r.top; row <= bottom && loaded; row++) {
      const d = s.indexOf(row)
      for (let col = r.left; col <= right; col++) {
        if (d === undefined || !w.slotAt(d, col)) {
          loaded = false
          break
        }
      }
    }
    if (loaded) {
      for (let row = r.top; row <= bottom; row++)
        for (let col = r.left; col <= right; col++) if (holds(w.textOfRow(row, col), w.valueOfRow(row, col), w.formulaOfRow(row, col)) && add(row)) return true
      return false
    }
    const source = s.source()
    const max = Math.max(1, source.maxCells ?? DEFAULT_MAX_CELLS)
    const step = Math.min(right - r.left + 1, max)
    for (let col = r.left; col <= right; col += step) {
      const cols = Math.min(step, right - col + 1)
      for (let row = r.top; row <= bottom; ) {
        const rows = rowsFor(max, bottom - row + 1, cols)
        let got
        try {
          got = await source.window({ sheet, row, rows, col, cols }, signal)
        } catch {
          return signal.aborted ? null : true
        }
        if (signal.aborted) return null
        if ('pending' in got || !got.text.length) return true
        for (let i = 0; i < got.text.length; i++)
          for (let j = 0; j < cols; j++) if (holds(got.text[i]?.[j], got.values[i]?.[j], got.formulas[cellAddress(row + i, col + j)]) && add(row + i)) return true
        row += got.text.length
      }
    }
    return false
  }

  /**
   * A merge command over the selection, as in Excel: none on a single cell outside a merged range; before a merge that loses
   * values, Excel's warning, whose Cancel sends nothing. The source's outcome, or null when nothing was sent.
   */
  async function run(command: MergeCommand): Promise<EditOutcome | null> {
    const r = sel.fullRect.value
    const a = s.selected.value
    if (!available.value || !r || !a) return null
    const sheet = s.sheetName.value
    const range = rangeAddress(r)
    if (r.top === r.bottom && r.left === r.right && !w.merges.value.some((m) => rectsMeet(m, r))) return null
    if (command === 'unmergeCells') return cmd.structure({ op: 'unmergeCells', sheet, range })
    const across = command === 'mergeAcross'
    // The range the merge covers: the selection extended by the loaded merged ranges it overlaps, as the library extends it.
    const lost = await loses(sheet, mergedArea(r, w.merges.value), across, read())
    if (lost === null) return null
    if (lost) {
      const answer = await deps.ask({ severity: 'merge', message: s.messages.value.mergeWarning, sheet, cell: cellAddress(a.row, a.col) })
      if (answer !== 'accept' || sheet !== s.sheetName.value) return null
    }
    const op: SheetOp = { op: 'mergeCells', sheet, range, ...(across ? { across: true } : {}), ...(command === 'mergeCenter' ? { center: true } : {}) }
    return cmd.structure(op)
  }

  return { available, area, run }
}

export type Merge = ReturnType<typeof useMerge>
