// Excel's Merge & Center and its list (Merge Across, Merge Cells, Unmerge Cells) over the selection through the source's
// `structure`, with Excel's warning before cells holding more than one value are merged: the library keeps the top-left one.
import { computed } from 'vue'
import { cellAddress, rangeAddress, rectContains, rectsMeet } from 'xlcalc/a1'
import type { CellRect, CellValue } from 'xlcalc/contract'
import type { EditOutcome, SheetOp } from '../source'
import type { Commands } from './commands'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { AlertAnswer, SheetAlert } from './validation'
import type { Windows } from './windows'

/** Excel's merge commands: the button's Merge & Center, and Merge Across, Merge Cells and Unmerge Cells of its list. */
export type MergeCommand = 'mergeCenter' | 'mergeAcross' | 'mergeCells' | 'unmergeCells'

interface MergeDeps {
  /** Show an alert and wait for the answer. */
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
}

const holds = (text: string | undefined, value: CellValue | undefined, formula: string | undefined): boolean =>
  formula !== undefined || (text ?? '') !== '' || (value !== null && value !== undefined && value !== '')

export function useMerge(s: SheetState, w: Windows, sel: Selection, cmd: Commands, deps: MergeDeps) {
  /** The commands act: the source has `structure`, the sheet is editable, and no display filter scatters its rows. */
  const available = computed(() => !!s.sheet.value && cmd.can('structure') && !s.visible.value)
  /** The merged area the active cell is in, from the loaded windows: Excel's button is pressed on it. */
  const area = computed<CellRect | null>(() => {
    const a = s.selected.value
    if (!a) return null
    return w.merges.value.find((m) => rectContains(m, { top: a.row, left: a.col, bottom: a.row, right: a.col })) ?? null
  })

  /** The rectangle grown by the loaded merged areas it meets, again for each it then meets, as the library grows it. */
  function grown(r: CellRect): CellRect {
    let out = r
    for (let again = true; again; ) {
      again = false
      for (const m of w.merges.value) {
        if (!rectsMeet(out, m) || rectContains(out, m)) continue
        out = { top: Math.min(out.top, m.top), left: Math.min(out.left, m.left), bottom: Math.max(out.bottom, m.bottom), right: Math.max(out.right, m.right) }
        again = true
      }
    }
    return out
  }

  /**
   * Whether a merge loses values, as Excel warns: an area (each row of it with `across`) where more than one cell holds a
   * value or a formula. Cells past the sheet's data are empty; the loaded windows answer when they hold the rectangle,
   * else the source's windows over it, a tile of `maxCells` at a time. A window not to be had warns, to be safe.
   */
  async function loses(sheet: string, r: CellRect, across: boolean): Promise<boolean> {
    const bottom = Math.min(r.bottom, s.dataRows.value)
    const right = Math.min(r.right, s.dataCols.value)
    if (bottom < r.top || right < r.left) return false
    const counts = new Map<number, number>()
    /** One more cell holding something in the row's area: true once an area has two. */
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
    const max = Math.max(1, source.maxCells ?? 20_000)
    const step = Math.min(right - r.left + 1, max)
    for (let col = r.left; col <= right; col += step) {
      const cols = Math.min(step, right - col + 1)
      for (let row = r.top; row <= bottom; ) {
        const rows = Math.min(bottom - row + 1, Math.max(1, Math.floor(max / cols)))
        let got
        try {
          got = await source.window({ sheet, row, rows, col, cols })
        } catch {
          return true
        }
        if ('pending' in got || !got.text.length) return true
        for (let i = 0; i < got.text.length; i++)
          for (let j = 0; j < cols; j++) if (holds(got.text[i]?.[j], got.values[i]?.[j], got.formulas[cellAddress(row + i, col + j)]) && add(row + i)) return true
        row += got.text.length
      }
    }
    return false
  }

  /**
   * A merge command over the selection, as Excel's: none on one cell outside a merged area; before a merge that loses values
   * Excel's warning, whose Cancel sends nothing. The source's outcome, or null when nothing was sent.
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
    if (await loses(sheet, grown(r), across)) {
      const answer = await deps.ask({ severity: 'merge', message: s.messages.value.mergeWarning, sheet, cell: cellAddress(a.row, a.col) })
      if (answer !== 'accept' || sheet !== s.sheetName.value) return null
    }
    const op: SheetOp = { op: 'mergeCells', sheet, range, ...(across ? { across: true } : {}), ...(command === 'mergeCenter' ? { center: true } : {}) }
    return cmd.structure(op)
  }

  return { available, area, run }
}

export type Merge = ReturnType<typeof useMerge>
