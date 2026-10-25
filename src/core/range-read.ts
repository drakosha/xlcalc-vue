// A range's cells read from the source's windows, a tile of `maxCells` at a time: for dialogs that look at cells beyond the loaded windows.
import { cellAddress } from 'xlcalc/a1'
import type { CellRect, CellValue } from 'xlcalc/contract'
import type { SheetSource } from '../contract/source'
import { DEFAULT_MAX_CELLS, rowsFor } from '../window-size'

/** A cell as a window gives it: its displayed text, value, formula and edit text. */
export interface ReadCell {
  text: string
  value: CellValue
  formula?: string
  editText?: string
}

/** The cells of `r` by address (empty ones without text, value or formula left out); null when a window cannot be read or the workbook is still recalculating. */
export async function readRange(source: SheetSource, sheet: string, r: CellRect, signal?: AbortSignal): Promise<Map<string, ReadCell> | null> {
  const out = new Map<string, ReadCell>()
  const max = Math.max(1, source.maxCells ?? DEFAULT_MAX_CELLS)
  const step = Math.min(r.right - r.left + 1, max)
  for (let col = r.left; col <= r.right; col += step) {
    const cols = Math.min(step, r.right - col + 1)
    for (let row = r.top; row <= r.bottom; ) {
      const rows = rowsFor(max, r.bottom - row + 1, cols)
      let got
      try {
        got = await source.window({ sheet, row, rows, col, cols }, signal)
      } catch {
        return null
      }
      if (signal?.aborted || 'pending' in got) return null
      // a window shorter than asked ends at the sheet's last row
      if (!got.text.length) break
      for (let i = 0; i < got.text.length; i++)
        for (let j = 0; j < cols; j++) {
          const addr = cellAddress(row + i, col + j)
          const text = got.text[i]?.[j] ?? ''
          const value = got.values[i]?.[j] ?? null
          const formula = got.formulas[addr]
          if (text === '' && (value === null || value === '') && formula === undefined) continue
          const editText = got.editText?.[addr]
          out.set(addr, { text, value, ...(formula !== undefined ? { formula } : {}), ...(editText !== undefined ? { editText } : {}) })
        }
      row += got.text.length
    }
  }
  return out
}
