// Locations as text: what the Name Box, a hyperlink, `focus` and a defined name's formula refer to, and the selection as Define Name writes it.
import { quoteSheetIfNeeded, rangeAddress } from 'xlcalc/a1'
import type { CellRect, FormulaReference, GridReference } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import { isGridRef } from '../grid-ref'

/** A location from the Name Box, a hyperlink or `focus`: a range with its sheet (`B12`, `'Sheet 2'!C7:D9`, `B:B`) or a sheet alone; null for anything else (`B12+C1`). */
export function parsePlace(text: string): { sheet?: string; range?: CellRect } | null {
  const raw = text.trim()
  if (!raw) return null
  const whole = (refs: FormulaReference[], end: number): GridReference | null => {
    const r = refs.length === 1 ? refs[0]! : null
    return r && isGridRef(r) && r.at === 0 && r.end === end && !r.spill ? r : null
  }
  const refs = scanReferences(raw)
  if (refs.length) {
    const r = whole(refs, raw.length)
    return r && { sheet: r.sheet, range: r }
  }
  // A sheet alone is parsed as a reference to its first cell, with its quotes as a reference has them.
  const sheet = whole(scanReferences(`${raw}!A1`), raw.length + 3)?.sheet
  return sheet === undefined ? null : { sheet }
}

/** A range as Excel's Define Name writes the selection: absolute (`$B$2:$D$9`, `$B:$B`, `$2:$5`). */
export function absoluteRange(rect: CellRect): string {
  return rangeAddress(rect)
    .split(':')
    .map((part) => part.replace(/^([A-Z]*)(\d*)$/, (_, col: string, row: string) => `${col ? `$${col}` : ''}${row ? `$${row}` : ''}`))
    .join(':')
}

/** A reference to a range of a sheet as a name's formula (without `=`): `'Sheet 2'!$B$2:$D$9`. */
export const rangeRef = (sheet: string, rect: CellRect): string => `${quoteSheetIfNeeded(sheet)}!${absoluteRange(rect)}`
