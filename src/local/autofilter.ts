// The range each sheet's AutoFilter acts on, as Excel reads it after data typed under it (xlcalc 0.67's
// `getAutoFilterRange`): listed with the sheets, and read again after an edit to tell whether the sheet must learn it.
import type { AsyncSession, SheetLayout } from 'xlcalc'
import { nameKey } from 'xlcalc/a1'

export function autoFilterRanges() {
  /** The ranges last listed, by the sheet's name key: those of sheets with an AutoFilter. */
  const listed = new Map<string, { sheet: string; range: string }>()

  /** The range a sheet's AutoFilter acts on: its stored range down to the last row of its current region. */
  async function rangeOf(s: AsyncSession, sheet: SheetLayout): Promise<string | undefined> {
    const stored = sheet.autoFilter?.range
    if (!stored) return undefined
    return (await s.getAutoFilterRange(sheet.name)) ?? stored
  }

  /** The sheets' ranges for the list the component reads, remembered to compare after an edit. */
  async function list(s: AsyncSession, sheets: SheetLayout[]): Promise<(string | undefined)[]> {
    const got = await Promise.all(sheets.map((x) => rangeOf(s, x)))
    listed.clear()
    sheets.forEach((x, i) => got[i] && listed.set(nameKey(x.name), { sheet: x.name, range: got[i] }))
    return got
  }

  /** Whether cell changes moved the range of an AutoFilter on their sheets; the new ranges are remembered. */
  async function moved(s: AsyncSession, changes: { sheet: string }[]): Promise<boolean> {
    const touched = new Set(changes.map((c) => nameKey(c.sheet)).filter((k) => listed.has(k)))
    let any = false
    for (const k of touched) {
      const was = listed.get(k)!
      const now = await s.getAutoFilterRange(was.sheet)
      if (!now || now === was.range) continue
      listed.set(k, { sheet: was.sheet, range: now })
      any = true
    }
    return any
  }

  return { list, moved }
}
