// The structure menu as Excel's: on row numbers and column letters (insert, delete, sizes, hide, unhide), on the sheet
// tabs (insert, rename, move, delete, hide, unhide) and "+" after them; a right click opens it, a finger's long press too.
import { sameName } from 'xlcalc/a1'
import type { Commands } from './commands'
import type { SheetState } from './state'

export function useMenu(s: SheetState, cmd: Commands) {
  // ── sheets ──
  /** The last number this sheet gave a new sheet: Excel counts on past one deleted since. */
  let issued = 0
  /** Excel's name for a new sheet: the interface's word and the next number past the sheets' count and their own numbers. */
  function nextSheetName(): string {
    const prefix = s.messages.value.sheetNamePrefix
    const own = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\d+)$`, 'i')
    let n = Math.max(issued, s.sheets.value.length, ...s.sheets.value.map((x) => Number(own.exec(x.name)?.[1] ?? 0)))
    do n++
    while (s.sheets.value.some((x) => sameName(x.name, `${prefix}${n}`)))
    issued = n
    return `${prefix}${n}`
  }
  /** "+": a new sheet at the end, shown once it is in. */
  const addSheet = () => cmd.structure({ op: 'insertSheet', name: nextSheetName() })

  return { nextSheetName, addSheet }
}

export type Menu = ReturnType<typeof useMenu>
