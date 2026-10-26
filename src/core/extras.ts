// Changes to what cells hold beside their values (Clear Formats and Clear All, hyperlinks, notes): each through the source's
// method as Ctrl+Z undoes it, then the windows over what it changed read again in place, since its outcome carries values only.
import { MAX_COLS, rangeAddress } from 'xlcalc/a1'
import type { CellRect } from 'xlcalc/contract'
import type { EditOutcome } from '../contract/source'
import type { ClearKind } from '../contract/workbook'
import type { Editing } from './editing'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

export interface ExtrasDeps {
  notify: (text: string) => void
  /** A change succeeded that may have changed the cells' rules too (Clear All removes data validation). */
  changed: () => void
}

/** The source's methods for what cells hold beside their values. */
type ExtraMethod = 'clear' | 'link' | 'note'

export function useExtras(s: SheetState, sel: Selection, w: Windows, ed: Editing, deps: ExtrasDeps) {
  /** A method the source accepts: the sheet is editable and the source has it; otherwise null. */
  const method = <K extends ExtraMethod>(name: K) => (s.editable.value ? s.sourceMethod(name) : null)

  /** Calls the source; after a success the windows over `ranges` of `sheet` are read again, whether or not the source sent `repaint`. */
  async function change(call: () => Promise<EditOutcome>, sheet: string, ranges: string[]): Promise<EditOutcome> {
    const outdated = s.outdated()
    const out = await ed.run(call, [], null, null)
    if (out.ok && !outdated()) {
      w.repaint(ranges.map((range) => ({ sheet, range })))
      deps.changed()
    }
    return out
  }

  /**
   * The selection's areas as Excel writes them (`B2:D5`, `B:D`, `2:5`); under a display filter, the runs of its visible rows,
   * since Excel changes only the visible cells of a filtered range.
   */
  function areas(): string[] {
    const full = sel.fullRect.value
    const shown = sel.selRect.value
    if (!full || !shown) return []
    if (!s.visible.value) return [sel.selectionRef.value]
    const out: string[] = []
    const run = (top: number, bottom: number) => out.push(rangeAddress({ top, bottom, left: full.left, right: full.right >= MAX_COLS ? MAX_COLS : full.right }))
    let start = 0
    const last = Math.min(shown.bottom, Math.max(1, s.dataRows.value))
    for (let row = shown.top; row <= last + 1; row++) {
      const on = row <= last && s.indexOf(row) !== undefined
      if (on && !start) start = row
      else if (!on && start) {
        run(start, row - 1)
        start = 0
      }
    }
    return out
  }

  /** A locked cell in the rectangle: its reason is displayed and nothing is sent. */
  function locked(rect: CellRect | null): boolean {
    const lock = rect ? ed.lockIn(rect) : null
    if (lock !== null) deps.notify(s.messages.value.locked(lock))
    return lock !== null
  }

  /** Excel's Home > Clear > Clear Formats or Clear All over the selection, as a single undo step. */
  async function clear(what: ClearKind): Promise<EditOutcome | null> {
    const call = method('clear')
    if (!call || !s.sheet.value || locked(sel.selRect.value)) return null
    const ranges = areas()
    if (!ranges.length) return null
    const sheet = s.sheetName.value
    return change(() => call(sheet, ranges, what), sheet, ranges)
  }

  return { method, change, locked, clear }
}

export type Extras = ReturnType<typeof useExtras>
