// Excel's Insert > PivotTable: its dialog (PivotTable from table or range) with the table at the active cell, the selection
// or the active cell's current region, a new sheet or a location on an existing one; OK sends xlcalc's `addPivotTable`
// through the source's `structure` and selects the PivotTable's top-left cell, where the PivotTable Fields pane shows.
import { computed, ref, watch } from 'vue'
import { cellAddress, quoteSheetIfNeeded, sameName } from 'xlcalc/a1'
import type { PivotSource } from 'xlcalc/contract'
import type { SheetOp } from '../contract/source'
import type { Commands } from './commands'
import { parsePlace, rangeRef } from './place'
import { useSheetReads, type SheetState } from './state'
import { tableNames, type Tables } from './tables'

/** The dialog: the source as typed, a new sheet or an existing one with its location, and the alert for what OK refused. */
export interface PivotCreateDraft {
  source: string
  place: 'new' | 'existing'
  location: string
  error: string
}

export interface PivotCreateDeps {
  tables: Tables
  /** Excel's name for a new sheet (`Sheet4`). */
  nextSheetName: () => string
  /** Goes to a location, switching sheets, as Go To does without its history. */
  goPlace: (text: string) => Promise<boolean>
}

type AddPivot = Extract<SheetOp, { op: 'addPivotTable' }>

/** Excel puts a PivotTable on a new sheet at A3. */
const NEW_SHEET_CELL = 'A3'
/** How long OK waits for the sheets read again after the PivotTable is added before it goes to its cell anyway. */
const RELIST_MS = 10_000

/**
 * The source typed in the dialog: a range with its sheet or on `sheet` (`Data!$A$1:$E$40`, `A1:E40`), a table of the
 * workbook by its name, or a defined name; null for text that is none of these.
 */
export function typedSource(text: string, sheet: string, tables: string[]): PivotSource | null {
  const raw = text.trim().replace(/^=/, '')
  if (!raw) return null
  const table = tables.find((t) => sameName(t, raw))
  if (table) return { type: 'table', table }
  const place = parsePlace(raw)
  if (place?.range) return place.range.top === place.range.bottom && place.range.left === place.range.right ? null : { type: 'range', sheet: place.sheet ?? sheet, range: rangeOf(place.range) }
  return /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(raw) ? { type: 'name', name: raw } : null
}

/** The location typed in the dialog: the top-left cell of a range with its sheet or on `sheet`; null for text that is no range. */
export function typedLocation(text: string, sheet: string): { sheet: string; cell: string } | null {
  const place = parsePlace(text.trim().replace(/^=/, ''))
  return place?.range ? { sheet: place.sheet ?? sheet, cell: cellAddress(place.range.top, place.range.left) } : null
}

const rangeOf = (r: { top: number; left: number; bottom: number; right: number }) => `${cellAddress(r.top, r.left)}:${cellAddress(r.bottom, r.right)}`

export function usePivotCreate(s: SheetState, cmd: Commands, deps: PivotCreateDeps) {
  const read = useSheetReads(s)
  const draft = ref<PivotCreateDraft | null>(null)
  /** Insert PivotTable is offered: the source has `structure` and the sheet is editable. */
  const available = computed(() => !!s.sheet.value && cmd.can('structure'))

  /** The source Excel's dialog proposes: the table at the active cell by its name, otherwise the selection or the active cell's current region; empty for an empty cell. */
  async function guessSource(signal: AbortSignal): Promise<string | null> {
    const table = deps.tables.activeTable.value
    if (table) return table.name
    const got = await deps.tables.guess(signal)
    if (!got) return null
    const { rect, cells } = got
    // an empty cell alone has no region: the cells read around it may hold data elsewhere
    return rect.top === rect.bottom && rect.left === rect.right && !cells.has(cellAddress(rect.top, rect.left)) ? '' : rangeRef(s.sheetName.value, rect)
  }

  /** Excel's Insert > PivotTable: the dialog with the source it proposes, a new sheet chosen; false where the sheet cannot take a PivotTable. */
  async function open(): Promise<boolean> {
    if (!available.value || !s.selected.value) return false
    const sheet = s.sheetName.value
    const source = await guessSource(read())
    if (source === null || sheet !== s.sheetName.value) return false
    draft.value = { source, place: 'new', location: '', error: '' }
    return true
  }

  function close(): void {
    draft.value = null
    s.focusGrid()
  }

  /**
   * OK: the PivotTable over the source, on a new sheet before the active one at A3 or at the location's top-left cell, as one
   * undo step; the dialog stays with Excel's alert for a source or a location it does not take. Over cells that hold
   * something the source makes it showing `#SPILL!`, as Excel 365; its refusal (a place over another PivotTable, no row
   * under the headers) is shown as the sheet shows any refusal.
   */
  async function ok(): Promise<boolean> {
    const d = draft.value
    if (!d) return false
    const m = s.messages.value
    const sheet = s.sheetName.value
    const signal = read()
    const source = typedSource(d.source, sheet, d.source.includes('!') || /[$:]/.test(d.source) ? [] : await tableNames(s, signal))
    if (signal.aborted || draft.value !== d) return false
    if (!source) {
      d.error = m.pivotSourceBad
      return false
    }
    let op: AddPivot
    if (d.place === 'existing') {
      const at = typedLocation(d.location, sheet)
      if (!at || !s.sheets.value.some((x) => sameName(x.name, at.sheet))) {
        d.error = m.pivotLocationBad
        return false
      }
      op = { op: 'addPivotTable', sheet: at.sheet, cell: at.cell, source }
    } else op = { op: 'addPivotTable', sheet: deps.nextSheetName(), newSheet: { before: sheet }, source }
    draft.value = null
    s.focusGrid()
    const listed = s.sheets.value
    const out = await cmd.structure(op)
    if (!out?.ok) return false
    // the source's `reload` reads the sheets again after the call returns: the new sheet is there once they are
    if (s.sheets.value === listed) await relisted(listed)
    await deps.goPlace(`${quoteSheetIfNeeded(op.sheet)}!${op.cell ?? NEW_SHEET_CELL}`)
    return true
  }

  /** Resolves once the sheets list is read again (or after `RELIST_MS`), after the reload that read it has kept or changed the active sheet. */
  function relisted(listed: unknown): Promise<void> {
    return new Promise((resolve) => {
      const done = () => {
        stop()
        clearTimeout(timer)
        resolve()
      }
      const stop = watch(s.sheets, (now) => now !== listed && done())
      const timer = setTimeout(done, RELIST_MS)
    })
  }

  return { available, draft, open, close, ok }
}

export type PivotCreate = ReturnType<typeof usePivotCreate>
