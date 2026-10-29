// Excel's sheet protection on the sheet: what the displayed protected sheet allows, which loaded cells are locked or hide
// their formula, Excel's alert for a change it refuses, which cells Excel lets the user select, and Protect Sheet and
// Unprotect Sheet through the source's `protect`. The source enforces protection; the sheet disables what Excel disables.
import { computed, ref } from 'vue'
import type { CellRect } from 'xlcalc/contract'
import { allows, hiddenStyle, lockedStyle, PROTECTION_DEFAULTS, protectionOf, type ProtectionAllowance } from '../protection'
import type { EditOutcome } from '../contract/source'
import type { SheetProtection } from 'xlcalc/contract'
import type { AlertAnswer, SheetAlert } from './validation'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** The most cells of a range checked among the loaded windows before a change is sent; past them the source decides. */
const MAX_CHECKED = 200_000

/** The Protect Sheet dialog: the sheet it protects, its first box (Protect the sheet and contents of locked cells) and the list's boxes. */
export interface ProtectDraft {
  sheet: string
  contents: boolean
  boxes: Record<ProtectionAllowance, boolean>
}

/** Changes the sheet's protection gates: each needs its box, or none allows it (null). */
const NEEDS: Record<string, ProtectionAllowance | null> = { format: 'formatCells', link: 'insertHyperlinks', note: 'objects', outline: null, table: null, rules: null, merge: null, setAutoFilter: null }

export interface ProtectionDeps {
  ask: (alert: SheetAlert) => Promise<AlertAnswer>
  /** Protects or unprotects a sheet through the source, after which the sheet reloads; null when nothing was sent. */
  protect: (sheet: string, protection: SheetProtection | null) => Promise<EditOutcome | null>
}

export function useProtection(s: SheetState, w: Windows, deps: ProtectionDeps) {
  /** The displayed sheet's protection; null when it is not protected. */
  const protection = computed(() => s.sheet.value?.protection ?? null)
  /** Whether the displayed sheet allows `what`: an unprotected one allows everything. */
  const can = (what: ProtectionAllowance): boolean => allows(protection.value, what)
  /** Whether the protection disables a kind of change (`NEEDS`): the source's method, or the merge commands. */
  const blocks = (kind: string): boolean => !!protection.value && kind in NEEDS && (NEEDS[kind] === null || !can(NEEDS[kind]!))

  const styleAt = (row: number, col: number) => w.lookOfRow(row, col)?.style
  /** A cell is locked on the protected sheet (its style does not unlock it); false on an unprotected sheet. */
  const lockedAt = (row: number, col: number): boolean => !!protection.value && lockedStyle(styleAt(row, col))
  /** A cell's formula is hidden on the protected sheet: the formula bar shows nothing, the cell its value. */
  const hiddenAt = (row: number, col: number): boolean => !!protection.value && hiddenStyle(styleAt(row, col))
  /** Whether a loaded cell of `rect` is locked; cells no window holds are left to the source, which refuses them. */
  function lockedIn(rect: CellRect): boolean {
    if (!protection.value) return false
    const bottom = Math.min(rect.bottom, s.rowCount.value)
    const right = Math.min(rect.right, s.colCount.value)
    let left = MAX_CHECKED
    for (let row = rect.top; row <= bottom && left > 0; row++) {
      for (let col = rect.left; col <= right && left-- > 0; col++) if (w.lookOfRow(row, col) && lockedAt(row, col)) return true
    }
    return false
  }
  /** Excel's alert for a change the protected sheet does not allow. */
  const alert = (): void => void deps.ask({ severity: 'notice', message: s.messages.value.protectedAlert })
  /** A locked cell in `rect`: Excel's alert, and true; nothing is sent. */
  function refuse(rect: CellRect): boolean {
    if (!lockedIn(rect)) return false
    alert()
    return true
  }
  /** Whether Excel lets the user select a cell: a locked one with Select locked cells, an unlocked one with Select unlocked cells. */
  const selectable = (row: number, col: number): boolean => !protection.value || (lockedAt(row, col) ? can('selectLockedCells') : can('selectUnlockedCells'))
  /** The protection limits which cells can be selected. */
  const limitsSelection = computed(() => !!protection.value && !(can('selectLockedCells') && can('selectUnlockedCells')))

  // ── Protect Sheet and Unprotect Sheet ──
  const available = computed(() => s.editable.value && !!s.sourceMethod('protect') && !!s.sheet.value)
  const draft = ref<ProtectDraft | null>(null)
  /** Opens Protect Sheet for the displayed sheet with Excel's defaults; false when it cannot be protected (or already is). */
  function open(): boolean {
    if (!available.value || protection.value) return false
    draft.value = { sheet: s.sheetName.value, contents: true, boxes: { ...PROTECTION_DEFAULTS } }
    return true
  }
  function close(): void {
    draft.value = null
    s.focusGrid()
  }
  /** OK (disabled while the first box is unchecked, as in Excel): the sheet protected with the boxes as checked. */
  function apply(): void {
    const d = draft.value
    if (!d) return
    close()
    void deps.protect(d.sheet, protectionOf(d.boxes))
  }
  /** Unprotect Sheet: no password, since the workbook holds none. */
  function unprotect(): Promise<EditOutcome | null> {
    if (!available.value || !protection.value) return Promise.resolve(null)
    return deps.protect(s.sheetName.value, null)
  }

  return { protection, can, blocks, lockedAt, hiddenAt, lockedIn, alert, refuse, selectable, limitsSelection, available, draft, open, close, apply, unprotect }
}

export type Protection = ReturnType<typeof useProtection>
