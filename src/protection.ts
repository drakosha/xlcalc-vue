// Excel's sheet protection without the engine: what a protected sheet allows (the boxes of its Protect Sheet dialog, with
// their defaults) and whether a cell's style locks it or hides its formula; shared by the sheet and the local source.
import type { CellStyle, SheetProtection } from 'xlcalc/contract'

/** What users may do on a protected sheet, as a box of Excel's Protect Sheet dialog. */
export type ProtectionAllowance = keyof SheetProtection

/** The boxes of Excel's Protect Sheet dialog in its order, each with its default: selecting cells is allowed, the rest is not. */
export const PROTECTION_DEFAULTS: Readonly<Required<SheetProtection>> = {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  objects: false,
  scenarios: false,
}
export const PROTECTION_BOXES = Object.keys(PROTECTION_DEFAULTS) as ProtectionAllowance[]

/** Whether a sheet with this protection lets users do `what`: an unprotected sheet (none) allows everything. */
export const allows = (protection: SheetProtection | null | undefined, what: ProtectionAllowance): boolean =>
  !protection || (protection[what] ?? PROTECTION_DEFAULTS[what])

/** A protection as the workbook holds it: only the boxes that differ from the defaults (`{}` is Excel's default protection). */
export function protectionOf(boxes: Readonly<Record<ProtectionAllowance, boolean>>): SheetProtection {
  const out: SheetProtection = {}
  for (const k of PROTECTION_BOXES) if (boxes[k] !== PROTECTION_DEFAULTS[k]) out[k] = boxes[k]
  return out
}

/** A cell with this style is locked, as every cell is unless its style unlocks it (Excel's Format Cells, Protection, Locked). */
export const lockedStyle = (style: CellStyle | undefined): boolean => style?.protection?.locked !== false
/** A cell with this style hides its formula on a protected sheet (Protection, Hidden). */
export const hiddenStyle = (style: CellStyle | undefined): boolean => style?.protection?.hidden === true
