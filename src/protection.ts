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

/** What is done to a slicer or a timeline: its choice or settings (`choose`), its place (`place`), adding or removing it. */
export type ObjectChange = 'choose' | 'place' | 'add' | 'remove'

/**
 * A slicer's or a timeline's two locks: `position`, Excel's Disable resizing and moving, which no move or resize passes
 * on any sheet; `sheet`, Excel's Locked (true by default), which a protected sheet holds unless it allows Edit objects.
 */
export function objectLocks(object: object): { position: boolean; sheet: boolean } {
  const o = object as { locked?: boolean; positionLocked?: boolean }
  return { position: o.positionLocked === true, sheet: o.locked !== false }
}

/**
 * The boxes a change of a slicer or a timeline needs on protected sheets, by Microsoft's Protect a worksheet (not measured in
 * Excel): Use PivotTable reports on the sheets of the PivotTables it filters, for adding or removing it on its own sheet too;
 * Edit objects on its sheet for adding it, and for moving, resizing or removing it while it is locked (`objectLocks`), as
 * slicers and timelines are when added. The sheet and the local source both ask this.
 */
export function objectNeeds(change: ObjectChange, at: { sheet: string; locked: boolean }, pivots: readonly { sheet: string }[]): { sheet: string; allow: 'pivotTables' | 'objects' }[] {
  const out: { sheet: string; allow: 'pivotTables' | 'objects' }[] = pivots.map((p) => ({ sheet: p.sheet, allow: 'pivotTables' }))
  if (change === 'add' || change === 'remove') out.push({ sheet: at.sheet, allow: 'pivotTables' })
  if (change === 'add' || ((change === 'place' || change === 'remove') && at.locked)) out.push({ sheet: at.sheet, allow: 'objects' })
  return out
}
