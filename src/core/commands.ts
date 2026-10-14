// Changes of the book the host or the grid's menus and keys ask for (format, structure, lines, outline, frozen panes, undo and redo): each goes
// through the source as Control+Z goes, and the sheet is read again after one a window's cells cannot carry.
import { ref } from 'vue'
import type { CellStyle } from 'xlcalc/contract'
import type { EditOutcome, FormatPatch, LayoutChange, OutlineChange, SheetOp } from '../source'
import { fontToggle, type FontToggle } from '../toolbar'
import { shortcutLetter, type Editing } from './editing'
import type { MergeCommand } from './merge'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A patch of `SheetView.format`: over its own range, or over the selection without one. */
export type SheetFormatPatch = Omit<FormatPatch, 'range'> & { range?: string }

/** Excel's Freeze Panes commands: the panes of the active cell, Unfreeze Panes, Freeze Top Row, Freeze First Column. */
export type FreezeMode = 'panes' | 'unfreeze' | 'topRow' | 'firstColumn'

/**
 * What `SheetView` exposes to its host (a template `ref`): changes made as the sheet's own keys make them, the outcome's
 * selection set without a flash or a step of the way back. Each gives the source's outcome, or null when nothing was sent
 * (the source has no such method, the sheet is read-only, or an undo is still out).
 */
export interface SheetViewApi {
  undo(): Promise<EditOutcome | null>
  redo(): Promise<EditOutcome | null>
  /** Formatting of the active sheet in one call of the source's `format`. */
  format(patches: SheetFormatPatch[]): Promise<EditOutcome | null>
  /** A structure operation of the source; the sheet shown after it is the one Excel shows. */
  structure(op: SheetOp): Promise<EditOutcome | null>
  /** An outline command of the source (Group, Ungroup, a group's button, a level button); the sheet is read again after it. */
  outline(change: OutlineChange): Promise<EditOutcome | null>
  /** Excel's Format Cells dialog (Control+1) over the selection, at the active cell's look; false when the sheet cannot format. */
  formatCells(): boolean
  /**
   * Excel's Freeze Panes through the source's `freeze`, one step of undo, the sheet read again after it: `panes` (the
   * default) keeps the rows above and the columns left of the active cell in view (from A1 the frame's middle, as Excel
   * splits it), `unfreeze` drops the book's panes, `topRow` and `firstColumn` freeze one line. A freeze shows whatever
   * the lock said before; null when nothing was sent (no `freeze`, a read-only sheet, nothing to freeze or unfreeze).
   */
  freezePanes(mode?: FreezeMode): Promise<EditOutcome | null>
  /** Excel's Reapply (Control+Alt+L): the display filter's conditions and sort again over the data as they are now; false with none in force. */
  reapplyFilter(): Promise<boolean>
  /** Excel's Trace Precedents as a panel beside the grid: the source's tree of the active cell's precedents, following the cell; false without the source's `explain`. */
  explain(): boolean
  /**
   * Excel's Merge & Center (`mergeCenter`), Merge Across, Merge Cells and Unmerge Cells over the selection through the
   * source's `structure`, one step of undo; Excel's warning first when more than one cell would lose its value. null when
   * nothing was sent: no `structure`, a read-only sheet, a display filter, one cell outside a merged area, the warning's Cancel.
   */
  merge(command: MergeCommand): Promise<EditOutcome | null>
}

/** The sheet to show once the sheets are read again: by name, keeping the selection or starting over on it. */
export interface SheetTarget {
  name: string
  keep: boolean
}

export interface CommandsDeps {
  /** Read the sheets and windows again; how many times it has started, to tell whether the source sent `reload` during a call. */
  reload: () => Promise<void>
  reloads: () => number
  /** The sheet the next read shows. */
  target: (t: SheetTarget | null) => void
  /** The style the active cell shows. */
  activeStyle: () => CellStyle | undefined
  /** Whether the book freezes panes on the sheet shown, the panes Excel freezes from A1, and the view after a freeze made. */
  frozen: () => boolean
  middlePanes: () => { rows: number; cols: number }
  showFreeze: (sheet: string, on: boolean) => void
  /** The display filter's Reapply and the precedents' panel, for the host's methods. */
  reapplyFilter: () => Promise<boolean>
  explain: () => boolean
  /** Excel's merge commands over the selection. */
  merge: (command: MergeCommand) => Promise<EditOutcome | null>
  /** A structure operation went through: lines may have moved, so what the sheet keeps by line number no longer holds. */
  structured: () => void
}

const FONT_KEYS: Record<string, FontToggle> = { b: 'bold', i: 'italic', u: 'underline' }

export function useCommands(s: SheetState, sel: Selection, ed: Editing, deps: CommandsDeps) {
  /** A change the source takes: the sheet is editable (not read-only, with `edit`) and the source has the method. */
  const can = (method: 'format' | 'structure' | 'layout' | 'outline' | 'freeze'): boolean => s.editable.value && typeof s.source()[method] === 'function'

  /** A change the windows cannot carry: the source's outcome, then the sheet read again unless the source sent `reload` meanwhile; formatting and lines leave the copy mode on (`endsCopy`). */
  async function change(call: () => Promise<EditOutcome>, target: SheetTarget | null = null, endsCopy = true): Promise<EditOutcome> {
    const before = deps.reloads()
    deps.target(target)
    const out = await ed.run(call, [], null, null, endsCopy)
    if (!out.ok) deps.target(null)
    else if (deps.reloads() === before) await deps.reload()
    return out
  }

  async function format(patches: SheetFormatPatch[]): Promise<EditOutcome | null> {
    const call = s.source().format
    const range = sel.selectionRef.value
    if (!call || !can('format') || !patches.length || !s.sheet.value) return null
    const sheet = s.sheetName.value
    const list: FormatPatch[] = []
    for (const p of patches) {
      if (p.range === undefined && !range) return null
      list.push({ ...p, range: p.range ?? range })
    }
    return change(() => call.call(s.source(), sheet, list), null, false)
  }

  /** A structure operation; a paste of the copied range keeps the copy mode on (`endsCopy` false), as Excel's does. */
  async function structure(op: SheetOp, endsCopy = true): Promise<EditOutcome | null> {
    const call = s.source().structure
    if (!call || !can('structure')) return null
    // Excel shows a sheet it inserts or unhides, and keeps the selection of the active sheet it renames.
    let target: SheetTarget | null = null
    if (op.op === 'insertSheet') target = { name: op.name, keep: false }
    else if (op.op === 'unhideSheet') target = { name: op.sheet, keep: false }
    else if (op.op === 'renameSheet' && op.sheet === s.sheetName.value) target = { name: op.name, keep: true }
    const out = await change(() => call.call(s.source(), op), target, endsCopy)
    if (out.ok) deps.structured()
    return out
  }

  async function layout(edit: LayoutChange): Promise<EditOutcome | null> {
    const call = s.source().layout
    if (!call || !can('layout')) return null
    return change(() => call.call(s.source(), edit), null, false)
  }

  async function outline(command: OutlineChange): Promise<EditOutcome | null> {
    const call = s.source().outline
    if (!call || !can('outline')) return null
    return change(() => call.call(s.source(), command), null, false)
  }

  /**
   * Excel's Freeze Panes commands on the active sheet; a selection of whole columns or rows freezes left of or above it, as its
   * active cell is in the first row or column. None under the display filter, whose rows on screen are not the sheet's.
   */
  async function freeze(mode: FreezeMode = 'panes'): Promise<EditOutcome | null> {
    const call = s.source().freeze
    const a = s.selected.value
    if (!call || !can('freeze') || !s.sheet.value || !a || s.visible.value) return null
    let panes: { rows: number; cols: number } | null = null
    if (mode === 'topRow') panes = { rows: 1, cols: 0 }
    else if (mode === 'firstColumn') panes = { rows: 0, cols: 1 }
    else if (mode === 'panes') panes = a.row === 1 && a.col === 1 ? deps.middlePanes() : { rows: a.row - 1, cols: a.col - 1 }
    if (panes ? panes.rows <= 0 && panes.cols <= 0 : !deps.frozen()) return null
    const sheet = s.sheetName.value
    const out = await change(() => call.call(s.source(), sheet, panes), null, false)
    if (out.ok) deps.showFreeze(sheet, panes !== null)
    return out
  }

  /** Control+B, Control+I and Control+U (Cmd on a Mac, any layout): the font toggled over the selection by the active cell; true when taken. */
  function fontKey(e: KeyboardEvent): boolean {
    if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || !can('format')) return false
    const what = FONT_KEYS[shortcutLetter(e)]
    if (!what) return false
    e.preventDefault()
    void format([{ style: fontToggle(deps.activeStyle(), what) }])
    return true
  }

  // ── Excel's Format Cells (Control+1): OK formats the selection in one call of `format` ──
  const formatOpen = ref(false)
  function openFormat(): boolean {
    if (!can('format') || !s.sheet.value || !sel.selectionRef.value) return false
    formatOpen.value = true
    return true
  }
  function closeFormat(): void {
    formatOpen.value = false
    s.scroller.value?.focus({ preventScroll: true })
  }

  const api: SheetViewApi = {
    undo: () => ed.stepHistory('undo'),
    redo: () => ed.stepHistory('redo'),
    format,
    structure: (op) => structure(op),
    outline,
    formatCells: openFormat,
    freezePanes: freeze,
    reapplyFilter: deps.reapplyFilter,
    explain: deps.explain,
    merge: deps.merge,
  }

  return { can, change, format, structure, layout, outline, freeze, fontKey, formatOpen, openFormat, closeFormat, api }
}

export type Commands = ReturnType<typeof useCommands>
