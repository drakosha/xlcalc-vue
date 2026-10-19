// Changes of the book the host or the grid's menus and keys ask for (format, structure, lines, outline, frozen panes, undo and redo): each goes
// through the source as Control+Z goes, and the sheet is read again after one a window's cells cannot carry.
import { computed, ref } from 'vue'
import type { CellStyle } from 'xlcalc/contract'
import type { EditOutcome, FormatPatch, LayoutChange, OutlineChange, SheetOp } from '../contract/source'
import { fontToggle, type FontToggle } from '../toolbar'
import type { Editing } from './editing'
import { isMod, shortcutLetter } from './keys'
import type { FreezeMode, MergeCommand, SheetFormatPatch } from '../contract/view'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** The sheet to show once the sheets are read again: by name, keeping the selection or starting over on it. */
export interface SheetTarget {
  name: string
  keep: boolean
}

export interface CommandsDeps {
  /** Read the sheets and windows again; how many times it has started, and how many `repaint` events came, to tell whether the source sent either during a call. */
  reload: () => Promise<void>
  reloads: () => number
  repaints: () => number
  /** The sheet the next read shows. */
  target: (t: SheetTarget | null) => void
  /** The style the active cell shows. */
  activeStyle: () => CellStyle | undefined
  /** Whether the book freezes panes on the sheet shown, the panes Excel freezes from A1, and the view after a freeze made. */
  frozen: () => boolean
  middlePanes: () => { rows: number; cols: number }
  showFreeze: (sheet: string, on: boolean) => void
  /** A structure operation went through: lines may have moved, so what the sheet keeps by line number no longer holds. */
  structured: () => void
}

const FONT_KEYS: Record<string, FontToggle> = { b: 'bold', i: 'italic', u: 'underline' }

/** The source's methods that change the book past cells. */
type ChangeMethod = 'format' | 'structure' | 'layout' | 'outline' | 'freeze'

export function useCommands(s: SheetState, sel: Selection, ed: Editing, deps: CommandsDeps) {
  /** A method of a change the source takes: the sheet is editable (not read-only, with `edit`) and the source has it; else null. */
  const method = <K extends ChangeMethod>(name: K) => (s.editable.value ? s.sourceMethod(name) : null)
  const can = (name: ChangeMethod): boolean => method(name) !== null
  /** Structure operations over ranges (a copy or a cut pasted, a fill, a merge): with `structure`, and no display filter scattering the rows. */
  const canRanges = computed(() => can('structure') && !s.visible.value)

  /**
   * A change the windows cannot carry: the source's outcome, then the sheet read again unless the source sent `reload`
   * meanwhile, or `repaint` for a change of `looks` alone (formatting), whose windows it reads again itself; formatting and
   * lines leave the copy mode on (`endsCopy`).
   */
  async function change(call: () => Promise<EditOutcome>, target: SheetTarget | null = null, endsCopy = true, looks = false): Promise<EditOutcome> {
    const before = { reloads: deps.reloads(), repaints: deps.repaints() }
    deps.target(target)
    const out = await ed.run(call, [], null, null, endsCopy)
    if (!out.ok) deps.target(null)
    else if (deps.reloads() === before.reloads && !(looks && deps.repaints() !== before.repaints)) await deps.reload()
    return out
  }

  async function format(patches: SheetFormatPatch[]): Promise<EditOutcome | null> {
    const call = method('format')
    const range = sel.selectionRef.value
    if (!call || !patches.length || !s.sheet.value) return null
    const sheet = s.sheetName.value
    const list: FormatPatch[] = []
    for (const p of patches) {
      if (p.range === undefined && !range) return null
      list.push({ ...p, range: p.range ?? range })
    }
    return change(() => call(sheet, list), null, false, true)
  }

  /** A structure operation; a paste of the copied range keeps the copy mode on (`endsCopy` false), as Excel's does. */
  async function structure(op: SheetOp, endsCopy = true): Promise<EditOutcome | null> {
    const call = method('structure')
    if (!call) return null
    // Excel shows a sheet it inserts or unhides, and keeps the selection of the active sheet it renames.
    let target: SheetTarget | null = null
    if (op.op === 'insertSheet') target = { name: op.name, keep: false }
    else if (op.op === 'unhideSheet') target = { name: op.sheet, keep: false }
    else if (op.op === 'renameSheet' && op.sheet === s.sheetName.value) target = { name: op.name, keep: true }
    const out = await change(() => call(op), target, endsCopy)
    if (out.ok) deps.structured()
    return out
  }

  /** A change of lines (sizes, hidden, outline): one call, the copy mode left on. */
  async function lineChange<A>(call: ((arg: A) => Promise<EditOutcome>) | null, arg: A): Promise<EditOutcome | null> {
    return call ? change(() => call(arg), null, false) : null
  }
  const layout = (edit: LayoutChange) => lineChange(method('layout'), edit)
  const outline = (command: OutlineChange) => lineChange(method('outline'), command)

  /**
   * Excel's Freeze Panes commands on the active sheet; a selection of whole columns or rows freezes left of or above it, as its
   * active cell is in the first row or column. None under the display filter, whose rows on screen are not the sheet's.
   */
  async function freeze(mode: FreezeMode = 'panes'): Promise<EditOutcome | null> {
    const call = method('freeze')
    const a = s.selected.value
    if (!call || !s.sheet.value || !a || s.visible.value) return null
    let panes: { rows: number; cols: number } | null = null
    if (mode === 'topRow') panes = { rows: 1, cols: 0 }
    else if (mode === 'firstColumn') panes = { rows: 0, cols: 1 }
    else if (mode === 'panes') panes = a.row === 1 && a.col === 1 ? deps.middlePanes() : { rows: a.row - 1, cols: a.col - 1 }
    if (panes ? panes.rows <= 0 && panes.cols <= 0 : !deps.frozen()) return null
    const sheet = s.sheetName.value
    const out = await change(() => call(sheet, panes), null, false)
    if (out.ok) deps.showFreeze(sheet, panes !== null)
    return out
  }

  /** Control+B, Control+I and Control+U (Cmd on a Mac, any layout): the font toggled over the selection by the active cell; true when taken. */
  function fontKey(e: KeyboardEvent): boolean {
    if (!isMod(e) || e.altKey || e.shiftKey || !can('format')) return false
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
    s.focusGrid()
  }

  return { can, canRanges, format, structure, layout, outline, freeze, fontKey, formatOpen, openFormat, closeFormat }
}

export type Commands = ReturnType<typeof useCommands>
