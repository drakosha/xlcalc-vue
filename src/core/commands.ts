// Changes of the book the host or the grid's menus and keys ask for (format, structure, lines, outline, undo and redo): each goes
// through the source as Control+Z goes, and the sheet is read again after one a window's cells cannot carry.
import { ref } from 'vue'
import type { CellStyle } from 'xlcalc/contract'
import type { EditOutcome, FormatPatch, LayoutChange, OutlineChange, SheetOp } from '../source'
import { fontToggle, type FontToggle } from '../toolbar'
import { shortcutLetter, type Editing } from './editing'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** A patch of `SheetView.format`: over its own range, or over the selection without one. */
export type SheetFormatPatch = Omit<FormatPatch, 'range'> & { range?: string }

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
}

const FONT_KEYS: Record<string, FontToggle> = { b: 'bold', i: 'italic', u: 'underline' }

export function useCommands(s: SheetState, sel: Selection, ed: Editing, deps: CommandsDeps) {
  const can = (method: 'format' | 'structure' | 'layout' | 'outline'): boolean => !s.props.readOnly && typeof s.source()[method] === 'function'

  /** A change the windows cannot carry: the source's outcome, then the sheet read again unless the source sent `reload` meanwhile. */
  async function change(call: () => Promise<EditOutcome>, target: SheetTarget | null = null): Promise<EditOutcome> {
    const before = deps.reloads()
    deps.target(target)
    const out = await ed.run(call, [], null, null)
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
    return change(() => call.call(s.source(), sheet, list))
  }

  async function structure(op: SheetOp): Promise<EditOutcome | null> {
    const call = s.source().structure
    if (!call || !can('structure')) return null
    // Excel shows a sheet it inserts or unhides, and keeps the selection of the active sheet it renames.
    let target: SheetTarget | null = null
    if (op.op === 'insertSheet') target = { name: op.name, keep: false }
    else if (op.op === 'unhideSheet') target = { name: op.sheet, keep: false }
    else if (op.op === 'renameSheet' && op.sheet === s.sheetName.value) target = { name: op.name, keep: true }
    return change(() => call.call(s.source(), op), target)
  }

  async function layout(edit: LayoutChange): Promise<EditOutcome | null> {
    const call = s.source().layout
    if (!call || !can('layout')) return null
    return change(() => call.call(s.source(), edit))
  }

  async function outline(command: OutlineChange): Promise<EditOutcome | null> {
    const call = s.source().outline
    if (!call || !can('outline')) return null
    return change(() => call.call(s.source(), command))
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
    structure,
    outline,
    formatCells: openFormat,
  }

  return { can, change, format, structure, layout, outline, fontKey, formatOpen, openFormat, closeFormat, api }
}

export type Commands = ReturnType<typeof useCommands>
