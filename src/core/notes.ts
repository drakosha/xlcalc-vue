// Excel's notes on cells: New Note, Edit Note and Delete Note in the cell menu, Shift+F2, and the note box being edited at its cell,
// which saves through the source's `note` when it is left (Esc or a press outside it), as Excel saves a note; and a cell's note
// shown in the same box, read-only, while the mouse rests on the cell, as Excel shows it; notes shown all the time (Show/Hide
// Note, saved with the note) and Excel's Show All Notes, a view setting.
import { computed, ref } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { SheetNote } from '../contract/workbook'
import type { Extras } from './extras'
import type { MenuItem } from './menu'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** The note being edited: its cell, its author (shown in bold above the text), its text and the text it opened with. */
export interface NoteDraft {
  sheet: string
  row: number
  col: number
  author?: string
  text: string
  initial: string
  /** A new note (New Note): saved even unchanged, as Excel keeps a note once it is inserted, unless it has neither author nor text. */
  fresh: boolean
  /** A note shown all the time stays so when edited. */
  shown?: true
}

export interface NotesDeps {
  /** The author's name for a new note: the sheet's `noteAuthor` prop. */
  author: () => string | undefined
}

export function useNotes(s: SheetState, w: Windows, extras: Extras, deps: NotesDeps) {
  const m = () => s.messages.value
  const draft = ref<NoteDraft | null>(null)
  const available = computed(() => extras.method('note') !== null)
  /** Excel's Show All Notes: every note drawn all the time, on every sheet, for the component's lifetime. */
  const all = ref(false)
  /** The notes drawn all the time on the displayed sheet, from its loaded windows; the one being edited is in its own box. */
  const shown = computed(() => {
    const d = draft.value
    return w.shownNotes(all.value).filter((n) => !d || d.row !== n.row || d.col !== n.col)
  })
  /** The cell with a note that the mouse is over, while no note is being edited. */
  const hovered = ref<{ sheet: string; row: number; col: number } | null>(null)
  /** The hovered cell's note, from its loaded window, for the read-only box. */
  const peek = computed(() => {
    const h = hovered.value
    const note = h && !draft.value && h.sheet === s.sheetName.value ? w.noteOfRow(h.row, h.col) : undefined
    return h && note ? { ...h, ...note } : null
  })
  /** The mouse over a cell (null: over none, or not a mouse): its note shows at once, and hides when the mouse leaves the cell. */
  function hover(row: number | null, col = 0): void {
    const sheet = s.sheetName.value
    const note = row !== null ? w.noteOfRow(row, col) : undefined
    // a note in view all the time needs no peek
    const at = row !== null && note && !note.shown && !all.value ? { sheet, row, col } : null
    const h = hovered.value
    if (at?.row !== h?.row || at?.col !== h?.col || at?.sheet !== h?.sheet) hovered.value = at
  }

  /** The active cell's note from its loaded window. */
  function activeNote(): SheetNote | undefined {
    const a = s.selected.value
    return a ? w.noteOfRow(a.row, a.col) : undefined
  }

  /** Excel's New Note, or Edit Note on a cell with one (Shift+F2 does either): the box opens at the active cell; false when notes cannot be edited. */
  function open(): boolean {
    const a = s.selected.value
    if (!available.value || !s.sheet.value || !a || extras.locked({ top: a.row, left: a.col, bottom: a.row, right: a.col })) return false
    const note = activeNote()
    const author = note ? note.author : deps.author()?.trim() || undefined
    draft.value = { sheet: s.sheetName.value, row: a.row, col: a.col, ...(author ? { author } : {}), text: note?.text ?? '', initial: note?.text ?? '', fresh: !note, ...(note?.shown ? { shown: true as const } : {}) }
    return true
  }

  /** Leaves the box: a changed note, or a new one, goes to the source; the grid takes focus back unless the press went elsewhere. */
  async function save(refocus = true): Promise<void> {
    const d = draft.value
    const call = extras.method('note')
    if (!d) return
    draft.value = null
    if (refocus) s.focusGrid()
    if (!call || (d.text === d.initial && !d.fresh) || (d.fresh && !d.text && !d.author)) return
    const cell = cellAddress(d.row, d.col)
    const note: SheetNote = { ...(d.author ? { author: d.author } : {}), text: d.text, ...(d.shown ? { shown: true } : {}) }
    await extras.change(() => call(d.sheet, cell, note), d.sheet, [cell])
  }

  /** Delete Note on the active cell. */
  async function remove(): Promise<void> {
    const a = s.selected.value
    const call = extras.method('note')
    if (!a || !call || extras.locked({ top: a.row, left: a.col, bottom: a.row, right: a.col })) return
    const sheet = s.sheetName.value
    const cell = cellAddress(a.row, a.col)
    await extras.change(() => call(sheet, cell, null), sheet, [cell])
  }

  /** Excel's Show/Hide Note on the active cell: the note saved shown all the time, or no longer, through `note`. */
  async function toggleShown(): Promise<void> {
    const a = s.selected.value
    const call = extras.method('note')
    const note = activeNote()
    if (!a || !call || !note || extras.locked({ top: a.row, left: a.col, bottom: a.row, right: a.col })) return
    const sheet = s.sheetName.value
    const cell = cellAddress(a.row, a.col)
    const { shown: was, ...rest } = note
    await extras.change(() => call(sheet, cell, was ? rest : { ...rest, shown: true }), sheet, [cell])
  }

  /** Excel's Show All Notes, toggled without `on`: whether it is on afterwards. */
  function showAll(on = !all.value): boolean {
    all.value = on
    if (on) hovered.value = null
    return on
  }

  /** The cell menu's note items, as Excel's: New Note on a cell without one, Edit Note, Delete Note and Show/Hide Note on a cell with one. */
  function items(): MenuItem[] {
    if (!available.value) return []
    const out: MenuItem[] = activeNote()
      ? [
          { id: 'editNote', label: m().menuEditNote, run: () => void open() },
          { id: 'deleteNote', label: m().menuDeleteNote, run: () => void remove() },
          { id: 'showNote', label: m().menuShowNote, run: () => void toggleShown() },
        ]
      : [{ id: 'newNote', label: m().menuNewNote, run: () => void open() }]
    out[0]!.group = true
    return out
  }

  return { draft, available, all, shown, peek, hover, open, save, remove, toggleShown, showAll, items }
}

export type Notes = ReturnType<typeof useNotes>
