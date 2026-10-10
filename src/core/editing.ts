// Editing cells: in place or in the formula bar, Excel's keys, clearing and pasting, edits in flight and the source's outcomes.
import { computed, nextTick, ref, shallowRef, type Ref } from 'vue'
import type { CellEntry, EditOutcome } from '../source'
import { cellAddress, parseRect, rectsMeet, type Rect } from './a1'
import { parseTsv } from './clipboard'
import { numberEditText, textEditText } from './edit-text'
import type { Formula } from './formula'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** The cell being edited and its draft. `enter` was started by typing (arrows commit), `edit` by F2 or a double click (arrows move the caret). */
export interface EditorState {
  sheet: string
  row: number
  col: number
  text: string
  initial: string
  mode: 'enter' | 'edit'
  place: 'cell' | 'bar'
  /** Where to put the caret once the field shows; null after that. */
  caret: number | null
}

/** Cells one edit may carry. */
const MAX_EDIT_CELLS = 50_000
const ARROWS: Record<string, [-1 | 0 | 1, -1 | 0 | 1]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }

export interface EditingDeps {
  notify: (text: string) => void
  /** Read the sheets and windows again (a conflict). */
  reload: () => Promise<void>
  cellInput: Ref<HTMLTextAreaElement | null>
  barInput: Ref<HTMLTextAreaElement | null>
}

export function useEditing(s: SheetState, w: Windows, sel: Selection, f: Formula, editor: Ref<EditorState | null>, deps: EditingDeps) {
  /** Edits in flight by cell: the typed text the cell shows meanwhile (null: clearing). */
  const pending = shallowRef(new Map<string, { text: string | null; seq: number }>())
  let seq = 0
  const pendingKey = (sheet: string, cell: string) => `${sheet}!${cell}`
  const pendingAt = (row: number, col: number): string | null | undefined => pending.value.get(pendingKey(s.sheetName.value, cellAddress(row, col)))?.text

  /** The sheet's locked rectangles with their reasons. */
  const locks = computed(() =>
    (s.sheet.value?.locked ?? []).flatMap((l) => {
      const rect = parseRect(l.range)
      return rect ? [{ rect, reason: l.reason }] : []
    }),
  )
  const lockIn = (rect: Rect): string | null => locks.value.find((l) => rectsMeet(l.rect, rect))?.reason ?? null
  const lockAt = (row: number, col: number): string | null => lockIn({ r1: row, c1: col, r2: row, c2: col })

  /** What the field opens with: the formula, the text, a number in the book's locale; shown text otherwise. */
  function editText(row: number, col: number): string {
    const formula = f.formulaAt(row, col)
    if (formula !== undefined) return `=${formula}`
    const value = w.valueOfRow(row, col)
    const shown = w.textOfRow(row, col)
    if (typeof value === 'string') return textEditText(value)
    if (typeof value === 'number') return numberEditText(value, shown, s.locale.value)
    return shown
  }

  /** Open the field on the active cell; a locked cell is explained instead. False when nothing opened. */
  function startEdit(opts: { text?: string; mode: 'enter' | 'edit'; place?: 'cell' | 'bar'; caret?: number }): boolean {
    const a = s.selected.value
    if (!s.editable.value || !a) return false
    const lock = lockAt(a.row, a.col)
    if (lock !== null) {
      deps.notify(s.messages.value.locked(lock))
      return false
    }
    const initial = editText(a.row, a.col)
    const text = opts.text ?? initial
    editor.value = { sheet: s.sheetName.value, row: a.row, col: a.col, text, initial, mode: opts.mode, place: opts.place ?? 'cell', caret: opts.caret ?? text.length }
    focusEditor()
    return true
  }
  /** Focus the field at once (a phone opens its keyboard only inside the tap) and put the caret once it shows the draft. */
  function focusEditor(): void {
    const inputOf = () => (editor.value?.place === 'bar' ? deps.barInput.value : deps.cellInput.value)
    inputOf()?.focus({ preventScroll: true })
    void nextTick(() => {
      const ed = editor.value
      const el = inputOf()
      if (!ed || !el || ed.caret === null) return
      if (document.activeElement !== el) el.focus({ preventScroll: true })
      el.value = ed.text
      el.setSelectionRange(ed.caret, ed.caret)
      ed.caret = null
    })
  }
  function onEditorInput(e: Event): void {
    if (editor.value) editor.value.text = (e.target as HTMLTextAreaElement).value
  }

  /** Finish the edit: the cursor moves at once, the source gets the text unless it did not change. */
  function commit(dr: -1 | 0 | 1, dc: -1 | 0 | 1, refocus = true): void {
    const ed = editor.value
    if (!ed) return
    editor.value = null
    if (refocus) s.scroller.value?.focus({ preventScroll: true })
    if (ed.sheet === s.sheetName.value && sel.isActive(ed.row, ed.col) && (dr || dc)) sel.step(dr, dc)
    if (ed.text === ed.initial) return
    void send([{ sheet: ed.sheet, cell: cellAddress(ed.row, ed.col), text: ed.text === '' ? null : ed.text }], ed)
  }
  function cancel(): void {
    editor.value = null
    s.scroller.value?.focus({ preventScroll: true })
  }
  /** Open the field again with the draft, after a conflict or a formula that did not parse. */
  function reopen(ed: EditorState, caret: number | null): void {
    if (editor.value || ed.sheet !== s.sheetName.value) return
    s.selected.value = { row: ed.row, col: ed.col }
    s.extent.value = null
    editor.value = { ...ed, mode: 'edit', caret: caret ?? ed.text.length }
    sel.revealCell(ed.row, ed.col)
    focusEditor()
  }

  /** Send entries; the cells show the typed text in the "pending" style until the outcome. */
  async function send(entries: CellEntry[], draft: EditorState | null): Promise<void> {
    const edit = s.source().edit
    if (!edit) return
    const mine = ++seq
    const next = new Map(pending.value)
    for (const e of entries) next.set(pendingKey(e.sheet, e.cell), { text: e.text, seq: mine })
    pending.value = next
    let outcome: EditOutcome
    try {
      outcome = await edit.call(s.source(), entries)
    } catch (e) {
      outcome = { ok: false, reason: 'failed', message: e instanceof Error ? e.message : '' }
    }
    const left = new Map(pending.value)
    for (const e of entries) if (left.get(pendingKey(e.sheet, e.cell))?.seq === mine) left.delete(pendingKey(e.sheet, e.cell))
    pending.value = left
    if (outcome.ok) {
      w.applyChanges(outcome.changes)
      return
    }
    const m = s.messages.value
    if (outcome.reason === 'conflict') {
      deps.notify(m.conflict)
      await deps.reload()
      if (draft) reopen(draft, null)
    } else if (outcome.reason === 'invalid') {
      deps.notify(outcome.message || m.editFailed)
      if (draft) reopen(draft, outcome.at ?? null)
    } else deps.notify(outcome.message || m.editFailed)
  }

  /** The selection cut to the data: past it there is nothing to clear. */
  function dataRect(): Rect | null {
    const r = sel.selRect.value
    if (!r) return null
    const out = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, s.dataRows.value), c2: Math.min(r.c2, s.dataCols.value) }
    return out.r1 <= out.r2 && out.c1 <= out.c2 ? out : null
  }
  /** Delete and Backspace: clear the selection's cells (those known empty and rows the filter hides are skipped, as in Excel). */
  function clearSelection(): void {
    if (!s.editable.value || !sel.selRect.value) return
    const lock = lockIn(sel.selRect.value)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const r = dataRect()
    if (!r) return
    if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > MAX_EDIT_CELLS) return deps.notify(s.messages.value.tooManyCells)
    const entries: CellEntry[] = []
    for (let row = r.r1; row <= r.r2; row++) {
      if (s.indexOf(row) === undefined) continue
      for (let col = r.c1; col <= r.c2; col++) {
        if (w.emptyAtRow(row, col) && f.formulaAt(row, col) === undefined) continue
        entries.push({ sheet: s.sheetName.value, cell: cellAddress(row, col), text: null })
      }
    }
    if (entries.length) void send(entries, null)
  }
  /** Paste TSV as a rectangle from the active cell; one value fills the whole selection. */
  function pasteText(text: string): void {
    const a = sel.selRect.value
    if (!s.editable.value || !a) return
    const rows = parseTsv(text)
    if (!rows.length) return
    const one = rows.length === 1 && rows[0]!.length === 1
    const width = Math.max(...rows.map((line) => line.length))
    const target: Rect = one ? a : { r1: a.r1, c1: a.c1, r2: a.r1 + rows.length - 1, c2: a.c1 + width - 1 }
    const lock = lockIn(target)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    if ((target.r2 - target.r1 + 1) * (target.c2 - target.c1 + 1) > MAX_EDIT_CELLS) return deps.notify(s.messages.value.tooManyCells)
    const entries: CellEntry[] = []
    for (let row = target.r1; row <= target.r2; row++) {
      for (let col = target.c1; col <= target.c2; col++) {
        const t = one ? rows[0]![0]! : (rows[row - target.r1]?.[col - target.c1] ?? '')
        entries.push({ sheet: s.sheetName.value, cell: cellAddress(row, col), text: t === '' ? null : t })
      }
    }
    s.stretch(target.r2, target.c2)
    s.selected.value = { row: target.r1, col: target.c1 }
    s.extent.value = target.r1 === target.r2 && target.c1 === target.c2 ? null : { row: target.r2, col: target.c2 }
    void send(entries, null)
  }

  /** Keys on the grid: Enter and Tab move, F2 edits, Delete clears, a character starts a new text, the rest navigates. */
  function onGridKey(e: KeyboardEvent): void {
    if (e.target !== s.scroller.value || e.isComposing) return
    const ctrl = e.ctrlKey || e.metaKey
    if (ctrl && /^[aAфФ]$/.test(e.key)) {
      e.preventDefault()
      sel.selectAll()
      return
    }
    if (e.key === 'Enter' && !ctrl && !e.altKey) {
      e.preventDefault()
      sel.step(e.shiftKey ? -1 : 1, 0)
      return
    }
    if (e.key === 'Tab' && !ctrl) {
      e.preventDefault()
      sel.step(0, e.shiftKey ? -1 : 1)
      return
    }
    if (e.key === 'F2') {
      e.preventDefault()
      startEdit({ mode: 'edit' })
      return
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && !ctrl) {
      e.preventDefault()
      clearSelection()
      return
    }
    if (sel.navKey(e)) return
    // AltGr comes as Ctrl+Alt and types a character (`@` on many layouts): only Ctrl alone or Cmd is a shortcut.
    if (e.key.length === 1 && !((e.ctrlKey && !e.altKey) || e.metaKey) && s.editable.value) {
      e.preventDefault()
      startEdit({ text: e.key, mode: 'enter' })
    }
  }
  /** Keys in the field: Enter, Tab and their Shift commit and move, Esc cancels, Alt+Enter breaks the line, F2 toggles arrows. */
  function onEditorKey(e: KeyboardEvent): void {
    const ed = editor.value
    if (!ed || e.isComposing) return
    e.stopPropagation()
    if (e.key === 'Escape') {
      e.preventDefault()
      cancel()
    } else if (e.key === 'Enter' && e.altKey) {
      e.preventDefault()
      const el = e.target as HTMLTextAreaElement
      const at = el.selectionStart ?? ed.text.length
      ed.text = `${ed.text.slice(0, at)}\n${ed.text.slice(el.selectionEnd ?? at)}`
      ed.caret = at + 1
      focusEditor()
    } else if (e.key === 'Enter') {
      e.preventDefault()
      commit(e.shiftKey ? -1 : 1, 0)
    } else if (e.key === 'Tab') {
      e.preventDefault()
      commit(0, e.shiftKey ? -1 : 1)
    } else if (e.key === 'F2') {
      e.preventDefault()
      ed.mode = ed.mode === 'enter' ? 'edit' : 'enter'
    } else if (ed.mode === 'enter' && ed.place === 'cell' && !ed.text.startsWith('=') && ARROWS[e.key]) {
      e.preventDefault()
      const [dr, dc] = ARROWS[e.key]!
      commit(dr, dc)
    }
  }
  /** The field lost focus: moving to the other field keeps editing, anything else in the page commits in place. */
  function onEditorBlur(e: FocusEvent, place: 'cell' | 'bar'): void {
    const ed = editor.value
    if (!ed || ed.place !== place) return
    const to = e.relatedTarget
    if (to && (to === deps.cellInput.value || to === deps.barInput.value)) {
      ed.place = to === deps.barInput.value ? 'bar' : 'cell'
      return
    }
    // Another window took the focus: the edit waits, as in Excel.
    if (typeof document.hasFocus === 'function' && !document.hasFocus()) return
    commit(0, 0, false)
  }
  /** Paste on the grid (the field pastes into itself). */
  function onPaste(e: ClipboardEvent): void {
    if (e.target !== s.scroller.value || !s.editable.value) return
    const text = e.clipboardData?.getData('text/plain')
    if (!text) return
    e.preventDefault()
    pasteText(text)
  }

  return { pending, pendingAt, locks, lockAt, editText, startEdit, focusEditor, onEditorInput, commit, cancel, send, clearSelection, pasteText, onGridKey, onEditorKey, onEditorBlur, onPaste }
}

export type Editing = ReturnType<typeof useEditing>
