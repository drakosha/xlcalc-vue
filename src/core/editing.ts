// Editing cells: in place or in the formula bar, Excel's keys, clearing and pasting, Control+Return, undo and redo, calls in flight and the source's outcomes.
import { computed, nextTick, ref, shallowRef, type Ref } from 'vue'
import { cellAddress, MAX_COLS, MAX_ROWS, parseRange, quoteSheetIfNeeded, rangeAddress, rectsMeet } from 'xlcalc/a1'
import type { CellPosition, CellRect } from 'xlcalc/contract'
import { parseClipboard } from 'xlcalc/text'
import type { CellEntry, EditOutcome } from '../source'
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
/** Cells a fill or a paste through the source shows as pending; past that the rest wait unmarked. */
const MAX_PENDING_CELLS = 10_000
const ARROWS: Record<string, [-1 | 0 | 1, -1 | 0 | 1]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }

export interface EditingDeps {
  notify: (text: string) => void
  /** Read the sheets and windows again (a conflict). */
  reload: () => Promise<void>
  /** Go to a place as the address box goes there (`'Sheet 2'!B3:C4`): an outcome's selection on another sheet. */
  goPlace: (place: string) => Promise<boolean>
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
      const rect = parseRange(l.range)
      return rect ? [{ rect, reason: l.reason }] : []
    }),
  )
  const lockIn = (rect: CellRect): string | null => locks.value.find((l) => rectsMeet(l.rect, rect))?.reason ?? null
  const lockAt = (row: number, col: number): string | null => lockIn({ top: row, left: col, bottom: row, right: col })

  /** What the field opens with: the window's `editText`; without it the formula, the text, a number in the book's locale, shown text otherwise. */
  function editText(row: number, col: number): string {
    const own = w.editTextOfRow(row, col)
    if (own !== undefined) return own
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
  /** Control+Return: with `fill` the text goes into every cell of the selection, else into the active cell; nothing moves. */
  function commitFill(): void {
    const ed = editor.value
    const fill = s.source().fill
    const r = sel.selRect.value
    if (!ed || !fill || !r || !sel.multi.value || ed.sheet !== s.sheetName.value || !sel.isActive(ed.row, ed.col)) return commit(0, 0)
    // A locked cell anywhere in the range: the field stays open with the draft.
    const lock = lockIn(r)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const extent = s.extent.value
    editor.value = null
    s.scroller.value?.focus({ preventScroll: true })
    const marks = marksOf(r, () => (ed.text === '' ? null : ed.text))
    void run(() => fill.call(s.source(), ed.sheet, rangeAddress(r), ed.text, cellAddress(ed.row, ed.col)), marks, ed, extent)
  }
  /** Open the field again with the draft, after a conflict or a formula that did not parse; `extent` gives a range back its corner. */
  function reopen(ed: EditorState, caret: number | null, extent: CellPosition | null): void {
    if (editor.value || ed.sheet !== s.sheetName.value) return
    s.selected.value = { row: ed.row, col: ed.col }
    s.extent.value = extent
    editor.value = { ...ed, mode: 'edit', caret: caret ?? ed.text.length }
    sel.revealCell(ed.row, ed.col)
    focusEditor()
  }

  /** Send entries; the cells show the typed text in the "pending" style until the outcome. */
  async function send(entries: CellEntry[], draft: EditorState | null): Promise<void> {
    const edit = s.source().edit
    if (edit) await run(() => edit.call(s.source(), entries), entries, draft, null)
  }
  /** A rectangle's shown cells on the sheet (rows a filter hides are not drawn) with their pending text, row by row, up to the limit. */
  function marksOf(r: CellRect, textAt: (row: number, col: number) => string | null): CellEntry[] {
    const out: CellEntry[] = []
    for (let row = r.top; row <= Math.min(r.bottom, MAX_ROWS) && out.length < MAX_PENDING_CELLS; row++) {
      if (s.indexOf(row) === undefined) continue
      for (let col = r.left; col <= Math.min(r.right, MAX_COLS) && out.length < MAX_PENDING_CELLS; col++) out.push({ sheet: s.sheetName.value, cell: cellAddress(row, col), text: textAt(row, col) })
    }
    return out
  }
  /** Call the source: `marks` show in the "pending" style until the outcome; a failure opens `draft` again, over its range with `extent`. */
  async function run(call: () => Promise<EditOutcome>, marks: CellEntry[], draft: EditorState | null, extent: CellPosition | null): Promise<void> {
    const mine = ++seq
    if (marks.length) {
      const next = new Map(pending.value)
      for (const e of marks) next.set(pendingKey(e.sheet, e.cell), { text: e.text, seq: mine })
      pending.value = next
    }
    const before = { sheet: s.sheetName.value, selected: s.selected.value, extent: s.extent.value }
    let outcome: EditOutcome
    try {
      outcome = await call()
    } catch (e) {
      outcome = { ok: false, reason: 'failed', message: e instanceof Error ? e.message : '' }
    }
    if (marks.length) {
      const left = new Map(pending.value)
      for (const e of marks) if (left.get(pendingKey(e.sheet, e.cell))?.seq === mine) left.delete(pendingKey(e.sheet, e.cell))
      pending.value = left
    }
    if (outcome.ok) {
      w.applyChanges(outcome.changes)
      // The outcome's selection, unless the person moved on or opened a field meanwhile.
      const still = before.sheet === s.sheetName.value && before.selected === s.selected.value && before.extent === s.extent.value && !editor.value
      if (outcome.select && still) await selectPlace(outcome.select)
      return
    }
    const m = s.messages.value
    if (outcome.reason === 'conflict') {
      deps.notify(m.conflict)
      await deps.reload()
      if (draft) reopen(draft, null, extent)
    } else if (outcome.reason === 'invalid') {
      deps.notify(outcome.message || m.editFailed)
      if (draft) reopen(draft, outcome.at ?? null, extent)
    } else deps.notify(outcome.message || m.editFailed)
  }
  /** Put the selection where an outcome says: on this sheet at once, on another as the address box goes there. */
  async function selectPlace(place: { sheet: string; range: string }): Promise<void> {
    if (place.sheet !== s.sheetName.value) {
      await deps.goPlace(`${quoteSheetIfNeeded(place.sheet)}!${place.range}`)
      return
    }
    const r = parseRange(place.range)
    if (!r) return
    s.stretch(r.bottom, r.right)
    s.selected.value = { row: r.top, col: r.left }
    s.extent.value = r.top === r.bottom && r.left === r.right ? null : { row: r.bottom, col: r.right }
    void nextTick(() => sel.revealCell(r.top, r.left))
  }

  /** Undo and redo of the source, one call at a time: a second press while one is out sends nothing. */
  let stepping = false
  async function stepHistory(kind: 'undo' | 'redo'): Promise<void> {
    const call = s.source()[kind]
    if (!call || stepping) return
    stepping = true
    try {
      await run(() => call.call(s.source()), [], null, null)
    } finally {
      stepping = false
    }
  }
  /** Control+Z, Control+Y and Control+Shift+Z (Cmd on a Mac); another layout's letters by the key's place. */
  function historyKey(e: KeyboardEvent): 'undo' | 'redo' | null {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return null
    const letter = /^[a-z]$/i.test(e.key) ? e.key.toLowerCase() : e.code === 'KeyZ' ? 'z' : e.code === 'KeyY' ? 'y' : ''
    if (letter === 'z') return e.shiftKey ? 'redo' : 'undo'
    return letter === 'y' && !e.shiftKey ? 'redo' : null
  }

  /** The selection cut to the data: past it there is nothing to clear. */
  function dataRect(): CellRect | null {
    const r = sel.selRect.value
    if (!r) return null
    const out = { top: r.top, left: r.left, bottom: Math.min(r.bottom, s.dataRows.value), right: Math.min(r.right, s.dataCols.value) }
    return out.top <= out.bottom && out.left <= out.right ? out : null
  }
  /** Delete and Backspace: clear the selection's cells (those known empty and rows the filter hides are skipped, as in Excel). */
  function clearSelection(): void {
    if (!s.editable.value || !sel.selRect.value) return
    const lock = lockIn(sel.selRect.value)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const r = dataRect()
    if (!r) return
    if ((r.bottom - r.top + 1) * (r.right - r.left + 1) > MAX_EDIT_CELLS) return deps.notify(s.messages.value.tooManyCells)
    const entries: CellEntry[] = []
    for (let row = r.top; row <= r.bottom; row++) {
      if (s.indexOf(row) === undefined) continue
      for (let col = r.left; col <= r.right; col++) {
        if (w.emptyAtRow(row, col) && f.formulaAt(row, col) === undefined) continue
        entries.push({ sheet: s.sheetName.value, cell: cellAddress(row, col), text: null })
      }
    }
    if (entries.length) void send(entries, null)
  }
  /** Paste text as the rectangle Excel pastes from the selection's top-left cell: with `paste` the source lays it out in one call, else one value fills the selection. */
  function pasteText(text: string): void {
    const a = sel.selRect.value
    if (!s.editable.value || !a) return
    const rows = parseClipboard(text)
    const paste = s.source().paste
    const one = !paste && rows.length === 1 && rows[0]!.length === 1
    const width = rows[0]!.length
    const target: CellRect = one ? a : { top: a.top, left: a.left, bottom: a.top + rows.length - 1, right: a.left + width - 1 }
    const lock = lockIn(target)
    if (lock !== null) return deps.notify(s.messages.value.locked(lock))
    const textAt = (row: number, col: number): string | null => {
      const t = one ? rows[0]![0]! : (rows[row - target.top]?.[col - target.left] ?? '')
      return t === '' ? null : t
    }
    if (paste) {
      const sheet = s.sheetName.value
      void run(() => paste.call(s.source(), sheet, cellAddress(a.top, a.left), text), marksOf(target, textAt), null, null)
      return
    }
    // A rectangle past the sheet's edge is refused whole, as Excel refuses it; a source's `paste` says so itself.
    if (target.bottom > MAX_ROWS || target.right > MAX_COLS) return deps.notify(s.messages.value.editFailed)
    if ((target.bottom - target.top + 1) * (target.right - target.left + 1) > MAX_EDIT_CELLS) return deps.notify(s.messages.value.tooManyCells)
    const entries: CellEntry[] = []
    for (let row = target.top; row <= target.bottom; row++) {
      for (let col = target.left; col <= target.right; col++) entries.push({ sheet: s.sheetName.value, cell: cellAddress(row, col), text: textAt(row, col) })
    }
    s.stretch(target.bottom, target.right)
    s.selected.value = { row: target.top, col: target.left }
    s.extent.value = target.top === target.bottom && target.left === target.right ? null : { row: target.bottom, col: target.right }
    void send(entries, null)
  }

  /** Keys on the grid: Enter and Tab move, F2 edits, Delete clears, Control+Z and Control+Y undo and redo, a character starts a new text, the rest navigates. */
  function onGridKey(e: KeyboardEvent): void {
    if (e.target !== s.scroller.value || e.isComposing) return
    const ctrl = e.ctrlKey || e.metaKey
    const history = historyKey(e)
    if (history && s.editable.value && s.source()[history]) {
      e.preventDefault()
      void stepHistory(history)
      return
    }
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
  /** Keys in the field: Enter, Tab and their Shift commit and move, Control+Return (Cmd on a Mac) fills, Esc cancels, Alt+Enter breaks the line, F2 toggles arrows. */
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
    } else if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      e.preventDefault()
      commitFill()
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
