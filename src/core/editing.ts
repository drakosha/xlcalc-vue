// Editing cells: in place or in the formula bar, Excel's keys, clearing and pasting text, Control+Return, undo and redo, calls in flight and the source's outcomes.
import { computed, nextTick, ref, shallowRef, type Ref } from 'vue'
import { cellAddress, MAX_COLS, MAX_ROWS, parseRange, rangeAddress, rectsMeet } from 'xlcalc/a1'
import type { CellPosition, CellRect } from 'xlcalc/contract'
import { parseClipboard } from 'xlcalc/text'
import { loadEntry, type Entry } from '../entry'
import type { CellEntry, EditOutcome, EditOptions, SheetSource } from '../source'
import { useAutocomplete } from './autocomplete'
import { useFunctionTip } from './function-tip'
import type { AlertAnswer, SheetAlert } from './validation'
import type { Formula } from './formula'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Windows } from './windows'

/** The cell being edited and its draft. `enter` was started by typing (arrows point or commit), `edit` by F2 or a double click (arrows move the caret). */
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
  /** The end of a selection to put with the caret (the formula bar's text dragged over), else the caret alone. */
  caretEnd?: number
  /** The field opened with the cell's own text (F2, a double click, the formula bar), not with a key typed over the cell. */
  inPlace: boolean
}

/** Cells one edit may carry. */
const MAX_EDIT_CELLS = 50_000
/** Cells a fill or a paste through the source shows as pending; past that the rest wait unmarked. */
const MAX_PENDING_CELLS = 10_000

/** A shortcut's Latin letter: the key's own, else its place on the keyboard (`KeyB` under Russian `и`); '' for anything else. */
export function shortcutLetter(e: KeyboardEvent): string {
  if (/^[a-z]$/i.test(e.key)) return e.key.toLowerCase()
  return /^Key[A-Z]$/.test(e.code) ? e.code.slice(3).toLowerCase() : ''
}

/** The scroll that shows a field's caret line (setSelectionRange does not scroll a field of fixed height, as the collapsed formula bar's): the text up to the caret measured in the field itself; null when it all shows. */
function caretTop(el: HTMLTextAreaElement, at: number): number | null {
  if (el.scrollHeight <= el.clientHeight) return null
  const cs = getComputedStyle(el)
  const line = parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.2
  const padTop = parseFloat(cs.paddingTop) || 0
  const padBottom = parseFloat(cs.paddingBottom) || 0
  const text = el.value
  el.value = text.slice(0, at)
  // The text up to the caret within the first screen: the top shows it.
  const first = el.scrollHeight <= el.clientHeight
  const bottom = el.scrollHeight - padBottom
  el.value = text
  if (first) return 0
  const view = el.clientHeight - padTop - padBottom
  const now = el.scrollTop
  if (bottom - line < now + padTop) return bottom - line - padTop
  return bottom > now + padTop + view ? bottom - padTop - view : now
}

export interface EditingDeps {
  notify: (text: string) => void
  /** Read the sheets and windows again (a conflict). */
  reload: () => Promise<void>
  /** Make another sheet active: an outcome's selection on it. */
  pickSheet: (index: number) => Promise<void>
  cellInput: Ref<HTMLTextAreaElement | null>
  barInput: Ref<HTMLTextAreaElement | null>
  /** A key that is the sheet's own in the field too (Control+Shift+U for the formula bar): true when taken. */
  barKey?: (e: KeyboardEvent) => boolean
  /** Excel's Enter mode by a key in the field: an arrow points, enters the draft or does nothing (`refAt`); true when taken. */
  point?: (e: KeyboardEvent) => boolean
  /** Excel's alert of a rule of data validation for a refused value, or the question on a formula that does not read, and the person's answer; without it the refusal is said in the message line. */
  alert?: (alert: SheetAlert) => Promise<AlertAnswer>
}

export function useEditing(s: SheetState, w: Windows, sel: Selection, f: Formula, editor: Ref<EditorState | null>, deps: EditingDeps) {
  /** Edits in flight by cell: the typed text the cell shows meanwhile (null: clearing). */
  const pending = shallowRef(new Map<string, { text: string | null; seq: number }>())
  let seq = 0
  const pendingKey = (sheet: string, cell: string) => `${sheet}!${cell}`
  /** A cell's pending text; without edits in flight no key is made (the grid asks for every drawn cell). */
  function pendingAt(row: number, col: number): string | null | undefined {
    const map = pending.value
    return map.size ? map.get(pendingKey(s.sheetName.value, cellAddress(row, col)))?.text : undefined
  }

  /** The sheet's locked rectangles with their reasons. */
  const locks = computed(() =>
    (s.sheet.value?.locked ?? []).flatMap((l) => {
      const rect = parseRange(l.range)
      return rect ? [{ rect, reason: l.reason }] : []
    }),
  )
  const lockIn = (rect: CellRect): string | null => locks.value.find((l) => rectsMeet(l.rect, rect))?.reason ?? null
  const lockAt = (row: number, col: number): string | null => (locks.value.length ? lockIn({ top: row, left: col, bottom: row, right: col }) : null)

  /** xlcalc's typed entry for a source without `editText`, asked for by the first edit text that needs it; null until it comes. */
  const entry = shallowRef<Entry | null>(null)
  let entryAsked = false
  /**
   * What the field opens with: the window's `editText`; without it the formula, else the value's `entryText` in the book's
   * locale by the cell's number format, as Excel's formula bar shows it. Until xlcalc/entry and the locale are in (a moment
   * after the first cell is selected) the shown text, so the field opens at once, as a phone's keyboard needs.
   */
  function editText(row: number, col: number): string {
    const own = w.editTextOfRow(row, col)
    if (own !== undefined) return own
    const formula = f.formulaAt(row, col)
    if (formula !== undefined) return `=${formula}`
    const value = w.valueOfRow(row, col)
    const shown = w.textOfRow(row, col)
    if (value === null || value === undefined) return shown
    const lib = entry.value
    if (!lib && !entryAsked) {
      entryAsked = true
      // A failed import leaves the shown text and is asked again by the next cell.
      loadEntry().then((m) => (entry.value = m), () => (entryAsked = false))
    }
    const opts = f.scanOptions()
    if (!lib || !opts) return shown
    try {
      return lib.entryText(value, { locale: opts.locale, numberFormat: w.lookOfRow(row, col)?.numberFormat })
    } catch {
      // a format code from the source that the library refuses
      return shown
    }
  }

  /** Open the field on the active cell; a locked cell is explained instead. False when nothing opened. */
  function startEdit(opts: { text?: string; mode: 'enter' | 'edit'; place?: 'cell' | 'bar'; caret?: number; caretEnd?: number }): boolean {
    const a = s.selected.value
    if (!s.editable.value || !a) return false
    const lock = lockAt(a.row, a.col)
    if (lock !== null) {
      deps.notify(s.messages.value.locked(lock))
      return false
    }
    const initial = editText(a.row, a.col)
    const text = opts.text ?? initial
    editor.value = { sheet: s.sheetName.value, row: a.row, col: a.col, text, initial, mode: opts.mode, place: opts.place ?? 'cell', caret: opts.caret ?? text.length, caretEnd: opts.caretEnd, inPlace: opts.text === undefined }
    focusEditor()
    return true
  }
  const inputOf = () => (editor.value?.place === 'bar' ? deps.barInput.value : deps.cellInput.value)
  /** Focus the field at once (a phone opens its keyboard only inside the tap) and put the caret once it shows the draft. */
  function focusEditor(): void {
    inputOf()?.focus({ preventScroll: true })
    void nextTick(() => {
      const ed = editor.value
      const el = inputOf()
      if (!ed || !el || ed.caret === null) return
      if (document.activeElement !== el) el.focus({ preventScroll: true })
      el.value = ed.text
      const top = caretTop(el, ed.caret)
      const end = Math.max(ed.caret, ed.caretEnd ?? ed.caret)
      el.setSelectionRange(ed.caret, end)
      if (top !== null) el.scrollTop = top
      tip.update(ed.text, end === ed.caret ? ed.caret : null)
      ed.caret = null
      ed.caretEnd = undefined
    })
  }
  /** Excel's Insert Function: `name(` at the open draft's caret (`=` first in an empty draft), else the active cell's edit opened in the formula bar with `=name(`; the caret after the bracket. */
  function insertFunction(name: string): void {
    const ed = editor.value
    if (!ed) return void startEdit({ text: `=${name}(`, mode: 'edit', place: 'bar' })
    const el = inputOf()
    const from = el?.selectionStart ?? ed.text.length
    const piece = `${ed.text === '' ? '=' : ''}${name}(`
    ed.text = `${ed.text.slice(0, from)}${piece}${ed.text.slice(el?.selectionEnd ?? from)}`
    ed.caret = from + piece.length
    focusEditor()
  }
  /** Excel's Formula AutoComplete over the draft: the list of functions under the field. */
  const ac = useAutocomplete({ source: s.source, editor, input: inputOf, show: focusEditor, scan: f.scanOptions })
  /** Excel's tip with the arguments of the function the caret stands in, under the field; the list goes first. */
  const tip = useFunctionTip({ source: s.source, editor, input: inputOf, scan: f.scanOptions, listOpen: () => ac.open.value })
  function onEditorInput(e: Event): void {
    const el = e.target as HTMLTextAreaElement
    if (!editor.value) return
    editor.value.text = el.value
    const caret = el.selectionStart === el.selectionEnd ? el.selectionStart : null
    ac.update(el.value, caret)
    tip.update(el.value, caret)
  }
  /** A click in the field moved the caret: the list and the tip follow it. */
  function onEditorClick(): void {
    ac.follow()
    tip.follow()
  }

  /** Finish the edit: the cursor moves at once, the source gets the text unless it did not change. */
  function commit(dr: -1 | 0 | 1, dc: -1 | 0 | 1, refocus = true): void {
    const ed = editor.value
    if (!ed) return
    editor.value = null
    if (refocus) s.scroller.value?.focus({ preventScroll: true })
    if (ed.sheet === s.sheetName.value && sel.isActive(ed.row, ed.col) && (dr || dc)) sel.step(dr, dc)
    if (ed.text === ed.initial) return
    void enterDraft(ed)
  }
  /** Send the draft into its cell; Excel's correction of it is sent the same way. */
  function enterDraft(ed: EditorState): Promise<EditOutcome | undefined> {
    const entries = [{ sheet: ed.sheet, cell: cellAddress(ed.row, ed.col), text: ed.text === '' ? null : ed.text }]
    return send(entries, ed, ed.inPlace ? { inPlace: true } : undefined, (text) => enterDraft({ ...ed, text }))
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
    void fillDraft(ed, r, extent, fill)
  }
  /** Send the draft into every cell of the range from its cell; Excel's correction of it is sent the same way. */
  function fillDraft(ed: EditorState, r: CellRect, extent: CellPosition | null, fill: NonNullable<SheetSource['fill']>): Promise<EditOutcome> {
    const marks = marksOf(r, () => (ed.text === '' ? null : ed.text))
    const args = [ed.sheet, rangeAddress(r), ed.text, cellAddress(ed.row, ed.col)] as const
    const call = (o?: EditOptions) => withOptions(ed.inPlace ? { inPlace: true } : undefined, o, (all) => (all ? fill.call(s.source(), ...args, all) : fill.call(s.source(), ...args)))
    return run(call, marks, ed, extent, true, (text) => fillDraft({ ...ed, text }, r, extent, fill))
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

  /** Send entries; the cells show the typed text in the "pending" style until the outcome. `opts`: `accept` writes past the rules (a list's item), `inPlace` the cell's own text edited; `retype` sends the draft's correction. */
  async function send(entries: CellEntry[], draft: EditorState | null, opts?: EditOptions, retype?: (text: string) => Promise<EditOutcome | undefined>): Promise<EditOutcome | undefined> {
    const edit = s.source().edit
    if (edit) return run((o) => withOptions(opts, o, (all) => (all ? edit.call(s.source(), entries, all) : edit.call(s.source(), entries))), entries, draft, null, true, retype)
  }
  /** A call with the options of the edit and of the alert's answer together; without any, a call as before options were. */
  function withOptions(base: EditOptions | undefined, more: EditOptions | undefined, call: (all: EditOptions | undefined) => Promise<EditOutcome>): Promise<EditOutcome> {
    const all = { ...base, ...more }
    return call(Object.keys(all).length ? all : undefined)
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
  /**
   * Call the source: `marks` show in the "pending" style until the outcome; a failure opens `draft` again, over its range
   * with `extent`; a change of cells ends the copy mode unless `endsCopy` is false. A value a rule of data validation does
   * not take shows Excel's alert, the marks kept: Yes or OK calls again with `accept`, Retry or No opens the draft again,
   * Cancel leaves the cell as it was. A formula the source offers to correct asks to accept it: Yes sends the correction by
   * `retype`, the selection put back where the edit left it.
   */
  async function run(call: (opts?: EditOptions) => Promise<EditOutcome>, marks: CellEntry[], draft: EditorState | null, extent: CellPosition | null, endsCopy = true, retype?: (text: string) => Promise<EditOutcome | undefined>): Promise<EditOutcome> {
    const mine = ++seq
    if (endsCopy) s.hooks.changing()
    if (marks.length) {
      const next = new Map(pending.value)
      for (const e of marks) next.set(pendingKey(e.sheet, e.cell), { text: e.text, seq: mine })
      pending.value = next
    }
    const before = { sheet: s.sheetName.value, selected: s.selected.value, extent: s.extent.value }
    const failure = (e: unknown): EditOutcome => ({ ok: false, reason: 'failed', message: e instanceof Error ? e.message : '' })
    let outcome: EditOutcome
    try {
      outcome = await call()
    } catch (e) {
      outcome = failure(e)
    }
    let answer: AlertAnswer | null = null
    if (!outcome.ok && outcome.alert && deps.alert) {
      // Excel keeps the edited cell active while its alert is up.
      if (draft && draft.sheet === s.sheetName.value) {
        s.selected.value = { row: draft.row, col: draft.col }
        s.extent.value = extent
      }
      answer = await deps.alert(outcome.alert)
      if (answer === 'accept') outcome = await call({ accept: true }).catch(failure)
      if (answer !== 'retry') s.scroller.value?.focus({ preventScroll: true })
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
      return outcome
    }
    if (answer === 'retry') {
      if (draft) reopen(draft, null, extent)
      return outcome
    }
    // Cancel: the cell keeps what it had and stays the active one, as Excel leaves it.
    if (answer === 'cancel') return outcome
    const m = s.messages.value
    if (outcome.reason === 'conflict') {
      deps.notify(m.conflict)
      await deps.reload()
      if (draft) reopen(draft, null, extent)
    } else if (outcome.reason === 'invalid' && draft && deps.alert && draft.sheet === s.sheetName.value) {
      // The edited cell is active while the question is up; where the edit moved the selection is kept for Excel's correction.
      const moved = { selected: s.selected.value, extent: s.extent.value }
      s.selected.value = { row: draft.row, col: draft.col }
      s.extent.value = extent
      const place = { sheet: draft.sheet, cell: cellAddress(draft.row, draft.col) }
      if (outcome.correction !== undefined && retype) {
        // Excel's correction of a formula left open: Yes enters it as the same key would, No (Esc) says what is missing and opens the field again.
        const said = await deps.alert({ severity: 'typo', message: m.typoFound(outcome.correction), ...place })
        if (said === 'accept') {
          s.selected.value = moved.selected
          s.extent.value = moved.extent
          s.scroller.value?.focus({ preventScroll: true })
          return (await retype(outcome.correction)) ?? outcome
        }
        await deps.alert({ severity: 'parenthesis', message: m.missingParenthesis, ...place })
        reopen(draft, outcome.at ?? null, extent)
        return outcome
      }
      // A formula that does not read: Fix opens it again, Restore leaves the cell as it was.
      const said = await deps.alert({ severity: 'formula', message: outcome.message || m.editFailed, ...place })
      if (said === 'retry') reopen(draft, outcome.at ?? null, extent)
      else s.scroller.value?.focus({ preventScroll: true })
    } else if (outcome.reason === 'invalid') {
      deps.notify(outcome.message || m.editFailed)
      if (draft) reopen(draft, outcome.at ?? null, extent)
    } else deps.notify(outcome.message || m.editFailed)
    return outcome
  }
  /** Put the selection where an outcome says, as keys put it: no flash and no step of the way back, on another sheet too. */
  async function selectPlace(place: { sheet: string; range: string }): Promise<void> {
    const r = parseRange(place.range)
    if (!r) return
    if (place.sheet !== s.sheetName.value) {
      const index = f.sheetIndexByName(place.sheet)
      if (index < 0) return
      await deps.pickSheet(index)
    }
    sel.selectRect(r)
    void nextTick(() => sel.revealCell(r.top, r.left))
  }

  /** Undo and redo of the source, one call at a time: a second press while one is out sends nothing (null). */
  let stepping = false
  async function stepHistory(kind: 'undo' | 'redo'): Promise<EditOutcome | null> {
    const call = s.source()[kind]
    if (!call || stepping || !s.editable.value) return null
    stepping = true
    try {
      return await run(() => call.call(s.source()), [], null, null)
    } finally {
      stepping = false
    }
  }
  /** Control+Z, Control+Y and Control+Shift+Z (Cmd on a Mac); another layout's letters by the key's place. */
  function historyKey(e: KeyboardEvent): 'undo' | 'redo' | null {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return null
    const letter = shortcutLetter(e)
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
    if (target.bottom > MAX_ROWS || target.right > MAX_COLS) return deps.notify(s.messages.value.pasteOffSheet)
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
  /** Keys in the field: the list of functions first when it is open, then an arrow of Enter mode (it points, enters or does nothing), Enter, Tab and their Shift commit and move, Control+Return (Cmd on a Mac) fills, Esc cancels, Alt+Enter breaks the line, F2 toggles arrows, Control+Shift+U the formula bar's height. */
  function onEditorKey(e: KeyboardEvent): void {
    const ed = editor.value
    if (!ed || e.isComposing) return
    e.stopPropagation()
    tip.follow()
    if (deps.barKey?.(e) || ac.onKey(e) || deps.point?.(e)) return
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
    // A dialog that works on the draft (Insert Function) keeps the edit open.
    if (to instanceof Element && to.closest('[data-xs-keep-edit]')) return
    // Another window took the focus: the edit waits, as in Excel.
    if (typeof document.hasFocus === 'function' && !document.hasFocus()) return
    commit(0, 0, false)
  }
  return { pending, pendingAt, locks, lockIn, lockAt, editText, startEdit, inputOf, focusEditor, insertFunction, ac, tip, onEditorInput, onEditorClick, commit, cancel, send, run, stepHistory, clearSelection, pasteText, onGridKey, onEditorKey, onEditorBlur }
}

export type Editing = ReturnType<typeof useEditing>
