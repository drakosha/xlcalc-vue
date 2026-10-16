// Excel's Formula AutoComplete in the field being edited: the functions whose names begin with the name typed at the
// caret, the source's names asked once per source; Tab or a double click enters one with its "(", Esc closes the list.
import { computed, getCurrentScope, onScopeDispose, ref, shallowRef, watch, type Ref } from 'vue'
import type { SheetSource } from '../source'
import { askSource } from './ask'
import type { EditorState } from './editing'
import { listStep, stepIndex } from './list-keys'

/** Items the list shows at once; Page Up and Page Down move by them. */
export const FUNCTION_ROWS = 8

/** A character of a function's name: a letter of any script, a digit, `.` or `_` (`СТАНДОТКЛОН.В`, `LOG10`). */
const NAME_CHAR = /[\p{L}\p{N}._]/u
/** What may stand right before a function's name: the formula's `=`, an operator, `(`, an argument separator, a space or a line break. */
const BEFORE_NAME = /[=+\-*/^&<>@(,;\s]/

/** Whether a formula's text before `upTo` leaves a string, a quoted sheet's name or a table's brackets open. */
function inLiteral(text: string, upTo: number): boolean {
  let quote: string | null = null
  let depth = 0
  for (let i = 1; i < upTo; i++) {
    const c = text[i]!
    if (quote) {
      if (c === quote) quote = null
    } else if (c === '"' || c === "'") quote = c
    else if (c === '[') depth++
    else if (c === ']') depth = Math.max(0, depth - 1)
  }
  return quote !== null || depth > 0
}

/** The name being typed at `caret` in a formula: where it begins and ends and the part before the caret; null where no function can stand (in a string, a quoted sheet's name, a table's brackets, after `!` or `$`). */
export function nameAt(text: string, caret: number): { from: number; to: number; typed: string } | null {
  if (!text.startsWith('=') || caret < 2 || caret > text.length) return null
  let from = caret
  while (from > 1 && NAME_CHAR.test(text[from - 1]!)) from--
  if (from === caret || !BEFORE_NAME.test(text[from - 1]!) || inLiteral(text, from)) return null
  let to = caret
  while (to < text.length && NAME_CHAR.test(text[to]!)) to++
  return { from, to, typed: text.slice(from, caret) }
}

/** A source's names (with their upper case, for matching) asked once; `list` is null while they come. */
interface Names {
  src: SheetSource
  list: { name: string; key: string }[] | null
  ctrl: AbortController
}

export interface AutocompleteDeps {
  source: () => SheetSource
  editor: Ref<EditorState | null>
  /** The open field's element, for its caret. */
  input: () => HTMLTextAreaElement | null
  /** Show the draft's new text with its caret in the field. */
  show: () => void
}

/** The list's state over the editor's draft: what it shows, the highlighted item, the keys it takes, the name it enters. */
export function useAutocomplete(deps: AutocompleteDeps) {
  /** The functions whose names begin with the typed name, in the source's order. */
  const items = shallowRef<readonly string[]>([])
  /** The highlighted item, the first one as the list opens. */
  const index = ref(0)
  const open = computed(() => items.value.length > 0)
  /** The name in the draft an item takes the place of. */
  let span: { from: number; to: number } | null = null
  /** The draft and caret the list was last asked for: names that come later are matched against them. */
  let last: { text: string; caret: number } | null = null
  /** The current source's names. */
  let names: Names | null = null
  let timer: ReturnType<typeof setTimeout> | null = null

  function close(): void {
    items.value = []
    index.value = 0
    span = null
  }
  /** Ask the current source for its names unless they are in or on the way; a failed request is asked again with the next field. */
  function load(): void {
    const src = deps.source()
    if (names?.src === src) return
    names?.ctrl.abort()
    names = null
    if (!src.functionNames) return
    const entry: Names = {
      src,
      list: null,
      ctrl: askSource(
        (signal) => src.functionNames!(signal),
        () => names === entry,
        (list) => {
          entry.list = list.map((name) => ({ name, key: name.toUpperCase() }))
          if (last && deps.editor.value?.text === last.text) match(last.text, last.caret)
        },
        () => (names = null),
      ),
    }
    names = entry
  }
  /** The list for the name at `caret` in `text`: the functions it begins, the first highlighted; closed when none. */
  function match(text: string, caret: number): void {
    const list = names?.src === deps.source() ? names.list : null
    const at = list ? nameAt(text, caret) : null
    const typed = at?.typed.toUpperCase()
    const got = typed ? list!.filter((n) => n.key.startsWith(typed)).map((n) => n.name) : []
    if (!at || !got.length) return close()
    items.value = got
    index.value = 0
    span = { from: at.from, to: at.to }
  }
  /** The field's text changed as typed; `caret` is null when text is selected. */
  function update(text: string, caret: number | null): void {
    last = caret === null ? null : { text, caret }
    if (caret === null) close()
    else match(text, caret)
  }
  /** The caret moved in the field (Left, Right, a click): the open list follows it once the move is done. */
  function follow(): void {
    if (!open.value) return
    if (timer) clearTimeout(timer)
    timer = setTimeout(() => {
      timer = null
      const el = deps.input()
      const ed = deps.editor.value
      const caret = el && el.selectionStart === el.selectionEnd ? el.selectionStart : null
      // The input's update already matched this draft and caret: a key's follow does it once.
      if (open.value && el && ed && !(last && last.text === ed.text && last.caret === caret)) update(ed.text, caret)
    })
  }
  /** Highlight the item `by` steps away, within the list. */
  function move(by: number): void {
    const n = items.value.length
    if (n) index.value = stepIndex(index.value, by, n)
  }
  /** Enter a function (the highlighted one by default) in place of the typed name, with its "(" unless one follows, and close the list; false for none. */
  function accept(i = index.value): boolean {
    const ed = deps.editor.value
    const name = items.value[i]
    if (!ed || !span || name === undefined) return false
    const rest = ed.text.slice(span.to)
    ed.text = ed.text.slice(0, span.from) + name + (rest.startsWith('(') ? '' : '(') + rest
    ed.caret = span.from + name.length + 1
    close()
    last = null
    deps.show()
    return true
  }

  /**
   * Keys in the field while the list is open, as Excel's: the arrows, Page Up and Down, Home and End move, Tab enters,
   * Esc closes; Enter closes it and commits the cell as typed, Left and Right move the caret and the list follows. True when taken.
   */
  function onKey(e: KeyboardEvent): boolean {
    if (!open.value) return false
    const plain = !(e.altKey || e.ctrlKey || e.metaKey || e.shiftKey)
    const by = plain ? listStep(e.key, items.value.length, FUNCTION_ROWS) : null
    if (by !== null) move(by)
    else if (plain && e.key === 'Tab') accept()
    else if (plain && e.key === 'Escape') close()
    else {
      if (e.key === 'Enter') close()
      else follow()
      return false
    }
    e.preventDefault()
    return true
  }

  // A field opened, moved between the cell and the bar or closed, or a new source: the list goes; the names are asked once a field opens.
  watch([() => deps.editor.value, () => deps.editor.value?.place, deps.source], ([ed]) => {
    close()
    last = null
    if (ed) load()
  })
  if (getCurrentScope())
    onScopeDispose(() => {
      names?.ctrl.abort()
      if (timer) clearTimeout(timer)
    })

  return { items, index, open, update, follow, accept, close, onKey }
}

export type Autocomplete = ReturnType<typeof useAutocomplete>
