// Excel's Formula AutoComplete in the editor: the functions whose names begin with the name typed at the caret, with the
// source's names requested once per source; Tab or a double click inserts one with its "(", Esc closes the list.
import { computed, getCurrentScope, onScopeDispose, ref, shallowRef, watch, type Ref } from 'vue'
import type { ScanOptions } from 'xlcalc/contract'
import { nameAt } from 'xlcalc/text'
import type { SheetSource } from '../contract/source'
import { askSource } from './ask'
import type { EditorState } from './editing'
import { listStep, stepIndex } from './list-keys'

/** Items the list displays at once; Page Up and Page Down move by them. */
export const FUNCTION_ROWS = 8

/** A source's names (with their uppercase forms, for matching), requested once; `list` is null while they load. */
interface Names {
  src: SheetSource
  list: { name: string; key: string }[] | null
  ctrl: AbortController
}

export interface AutocompleteDeps {
  source: () => SheetSource
  editor: Ref<EditorState | null>
  /** The open editor's element, for its caret. */
  input: () => HTMLTextAreaElement | null
  /** Shows the draft's new text with its caret in the editor. */
  show: () => void
  /** How the workbook's formula text is parsed (its list separator); null while its locale loads. */
  scan: () => ScanOptions | null
}

/** The list's state over the editor's draft: what it displays, the highlighted item, the keys it handles, the name it inserts. */
export function useAutocomplete(deps: AutocompleteDeps) {
  /** The functions whose names begin with the typed name, in the source's order. */
  const items = shallowRef<readonly string[]>([])
  /** The highlighted item, the first one when the list opens. */
  const index = ref(0)
  const open = computed(() => items.value.length > 0)
  /** The name in the draft that an item replaces. */
  let span: { from: number; to: number } | null = null
  /** The draft and caret the list was last requested for: names that arrive later are matched against them. */
  let last: { text: string; caret: number } | null = null
  /** The current source's names. */
  let names: Names | null = null
  let timer: ReturnType<typeof setTimeout> | null = null

  function close(): void {
    items.value = []
    index.value = 0
    span = null
  }
  /** Requests the current source's names unless they are loaded or loading; a failed request is retried when the next editor opens. */
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
  /** The list for the name at `caret` in `text`: the functions that begin with it, the first highlighted; closed when there are none. */
  function match(text: string, caret: number): void {
    const list = names?.src === deps.source() ? names.list : null
    const opts = deps.scan()
    // the library rejects a caret beyond the text, which an editor's stale caret may be
    const at = list && opts && caret <= text.length ? nameAt(text, caret, opts) : null
    const typed = at?.typed.toUpperCase()
    const got = typed ? list!.filter((n) => n.key.startsWith(typed)).map((n) => n.name) : []
    if (!at || !got.length) return close()
    items.value = got
    index.value = 0
    span = { from: at.from, to: at.to }
  }
  /** The editor's text changed by typing; `caret` is null when text is selected. */
  function update(text: string, caret: number | null): void {
    last = caret === null ? null : { text, caret }
    if (caret === null) close()
    else match(text, caret)
  }
  /** The caret moved in the editor (Left, Right, a click): the open list follows it once the move is done. */
  function follow(): void {
    if (!open.value) return
    if (timer) clearTimeout(timer)
    timer = setTimeout(() => {
      timer = null
      const el = deps.input()
      const ed = deps.editor.value
      const caret = el && el.selectionStart === el.selectionEnd ? el.selectionStart : null
      // The input's update already matched this draft and caret: a keystroke's follow-up does it once.
      if (open.value && el && ed && !(last && last.text === ed.text && last.caret === caret)) update(ed.text, caret)
    })
  }
  /** Highlights the item `by` steps away, within the list. */
  function move(by: number): void {
    const n = items.value.length
    if (n) index.value = stepIndex(index.value, by, n)
  }
  /** Inserts a function (the highlighted one by default) in place of the typed name, with its "(" unless one follows, and closes the list; false for none. */
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
   * Keys in the editor while the list is open, as in Excel: the arrows, Page Up and Page Down, Home and End move, Tab inserts,
   * Esc closes; Enter closes the list and commits the cell as typed, Left and Right move the caret and the list follows. True when handled.
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

  // An editor opened, moved between the cell and the formula bar, or closed, or a new source: the list goes away; the names are requested once an editor opens.
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
