// Excel's Formula AutoComplete in the field being edited: the functions whose names begin with the name typed at the
// caret, the source's names asked once per source; Tab or a double click enters one with its "(", Esc closes the list.
import { computed, getCurrentScope, onScopeDispose, ref, shallowRef, watch, type Ref } from 'vue'
import type { ScanOptions } from 'xlcalc/contract'
import { nameAt } from 'xlcalc/text'
import type { SheetSource } from '../source'
import type { EditorState } from './editing'

/** Items the list shows at once; Page Up and Page Down move by them. */
export const FUNCTION_ROWS = 8

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
  /** How the book's formula text reads (its list separator); null while its locale loads. */
  scan: () => ScanOptions | null
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
    const entry: Names = { src, list: null, ctrl: new AbortController() }
    names = entry
    Promise.resolve()
      .then(() => src.functionNames!(entry.ctrl.signal))
      .then(
        (list) => {
          if (names !== entry) return
          entry.list = list.map((name) => ({ name, key: name.toUpperCase() }))
          if (last && deps.editor.value?.text === last.text) match(last.text, last.caret)
        },
        () => {
          if (names === entry) names = null
        },
      )
  }
  /** The list for the name at `caret` in `text`: the functions it begins, the first highlighted; closed when none. */
  function match(text: string, caret: number): void {
    const list = names?.src === deps.source() ? names.list : null
    const opts = deps.scan()
    // the library refuses a caret past the text, which a field's stale caret may be
    const at = list && opts && caret <= text.length ? nameAt(text, caret, opts) : null
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
    if (n) index.value = Math.min(n - 1, Math.max(0, index.value + by))
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
    const n = items.value.length
    const keys: Record<string, () => unknown> = {
      ArrowDown: () => move(1),
      ArrowUp: () => move(-1),
      PageDown: () => move(FUNCTION_ROWS),
      PageUp: () => move(-FUNCTION_ROWS),
      Home: () => move(-n),
      End: () => move(n),
      Tab: () => accept(),
      Escape: close,
    }
    const act = e.altKey || e.ctrlKey || e.metaKey || e.shiftKey ? undefined : keys[e.key]
    if (act) {
      e.preventDefault()
      act()
      return true
    }
    if (e.key === 'Enter') close()
    else follow()
    return false
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

  return { items, index, open, update, follow, move, accept, close, onKey }
}

export type Autocomplete = ReturnType<typeof useAutocomplete>
