// Excel's Define Name (the New Name and Edit Name dialog) and Name Manager (Ctrl+F3) over the sheet's names: each OK, Delete
// and Commit replaces the workbook's names in one call to `setNames`, and a refusal is shown with its message, as Excel's alert.
import { computed, ref, shallowRef } from 'vue'
import type { SheetName } from '../contract/workbook'
import { nameKey, type Names } from './names'
import { rangeRef } from './place'
import type { Selection } from './selection'
import type { SheetState } from './state'

/** The New Name or Edit Name dialog's fields: the name, its scope (a sheet's name, '' for the workbook) and Refers to, with `=`. */
export interface NameDraft {
  /** The name being edited, by its key; null for a new one. */
  editing: string | null
  name: string
  scope: string
  ref: string
}

/** The Name Manager's state: the selected name by its key, and Refers to as edited in the box at the bottom (null: unchanged). */
export interface NameManagerState {
  selected: string | null
  ref: string | null
}

/** An alert over the dialogs: a refusal (OK), the question before a deletion (OK, Cancel), before closing with an edited reference (Yes, No). */
export interface NameAsk {
  kind: 'refused' | 'delete' | 'save'
  message: string
}

/** A formula as typed in Refers to, without its `=`, as the source takes it. */
const formulaOf = (typed: string): string => typed.trim().replace(/^=/, '')

export function useNameDialogs(s: SheetState, sel: Selection, names: Names, deps: { editing: () => boolean }) {
  const draft = ref<NameDraft | null>(null)
  const manager = ref<NameManagerState | null>(null)
  const ask = shallowRef<NameAsk | null>(null)
  let settle: ((yes: boolean) => void) | null = null
  /** A call to the source is under way: OK, Delete and Commit wait for it. */
  const busy = ref(false)

  /** Asks and waits for the answer: true for OK or Yes. */
  function question(kind: NameAsk['kind'], message: string): Promise<boolean> {
    answer(false)
    ask.value = { kind, message }
    return new Promise((resolve) => (settle = resolve))
  }
  function answer(yes: boolean): void {
    const done = settle
    settle = null
    ask.value = null
    done?.(yes)
  }

  /** The Name Manager's list: every name of the workbook, by name and then by scope, as Excel sorts it. */
  const rows = computed(() => [...names.list.value].sort((a, b) => names.compare(a.name, b.name) || names.compare(a.scope ?? '', b.scope ?? '')))
  const selected = computed(() => (manager.value?.selected ? (rows.value.find((n) => nameKey(n) === manager.value!.selected) ?? null) : null))
  /** The box at the bottom: the selected name's Refers to, as edited or as it is. */
  const refText = computed(() => manager.value?.ref ?? (selected.value ? `=${selected.value.ref}` : ''))
  const refChanged = computed(() => !!selected.value && manager.value?.ref != null && formulaOf(manager.value.ref) !== selected.value.ref)

  /** Excel's Define Name: the New Name dialog for the selection, with Refers to the selection as an absolute reference with its sheet; false when names cannot be defined or a cell is being edited. */
  function define(): boolean {
    const rect = sel.fullRect.value
    if (!names.editable.value || !s.sheet.value || !rect || deps.editing()) return false
    draft.value = { editing: null, name: '', scope: '', ref: `=${rangeRef(s.sheetName.value, rect)}` }
    return true
  }
  /** Excel's Name Manager; false without the source's names or while a cell is being edited, when Excel does not open it either. */
  function openManager(): boolean {
    if (!names.available.value || !s.sheet.value || deps.editing()) return false
    manager.value = { selected: rows.value[0] ? nameKey(rows.value[0]) : null, ref: null }
    return true
  }
  /** Selects a name in the list; an edit of the previous name's reference is dropped. */
  function select(n: SheetName): void {
    if (manager.value) manager.value = { selected: nameKey(n), ref: null }
  }
  function edit(): void {
    const n = selected.value
    if (!n || !names.editable.value) return
    draft.value = { editing: nameKey(n), name: n.name, scope: n.scope ?? '', ref: `=${n.ref}` }
  }

  /** Replaces the names; a refusal or a failure is shown with its message. True when the source accepted them. */
  async function commit(next: SheetName[]): Promise<boolean> {
    if (busy.value) return false
    busy.value = true
    try {
      const out = await names.save(next)
      if (!out) return false
      if (!out.ok) await question('refused', out.message || s.messages.value.editFailed)
      return out.ok
    } finally {
      busy.value = false
    }
  }
  /** OK in New Name or Edit Name: the name added or put in place of the edited one; the dialog stays open after a refusal. */
  async function submit(): Promise<void> {
    const d = draft.value
    if (!d) return
    const entry: SheetName = { name: d.name.trim(), ref: formulaOf(d.ref), ...(d.scope ? { scope: d.scope } : {}) }
    const list = names.list.value
    const next = d.editing === null ? [...list, entry] : list.map((n) => (nameKey(n) === d.editing ? entry : n))
    if (!(await commit(next))) return
    draft.value = null
    if (manager.value) manager.value = { selected: nameKey(entry), ref: null }
    else s.focusGrid()
  }
  function closeDraft(): void {
    draft.value = null
    if (!manager.value) s.focusGrid()
  }
  /** Delete, after Excel's question. */
  async function remove(): Promise<void> {
    const n = selected.value
    if (!n || !names.editable.value || busy.value || !(await question('delete', s.messages.value.nameDeleteAsk(n.name)))) return
    const key = nameKey(n)
    if (await commit(names.list.value.filter((x) => nameKey(x) !== key)) && manager.value) manager.value = { selected: null, ref: null }
  }
  /** Refers to edited in the box at the bottom. */
  function editRef(text: string): void {
    if (manager.value && selected.value && names.editable.value) manager.value = { ...manager.value, ref: text }
  }
  /** Cancel (the cross): the box shows the name's reference again. */
  function cancelRef(): void {
    if (manager.value) manager.value.ref = null
  }
  /** Commit (the check mark): the selected name with the edited reference; true when there was nothing to commit or the source accepted it. */
  async function commitRef(): Promise<boolean> {
    const n = selected.value
    const typed = manager.value?.ref
    if (!n || typed == null || !refChanged.value) {
      cancelRef()
      return true
    }
    const key = nameKey(n)
    const ok = await commit(names.list.value.map((x) => (nameKey(x) === key ? { ...x, ref: formulaOf(typed) } : x)))
    if (ok && manager.value) manager.value.ref = null
    return ok
  }
  /** Close: an edited reference not yet committed is offered for saving first, as Excel asks. */
  async function closeManager(): Promise<void> {
    if (!manager.value || draft.value || busy.value) return
    if (refChanged.value && (await question('save', s.messages.value.nameSaveAsk)) && !(await commitRef())) return
    manager.value = null
    s.focusGrid()
  }

  return { draft, manager, ask, answer, busy, rows, selected, refText, refChanged, define, openManager, select, edit, submit, closeDraft, remove, editRef, commitRef, cancelRef, closeManager }
}

export type NameDialogs = ReturnType<typeof useNameDialogs>
