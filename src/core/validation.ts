// Excel's data validation on the sheet: the rule of the active cell asked of the source as the cell changes (its dropdown
// list, open or not, with its item, and its input message), and the alert of a refused value with the person's answer.
import { computed, getCurrentScope, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { SheetSource, SheetValidation, SheetValidationAlert } from '../source'

/** How long the active cell stays before its rule is asked: arrows held over a column ask once, at the end. */
export const VALIDATION_DELAY_MS = 60
/** Items the list shows at once, as Excel's; Page Up and Page Down move by them. */
export const LIST_ROWS = 8

/** A rectangle in the coordinates of the element the components are placed in (the grid's content), in px. */
export interface ValidationRect {
  top: number
  left: number
  width: number
  height: number
}

export interface ValidationPlace {
  sheet: string
  cell: string
}

export interface ValidationDeps {
  source: () => SheetSource
  sheet: () => string
  /** The active cell; null for none. */
  active: () => { row: number; col: number } | null
  /** The active cell's shown text: the list opens at the item it equals. */
  text?: () => string
  /** An item chosen from the list (a click, Enter), by its text and its place in the list: the sheet enters it into the cell. */
  pick: (text: string, place: ValidationPlace, item: number) => void
  /** Stops asking (the sheet's end); the end of the scope the composable runs in stops it too. */
  signal?: AbortSignal
  /** Milliseconds the active cell stays before it is asked, `VALIDATION_DELAY_MS` by default. */
  delay?: number
}

/** The rule of the active cell as Excel shows it: the source asked once the cell stays, an older answer dropped, the list's state. */
export function useValidation(deps: ValidationDeps) {
  /** What the source gave for `place`; null while it is asked, without a rule, or without the source's `validation`. */
  const current = shallowRef<SheetValidation | null>(null)
  const place = shallowRef<ValidationPlace | null>(null)
  const open = ref(false)
  /** The list's highlighted item; -1 for none. */
  const index = ref(-1)
  const list = computed(() => current.value?.list ?? null)
  const prompt = computed(() => current.value?.prompt ?? null)
  const delay = deps.delay ?? VALIDATION_DELAY_MS

  let timer: ReturnType<typeof setTimeout> | null = null
  let ctrl: AbortController | null = null
  let stopped = false

  const where = (): ValidationPlace | null => {
    const a = deps.active()
    return a ? { sheet: deps.sheet(), cell: cellAddress(a.row, a.col) } : null
  }
  function drop(): void {
    if (timer) clearTimeout(timer)
    timer = null
    ctrl?.abort()
    ctrl = null
  }
  /** Ask the source for the active cell's rule after `wait` ms; a request for an earlier cell is aborted, its answer dropped. */
  function ask(wait: number): void {
    drop()
    const at = where()
    const src = deps.source()
    if (stopped || !at || !src.validation) return
    const run = async (): Promise<void> => {
      timer = null
      const mine = (ctrl = new AbortController())
      let got: SheetValidation | null = null
      try {
        got = await src.validation!(at.sheet, at.cell, mine.signal)
      } catch {
        got = null
      }
      if (mine.signal.aborted) return
      ctrl = null
      current.value = got
      place.value = at
      const items = got?.list
      if (!items) close()
      else if (index.value >= items.length) index.value = items.length - 1
    }
    if (wait > 0) timer = setTimeout(() => void run(), wait)
    else void run()
  }
  /** Ask again after `wait` ms (now by default), the shown rule kept until the answer (a list of cells after an edit). */
  const refresh = (wait = 0) => ask(wait)

  function close(): void {
    open.value = false
    index.value = -1
  }
  /** Open the list at the item the cell shows (Alt+Down, the arrow); its items are asked again, as Excel computes them on opening. False without a list. */
  function openList(): boolean {
    const items = list.value
    if (!items) return false
    const shown = deps.text?.() ?? ''
    index.value = items.indexOf(shown)
    open.value = true
    refresh()
    return true
  }
  const toggle = () => (open.value ? close() : openList())
  /** Highlight the item `by` steps away, within the list; from none, down goes to the first. */
  function move(by: number): void {
    const n = list.value?.length ?? 0
    if (!n) return
    const from = index.value < 0 ? (by > 0 ? -1 : n) : index.value
    index.value = Math.min(n - 1, Math.max(0, from + by))
  }
  /** Enter an item into the cell (the highlighted one by default) and close the list; false for no item. */
  function choose(i = index.value): boolean {
    const items = list.value
    const at = place.value
    if (!items || !at || i < 0 || i >= items.length) return false
    const text = items[i]!
    close()
    deps.pick(text, at, i)
    return true
  }

  /** Keys on the grid: Alt+Down opens the list; while it is open the arrows, Page keys, Home and End move, Enter enters, Esc and Alt+Up close. True when taken. */
  function onKey(e: KeyboardEvent): boolean {
    if (e.isComposing) return false
    if (!open.value) {
      if (!(e.altKey && e.key === 'ArrowDown' && openList())) return false
      e.preventDefault()
      return true
    }
    const n = list.value?.length ?? 0
    const keys: Record<string, () => void> = {
      ArrowDown: () => (e.altKey ? undefined : move(1)),
      ArrowUp: () => (e.altKey ? close() : move(-1)),
      PageDown: () => move(LIST_ROWS),
      PageUp: () => move(-LIST_ROWS),
      Home: () => move(-n),
      End: () => move(n),
      Enter: () => void (choose() || close()),
      Escape: close,
    }
    const act = keys[e.key]
    if (!act) {
      // anything else closes the list and goes on to the grid (a character starts typing, Tab moves)
      close()
      return false
    }
    e.preventDefault()
    act()
    return true
  }

  function stop(): void {
    stopped = true
    drop()
  }
  // A new cell, sheet or source: what was shown goes at once, the new rule is asked once the cell stays.
  watch(
    () => [where(), deps.source()] as const,
    ([at, src], old) => {
      if (old && old[1] === src && at?.sheet === old[0]?.sheet && at?.cell === old[0]?.cell) return
      close()
      current.value = null
      place.value = null
      ask(delay)
    },
    { immediate: true },
  )
  deps.signal?.addEventListener('abort', stop, { once: true })
  if (getCurrentScope()) onScopeDispose(stop)

  return { current, place, list, prompt, open, index, refresh, openList, close, toggle, move, choose, onKey, stop }
}

export type Validation = ReturnType<typeof useValidation>

/** The sheet's own question on a formula that does not read: Fix opens the field again, Restore leaves the cell as it was, so Enter or a click away is a way out. */
export interface FormulaAlert {
  severity: 'formula'
  title?: undefined
  /** The source's reason. */
  message: string
  sheet: string
  cell: string
}
/** Excel's warning before cells holding more than one value are merged: OK merges, Cancel leaves them. */
export interface MergeAlert {
  severity: 'merge'
  title?: undefined
  /** The warning's text. */
  message: string
  sheet: string
  cell: string
}
/** Excel's window on a formula left open that it offers to correct: Yes enters the correction, No (or Esc) says what is missing. */
export interface TypoAlert {
  severity: 'typo'
  title?: undefined
  /** The question with the correction. */
  message: string
  sheet: string
  cell: string
}
/** Excel's message after No to a correction, with OK only: the field opens again with the text typed. */
export interface ParenthesisAlert {
  severity: 'parenthesis'
  title?: undefined
  message: string
  sheet: string
  cell: string
}
/** What the alert shows: a rule's refusal, the question on a formula, Excel's correction of one and what is missing, or the warning before a merge. */
export type SheetAlert = SheetValidationAlert | FormulaAlert | TypoAlert | ParenthesisAlert | MergeAlert

/** What the person's answer to an alert does with the refused edit: send it again with `accept`, open the field again with the text, or leave the cell as it was. */
export type AlertAnswer = 'accept' | 'retry' | 'cancel'
export type AlertButton = 'retry' | 'cancel' | 'yes' | 'no' | 'ok' | 'fix' | 'revert'

/** Excel's buttons of an alert by its style, in Excel's order, and what each answers (Help is left out); the formula's question after them. */
export const ALERT_BUTTONS: Record<SheetAlert['severity'], { button: AlertButton; answer: AlertAnswer }[]> = {
  stop: [
    { button: 'retry', answer: 'retry' },
    { button: 'cancel', answer: 'cancel' },
  ],
  warning: [
    { button: 'yes', answer: 'accept' },
    { button: 'no', answer: 'retry' },
    { button: 'cancel', answer: 'cancel' },
  ],
  information: [
    { button: 'ok', answer: 'accept' },
    { button: 'cancel', answer: 'cancel' },
  ],
  formula: [
    { button: 'fix', answer: 'retry' },
    { button: 'revert', answer: 'cancel' },
  ],
  typo: [
    { button: 'yes', answer: 'accept' },
    { button: 'no', answer: 'retry' },
  ],
  parenthesis: [{ button: 'ok', answer: 'accept' }],
  merge: [
    { button: 'ok', answer: 'accept' },
    { button: 'cancel', answer: 'cancel' },
  ],
}
/** The button an alert opens on, which Enter presses: Retry, No (a warning keeps a value only by a click on Yes), OK; Yes for Excel's correction. */
export const ALERT_DEFAULT: Record<SheetAlert['severity'], AlertButton> = { stop: 'retry', warning: 'no', information: 'ok', formula: 'fix', typo: 'yes', parenthesis: 'ok', merge: 'ok' }

/** The alert of a refused value: shown until the person answers; one at a time, a new one answering the old with Cancel. */
export function useValidationAlert() {
  const alert = shallowRef<SheetAlert | null>(null)
  let settle: ((a: AlertAnswer) => void) | null = null

  /** Show the alert and wait for the answer. */
  function ask(next: SheetAlert): Promise<AlertAnswer> {
    answer('cancel')
    alert.value = next
    return new Promise((resolve) => (settle = resolve))
  }
  /** The person's answer (the alert's `answer` event): the alert goes and `ask` resolves. */
  function answer(a: AlertAnswer): void {
    const done = settle
    settle = null
    alert.value = null
    done?.(a)
  }
  if (getCurrentScope()) onScopeDispose(() => answer('cancel'))

  return { alert, ask, answer }
}

export type ValidationAlertState = ReturnType<typeof useValidationAlert>
