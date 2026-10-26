// Excel's data validation on the sheet: the active cell's rule requested from the source as the cell changes (its dropdown
// list, open or closed, with its highlighted item, and its input message), and the alert for a rejected value with the user's answer.
import { computed, getCurrentScope, onScopeDispose, ref, shallowRef, watch } from 'vue'
import { cellAddress } from 'xlcalc/a1'
import type { SheetSource, SheetValidation, SheetValidationAlert } from '../contract/source'
import { listStep, stepIndex } from './list-keys'

/** How long the active cell stays before its rule is requested: holding an arrow key down a column requests once, at the end. */
export const VALIDATION_DELAY_MS = 60
/** Items the list displays at once, as in Excel; Page Up and Page Down move by them. */
export const LIST_ROWS = 8

export interface ValidationPlace {
  sheet: string
  cell: string
}

export interface ValidationDeps {
  source: () => SheetSource
  sheet: () => string
  /** The active cell; null for none. */
  active: () => { row: number; col: number } | null
  /** The active cell's displayed text: the list opens at the item equal to it. */
  text?: () => string
  /** An item chosen from the list (a click, Enter), by its text and position in the list: the sheet enters it into the cell. */
  pick: (text: string, place: ValidationPlace, item: number) => void
}

/** The active cell's rule as Excel displays it: requested from the source once the cell stays, an older response discarded, and the list's state. */
export function useValidation(deps: ValidationDeps) {
  /** What the source returned for `place`; null while it is requested, without a rule, or without the source's `validation`. */
  const current = shallowRef<SheetValidation | null>(null)
  const place = shallowRef<ValidationPlace | null>(null)
  const open = ref(false)
  /** The list's highlighted item; -1 for none. */
  const index = ref(-1)
  const list = computed(() => current.value?.list ?? null)
  const prompt = computed(() => current.value?.prompt ?? null)

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
  /** Requests the active cell's rule from the source after `wait` ms; a request for an earlier cell is aborted and its response discarded. */
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
  /** Requests again after `wait` ms (immediately by default), keeping the displayed rule until the response (a list of cells after an edit). */
  const refresh = (wait = 0) => ask(wait)

  function close(): void {
    open.value = false
    index.value = -1
  }
  /** Opens the list at the item the cell displays (Alt+Down, the arrow); its items are requested again, as Excel recalculates them on opening. False without a list. */
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
  /** Highlights the item `by` steps away, within the list; from none, Down goes to the first. */
  function move(by: number): void {
    const n = list.value?.length ?? 0
    if (n) index.value = stepIndex(index.value, by, n)
  }
  /** Enters an item into the cell (the highlighted one by default) and closes the list; false for no item. */
  function choose(i = index.value): boolean {
    const items = list.value
    const at = place.value
    if (!items || !at || i < 0 || i >= items.length) return false
    const text = items[i]!
    close()
    deps.pick(text, at, i)
    return true
  }

  /** Keys on the grid: Alt+Down opens the list; while it is open the arrows, Page Up and Page Down, Home and End move, Enter enters, Esc and Alt+Up close. True when handled. */
  function onKey(e: KeyboardEvent): boolean {
    if (e.isComposing) return false
    if (!open.value) {
      if (!(e.altKey && e.key === 'ArrowDown' && openList())) return false
      e.preventDefault()
      return true
    }
    // Alt+Down does nothing more on an open list, and Alt+Up closes it
    const by = e.altKey && (e.key === 'ArrowDown' || e.key === 'ArrowUp') ? null : listStep(e.key, list.value?.length ?? 0, LIST_ROWS)
    if (by !== null) move(by)
    else if (e.key === 'ArrowUp' || e.key === 'Escape') close()
    else if (e.key === 'Enter') void (choose() || close())
    else if (e.key !== 'ArrowDown') {
      // any other key closes the list and goes on to the grid (a character starts typing, Tab moves)
      close()
      return false
    }
    e.preventDefault()
    return true
  }

  function stop(): void {
    stopped = true
    drop()
  }
  // A new cell, sheet or source: the displayed rule goes at once, and the new rule is requested once the cell stays.
  watch(
    () => [where(), deps.source()] as const,
    ([at, src], old) => {
      if (old && old[1] === src && at?.sheet === old[0]?.sheet && at?.cell === old[0]?.cell) return
      close()
      current.value = null
      place.value = null
      ask(VALIDATION_DELAY_MS)
    },
    { immediate: true },
  )
  // When the scope it runs in (the sheet's) ends, requests stop.
  if (getCurrentScope()) onScopeDispose(stop)

  return { current, place, list, prompt, open, index, refresh, openList, close, toggle, choose, onKey }
}

export type Validation = ReturnType<typeof useValidation>

/** The sheet's own prompt for a formula that cannot be parsed: Fix reopens the editor, Restore leaves the cell unchanged, so Enter or a click elsewhere is a way out. */
export interface FormulaAlert {
  severity: 'formula'
  title?: undefined
  /** The source's reason. */
  message: string
  sheet: string
  cell: string
}
/** Excel's warning before cells containing more than one value are merged: OK merges, Cancel leaves them. */
export interface MergeAlert {
  severity: 'merge'
  title?: undefined
  /** The warning's text. */
  message: string
  sheet: string
  cell: string
}
/** Excel's dialog for an unclosed formula that it offers to correct: Yes enters the correction, No (or Esc) says what is missing. */
export interface TypoAlert {
  severity: 'typo'
  title?: undefined
  /** The question with the correction. */
  message: string
  sheet: string
  cell: string
}
/** Excel's message after No to a correction, with OK only: the editor reopens with the typed text. */
export interface ParenthesisAlert {
  severity: 'parenthesis'
  title?: undefined
  message: string
  sheet: string
  cell: string
}
/** Excel's message with OK alone over a dialog, such as Go To's for a bad reference. */
export interface NoticeAlert {
  severity: 'notice'
  title?: undefined
  message: string
}
/** Excel's questions before Data Validation opens: rules of more than one kind in the selection (OK erases them, Cancel), or cells without the rule (Yes extends it, No starts afresh, Cancel). */
export interface RulesAlert {
  severity: 'rulesMany' | 'rulesExtend'
  title?: undefined
  message: string
  sheet: string
  cell: string
}
/** What the alert displays: a rule's rejection, the prompt for a formula, Excel's correction of one and what is missing, the warning before a merge, a dialog's message, or Data Validation's questions. */
export type SheetAlert = SheetValidationAlert | FormulaAlert | TypoAlert | ParenthesisAlert | MergeAlert | NoticeAlert | RulesAlert

/** What the user's answer to an alert does with the rejected edit: send it again with `accept`, reopen the editor with the text, or leave the cell unchanged. */
export type AlertAnswer = 'accept' | 'retry' | 'cancel'
export type AlertButton = 'retry' | 'cancel' | 'yes' | 'no' | 'ok' | 'fix' | 'revert'

/** Excel's alert buttons by style, in Excel's order, and the answer each gives (Help is omitted); the formula prompt's buttons after them. */
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
  notice: [{ button: 'ok', answer: 'accept' }],
  merge: [
    { button: 'ok', answer: 'accept' },
    { button: 'cancel', answer: 'cancel' },
  ],
  rulesMany: [
    { button: 'ok', answer: 'accept' },
    { button: 'cancel', answer: 'cancel' },
  ],
  rulesExtend: [
    { button: 'yes', answer: 'accept' },
    { button: 'no', answer: 'retry' },
    { button: 'cancel', answer: 'cancel' },
  ],
}
/** The button an alert opens with, which Enter presses: Retry, No (a warning keeps a value only on a click on Yes), OK; Yes for Excel's correction and for extending a rule. */
export const ALERT_DEFAULT: Record<SheetAlert['severity'], AlertButton> = { stop: 'retry', warning: 'no', information: 'ok', formula: 'fix', typo: 'yes', parenthesis: 'ok', notice: 'ok', merge: 'ok', rulesMany: 'ok', rulesExtend: 'yes' }

/** The alert for a rejected value: shown until the user answers; one at a time, and a new one answers the old one with Cancel. */
export function useValidationAlert() {
  const alert = shallowRef<SheetAlert | null>(null)
  let settle: ((a: AlertAnswer) => void) | null = null

  /** Shows the alert and waits for the answer. */
  function ask(next: SheetAlert): Promise<AlertAnswer> {
    answer('cancel')
    alert.value = next
    return new Promise((resolve) => (settle = resolve))
  }
  /** The user's answer (the alert's `answer` event): the alert closes and `ask` resolves. */
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
