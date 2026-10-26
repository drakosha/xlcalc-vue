// Excel's Spelling (F7) over the source's `spelling` and `learnWord`: from the active cell on, within a selection of several
// cells or the entire sheet, a misspelt word at a time with the cell selected, until the check comes back to where it started.
import { computed, ref } from 'vue'
import { cellAddress, parseCell } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import type { SpellingHit } from '../contract/workbook'
import type { Editing } from './editing'
import { readRange } from './range-read'
import type { Selection } from './selection'
import { useSheetReads, type SheetState } from './state'

/** The most words one check goes through, Change All's own changes included, so a source that never ends cannot hold the sheet. */
const MAX_STEPS = 10_000

/** The word shown in the dialog: the hit, its cell's text (and edit text) and the suggestion chosen for Change. */
export interface SpellingShown {
  hit: SpellingHit
  at: CellPosition
  text: string
  editText?: string
  choice: number
}

/** A position in the check's order: by rows, then columns, then the place in the cell's text. */
type Place = [number, number, number]
const before = (a: Place, b: Place) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]

export function useSpellCheck(s: SheetState, sel: Selection, ed: Editing, notify: (text: string) => void) {
  const read = useSheetReads(s)
  const available = computed(() => !!s.sheet.value && typeof s.source().spelling === 'function')
  /** Change and Change All write through `edit`: not on a read-only sheet. */
  const canChange = computed(() => s.editable.value)
  const canLearn = computed(() => typeof s.source().learnWord === 'function')
  const shown = ref<SpellingShown | null>(null)
  /** Excel's message that the check is complete. */
  const done = ref(false)
  const busy = ref(false)

  /** The check under way: its sheet and range, where it started, the last word's place and whether it passed the end. */
  let run: { sheet: string; range?: string; start: Place; last: Place | null; wrapped: boolean; ignore: string[]; changeAll: Map<string, string>; signal: AbortSignal; steps: number } | null = null

  /** Excel's Spelling from the active cell; false without the source's `spelling`. */
  function check(): boolean {
    const a = s.selected.value
    const r = sel.fullRect.value
    if (!available.value || !a || busy.value) return false
    const multi = !!r && (r.top !== r.bottom || r.left !== r.right)
    run = { sheet: s.sheetName.value, ...(multi ? { range: sel.selectionRef.value } : {}), start: [a.row, a.col, 0], last: null, wrapped: false, ignore: [], changeAll: new Map(), signal: read(), steps: 0 }
    done.value = false
    shown.value = null
    void next(cellAddress(a.row, a.col), 0)
    return true
  }

  /** The next misspelt word after position `from` of `after`'s text; past the place the check started, the check is complete. */
  async function next(after: string, from: number): Promise<void> {
    const call = s.sourceMethod('spelling')
    const r = run
    if (!call || !r) return
    busy.value = true
    try {
      for (;;) {
        if (++r.steps > MAX_STEPS) return finish()
        const hit = await call({ sheet: r.sheet, ...(r.range ? { range: r.range } : {}), after, from, ...(r.ignore.length ? { ignore: [...r.ignore] } : {}) }, r.signal)
        if (r !== run || r.signal.aborted) return
        const at = hit && parseCell(hit.cell)
        if (!hit || !at) return finish()
        const place: Place = [at.row, at.col, hit.at]
        if (before(place, r.last ?? r.start) < 0) r.wrapped = true
        if (r.wrapped && before(place, r.start) >= 0) return finish()
        r.last = place
        const cells = await readRange(s.source(), r.sheet, { top: at.row, left: at.col, bottom: at.row, right: at.col }, r.signal)
        if (r !== run || r.signal.aborted) return
        const cell = cells?.get(cellAddress(at.row, at.col))
        const text = typeof cell?.value === 'string' ? cell.value : (cell?.text ?? hit.word)
        // Change All: the word is replaced where it is found later, as Excel goes on.
        const by = r.changeAll.get(hit.word.toLocaleLowerCase())
        if (by !== undefined) {
          if (!(await write(hit, at, text, by, cell?.editText))) return close()
          after = hit.cell
          from = hit.at + by.length
          continue
        }
        sel.pickCell(at.row, at.col)
        sel.revealCell(at.row, at.col)
        shown.value = { hit, at, text, ...(cell?.editText !== undefined ? { editText: cell.editText } : {}), choice: 0 }
        return
      }
    } catch {
      if (r === run && !r.signal.aborted) {
        notify(s.messages.value.spellFailed)
        close()
      }
    } finally {
      if (r === run) busy.value = false
    }
  }

  /** The cell's text with the word replaced, through `edit` as typed text (a text that opens with `'` keeps it); false when it was not written. */
  async function write(hit: SpellingHit, at: CellPosition, text: string, by: string, editText: string | undefined): Promise<boolean> {
    const r = run!
    const lock = ed.lockAt(at.row, at.col)
    if (lock !== null) {
      notify(s.messages.value.locked(lock))
      return false
    }
    const changed = text.slice(0, hit.at) + by + text.slice(hit.at + hit.word.length)
    const quoted = editText?.startsWith("'") || /^[=+-]/.test(changed) ? `'${changed}` : changed
    const out = await ed.send([{ sheet: r.sheet, cell: hit.cell, text: quoted }], null)
    return !!out?.ok && r === run
  }

  /** The dialog's buttons: Ignore Once, Ignore All, Add to Dictionary, Change and Change All, each going on to the next word. */
  async function act(kind: 'ignoreOnce' | 'ignoreAll' | 'add' | 'change' | 'changeAll'): Promise<void> {
    const sh = shown.value
    const r = run
    if (!sh || !r || busy.value) return
    const { hit } = sh
    let from = hit.at + hit.word.length
    if (kind === 'ignoreAll') r.ignore.push(hit.word)
    else if (kind === 'add') {
      const learn = s.sourceMethod('learnWord')
      if (!learn) return
      busy.value = true
      try {
        await learn(hit.word)
      } catch {
        notify(s.messages.value.spellFailed)
      }
      busy.value = false
      // kept for the session too, for a checker that does not learn words
      r.ignore.push(hit.word)
    } else if (kind === 'change' || kind === 'changeAll') {
      const by = hit.suggestions[sh.choice]
      if (by === undefined || !canChange.value) return
      if (kind === 'changeAll') r.changeAll.set(hit.word.toLocaleLowerCase(), by)
      busy.value = true
      const ok = await write(hit, sh.at, sh.text, by, sh.editText)
      busy.value = false
      if (!ok) return close()
      from = hit.at + by.length
    }
    // the dialog stays with this word until the next one replaces it, or the message does
    if (r === run) await next(hit.cell, from)
  }

  function finish(): void {
    shown.value = null
    done.value = true
    run = null
    busy.value = false
  }
  /** Cancel, or OK on the message: the check ends. */
  function close(): void {
    run = null
    shown.value = null
    done.value = false
    busy.value = false
    s.focusGrid()
  }

  /** F7, Excel's Spelling; true when handled. */
  function key(e: KeyboardEvent): boolean {
    if (e.key !== 'F7' || e.ctrlKey || e.altKey || e.metaKey || e.shiftKey || !available.value) return false
    e.preventDefault()
    check()
    return true
  }

  return { available, canChange, canLearn, shown, done, busy, check, act, close, key }
}

export type SpellCheck = ReturnType<typeof useSpellCheck>
