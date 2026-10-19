// Excel's Point mode: while a formula is being edited and the caret is where a reference may go (`refAt` returns `insert`),
// a press on the sheet inserts the cell's address at the caret, a drag a range's, a column letter or a row number an entire
// column's or row's; in an editor opened by typing over a cell an arrow does what `refAt` says (a reference to the cell next
// to the edited one, the text entered, or nothing), and the next press or key before anything is typed moves the reference as on the sheet.
import type { Ref } from 'vue'
import { MAX_COLS, MAX_ROWS, rangeAddress } from 'xlcalc/a1'
import type { CellPosition, PointAction, ScanOptions } from 'xlcalc/contract'
import { refAt } from 'xlcalc/text'
import { rectOf } from '../rect'
import type { EditorState } from './editing'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** What a press points at: cells, entire columns (a letter) or entire rows (a number). */
export type PointKind = 'cells' | 'columns' | 'rows'

/** The arrows' steps: rows and columns. */
const ARROWS: Record<string, [-1 | 0 | 1, -1 | 0 | 1]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }
/** Keys that point besides the arrows, as they move on the sheet (measured in Excel: Home and Page Down). */
const MOVES = new Set(['Home', 'PageDown', 'PageUp'])

interface PointDeps {
  /** The open editor's element. */
  input: () => HTMLTextAreaElement | null
  /** Shows the draft's new text with its caret in the editor. */
  show: () => void
  /** Closes the function list. */
  closeList: () => void
  /** How the workbook's formula text is parsed (its locale); null while the locale loads. */
  scan: () => ScanOptions | null
  /** Enters the draft and moves the selection in the arrow's direction. */
  commit: (dr: -1 | 0 | 1, dc: -1 | 0 | 1) => void
}

export function usePoint(s: SheetState, sel: Selection, v: Viewport, editor: Ref<EditorState | null>, deps: PointDeps) {
  /** The reference the last press wrote: its editor, its position in the draft, the draft with it, what it covers and where it started. */
  let run: { ed: EditorState; from: number; to: number; text: string; kind: PointKind; anchor: CellPosition; far: CellPosition } | null = null
  let dragging = false
  /** The last press was handled: its mousedown keeps the editor's focus, and its click and double click select nothing. */
  let taken = false

  /** A press's or a move's corners: a column runs from row 1 to the last, a row from column 1 to the last. */
  const corners = (kind: PointKind, row: number, col: number): [CellPosition, CellPosition] =>
    kind === 'columns' ? [{ row: 1, col }, { row: MAX_ROWS, col }] : kind === 'rows' ? [{ row, col: 1 }, { row, col: MAX_COLS }] : [{ row, col }, { row, col }]

  /** A mouse press on the sheet (`at` null: on nothing to point at): true when it went into the formula. */
  function down(e: PointerEvent, at: { kind: PointKind; row: number; col: number } | null): boolean {
    taken = false
    const ed = editor.value
    const el = deps.input()
    if (!at || !ed || !el || e.button !== 0 || e.pointerType === 'touch' || ed.sheet !== s.sheetName.value || (e.target as HTMLElement).closest('a,button')) return false
    const start = el.selectionStart ?? ed.text.length
    const end = el.selectionEnd ?? start
    const last = going(ed, start, end)
    const act = last ? null : actionAt(ed.text, start)
    if (act && act.action !== 'insert') return false
    const [first, far] = corners(at.kind, at.row, at.col)
    const anchor = last && e.shiftKey && last.kind === at.kind ? last.anchor : first
    run = last ? { ...last, kind: at.kind, anchor } : { ...begin(ed, start, end, act), kind: at.kind, anchor, far }
    write(far)
    taken = dragging = true
    return true
  }
  /** The run the caret is at the end of, with nothing typed since: the next press or arrow replaces its reference. */
  const going = (ed: EditorState, start: number, end: number) => (run !== null && run.ed === ed && run.text === ed.text && start === run.to && end === run.to ? run : null)
  /** What an arrow does at `caret` with no reference inserted yet (`refAt`); nothing while the locale loads. */
  const actionAt = (text: string, caret: number): PointAction => {
    const scan = deps.scan()
    return scan ? refAt(text, caret, scan) : { action: 'none' }
  }
  /** A new reference's position in the draft; before the formula's first character it comes with its own `=` (`equals`), as Excel inserts it. */
  function begin(ed: EditorState, start: number, end: number, act: PointAction | null): { ed: EditorState; from: number; to: number; text: string } {
    if (act?.action !== 'insert' || !act.equals) return { ed, from: start, to: end, text: ed.text }
    ed.text = ed.text.slice(0, start) + '=' + ed.text.slice(end)
    return { ed, from: start + 1, to: start + 1, text: ed.text }
  }
  let seq = 0
  /**
   * A key in an editor opened by typing over a cell (Enter mode; with F2 or in the formula bar keys move the caret). An
   * inserted reference moves as the selection moves on the sheet: from its first cell, or with Shift its far corner. With no
   * reference inserted yet, an arrow does what `refAt` says: a reference from the edited cell (with Shift a range from it),
   * the text entered, or nothing; Home and Page Down or Page Up point only where a reference can go. True when handled.
   */
  function key(e: KeyboardEvent): boolean {
    const way = ARROWS[e.key]
    const ed = editor.value
    const el = deps.input()
    const moves = MOVES.has(e.key) && !e.ctrlKey && !e.metaKey
    if ((!way && !moves) || !ed || !el || e.altKey || ed.mode !== 'enter' || ed.place !== 'cell') return false
    const start = el.selectionStart ?? ed.text.length
    const end = el.selectionEnd ?? start
    const last = ed.sheet === s.sheetName.value ? going(ed, start, end) : null
    const action = last ? null : actionAt(ed.text, start)
    if (action) {
      const act = action.action
      if (way && act === 'commit') {
        e.preventDefault()
        deps.commit(...way)
        return true
      }
      if (way && act === 'none') {
        e.preventDefault()
        return true
      }
      if (act !== 'insert' || ed.sheet !== s.sheetName.value) return false
    }
    const extend = e.shiftKey && (!last || last.kind === 'cells')
    const from = last ? (extend ? last.far : last.anchor) : { row: ed.row, col: ed.col }
    const to = sel.keyTarget(from, e)
    if (!to) return false
    e.preventDefault()
    const mine = ++seq
    const text = ed.text
    // the source's `end` responds later: a key, a press or typing meanwhile takes precedence
    const point = (p: CellPosition | null): void => {
      if (!p || mine !== seq || editor.value !== ed || ed.text !== text) return
      const anchor = extend ? (last ? last.anchor : from) : p
      run = last ? { ...last, kind: 'cells', anchor } : { ...begin(ed, start, end, action), kind: 'cells', anchor, far: p }
      write(p)
      if (!s.visible.value) s.stretch(p.row, p.col)
      sel.revealCell(p.row, p.col)
    }
    if (to instanceof Promise) void to.then(point)
    else point(to)
    return true
  }
  /** The drag's far corner under the pointer. */
  function move(e: PointerEvent): void {
    const box = dragging && run ? s.scroller.value?.getBoundingClientRect() : null
    if (!box || !run) return
    const p = v.cellAt(e.clientX - box.left, e.clientY - box.top)
    const [, far] = corners(run.kind, s.rowOf(p.d), p.col)
    if (far.row !== run.far.row || far.col !== run.far.col) write(far)
  }
  /** Replaces the run's reference with the one from the anchor to `far`, with the caret after it; a draft changed meanwhile ends the run. */
  function write(far: CellPosition): void {
    const ed = editor.value
    if (!run || ed !== run.ed || ed.text !== run.text) {
      run = null
      dragging = false
      return
    }
    const a = run.anchor
    const ref = rangeAddress(rectOf(a, far))
    ed.text = run.text.slice(0, run.from) + ref + run.text.slice(run.to)
    run = { ...run, to: run.from + ref.length, text: ed.text, far }
    ed.caret = run.to
    deps.closeList()
    deps.show()
  }
  const up = (): void => void (dragging = false)
  /** The mousedown of a handled press must not move focus: the editor would lose it and commit. */
  function onMouseDown(e: MouseEvent): void {
    if (taken) e.preventDefault()
  }
  const took = (): boolean => taken

  return { down, move, up, key, onMouseDown, took }
}

export type Point = ReturnType<typeof usePoint>
