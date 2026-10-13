// Excel's Point mode: while a formula is edited and the caret stands where a reference may (after `=`, an operator, `(`,
// `:` or the list separator), a press on the sheet puts the cell's address at the caret, a drag a range's, a letter or a
// number a whole column's or row's, and in a field typing opened over a cell an arrow the address of the cell beside the
// edited one; the next press or arrow before anything is typed replaces the reference, Shift extends it.
import type { Ref } from 'vue'
import { MAX_COLS, MAX_ROWS, rangeAddress } from 'xlcalc/a1'
import type { CellPosition } from 'xlcalc/contract'
import { inLiteral } from './autocomplete'
import type { EditorState } from './editing'
import type { Geometry } from './geometry'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

/** What a press points at: cells, whole columns (a letter) or whole rows (a number). */
export type PointKind = 'cells' | 'columns' | 'rows'

/** The arrows' steps: rows and columns. */
const ARROWS: Record<string, [-1 | 0 | 1, -1 | 0 | 1]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] }

/** What a reference may follow besides the list separator: the formula's `=`, the operators, `(` and the range's `:`. */
const BEFORE_REF = new Set([...'=+-*/^&<>(:@'])

/** Whether a reference may stand at `at` in formula text: after one of `BEFORE_REF` or `separator` (spaces between), not in a string, a quoted sheet's name or a table's brackets. */
export function refFits(text: string, at: number, separator: string): boolean {
  if (!text.startsWith('=') || at < 1 || at > text.length) return false
  let i = at
  while (i > 1 && /\s/.test(text[i - 1]!)) i--
  const c = text[i - 1]!
  return (BEFORE_REF.has(c) || c === separator) && !inLiteral(text, i - 1)
}

export interface PointDeps {
  /** The open field's element. */
  input: () => HTMLTextAreaElement | null
  /** Show the draft's new text with its caret in the field. */
  show: () => void
  /** Close the list of functions. */
  closeList: () => void
  /** The book's separator of arguments. */
  separator: () => string
  /** Scroll a cell into the frame. */
  reveal: (row: number, col: number) => void
}

export function usePoint(s: SheetState, geo: Geometry, v: Viewport, editor: Ref<EditorState | null>, deps: PointDeps) {
  /** The reference the last press wrote: its field, its place in the draft, the draft with it, what it covers and where it began. */
  let run: { ed: EditorState; from: number; to: number; text: string; kind: PointKind; anchor: CellPosition; far: CellPosition } | null = null
  let dragging = false
  /** The last press was taken: its mousedown keeps the field's focus, its click and double click pick nothing. */
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
    if (!last && !refFits(ed.text, start, deps.separator())) return false
    const [first, far] = corners(at.kind, at.row, at.col)
    const anchor = last && e.shiftKey && last.kind === at.kind ? last.anchor : first
    run = last ? { ...last, kind: at.kind, anchor } : { ed, from: start, to: end, text: ed.text, kind: at.kind, anchor, far }
    write(far)
    taken = dragging = true
    return true
  }
  /** The run the caret stands at the end of, nothing typed since: the next press or arrow replaces its reference. */
  const going = (ed: EditorState, start: number, end: number) => (run !== null && run.ed === ed && run.text === ed.text && start === run.to && end === run.to ? run : null)
  /** The cell one shown line from `p` the arrow's way; at the grid's edge `p`. */
  function stepFrom(p: CellPosition, [dr, dc]: [-1 | 0 | 1, -1 | 0 | 1]): CellPosition {
    const row = dr ? s.rowOf(geo.nextRow(s.indexOf(p.row) ?? p.row, dr)) : p.row
    return { row, col: dc ? geo.nextCol(p.col, dc) : p.col }
  }
  /**
   * An arrow in a field typing opened over a cell (Enter mode; F2 or the formula bar move the caret): the cell beside the
   * edited one, or the reference moved from its first cell, or with Shift its far corner moved. True when taken.
   */
  function key(e: KeyboardEvent): boolean {
    const way = ARROWS[e.key]
    const ed = editor.value
    const el = deps.input()
    if (!way || !ed || !el || e.ctrlKey || e.metaKey || e.altKey || ed.mode !== 'enter' || ed.place !== 'cell' || ed.sheet !== s.sheetName.value) return false
    const start = el.selectionStart ?? ed.text.length
    const end = el.selectionEnd ?? start
    const last = going(ed, start, end)
    if (!last && !refFits(ed.text, start, deps.separator())) return false
    e.preventDefault()
    const extend = last && e.shiftKey && last.kind === 'cells'
    const anchor = extend ? last.anchor : stepFrom(last ? last.anchor : { row: ed.row, col: ed.col }, way)
    const far = extend ? stepFrom(last.far, way) : anchor
    run = last ? { ...last, kind: 'cells', anchor } : { ed, from: start, to: end, text: ed.text, kind: 'cells', anchor, far }
    write(far)
    if (!s.visible.value) s.stretch(far.row, far.col)
    deps.reveal(far.row, far.col)
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
  /** The reference from the anchor to `far` in place of the run's, the caret after it; a draft changed meanwhile ends the run. */
  function write(far: CellPosition): void {
    const ed = editor.value
    if (!run || ed !== run.ed || ed.text !== run.text) {
      run = null
      dragging = false
      return
    }
    const a = run.anchor
    const ref = rangeAddress({ top: Math.min(a.row, far.row), left: Math.min(a.col, far.col), bottom: Math.max(a.row, far.row), right: Math.max(a.col, far.col) })
    ed.text = run.text.slice(0, run.from) + ref + run.text.slice(run.to)
    run = { ...run, to: run.from + ref.length, text: ed.text, far }
    ed.caret = run.to
    deps.closeList()
    deps.show()
  }
  const up = (): void => void (dragging = false)
  /** The mousedown of a taken press must not move the focus: the field would lose it and commit. */
  function onMouseDown(e: MouseEvent): void {
    if (taken) e.preventDefault()
  }
  const took = (): boolean => taken

  return { down, move, up, key, onMouseDown, took }
}

export type Point = ReturnType<typeof usePoint>
