// Excel's outline beside the grid (the sheet's `outline`): what each drawn line shows in the bars, the group and level
// buttons, and Group and Ungroup over whole lines from the menu and the keys (Shift+Alt+→ and ←, Cmd+Shift+K and J).
import { computed } from 'vue'
import { MAX_COLS, MAX_ROWS, rangeAddress } from 'xlcalc/a1'
import type { OutlineChange, OutlineGroup } from '../source'
import { linesOn, MAX_LEVEL } from '../lines'
import type { Commands } from './commands'
import { shortcutLetter } from './keys'
import type { Selection } from './selection'
import type { SheetState } from './state'
import type { Viewport } from './viewport'

export type OutlineAxis = 'rows' | 'columns'
export type OutlineOp = 'group' | 'ungroup'

/** A finger's tap on a bar presses the nearest button this close (px each way), a 44 px target around a 13 px button. */
const TOUCH_REACH_PX = 22

/** What a line shows in the slot of one level: a piece of a group's bracket, or the button of a group it is the summary line of. */
export interface OutlineMark {
  /** The slot, from 1 (the outer groups) at the bar's outer edge. */
  level: number
  /** The bracket runs along the line; `start` is its far end from the summary line, with a tick toward the headers. */
  bracket?: 'line' | 'start'
  /** The group whose button the line holds, with the bracket's end toward the group while it is expanded. */
  button?: OutlineGroup
}

/** One axis's groups by level (ordered by their first line) and by summary line. */
interface AxisIndex {
  levels: OutlineGroup[][]
  bySummary: Map<number, OutlineGroup[]>
  /** The summaries are after the groups (below, right). */
  after: boolean
}

function indexOf(groups: OutlineGroup[], after: boolean): AxisIndex {
  const levels: OutlineGroup[][] = []
  const bySummary = new Map<number, OutlineGroup[]>()
  for (const g of groups) {
    ;(levels[g.level - 1] ??= []).push(g)
    const s = after ? g.to + 1 : g.from - 1
    bySummary.set(s, [...(bySummary.get(s) ?? []), g])
  }
  for (let i = 0; i < levels.length; i++) (levels[i] ??= []).sort((a, b) => a.from - b.from)
  return { levels, bySummary, after }
}

/** The group of a sorted level's list that holds a line, by binary search. */
function holding(list: OutlineGroup[], line: number): OutlineGroup | null {
  let lo = 0
  let hi = list.length - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    const g = list[mid]!
    if (g.to < line) lo = mid + 1
    else if (g.from > line) hi = mid - 1
    else return g
  }
  return null
}

interface OutlineDeps {
  notify: (text: string) => void
  /** Ask whether a selection of cells groups its rows or its columns, as Excel's box does. */
  ask: (op: OutlineOp) => void
}

export function useOutline(s: SheetState, v: Viewport, sel: Selection, cmd: Commands, deps: OutlineDeps) {
  const outline = computed(() => s.sheet.value?.outline ?? null)
  /** The buttons press only with the source's `outline` on an editable sheet; without it the bars are drawn all the same. */
  const can = computed(() => cmd.can('outline'))
  const rows = computed(() => (outline.value && v.outlineSlots.value.rows ? indexOf(outline.value.rows, outline.value.summaryBelow) : null))
  const columns = computed(() => (outline.value && v.outlineSlots.value.cols ? indexOf(outline.value.columns, outline.value.summaryRight) : null))
  const m = () => s.messages.value

  /** What a line shows in each level's slot of its bar. */
  function marks(axis: OutlineAxis, line: number): OutlineMark[] {
    const idx = axis === 'rows' ? rows.value : columns.value
    if (!idx) return []
    const out: OutlineMark[] = []
    const summaries = idx.bySummary.get(line)
    for (let level = 1; level <= idx.levels.length; level++) {
      const button = summaries?.find((g) => g.level === level)
      if (button) {
        out.push({ level, button })
        continue
      }
      const g = holding(idx.levels[level - 1]!, line)
      if (g && !g.collapsed) out.push({ level, bracket: line === (idx.after ? g.from : g.to) ? 'start' : 'line' })
    }
    return out
  }
  /** A slot's middle from the bar's outer edge, px. */
  const at = (level: number) => (level - 0.5) * v.slotPx.value

  /** Lines as the contract names them: `3:6`, `B:D`. */
  const span = (axis: OutlineAxis, first: number, last: number) =>
    rangeAddress(axis === 'rows' ? { top: first, left: 1, bottom: last, right: MAX_COLS } : { top: 1, left: first, bottom: MAX_ROWS, right: last })
  function send(axis: OutlineAxis, change: (lines: { rows: string } | { columns: string }) => OutlineChange, first: number, last: number): void {
    const lines = linesOn(axis, span(axis, first, last))
    void cmd.outline(change(lines))
  }

  /** A group's button: Hide Detail when it is expanded, Show Detail when collapsed. */
  function toggle(axis: OutlineAxis, g: OutlineGroup): void {
    if (!can.value) return
    const op = g.collapsed ? 'expand' : 'collapse'
    send(axis, (lines) => ({ op, sheet: s.sheetName.value, level: g.level, ...lines }) as OutlineChange, g.from, g.to)
  }
  /** A level button: groups of that level and deeper collapsed, the shallower ones expanded. */
  function showLevel(axis: OutlineAxis, level: number): void {
    if (can.value) void cmd.outline({ op: 'showLevel', sheet: s.sheetName.value, axis, level })
  }

  /**
   * Group or Ungroup lines `first`..`last`. What Excel refuses is said here in the interface's language and not sent: a
   * Group of lines all at the deepest level (inside one group of it, as such groups are its longest runs), an Ungroup of
   * lines in no group.
   */
  function lines(op: OutlineOp, axis: OutlineAxis, first: number, last: number): void {
    if (!can.value) return
    const groups = outline.value?.[axis] ?? []
    const meets = (g: OutlineGroup) => g.from <= last && g.to >= first
    if (op === 'group' && groups.some((g) => g.level >= MAX_LEVEL && g.from <= first && g.to >= last)) return deps.notify(m().outlineDeepest)
    if (op === 'ungroup' && !groups.some(meets)) return deps.notify(m().outlineNone)
    send(axis, (lines) => ({ op, sheet: s.sheetName.value, ...lines }) as OutlineChange, first, last)
  }
  /** Group or Ungroup the selection: its whole rows or columns, else the box asks which, as Excel's does. */
  function run(op: OutlineOp): void {
    const r = sel.wholeRows.value
    const c = sel.wholeColumns.value
    if (r && !c) return lines(op, 'rows', r.first, Math.min(r.last, MAX_ROWS))
    if (c && !r) return lines(op, 'columns', c.first, c.last)
    deps.ask(op)
  }
  /** Excel's keys: Shift+Alt+→ groups and ← ungroups; on a Mac Cmd+Shift+K and J too. True when taken. */
  function key(e: KeyboardEvent): boolean {
    let op: OutlineOp | null = null
    if (e.shiftKey && e.altKey && !e.ctrlKey && !e.metaKey) op = e.key === 'ArrowRight' ? 'group' : e.key === 'ArrowLeft' ? 'ungroup' : null
    else if (e.metaKey && e.shiftKey && !e.ctrlKey && !e.altKey) {
      const letter = shortcutLetter(e)
      op = letter === 'k' ? 'group' : letter === 'j' ? 'ungroup' : null
    }
    if (!op || !can.value) return false
    e.preventDefault()
    run(op)
    return true
  }

  /** A finger's tap on a bar off its buttons presses the nearest one within reach. */
  function tap(e: MouseEvent): void {
    if (!v.touch.value || (e.target as HTMLElement).closest('button')) return
    let best: HTMLButtonElement | null = null
    let near = Infinity
    for (const b of s.scroller.value?.querySelectorAll<HTMLButtonElement>('.xs-ol-btn:not(:disabled)') ?? []) {
      const r = b.getBoundingClientRect()
      const dx = Math.abs(e.clientX - (r.left + r.width / 2))
      const dy = Math.abs(e.clientY - (r.top + r.height / 2))
      if (dx <= TOUCH_REACH_PX && dy <= TOUCH_REACH_PX && dx + dy < near) {
        best = b
        near = dx + dy
      }
    }
    best?.click()
  }

  return { outline, can, marks, at, toggle, showLevel, lines, key, tap }
}

export type Outline = ReturnType<typeof useOutline>
