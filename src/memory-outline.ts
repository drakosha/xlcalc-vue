// Excel's outline rules for `memorySource`, which has no engine: the groups Excel draws from one axis's outline levels and
// hidden lines, and the lines its commands change (Group, Ungroup, a group's button, a level button), as measured in Excel
// for Windows 16 (build 20430, 12.10.2026); `MAX_LEVEL` serves the grid too. `./local` takes xlcalc's groups and commands,
// whose measurements found more (its docs/outline.md): a summary line two groups share acts on its whole detail, Excel marks
// a summary line once its detail is all hidden and keeps the mark inside a hidden group, Group deepens the lines beside those
// at level 7, and a group at the sheet's edge has no button.
// TODO(xlcalc): the commands without the engine in `xlcalc/a1` would replace these, if the main entry has room (`outlineGroups` alone is 4.3 KB).
import { MAX_COLS, MAX_ROWS, parseLines, parseRange } from 'xlcalc/a1'
import type { OutlineChange, OutlineGroup, SheetOutline } from './source'

/** Excel's deepest outline level: an eighth Group of the same lines fails. */
export const MAX_LEVEL = 7

/** Lines `from`..`to` of one axis with the same outline level, hidden state and collapsed mark; lines left out have none. */
export interface OutlineLine {
  from: number
  to: number
  level: number
  hidden: boolean
  collapsed: boolean
}

/** What a command sets on lines `from`..`to`: the outline level (0: none), hidden or shown, the collapsed mark. */
interface OutlinePatch {
  from: number
  to: number
  level?: number
  hidden?: boolean
  collapsed?: boolean
}

export type OutlineAxis = 'rows' | 'columns'

/** A command's lines to change, or why it is refused. */
type OutlineEdit = { ok: true; axis: OutlineAxis; patches: OutlinePatch[] } | { ok: false; message: string }

/**
 * The groups of one axis: for each level, the longest runs of lines next to each other at that level or deeper. A group
 * is collapsed when all its lines are hidden, by its button or by hand: Excel reads it so (ShowDetail of its summary line
 * is False) and saves the collapsed mark then; the mark itself only remembers a group inside a collapsed one.
 */
export function groupsOf(lines: readonly OutlineLine[]): OutlineGroup[] {
  const out: OutlineGroup[] = []
  for (let level = 1; level <= MAX_LEVEL; level++) {
    let open: OutlineGroup | null = null
    for (const l of lines) {
      if (l.level < level) continue
      if (open && l.from === open.to + 1) {
        open.to = l.to
        open.collapsed &&= l.hidden
      } else {
        if (open) out.push(open)
        open = { from: l.from, to: l.to, level, collapsed: l.hidden }
      }
    }
    if (open) out.push(open)
  }
  return out.sort((a, b) => a.from - b.from || a.level - b.level)
}

/** A sheet's outline from its rows' and columns' lines; none for a sheet without groups. */
export function outlineOf(rows: readonly OutlineLine[], columns: readonly OutlineLine[], settings?: { summaryBelow?: boolean; summaryRight?: boolean } | null): SheetOutline | undefined {
  const r = groupsOf(rows)
  const c = groupsOf(columns)
  if (!r.length && !c.length) return undefined
  return { rows: r, columns: c, summaryBelow: settings?.summaryBelow ?? true, summaryRight: settings?.summaryRight ?? true }
}

/** Lines as the component and a book name them: `3:6`, `5` or `B:D`, `C`; null for anything else. */
function spanOf(text: string, axis: OutlineAxis): { first: number; last: number } | null {
  const lines = parseLines(text, axis)
  if (lines) return lines
  const r = parseRange(text)
  if (r?.kind !== axis) return null
  return axis === 'rows' ? { first: r.top, last: r.bottom } : { first: r.left, last: r.right }
}

/** Lines `a`..`b` in pieces: each with the line entry over it, or null where no entry is. */
function pieces(lines: readonly OutlineLine[], a: number, b: number): { from: number; to: number; line: OutlineLine | null }[] {
  const out: { from: number; to: number; line: OutlineLine | null }[] = []
  let at = a
  for (const l of lines) {
    if (l.to < at) continue
    if (l.from > b) break
    if (l.from > at) out.push({ from: at, to: l.from - 1, line: null })
    const to = Math.min(l.to, b)
    out.push({ from: Math.max(l.from, at), to, line: l })
    at = to + 1
    if (at > b) break
  }
  if (at <= b) out.push({ from: at, to: b, line: null })
  return out
}

/** The entry over a line, by binary search. */
function lineAt(lines: readonly OutlineLine[], i: number): OutlineLine | null {
  let lo = 0
  let hi = lines.length - 1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    const l = lines[mid]!
    if (l.to < i) lo = mid + 1
    else if (l.from > i) hi = mid - 1
    else return l
  }
  return null
}

/** Ranges `a`..`b` less the sorted, merged holes. */
function without(a: number, b: number, holes: { from: number; to: number }[]): { from: number; to: number }[] {
  const out: { from: number; to: number }[] = []
  let at = a
  for (const h of holes) {
    if (h.to < at) continue
    if (h.from > b) break
    if (h.from > at) out.push({ from: at, to: h.from - 1 })
    at = Math.max(at, h.to + 1)
  }
  if (at <= b) out.push({ from: at, to: b })
  return out
}

/**
 * The lines Excel's command changes on one axis, `after` its summaries' side (below or right). Group puts each line one
 * level deeper and is refused when a line is at the deepest level already; Ungroup takes a level off the lines that have
 * one and leaves hidden lines hidden, refused when none has one. Collapse hides a group's lines and marks its summary line;
 * Expand shows them, all but those of the groups inside it whose summary line keeps the mark (one sharing its summary line
 * is expanded with it), and takes the mark off. A level button k hides the lines of level k and deeper, shows the
 * shallower ones and marks the summary lines of the collapsed groups that stay in view; lines outside groups stay as they are.
 */
export function outlineEdit(lines: readonly OutlineLine[], after: boolean, change: OutlineChange): OutlineEdit {
  const axis: OutlineAxis = change.op === 'showLevel' ? change.axis : 'rows' in change ? 'rows' : 'columns'
  const words = axis === 'rows' ? 'rows' : 'columns'
  const last = axis === 'rows' ? MAX_ROWS : MAX_COLS
  const markAt = (i: number) => lineAt(lines, i)?.collapsed ?? false
  if (change.op === 'showLevel') {
    if (!Number.isInteger(change.level) || change.level < 1) return { ok: false, message: `No outline level ${change.level}` }
    const k = change.level
    const patches: OutlinePatch[] = []
    for (const l of lines) if (l.level > 0) patches.push({ from: l.from, to: l.to, hidden: l.level >= k })
    const shownAfter = (i: number) => {
      const l = lineAt(lines, i)
      return l && l.level > 0 ? l.level < k : !l?.hidden
    }
    const marks = new Map<number, boolean>()
    for (const g of groupsOf(lines)) {
      const s = after ? g.to + 1 : g.from - 1
      if (s >= 1 && s <= last) marks.set(s, (marks.get(s) ?? false) || (g.level >= k && shownAfter(s)))
    }
    for (const [s, mark] of marks) if (mark !== markAt(s)) patches.push({ from: s, to: s, collapsed: mark })
    return { ok: true, axis, patches }
  }
  const text = 'rows' in change ? change.rows : change.columns
  const span = spanOf(text, axis)
  if (!span) return { ok: false, message: `Not ${words}: ${text}` }
  const { first: a, last: b } = span
  if (change.op === 'group' || change.op === 'ungroup') {
    const parts = pieces(lines, a, b)
    if (change.op === 'group') {
      if (parts.some((p) => (p.line?.level ?? 0) >= MAX_LEVEL)) return { ok: false, message: `The ${words} ${text} cannot be grouped deeper than ${MAX_LEVEL} levels` }
      return { ok: true, axis, patches: parts.map((p) => ({ from: p.from, to: p.to, level: (p.line?.level ?? 0) + 1 })) }
    }
    const grouped = parts.filter((p) => (p.line?.level ?? 0) > 0)
    if (!grouped.length) return { ok: false, message: `The ${words} ${text} are in no group` }
    return { ok: true, axis, patches: grouped.map((p) => ({ from: p.from, to: p.to, level: p.line!.level - 1 })) }
  }
  const level = 'level' in change ? change.level : 0
  const groups = groupsOf(lines)
  const g = groups.find((x) => x.from === a && x.to === b && x.level === level)
  if (!g) return { ok: false, message: `No group of level ${level} over the ${words} ${text}` }
  const summary = after ? b + 1 : a - 1
  const mark = summary >= 1 && summary <= last
  if (change.op === 'collapse') return { ok: true, axis, patches: [{ from: a, to: b, hidden: true }, ...(mark ? [{ from: summary, to: summary, collapsed: true }] : [])] }
  // Groups inside that stay collapsed: their summary line has the mark and is not this group's own.
  const holes = groups
    .filter((h) => h.level > g.level && h.from >= a && h.to <= b)
    .filter((h) => {
      const s = after ? h.to + 1 : h.from - 1
      return s !== summary && markAt(s)
    })
    .map((h) => ({ from: h.from, to: h.to }))
    .sort((x, y) => x.from - y.from)
  const patches: OutlinePatch[] = without(a, b, holes).map((r) => ({ ...r, hidden: false }))
  if (mark && markAt(summary)) patches.push({ from: summary, to: summary, collapsed: false })
  return { ok: true, axis, patches }
}
