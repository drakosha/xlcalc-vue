// Excel's outline for the sources: the groups xlcalc draws (`outlineGroups` of `xlcalc/a1`, Excel's marks) in the
// contract's terms, and for `memorySource`, which has no engine, the lines Excel's commands change (Group, Ungroup, a
// group's button, a level button), as measured in Excel for Windows 16 (build 20430, 12.10.2026). `./local` takes
// xlcalc's commands, whose measurements found more (its docs/outline.md): a summary line two groups share acts on its
// whole detail, a summary line keeps the mark under a level button, and a group at the sheet's edge has no button.
// TODO(xlcalc): the commands without the engine in `xlcalc/a1` would replace these.
import type { ColumnInput, OutlineGroup as BookGroup, OutlineSettings } from 'xlcalc/contract'
import { linesKey, outlineGroups } from 'xlcalc/a1'
import { axisOf, lineSpan, type LineAxis, MAX_LEVEL } from './lines'
import type { OutlineChange, SheetOutline } from './source'

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

/** A command's lines to change, or why it is refused. */
type OutlineEdit = { ok: true; patches: OutlinePatch[] } | { ok: false; message: string }

/**
 * One axis's groups as xlcalc draws them from its lines, the summaries after the detail (below, right) or before it: a
 * group is collapsed when Excel keeps the mark on its summary line, which it sets once the lines under it are all hidden.
 */
export function groupsOf(lines: readonly OutlineLine[], axis: LineAxis, after: boolean): BookGroup[] {
  // the lines as `SheetInput.rows` or `columns` hold them, which `outlineGroups` reads
  const input: Record<string, ColumnInput> = {}
  for (const l of lines) input[linesKey(l.from, l.to, axis)] = { ...(l.level ? { outlineLevel: l.level } : {}), ...(l.hidden ? { hidden: true } : {}), ...(l.collapsed ? { collapsed: true } : {}) }
  const side = { summaryBelow: after, summaryRight: after }
  return outlineGroups(axis === 'rows' ? { rows: input, outline: side } : { columns: input, outline: side })[axis]
}

/** A sheet's outline in the contract's terms from xlcalc's groups of both axes; none for a sheet without groups. */
export function outlineOf(groups: Record<LineAxis, BookGroup[]>, settings?: OutlineSettings | null): SheetOutline | undefined {
  if (!groups.rows.length && !groups.columns.length) return undefined
  const list = (axis: LineAxis) => groups[axis].map((g) => ({ from: g.first, to: g.last, level: g.level, collapsed: g.collapsed }))
  return { rows: list('rows'), columns: list('columns'), summaryBelow: settings?.summaryBelow ?? true, summaryRight: settings?.summaryRight ?? true }
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
 * The lines Excel's command changes on one axis, its `groups` as `groupsOf` gives them. Group puts each line one level
 * deeper and is refused when a line is at the deepest level already; Ungroup takes a level off the lines that have one and
 * leaves hidden lines hidden, refused when none has one. Collapse hides a group's lines and marks its summary line; Expand
 * shows them, all but those of the groups inside it whose summary line keeps the mark (one sharing its summary line is
 * expanded with it), and takes the mark off. A level button k hides the lines of level k and deeper, shows the shallower
 * ones and marks the summary lines of the collapsed groups that stay in view; lines outside groups stay as they are.
 */
export function outlineEdit(lines: readonly OutlineLine[], groups: readonly BookGroup[], change: OutlineChange): OutlineEdit {
  const axis = axisOf(change)
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
    for (const { summary: s, level } of groups) if (s !== null) marks.set(s, (marks.get(s) ?? false) || (level >= k && shownAfter(s)))
    for (const [s, mark] of marks) if (mark !== markAt(s)) patches.push({ from: s, to: s, collapsed: mark })
    return { ok: true, patches }
  }
  const text = 'rows' in change ? change.rows : change.columns
  const span = lineSpan(text, axis)
  if (!span) return { ok: false, message: `Not ${axis}: ${text}` }
  const { first: a, last: b } = span
  if (change.op === 'group' || change.op === 'ungroup') {
    const parts = pieces(lines, a, b)
    if (change.op === 'group') {
      // Excel refuses only when every line is at the deepest level; those that are stay there
      const deeper = parts.filter((p) => (p.line?.level ?? 0) < MAX_LEVEL)
      if (!deeper.length) return { ok: false, message: `The ${axis} ${text} cannot be grouped deeper than ${MAX_LEVEL} levels` }
      return { ok: true, patches: deeper.map((p) => ({ from: p.from, to: p.to, level: (p.line?.level ?? 0) + 1 })) }
    }
    const grouped = parts.filter((p) => (p.line?.level ?? 0) > 0)
    if (!grouped.length) return { ok: false, message: `The ${axis} ${text} are in no group` }
    return { ok: true, patches: grouped.map((p) => ({ from: p.from, to: p.to, level: p.line!.level - 1 })) }
  }
  const level = 'level' in change ? change.level : 0
  const g = groups.find((x) => x.first === a && x.last === b && x.level === level)
  if (!g) return { ok: false, message: `No group of level ${level} over the ${axis} ${text}` }
  // a group at the sheet's edge has no summary line to mark
  const summary = g.summary
  if (change.op === 'collapse') return { ok: true, patches: [{ from: a, to: b, hidden: true }, ...(summary !== null ? [{ from: summary, to: summary, collapsed: true }] : [])] }
  // Groups inside that stay collapsed: their summary line has the mark and is not this group's own.
  const holes = groups
    .filter((h) => h.level > g.level && h.first >= a && h.last <= b && h.summary !== null && h.summary !== summary && markAt(h.summary))
    .map((h) => ({ from: h.first, to: h.last }))
    .sort((x, y) => x.from - y.from)
  const patches: OutlinePatch[] = without(a, b, holes).map((r) => ({ ...r, hidden: false }))
  if (summary !== null && markAt(summary)) patches.push({ from: summary, to: summary, collapsed: false })
  return { ok: true, patches }
}
