// Excel's outline for the sources: the groups xlcalc draws (`outlineGroups` of `xlcalc/a1`) in the contract's terms, and
// for `memorySource`, which has no engine, Excel's commands as xlcalc's `outlineEdit` runs them on a sheet's rows or
// columns alone, with the same commands and rejections as `./local`'s `editStructure` (xlcalc's docs/outline.md).
import type { OutlineEditOp, OutlineGroup as BookGroup, OutlineSettings, RowInput, SheetInput } from 'xlcalc/contract'
import { colLetters, linesKey, outlineEdit, outlineGroups } from 'xlcalc/a1'
import { axisOf, lineSpan, type LineAxis } from './lines'
import type { OutlineChange, SheetOutline } from './contract/source'

/** One axis's rows or columns as `memorySource` stores them, one by one: outline levels, collapsed flags and hidden ones. */
export interface AxisLines {
  levels: Map<number, number>
  marks: Set<number>
  hidden: Set<number>
}

/** The summary settings alone, as xlcalc takes them: an omitted field is Excel's default. */
export function outlineSettings(o?: OutlineSettings | null): OutlineSettings {
  return { ...(o?.summaryBelow !== undefined ? { summaryBelow: o.summaryBelow } : {}), ...(o?.summaryRight !== undefined ? { summaryRight: o.summaryRight } : {}) }
}

/** One axis's rows or columns as `SheetInput.rows` or `columns` hold them, one key per row or column, which xlcalc's outline reads. */
function inputOf(lines: AxisLines, axis: LineAxis): Record<string, RowInput> {
  const out: Record<string, RowInput> = {}
  for (const i of new Set([...lines.levels.keys(), ...lines.marks, ...lines.hidden])) {
    const level = lines.levels.get(i)
    out[linesKey(i, i, axis)] = { ...(level ? { outlineLevel: level } : {}), ...(lines.hidden.has(i) ? { hidden: true } : {}), ...(lines.marks.has(i) ? { collapsed: true } : {}) }
  }
  return out
}

/** A sheet's rows and columns on both axes as xlcalc's outline reads them, with its summaries where the settings put them. */
export function outlineSheet(axes: Record<LineAxis, AxisLines>, settings?: OutlineSettings | null): Pick<SheetInput, 'rows' | 'columns' | 'outline'> {
  return { rows: inputOf(axes.rows, 'rows'), columns: inputOf(axes.columns, 'columns'), outline: outlineSettings(settings) }
}

/** A sheet's outline in the contract's terms from xlcalc's groups on both axes; none for a sheet without groups. */
export function outlineOf(groups: Record<LineAxis, BookGroup[]>, settings?: OutlineSettings | null): SheetOutline | undefined {
  if (!groups.rows.length && !groups.columns.length) return undefined
  const list = (axis: LineAxis) => groups[axis].map((g) => ({ from: g.first, to: g.last, level: g.level, collapsed: g.collapsed }))
  return { rows: list('rows'), columns: list('columns'), summaryBelow: settings?.summaryBelow ?? true, summaryRight: settings?.summaryRight ?? true }
}

/**
 * The component's outline command as xlcalc's, or why there is none. A group's button is Show or Hide Detail at its summary
 * row or column, with the group found among the axis's groups by its rows or columns and level; a group at the edge of the
 * sheet has no summary row or column and, as in Excel, no button.
 */
export function outlineOp(change: OutlineChange, groups: (axis: LineAxis) => readonly BookGroup[]): OutlineEditOp | { refused: string } {
  if (change.op === 'showLevel') return change.axis === 'rows' ? { op: 'showLevels', rows: change.level } : { op: 'showLevels', columns: change.level }
  const axis = axisOf(change)
  const text = 'rows' in change ? change.rows : change.columns
  const span = lineSpan(text, axis)
  if (!span) return { refused: `Not ${axis}: ${text}` }
  if (!('level' in change)) {
    const lines = linesKey(span.first, span.last, axis)
    return axis === 'rows' ? { op: change.op, rows: lines } : { op: change.op, columns: lines }
  }
  const g = groups(axis).find((x) => x.first === span.first && x.last === span.last && x.level === change.level)
  if (!g) return { refused: `No group of level ${change.level} over the ${axis} ${text}` }
  if (g.summary === null) return { refused: `The group of ${axis} ${text} has no summary ${axis === 'rows' ? 'row' : 'column'} within the sheet: Excel draws no button for it` }
  const op = change.op === 'collapse' ? 'hideDetail' : 'showDetail'
  return axis === 'rows' ? { op, row: g.summary } : { op, column: colLetters(g.summary) }
}

/**
 * Runs an outline command on the rows or columns of its axis as xlcalc's `outlineEdit` does, changing them in place; returns
 * Excel's rejection in English, or null. A command xlcalc cannot parse (a level button beyond 8) throws its `InputError`.
 */
export function outlineCommand(lines: AxisLines, settings: OutlineSettings | null | undefined, change: OutlineChange): string | null {
  const axis = axisOf(change)
  const input = inputOf(lines, axis)
  const outline = outlineSettings(settings)
  const sheet = axis === 'rows' ? { rows: input, outline } : { columns: input, outline }
  const op = outlineOp(change, () => outlineGroups(sheet)[axis])
  if ('refused' in op) return op.refused
  const got = outlineEdit(sheet, op)
  if (!got.ok) return got.message
  lines.levels.clear()
  lines.marks.clear()
  lines.hidden.clear()
  for (const [key, line] of Object.entries(got[axis] ?? {})) {
    const span = lineSpan(key, axis)!
    for (let i = span.first; i <= span.last; i++) {
      if (line.outlineLevel) lines.levels.set(i, line.outlineLevel)
      if (line.collapsed) lines.marks.add(i)
      if (line.hidden) lines.hidden.add(i)
    }
  }
  return null
}
