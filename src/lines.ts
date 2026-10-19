// A sheet's lines as the component and a book name them, and an outline command's axis.
import { parseLines, parseRange } from 'xlcalc/a1'
import type { OutlineChange } from './contract/source'

/** Rows or columns of a sheet. */
export type LineAxis = 'rows' | 'columns'

/** The lines of `text` on one axis (`C`, `C:C`, `B:D`; `5`, `5:5`, `3:6`; any case); null for anything else. */
export function lineSpan(text: string, axis: LineAxis): { first: number; last: number } | null {
  const lines = parseLines(text, axis)
  if (lines) return lines
  const r = parseRange(text.includes(':') ? text : `${text}:${text}`)
  if (r?.kind !== axis) return null
  return axis === 'rows' ? { first: r.top, last: r.bottom } : { first: r.left, last: r.right }
}

/** The axis an outline command acts on. */
export const axisOf = (change: OutlineChange): LineAxis => (change.op === 'showLevel' ? change.axis : 'rows' in change ? 'rows' : 'columns')

/** Lines on an axis as a change of them names them: `{ rows: '3:6' }`, `{ columns: 'B:D' }`. */
export const linesOn = (axis: LineAxis, key: string): { rows: string } | { columns: string } => (axis === 'rows' ? { rows: key } : { columns: key })
