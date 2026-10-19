// Excel's Find and Replace over cells' texts for `memorySource`: a text matches and is replaced by the library's
// `findMatcher` and `findReplacer` (as its `findCells` and `replaceInCells`); this side keeps the range, the display
// filter's rows and the order by rows or by columns from a cell. `./local` finds and replaces by the session.
import { parseCell, parseRange } from 'xlcalc/a1'
import { findMatcher, findReplacer } from 'xlcalc/text'
import type { FindRequest, FindResult } from './contract/source'

/** Visits every cell a search reads with its text (what it opens with, or what it shows), in any order. */
export type EachCell = (visit: (row: number, col: number, text: string) => void) => void

/** A cell and its text: a hit, or a cell Replace All writes. */
export interface CellText {
  row: number
  col: number
  text: string
}

/** The request's options as the library's Find takes them; Excel's Find always reads wildcards. */
const optionsOf = (req: Pick<FindRequest, 'matchCase' | 'wholeCell'>) => ({ wildcards: true, caseSensitive: !!req.matchCase, wholeCell: !!req.wholeCell })

/** Whether a cell is where the request looks: inside its range and on a row the display filter shows. */
function placeOf(req: FindRequest, visible: Set<number> | null): (row: number, col: number) => boolean {
  const r = req.range ? parseRange(req.range) : null
  if (req.range && !r) return () => false
  return (row, col) => (!r || (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right)) && (!visible || visible.has(row))
}

/** Excel's order of cells for Find: row by row, or column by column. */
export const cellOrder =
  (byColumns: boolean | undefined) =>
  (a: { row: number; col: number }, b: { row: number; col: number }): number =>
    byColumns ? a.col - b.col || a.row - b.row : a.row - b.row || a.col - b.col

/** The cells whose text matches, in Excel's order from the first after `req.after` round past the end, at most `max` of them; `total` counts them all. */
export function findCells(each: EachCell, req: FindRequest, visible: Set<number> | null, max: number): FindResult {
  const test = findMatcher(req.query, optionsOf(req))
  if (!test) return { hits: [], total: 0 }
  const inside = placeOf(req, visible)
  const found: CellText[] = []
  each((row, col, text) => {
    if (text !== '' && inside(row, col) && test(text)) found.push({ row, col, text })
  })
  const order = cellOrder(req.byColumns)
  const sorted = found.sort(order)
  const after = req.after ? parseCell(req.after) : null
  const from = after ? sorted.findIndex((c) => order(c, after) > 0) : 0
  const round = from > 0 ? sorted.slice(from).concat(sorted.slice(0, from)) : sorted
  return { hits: round.slice(0, max).map(({ row, col }) => ({ row, col })), total: found.length }
}

/**
 * Replace All's new texts: each cell the request finds with its new text, in Excel's order. Cells whose text would not
 * change are left out: written again, a memory cell would lose the shown text it was given.
 */
export function replaceCells(each: EachCell, req: FindRequest, by: string, visible: Set<number> | null): CellText[] {
  const replace = findReplacer(req.query, by, optionsOf(req))
  if (!replace) return []
  const inside = placeOf(req, visible)
  const out: CellText[] = []
  each((row, col, text) => {
    const next = text === '' || !inside(row, col) ? null : replace(text)
    if (next !== null && next !== text) out.push({ row, col, text: next })
  })
  return out.sort(cellOrder(req.byColumns))
}
