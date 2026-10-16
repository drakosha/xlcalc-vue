// Excel's Find and Replace over cells' texts for `memorySource`, free of the engine: the query's wildcards, match case,
// the entire cell, a range, the display filter's rows, the order by rows or by columns from a cell. `./local` finds and
// replaces by the library.
// TODO(xlcalc): Excel's Find and Replace without the engine (`xlcalc/text`) would replace this matching, close to Excel's.
import { parseCell, parseRange } from 'xlcalc/a1'
import type { FindRequest, FindResult } from './source'

/** Visits every cell a search reads with its text (what it opens with, or what it shows), in any order. */
export type EachCell = (visit: (row: number, col: number, text: string) => void) => void

/** A cell and its text: a hit, or a cell Replace All writes. */
export interface CellText {
  row: number
  col: number
  text: string
}

/** A character of the query taken as it is. */
const literal = (ch: string) => ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/**
 * The query as a regular expression, as the library reads a lookup's wildcards (measured in Excel for lookups, not yet
 * for Find): `*` any text, `?` one character, `~` takes the next character as it is, a trailing `~` is dropped.
 */
function patternSource(query: string): string {
  let re = ''
  for (let i = 0; i < query.length; i++) {
    const ch = query[i]!
    if (ch === '~') re += i + 1 < query.length ? literal(query[++i]!) : ''
    else if (ch === '*') re += '[\\s\\S]*'
    else if (ch === '?') re += '[\\s\\S]'
    else re += literal(ch)
  }
  return re
}

/** A request's pattern: the whole text with Match entire cell contents, in any case without Match case; null for an empty query. */
export function findPattern(req: Pick<FindRequest, 'query' | 'matchCase' | 'wholeCell'>): RegExp | null {
  if (!req.query) return null
  const source = patternSource(req.query)
  return new RegExp(req.wholeCell ? `^(?:${source})$` : source, req.matchCase ? 'u' : 'iu')
}

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
  const re = findPattern(req)
  if (!re) return { hits: [], total: 0 }
  const inside = placeOf(req, visible)
  const found: CellText[] = []
  each((row, col, text) => {
    if (text !== '' && inside(row, col) && re.test(text)) found.push({ row, col, text })
  })
  const order = cellOrder(req.byColumns)
  const sorted = found.sort(order)
  const after = req.after ? parseCell(req.after) : null
  const from = after ? sorted.findIndex((c) => order(c, after) > 0) : 0
  const round = from > 0 ? sorted.slice(from).concat(sorted.slice(0, from)) : sorted
  return { hits: round.slice(0, max), total: found.length }
}

/**
 * A cell's new text under Replace All: every match (the whole text with Match entire cell contents) replaced by `by` as
 * it is; null where nothing matches or nothing would change. Null for an empty query.
 */
function replacer(req: Pick<FindRequest, 'query' | 'matchCase' | 'wholeCell'>, by: string): ((text: string) => string | null) | null {
  const re = findPattern(req)
  if (!re) return null
  const all = new RegExp(re.source, `${re.flags}g`)
  return (text) => {
    if (!re.test(text)) return null
    // An empty match (after `*` took the rest) replaces nothing, as a cell's text is replaced once per match.
    const next = text.replace(all, (m) => (m === '' ? '' : by))
    return next === text ? null : next
  }
}

/** Replace All's new texts: each cell the request finds with its new text, in Excel's order; cells whose text would not change are left out. */
export function replaceCells(each: EachCell, req: FindRequest, by: string, visible: Set<number> | null): CellText[] {
  const replace = replacer(req, by)
  if (!replace) return []
  const inside = placeOf(req, visible)
  const out: CellText[] = []
  each((row, col, text) => {
    const next = text === '' || !inside(row, col) ? null : replace(text)
    if (next !== null) out.push({ row, col, text: next })
  })
  return out.sort(cellOrder(req.byColumns))
}
