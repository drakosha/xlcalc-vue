// Excel's Find and Replace over cells' texts, shared by `./local` and `memorySource` and free of the engine: the query's
// wildcards, match case, the entire cell, a range, the display filter's rows, the order by rows or by columns.
// TODO(xlcalc): ask the library for Excel's Find (the match and what Look in reads) in the session or in its text subpath.
import { parseRange } from 'xlcalc/a1'
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

/** Cells in Excel's order: row by row, or column by column. */
function sortCells(list: CellText[], byColumns: boolean | undefined): CellText[] {
  return list.sort(byColumns ? (a, b) => a.col - b.col || a.row - b.row : (a, b) => a.row - b.row || a.col - b.col)
}

/** The cells whose text matches, in Excel's order, at most `max` of them; `total` counts them all. */
export function findCells(each: EachCell, req: FindRequest, visible: Set<number> | null, max: number): FindResult {
  const re = findPattern(req)
  if (!re) return { hits: [], total: 0 }
  const inside = placeOf(req, visible)
  const found: CellText[] = []
  each((row, col, text) => {
    if (text !== '' && inside(row, col) && re.test(text)) found.push({ row, col, text })
  })
  return { hits: sortCells(found, req.byColumns).slice(0, max), total: found.length }
}

/**
 * Replace All's new texts: in each cell the request finds, every match (the whole text with Match entire cell contents)
 * replaced by `by` as it is, in Excel's order; cells whose text would not change are left out.
 */
export function replaceCells(each: EachCell, req: FindRequest, by: string, visible: Set<number> | null): CellText[] {
  const re = findPattern(req)
  if (!re) return []
  const all = new RegExp(re.source, `${re.flags}g`)
  const inside = placeOf(req, visible)
  const out: CellText[] = []
  each((row, col, text) => {
    if (text === '' || !inside(row, col) || !re.test(text)) return
    // An empty match (after `*` took the rest) replaces nothing, as a cell's text is replaced once per match.
    const next = text.replace(all, (m) => (m === '' ? '' : by))
    if (next !== text) out.push({ row, col, text: next })
  })
  return sortCells(out, req.byColumns)
}
