// Excel's Find and Replace All over the session: xlcalc's `findCells` with the options of Excel's dialog gives the cells
// Find gives, in its order, from the cell Find Next goes on from, hidden ones skipped as Find skips them, the display
// filter's rows alone; `replaceInCells` replaces as Excel's Replace All. The session keeps the texts it searches.
import type { AsyncSession, FindCellsOptions, ReplaceResult } from 'xlcalc'
import { cellAddress, parseCell, parseRange } from 'xlcalc/a1'
import { cellRect, unionRect } from '../rect'
import type { FindRequest, FindResult } from '../contract/source'
import type { CellsBox } from './outcome'

/** Hits of a search in one answer; the total counts them all. */
const FIND_MAX_HITS = 500
const WHOLE_SHEET = 'A1:XFD1048576'

const rangesOf = (req: FindRequest) => [{ sheet: req.sheet, range: req.range ?? WHOLE_SHEET }]

/** The dialog's options both Find and Replace take, under the display filter its rows alone. */
function matchOf(req: FindRequest, rows: number[] | null) {
  return {
    wildcards: true,
    skipHidden: true,
    ...(req.matchCase ? { caseSensitive: true } : {}),
    ...(req.wholeCell ? { wholeCell: true } : {}),
    ...(rows ? { rows } : {}),
  }
}

/** The cell Find Next goes on from, when it is one the search reads: the library refuses a cell outside the range, as Excel does. */
function afterOf(req: FindRequest, rows: number[] | null): FindCellsOptions['after'] {
  const at = req.after ? parseCell(req.after) : null
  if (!at) return undefined
  const r = req.range ? parseRange(req.range) : null
  if (r && (at.row < r.top || at.row > r.bottom || at.col < r.left || at.col > r.right)) return undefined
  if (rows && !rows.includes(at.row)) return undefined
  return { sheet: req.sheet, cell: cellAddress(at.row, at.col) }
}

/** Excel's Find: at most `FIND_MAX_HITS` hits in its order, from the first after `after` round past the end, and the count of all; `rows`, the display filter's in ascending order. */
export async function findHits(s: AsyncSession, req: FindRequest, rows: number[] | null, signal?: AbortSignal): Promise<FindResult> {
  const values = req.lookIn === 'values'
  const after = afterOf(req, rows)
  const page = await s.findCells(rangesOf(req), req.query, {
    ...matchOf(req, rows),
    in: values ? 'text' : 'edit',
    ...(req.byColumns ? { byColumns: true } : {}),
    ...(after ? { after } : {}),
    fields: [],
    maxCells: FIND_MAX_HITS,
    count: true,
    ...(signal ? { signal } : {}),
  })
  return {
    hits: page.cells.map((c) => {
      const at = parseCell(c.cell)!
      return { row: at.row, col: at.col }
    }),
    total: page.total ?? page.cells.length,
  }
}

/** Excel's Replace All: the cells Find finds in the text they open with, whatever Look in says, each with the matches replaced, in one step of undo. */
export function replaceAll(s: AsyncSession, req: FindRequest, by: string, rows: number[] | null): Promise<ReplaceResult> {
  return s.replaceInCells(rangesOf(req), req.query, by, matchOf(req, rows))
}

/**
 * Cells scattered over a sheet (Replace All's) in rectangles of at most `max` cells, so that reading them fits windows;
 * the cells come in Find's order, which keeps neighbours together.
 */
export function boxesUnder(sheet: string, cells: { row: number; col: number }[], max: number): Required<CellsBox>[] {
  const out: Required<CellsBox>[] = []
  let box: Required<CellsBox> | null = null
  for (const c of cells) {
    const addr = cellAddress(c.row, c.col)
    const r = box?.rect
    const rect = r ? unionRect(r, cellRect(c.row, c.col)) : null
    if (box && rect && (rect.bottom - rect.top + 1) * (rect.right - rect.left + 1) <= max) {
      box.rect = rect
      box.only.add(addr)
    } else out.push((box = { sheet, rect: cellRect(c.row, c.col), only: new Set([addr]) }))
  }
  return out
}
