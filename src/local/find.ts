// Excel's Find and Replace All over the session: xlcalc's `findCells` with the options of Excel's dialog gives the cells
// Find gives, in its order, hidden ones skipped as Find skips them; the display filter keeps its rows of them, and
// Replace All writes each found cell's text with the query's matches replaced.
import type { AsyncSession, FindCellsOptions, ReadCell } from 'xlcalc'
import { cellAddress, parseCell, rangeAddress } from 'xlcalc/a1'
import { replacer, type CellText } from '../find-match'
import type { FindRequest, FindResult } from '../source'
import type { CellsBox } from './outcome'

/** Hits of a search in one answer; the total counts them all. */
export const FIND_MAX_HITS = 500
/** Cells of one page of `findCells`, the most it gives. */
const PAGE = 10_000
/** Times a search over pages starts again when an edit between two pages leaves its cursor stale. */
const RESTARTS = 3
const WHOLE_SHEET = 'A1:XFD1048576'

/** The dialog's options as `findCells` takes them: Look in Formulas reads the text a cell opens with, Values the text it shows. */
function optionsOf(req: FindRequest, lookIn: FindRequest['lookIn']): FindCellsOptions {
  return {
    wildcards: true,
    in: lookIn === 'values' ? 'text' : 'edit',
    skipHidden: true,
    ...(req.matchCase ? { caseSensitive: true } : {}),
    ...(req.wholeCell ? { wholeCell: true } : {}),
    ...(req.byColumns ? { byColumns: true } : {}),
  }
}

const rangesOf = (req: FindRequest) => [{ sheet: req.sheet, range: req.range ?? WHOLE_SHEET }]

/** What matched, as the contract has it: the formula (as xlcalc keeps it, in en-US) in formulas, else the shown text. */
function hitOf(cell: ReadCell): FindResult['hits'][number] {
  const at = parseCell(cell.cell)!
  return { row: at.row, col: at.col, text: cell.formula !== undefined ? `=${cell.formula}` : (cell.text ?? '') }
}

/**
 * Every cell the search finds in `visible` rows, page by page in Find's order; an edit between two pages leaves the
 * cursor stale, and the search starts again after `restart`.
 */
async function eachFound(s: AsyncSession, req: FindRequest, opts: FindCellsOptions, visible: Set<number> | null, signal: AbortSignal | undefined, visit: (cell: ReadCell, row: number) => void, restart: () => void): Promise<void> {
  for (let n = 0; ; n++) {
    try {
      let cursor: string | undefined
      do {
        const page = await s.findCells(rangesOf(req), req.query, { ...opts, maxCells: PAGE, ...(cursor ? { cursor } : {}), ...(signal ? { signal } : {}) })
        for (const cell of page.cells) {
          const row = parseCell(cell.cell)!.row
          if (!visible || visible.has(row)) visit(cell, row)
        }
        cursor = page.cursor
      } while (cursor)
      return
    } catch (e) {
      if (n >= RESTARTS || (e as { code?: string }).code !== 'stale_revision') throw e
      restart()
    }
  }
}

/** Excel's Find: at most `FIND_MAX_HITS` hits in its order and the count of all; with `visible` only those rows', read over every page. */
export async function findHits(s: AsyncSession, req: FindRequest, visible: Set<number> | null, signal?: AbortSignal): Promise<FindResult> {
  const opts: FindCellsOptions = { ...optionsOf(req, req.lookIn), fields: req.lookIn === 'values' ? ['text'] : ['formula', 'text'] }
  if (!visible) {
    const page = await s.findCells(rangesOf(req), req.query, { ...opts, maxCells: FIND_MAX_HITS, count: true, ...(signal ? { signal } : {}) })
    return { hits: page.cells.map(hitOf), total: page.total ?? page.cells.length }
  }
  const hits: FindResult['hits'] = []
  let total = 0
  const take = (cell: ReadCell) => {
    if (total++ < FIND_MAX_HITS) hits.push(hitOf(cell))
  }
  await eachFound(s, req, opts, visible, signal, take, () => {
    hits.length = 0
    total = 0
  })
  return { hits, total }
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
    const rect = r ? { top: Math.min(r.top, c.row), left: Math.min(r.left, c.col), bottom: Math.max(r.bottom, c.row), right: Math.max(r.right, c.col) } : null
    if (box && rect && (rect.bottom - rect.top + 1) * (rect.right - rect.left + 1) <= max) {
      box.rect = rect
      box.only.add(addr)
    } else out.push((box = { sheet, rect: { top: c.row, left: c.col, bottom: c.row, right: c.col }, only: new Set([addr]) }))
  }
  return out
}

/**
 * Replace All's new texts: the cells Find finds in formulas, as Excel's Replace looks whatever Look in says (in `visible`
 * rows), each written as the text it opens with, the matches replaced by `by`. Only the found cells' texts are read.
 * TODO(xlcalc): the library has no Replace; the matches are our own pattern's (`find-match.ts`, where `?` is a whole
 * character, not Find's UTF-16 unit), and a cell Find found where it matches nothing is not written.
 */
export async function replacements(s: AsyncSession, req: FindRequest, by: string, visible: Set<number> | null, maxCells: number): Promise<CellText[]> {
  const replace = replacer(req, by)
  if (!replace) return []
  const cells: { row: number; col: number }[] = []
  await eachFound(s, req, { ...optionsOf(req, 'formulas'), fields: [] }, visible, undefined, (cell, row) => cells.push({ row, col: parseCell(cell.cell)!.col }), () => (cells.length = 0))
  const boxes = boxesUnder(req.sheet, cells, maxCells)
  const views = await Promise.all(boxes.map((b) => s.getViewport(req.sheet, rangeAddress(b.rect), { editText: true, maxCells })))
  const out: CellText[] = []
  for (const [i, box] of boxes.entries()) {
    for (const addr of box.only) {
      const cell = views[i]!.cells[addr]
      const typed = cell?.editText ?? ''
      // Find reads a text as it is, without the apostrophe the formula bar shows before one that reads as something else; the new text keeps it
      const text = cell && cell.formula === undefined && typeof cell.value === 'string' ? cell.value : typed
      const next = replace(text)
      if (next === null) continue
      const at = parseCell(addr)!
      out.push({ row: at.row, col: at.col, text: (typed === `'${text}` ? "'" : '') + next })
    }
  }
  return out
}
