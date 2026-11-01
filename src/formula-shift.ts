// A formula's relative references moved by rows and columns, as copying it moves them: for a rule's formula, written for the
// top-left cell of its ranges, shown for the active cell, and written again for the top-left cell of its new ranges.
import { colLetters, colNumber, MAX_COLS, MAX_ROWS } from 'xlcalc/a1'
import type { ScanOptions } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import { formatCorner, parseCorner, refPieces } from './ref-parts'

/** `n` moved by `by` within 1..`max`, wrapping past either edge as the relative references of Excel's rules do. */
const wrap = (n: number, by: number, max: number): number => ((((n - 1 + by) % max) + max) % max) + 1

/** A corner with its relative row and column moved; null for text that is not one. */
function shiftPart(part: string, rows: number, cols: number): string | null {
  const c = parseCorner(part)
  if (!c) return null
  let col = c.col
  if (col && !c.colAbs) {
    const n = colNumber(col)
    if (n === null) return null
    col = colLetters(wrap(n, cols, MAX_COLS))
  }
  const row = c.row && !c.rowAbs ? String(wrap(Number(c.row), rows, MAX_ROWS)) : c.row
  return formatCorner({ ...c, col, row })
}

/** The text of a reference moved by `rows` and `cols`: its address after the sheet's `!`, one corner or two, a spill's `#` and a trim reference's points kept. */
function shiftRef(text: string, rows: number, cols: number): string {
  const { sheet, pieces, spill } = refPieces(text)
  const moved = pieces.map((p, i) => (i % 2 ? p : shiftPart(p, rows, cols)))
  return moved.some((p) => p === null) ? text : `${sheet}${moved.join('')}${spill}`
}

/** `formula` (in the locale `opts` parses) with its relative references moved `rows` down and `cols` right; absolute parts, names and table references stay. */
export function shiftFormula(formula: string, rows: number, cols: number, opts: ScanOptions): string {
  if (!rows && !cols) return formula
  let out = ''
  let at = 0
  for (const ref of scanReferences(formula, opts)) {
    if (ref.kind === 'table') continue
    out += formula.slice(at, ref.at) + shiftRef(ref.text, rows, cols)
    at = ref.end
  }
  return out + formula.slice(at)
}
