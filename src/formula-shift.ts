// A formula's relative references moved by rows and columns, as copying it moves them: for a rule's formula, written for the
// top-left cell of its ranges, shown for the active cell, and written again for the top-left cell of its new ranges.
import { colLetters, colNumber, MAX_COLS, MAX_ROWS } from 'xlcalc/a1'
import type { ScanOptions } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'

/** A corner of a reference as written: `$A$1`, `B2`, `C` of `C:C`, `5` of `5:9`. */
const PART = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/

/** `n` moved by `by` within 1..`max`, wrapping past either edge as the relative references of Excel's rules do. */
const wrap = (n: number, by: number, max: number): number => ((((n - 1 + by) % max) + max) % max) + 1

/** A corner with its relative row and column moved; null for text that is not one. */
function shiftPart(part: string, rows: number, cols: number): string | null {
  const m = PART.exec(part)
  if (!m || (!m[2] && !m[4])) return null
  const [, colAbs, letters, rowAbs, digits] = m
  let col = letters ?? ''
  if (letters && !colAbs) {
    const n = colNumber(letters)
    if (n === null) return null
    col = colLetters(wrap(n, cols, MAX_COLS))
  }
  const row = digits && !rowAbs ? String(wrap(Number(digits), rows, MAX_ROWS)) : (digits ?? '')
  return `${colAbs}${col}${rowAbs}${row}`
}

/** The text of a reference moved by `rows` and `cols`: its address after the sheet's `!`, one corner or two, a spill's `#` kept. */
function shiftRef(text: string, rows: number, cols: number): string {
  const bang = text.lastIndexOf('!')
  const sheet = text.slice(0, bang + 1)
  let address = text.slice(bang + 1)
  const spill = address.endsWith('#') ? '#' : ''
  if (spill) address = address.slice(0, -1)
  const parts = address.split(':').map((p) => shiftPart(p, rows, cols))
  return parts.some((p) => p === null) ? text : `${sheet}${parts.join(':')}${spill}`
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
