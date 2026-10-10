// References in a formula (`B1:B4`, `$A$1`, `'Sheet 2'!A1`, `B:B`, `1:1`) for clickable links and range highlights.
// A lexical scan, so any formula locale works (addresses are never translated); names and structured refs stay text.

/** A range of a formula: bounds 1-based inclusive; a whole column or row runs to Excel's limit. */
export interface FormulaRef {
  /** The sheet named in the reference; null is the formula's own sheet. */
  sheet: string | null
  r1: number
  c1: number
  r2: number
  c2: number
}

/** A piece of a formula as shown; a reference carries its range. */
export interface FormulaToken {
  text: string
  ref?: FormulaRef
}

export const EXCEL_MAX_ROWS = 1_048_576
export const EXCEL_MAX_COLS = 16_384

const SHEET = String.raw`(?:'((?:[^']|'')+)'|([^\s!'"()+\-*/^&=<>,;:%\[\]]+))!`
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d{1,7}`
const REF_RE = new RegExp(
  `(?:${SHEET})?(?:(${CELL})(?::(${CELL}))?|\\$?([A-Za-z]{1,3}):\\$?([A-Za-z]{1,3})|\\$?(\\d{1,7}):\\$?(\\d{1,7}))`,
  'y',
)
const IDENT = /[\p{L}\p{N}_.]/u

function colOf(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}
function cellOf(ref: string): { r: number; c: number } {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref)!
  return { r: Number(m[2]), c: colOf(m[1]!) }
}

/** Split a formula (without the leading `=`) into text and references. */
export function formulaTokens(formula: string): FormulaToken[] {
  const out: FormulaToken[] = []
  let plain = ''
  let i = 0
  while (i < formula.length) {
    const ch = formula[i]!
    if (ch === '"') {
      // A string literal runs to the closing quote; a doubled quote belongs to the string.
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2
          else break
        } else j++
      }
      plain += formula.slice(i, j + 1)
      i = j + 1
      continue
    }
    const prev = i > 0 ? formula[i - 1]! : ''
    if (prev && IDENT.test(prev)) {
      plain += ch
      i++
      continue
    }
    REF_RE.lastIndex = i
    const m = REF_RE.exec(formula)
    const next = m ? (formula[i + m[0].length] ?? '') : ''
    const ref = m && next !== '(' && !(next && IDENT.test(next)) ? refOf(m) : null
    if (!m || !ref) {
      plain += ch
      i++
      continue
    }
    if (plain) out.push({ text: plain })
    plain = ''
    out.push({ text: m[0], ref })
    i += m[0].length
  }
  if (plain) out.push({ text: plain })
  return out
}

function refOf(m: RegExpExecArray): FormulaRef | null {
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : (m[2] ?? null)
  if (m[3]) {
    const a = cellOf(m[3])
    const b = m[4] ? cellOf(m[4]) : a
    if (!a.r || !a.c || !b.r || !b.c) return null
    if (a.r > EXCEL_MAX_ROWS || b.r > EXCEL_MAX_ROWS || a.c > EXCEL_MAX_COLS || b.c > EXCEL_MAX_COLS) return null
    return { sheet, r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) }
  }
  if (m[5]) {
    const a = colOf(m[5])
    const b = colOf(m[6]!)
    if (!a || !b || a > EXCEL_MAX_COLS || b > EXCEL_MAX_COLS) return null
    return { sheet, r1: 1, c1: Math.min(a, b), r2: EXCEL_MAX_ROWS, c2: Math.max(a, b) }
  }
  const a = Number(m[7])
  const b = Number(m[8])
  if (!a || !b || a > EXCEL_MAX_ROWS || b > EXCEL_MAX_ROWS) return null
  return { sheet, r1: Math.min(a, b), c1: 1, r2: Math.max(a, b), c2: EXCEL_MAX_COLS }
}

/** A range's key: equal references in a formula get one colour. */
export const refKey = (r: FormulaRef): string => `${r.sheet ?? ''}!${r.r1}:${r.c1}:${r.r2}:${r.c2}`

/** A place in the book from a link or the address box: `B12`, `B12:D15`, `Sheet!B12`, `'Sheet 2'!C7`, `B:B`, or a sheet alone; null when it is not one. */
export function parseCellRef(text: string): { sheet: string | null; range: FormulaRef | null } | null {
  const raw = text.trim()
  if (!raw) return null
  const tokens = formulaTokens(raw)
  if (tokens.length === 1 && tokens[0]!.ref) return { sheet: tokens[0]!.ref.sheet, range: tokens[0]!.ref }
  // Addresses inside other text (`B12+C1`) make a formula, not a place.
  if (tokens.some((t) => t.ref)) return null
  const m = /^'((?:[^']|'')+)'$|^([^!'"]+)$/.exec(raw)
  if (!m) return null
  return { sheet: m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]!.trim(), range: null }
}

/** A sheet's name for an address: quoted when it has anything but letters, digits and `_`. */
export function sheetRefName(name: string): string {
  return /^[\p{L}\p{N}_]+$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}
