// Excel's F4 while a formula is edited: a reference's `$` cycle A1, $A$1, A$1, $A1, A1 (a column or a row: C, $C), as Excel for Mac
// 16.114 cycles it (measured 31.10 with Cmd+T, its F4): the cell of a typed range at the caret alone, a pointed or selected range whole.
import type { ScanOptions } from 'xlcalc/contract'
import { scanReferences } from 'xlcalc/text'
import { formatCorner, parseCorner, refPieces, type Corner } from './ref-parts'

/** The `$` of a column and a row: a column alone or a row alone has one of them. */
export interface Anchoring {
  colAbs: boolean
  rowAbs: boolean
}

/** The state after a corner's in Excel's cycle: A1 → $A$1 → A$1 → $A1 → A1; a column or a row alone goes C → $C → C. */
function nextAnchoring(c: Corner): Anchoring {
  if (!c.row) return { colAbs: !c.colAbs, rowAbs: false }
  if (!c.col) return { colAbs: false, rowAbs: !c.rowAbs }
  if (!c.colAbs && !c.rowAbs) return { colAbs: true, rowAbs: true }
  if (c.colAbs && c.rowAbs) return { colAbs: false, rowAbs: true }
  if (!c.colAbs) return { colAbs: true, rowAbs: false }
  return { colAbs: false, rowAbs: false }
}

/** A reference's text with every corner anchored as `a`; null for text whose corners are not addresses. */
export function anchorRef(text: string, a: Anchoring): string | null {
  const { sheet, pieces, spill } = refPieces(text)
  const out: string[] = []
  for (let i = 0; i < pieces.length; i++) {
    if (i % 2) {
      out.push(pieces[i]!)
      continue
    }
    const c = parseCorner(pieces[i]!)
    if (!c) return null
    out.push(formatCorner({ ...c, ...a }))
  }
  return `${sheet}${out.join('')}${spill}`
}

/** The anchoring F4 gives a whole reference: the next one after its first corner's, for all its corners, as Excel cycles a range. */
export function nextOfRef(text: string): Anchoring | null {
  const c = parseCorner(refPieces(text).pieces[0]!)
  return c && nextAnchoring(c)
}

/** A reference cycled whole: its text with the next anchoring; null for one F4 leaves. */
const cycleWhole = (text: string): string | null => {
  const a = nextOfRef(text)
  return a && anchorRef(text, a)
}

/** The draft after F4, with the span to select (the caret alone when `from` equals `to`); null when F4 changes nothing. */
export interface Cycled {
  text: string
  from: number
  to: number
}

/**
 * F4 over `text` (a formula typed in the locale `opts` parses, with its `=`) with `from`..`to` selected. A selection cycles
 * each reference it touches, a range whole; a caret in or right after a reference cycles it, a range's cell at the caret
 * alone but a range of columns or rows whole, and goes to the end of what changed. Table references stay; text that is not
 * a formula has none.
 */
export function cycleRefs(text: string, from: number, to: number, opts: ScanOptions): Cycled | null {
  if (!/^[=+-]/.test(text)) return null
  const refs = scanReferences(text, opts).filter((r) => r.kind !== 'table')
  if (from !== to) {
    let out = ''
    let at = 0
    let first = from
    for (const r of refs) {
      if (r.end <= from || r.at >= to) continue
      const next = cycleWhole(r.text)
      if (next === null) continue
      if (!at) first = Math.min(from, r.at)
      out += text.slice(at, r.at) + next
      at = r.end
    }
    if (!at) return null
    out += text.slice(at)
    // the selection widens to the whole references it touched and takes in what grew or shrank within it
    return { text: out, from: first, to: Math.max(to, at) + out.length - text.length }
  }
  const r = refs.find((x) => x.at <= from && from <= x.end)
  if (!r) return null
  const { sheet, pieces } = refPieces(r.text)
  if (r.kind !== 'cells' || pieces.length === 1) {
    const next = cycleWhole(r.text)
    if (next === null) return null
    const end = r.at + next.length
    return { text: text.slice(0, r.at) + next + text.slice(r.end), from: end, to: end }
  }
  // the corner of a typed range the caret is in or right after: the first that ends at or past the caret
  let at = r.at + sheet.length
  for (let i = 0; i < pieces.length; i += 2) {
    const piece = pieces[i]!
    const end = at + piece.length
    if (from <= end || i === pieces.length - 1) {
      const c = parseCorner(piece)
      if (!c) return null
      const next = formatCorner({ ...c, ...nextAnchoring(c) })
      const stop = at + next.length
      return { text: text.slice(0, at) + next + text.slice(end), from: stop, to: stop }
    }
    at = end + (pieces[i + 1]?.length ?? 0)
  }
  return null
}
