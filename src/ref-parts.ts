// The pieces of a reference's text as `scanReferences` finds it: its sheet, its corners and the colons between them, a spill's `#`.

/** A corner of a reference as written: `$A$1`, `B2`, `C` of `C:C`, `5` of `5:9`. */
const PART = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/

/** A range's colon with a trim reference's points around it (`:.`, `.:`, `.:.`), kept as written. */
const COLON = /(\.?:\.?)/

/** A corner's column and row with their `$`; null for text that is not one. */
export interface Corner {
  colAbs: boolean
  col: string
  rowAbs: boolean
  row: string
}

export function parseCorner(part: string): Corner | null {
  const m = PART.exec(part)
  if (!m || (!m[2] && !m[4])) return null
  // a row alone (`$5`) takes the first `$` as its own
  return { colAbs: !!m[1] && !!m[2], col: m[2] ?? '', rowAbs: !!m[4] && (!!m[3] || (!m[2] && !!m[1])), row: m[4] ?? '' }
}

export const formatCorner = (c: Corner): string => `${c.col && c.colAbs ? '$' : ''}${c.col}${c.row && c.rowAbs ? '$' : ''}${c.row}`

/** A reference's text in pieces: `sheet` up to its `!`, then corners at even places and colons at odd ones, then a spill's `#`. */
export function refPieces(text: string): { sheet: string; pieces: string[]; spill: string } {
  const bang = text.lastIndexOf('!')
  let address = text.slice(bang + 1)
  const spill = address.endsWith('#') ? '#' : ''
  if (spill) address = address.slice(0, -1)
  return { sheet: text.slice(0, bang + 1), pieces: address.split(COLON), spill }
}
