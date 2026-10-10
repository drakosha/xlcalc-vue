/** A1 addresses: column letters, cells and rectangles; rows and columns are 1-based. */

export interface Cell {
  row: number
  col: number
}

/** A rectangle, bounds inclusive. */
export interface Rect {
  r1: number
  c1: number
  r2: number
  c2: number
}

const labels: string[] = ['']

/** 1 → A, 27 → AA, as Excel labels columns. */
export function colLabel(n: number): string {
  const cached = labels[n]
  if (cached !== undefined) return cached
  let s = ''
  let v = n
  while (v > 0) {
    s = String.fromCharCode(65 + ((v - 1) % 26)) + s
    v = Math.floor((v - 1) / 26)
  }
  if (n < 20_000) labels[n] = s
  return s
}

/** A → 1, AA → 27; any case. */
export function colNumber(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

export const cellAddress = (row: number, col: number): string => `${colLabel(col)}${row}`

/** `B12` or `$B$12` → the cell; null when it is not one. */
export function parseCell(text: string): Cell | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(text.trim())
  if (!m) return null
  const row = Number(m[2])
  const col = colNumber(m[1]!)
  return row > 0 && col > 0 ? { row, col } : null
}

/** `B2:D9` or `B2` → the rectangle with its corners in order; null when it is not one. */
export function parseRect(text: string): Rect | null {
  const [a, b, extra] = text.split(':')
  if (extra !== undefined || a === undefined) return null
  const p = parseCell(a)
  const q = b === undefined ? p : parseCell(b)
  if (!p || !q) return null
  return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) }
}

export const inRect = (rect: Rect, row: number, col: number): boolean =>
  row >= rect.r1 && row <= rect.r2 && col >= rect.c1 && col <= rect.c2

export const rectsMeet = (a: Rect, b: Rect): boolean => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2

/** `B12` for one cell, `B12:D15` for more. */
export const rectAddress = (r: Rect): string =>
  r.r1 === r.r2 && r.c1 === r.c2 ? cellAddress(r.r1, r.c1) : `${cellAddress(r.r1, r.c1)}:${cellAddress(r.r2, r.c2)}`

/** Column letters of a span for a layout change: `B` or `B:D`. */
export const colsAddress = (c1: number, c2: number): string => (c1 === c2 ? colLabel(c1) : `${colLabel(c1)}:${colLabel(c2)}`)
