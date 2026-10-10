// The book's layout kept beside the session: lines, merges, tables, frozen panes, styles and each sheet's extent.
// The session answers cells only; a window also needs these, so the source keeps them and patches them as it edits.
import type { BookInput, CellStyle, ColumnInput, FrozenPanes, LinePatch, NumberFormat, RowInput, SheetInput } from 'xlcalc'
import { colNumber, parseCell, parseRect, type Rect } from '../core/a1'

interface Span<T> {
  lo: number
  hi: number
  v: T
}

/** A value sets a field, null removes it, a field left out stays: `setLayout`'s patch. */
function patched<T extends object>(v: T, set: LinePatch<T>): T {
  const out = { ...v } as Record<string, unknown>
  for (const [k, x] of Object.entries(set)) {
    if (x === null) delete out[k]
    else if (x !== undefined) out[k] = x
  }
  return out as T
}

/** Columns or rows by spans, as the book keeps them (`B:D`, `5:9`): sorted, never overlapping. */
export class Lines<T extends object> {
  private spans: Span<T>[] = []

  constructor(record: Record<string, T> | undefined, parse: (s: string) => number) {
    for (const [key, v] of Object.entries(record ?? {})) {
      const [a, b] = key.split(':')
      const lo = parse(a!)
      const hi = b === undefined ? lo : parse(b)
      this.spans.push({ lo: Math.min(lo, hi), hi: Math.max(lo, hi), v: { ...v } })
    }
    this.spans.sort((x, y) => x.lo - y.lo)
  }

  get(i: number): T | undefined {
    let lo = 0
    let hi = this.spans.length - 1
    while (lo <= hi) {
      const mid = (lo + hi) >> 1
      const s = this.spans[mid]!
      if (i < s.lo) hi = mid - 1
      else if (i > s.hi) lo = mid + 1
      else return s.v
    }
    return undefined
  }

  /** Lines `lo`…`hi` patched as `setLayout` patches them: spans are split where the patch cuts them. */
  patch(lo: number, hi: number, set: LinePatch<T>): void {
    const out: Span<T>[] = []
    let next = lo
    for (const s of this.spans) {
      if (s.hi < lo || s.lo > hi) {
        out.push(s)
        continue
      }
      if (s.lo < lo) out.push({ lo: s.lo, hi: lo - 1, v: s.v })
      if (s.lo > next) out.push({ lo: next, hi: s.lo - 1, v: patched({} as T, set) })
      out.push({ lo: Math.max(s.lo, lo), hi: Math.min(s.hi, hi), v: patched(s.v, set) })
      if (s.hi > hi) out.push({ lo: hi + 1, hi: s.hi, v: s.v })
      next = Math.min(s.hi, hi) + 1
    }
    if (next <= hi) out.push({ lo: next, hi, v: patched({} as T, set) })
    this.spans = out.filter((s) => Object.keys(s.v).length > 0).sort((x, y) => x.lo - y.lo)
  }
}

export interface TableBox {
  name: string
  range: string
  rect: Rect
  /** The header row's number, 0 without one. */
  headerRow: number
}

export interface SheetLayout {
  name: string
  columns: Lines<ColumnInput>
  rows: Lines<RowInput>
  merges: { text: string; rect: Rect }[]
  tables: TableBox[]
  frozen?: FrozenPanes
  /** The extent: the last row and column with a cell or a merged area. */
  maxRow: number
  maxCol: number
}

export interface BookLayout {
  sheets: SheetLayout[]
  /** The book's style table, as the session has it. */
  styles: CellStyle[]
  locale: string
  dateSystem?: 1900 | 1904
}

function sheetLayout(sheet: SheetInput): SheetLayout {
  const merges = (sheet.merges ?? []).flatMap((text) => {
    const rect = parseRect(text)
    return rect ? [{ text, rect }] : []
  })
  const tables = (sheet.tables ?? []).flatMap((t) => {
    const rect = parseRect(t.range)
    return rect ? [{ name: t.name, range: t.range, rect, headerRow: t.headers === false ? 0 : rect.r1 }] : []
  })
  const out: SheetLayout = {
    name: sheet.name,
    columns: new Lines(sheet.columns, colNumber),
    rows: new Lines(sheet.rows, Number),
    merges,
    tables,
    ...(sheet.frozen ? { frozen: { ...sheet.frozen } } : {}),
    maxRow: 0,
    maxCol: 0,
  }
  for (const addr of Object.keys(sheet.cells)) {
    const at = parseCell(addr)
    if (at) grow(out, at.row, at.col)
  }
  for (const m of merges) grow(out, m.rect.r2, m.rect.c2)
  return out
}

/** The layout of a book; the book's objects are not kept. */
export function readLayout(book: BookInput): BookLayout {
  return {
    sheets: book.sheets.map(sheetLayout),
    styles: [...(book.styles ?? [])],
    locale: book.locale ?? 'en-US',
    ...(book.dateSystem ? { dateSystem: book.dateSystem } : {}),
  }
}

/** A sheet by name: as written, else without regard to case, as the library matches sheets. */
export function sheetOf(book: BookLayout, name: string): SheetLayout | undefined {
  const key = name.toUpperCase()
  return book.sheets.find((s) => s.name === name) ?? book.sheets.find((s) => s.name.toUpperCase() === key)
}

export function grow(l: SheetLayout, row: number, col: number): void {
  if (row > l.maxRow) l.maxRow = row
  if (col > l.maxCol) l.maxCol = col
}

/** The style and format of a cell with none of its own: its row's when the row has either, else its column's, as Excel resolves them. */
export function lineLook(l: SheetLayout, row: number, col: number): { style?: number; format?: NumberFormat } {
  const r = l.rows.get(row)
  if (r && (r.style !== undefined || r.format !== undefined)) return { style: r.style, format: r.format }
  const c = l.columns.get(col)
  return c ? { style: c.style, format: c.format } : {}
}

/** A column's width for a window: null by default, 0 when hidden. */
export function widthOf(l: SheetLayout, col: number): number | null {
  const c = l.columns.get(col)
  if (!c) return null
  return c.hidden ? 0 : (c.width ?? null)
}

/** A row's height in points, 0 when hidden, undefined by default. */
export function heightOf(l: SheetLayout, row: number): number | undefined {
  const r = l.rows.get(row)
  if (!r) return undefined
  return r.hidden ? 0 : r.height
}

/** The table a cell is in, for showing its formula as Excel does inside a table. */
export function tableAt(l: SheetLayout, row: number, col: number): string | undefined {
  for (const t of l.tables) if (row >= t.rect.r1 && row <= t.rect.r2 && col >= t.rect.c1 && col <= t.rect.c2) return t.name
  return undefined
}

/** A style's key for finding an equal one in the table: its fields in a fixed order. */
function styleKey(style: unknown): string {
  if (style === null || typeof style !== 'object') return JSON.stringify(style)
  const o = style as Record<string, unknown>
  return `{${Object.keys(o)
    .filter((k) => o[k] !== undefined)
    .sort()
    .map((k) => `${JSON.stringify(k)}:${styleKey(o[k])}`)
    .join(',')}}`
}

/** The index of a style equal to `style` in the table, and the table with it appended when it had none (else null). */
export function styleIndex(styles: CellStyle[], style: CellStyle): { index: number; table: CellStyle[] | null } {
  const key = styleKey(style)
  const found = styles.findIndex((s) => styleKey(s) === key)
  return found >= 0 ? { index: found, table: null } : { index: styles.length, table: [...styles, style] }
}
