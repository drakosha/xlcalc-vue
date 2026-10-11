// A window of a sheet from xlcalc's viewport: shown text, values, formulas in the book's locale, the text each cell
// opens with, looks, sizes, merged areas, tables, notes and links. Under a display filter only merged areas of one
// shown row stay, as Excel shows them; that rule is the one part of the window xlcalc does not give.
import { toLocal, type CellStyle, type CellView, type Locale, type Viewport } from 'xlcalc'
import { cellAddress, parseRange } from 'xlcalc/a1'
import type { SheetWindow } from '../source'
import { tableAt, type SheetLayout } from './layout'

const TRANSLATIONS_MAX = 100_000

/** Stored formulas as the locale shows them, remembered: windows ask for the same formulas again on every scroll. */
export class Formulas {
  private cache = new Map<string, string>()

  constructor(private readonly locale: Locale) {}

  /** A stored formula in the locale; inside its own table without the table's name (`[@Сумма]`), as Excel shows it. */
  show(formula: string, table?: string): string {
    const key = table ? `${table}\u0000${formula}` : formula
    let out = this.cache.get(key)
    if (out === undefined) {
      const t = toLocal(formula, this.locale, table ? { table } : undefined)
      out = t.ok ? t.formula : formula
      if (this.cache.size >= TRANSLATIONS_MAX) this.cache.clear()
      this.cache.set(key, out)
    }
    return out
  }

  /** A cell's formula as shown, or null without one. */
  of(l: SheetLayout, row: number, col: number, view: CellView | undefined): string | null {
    return view?.formula === undefined ? null : this.show(view.formula, tableAt(l, row, col))
  }
}

/** The height xlcalc gives a row without one (Excel's default): such rows are left out of `rowHeights`. */
const DEFAULT_ROW_PT = 15

export interface WindowShape {
  /** The window's first row as asked: a sheet row, or a display index under a filter. */
  row: number
  /** The sheet rows of the window in display order. */
  rows: number[]
  col: number
  cols: number
  /** The rows come from a display filter: merged areas over several rows are left out, as Excel shows them then. */
  filtered: boolean
  notice?: string
}

/** The component's window from a viewport of the shape's rows and columns. */
export function windowOf(v: Viewport, a: WindowShape): SheetWindow {
  const styles: CellStyle[] = []
  for (const [i, style] of Object.entries(v.styles)) styles[Number(i)] = style
  const win: SheetWindow = {
    row: a.row,
    col: a.col,
    text: [],
    values: [],
    formulas: {},
    editText: {},
    colors: {},
    styleIds: {},
    styles,
    widths: v.columns.slice(0, a.cols).map((c) => (c.hidden ? 0 : c.width)),
    rowHeights: {},
    merges: [],
    tables: v.tables.map((t) => ({ name: t.name, range: t.range, headerRow: t.headers ? (parseRange(t.range)?.top ?? 0) : 0 })),
    links: {},
    notes: {},
    ...(a.notice ? { notice: a.notice } : {}),
  }
  for (const r of v.rows) if (r.hidden || r.height !== DEFAULT_ROW_PT) win.rowHeights[r.row] = r.hidden ? 0 : r.height
  const lastCol = a.col + a.cols - 1
  for (const r of a.rows) {
    const text: string[] = []
    const values: SheetWindow['values'][number] = []
    for (let c = a.col; c <= lastCol; c++) {
      const addr = cellAddress(r, c)
      const cell = v.cells[addr]
      text.push(cell?.text ?? '')
      values.push(cell ? cell.value : null)
      if (!cell) continue
      if (cell.effectiveStyle !== undefined) win.styleIds[addr] = cell.effectiveStyle
      if (cell.formulaLocal !== undefined) win.formulas[addr] = cell.formulaLocal
      if (cell.editText) win.editText![addr] = cell.editText
      if (cell.color) win.colors[addr] = cell.color
      if (cell.note) win.notes![addr] = cell.note
      if (cell.link) {
        // with an address the place is in what it points to (Excel's SubAddress), not in this book
        const { url, location, tooltip } = cell.link
        const to = url ? { href: location ? `${url}#${location}` : url } : location ? { location } : {}
        win.links![addr] = { ...to, ...(tooltip ? { tooltip } : {}) }
      }
    }
    win.text.push(text)
    win.values.push(values)
  }
  if (!a.filtered) win.merges = v.merges
  else {
    const shown = new Set(a.rows)
    win.merges = v.merges.filter((m) => {
      const rect = parseRange(m)
      return !!rect && rect.top === rect.bottom && shown.has(rect.top)
    })
  }
  return win
}
