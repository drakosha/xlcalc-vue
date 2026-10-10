// A window of a sheet from the session's cells and the kept layout: shown text, values, formulas in the book's
// locale, styles resolved as Excel resolves them, notes, links, sizes, merges and tables.
import { toLocal, type CellStyle, type CellView, type Locale } from 'xlcalc'
import { cellAddress, rectsMeet } from '../core/a1'
import type { SheetWindow } from '../source'
import { heightOf, lineLook, tableAt, widthOf, type SheetLayout } from './layout'

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

export interface WindowInput {
  layout: SheetLayout
  styles: CellStyle[]
  /** The session's cells of the window's rows and columns, by address. */
  views: Record<string, CellView>
  /** The window's first row as asked: a sheet row, or a display index under a filter. */
  row: number
  /** The sheet rows of the window in display order. */
  rows: number[]
  col: number
  cols: number
  /** The rows come from a display filter: merged areas over several rows are left out, as Excel shows them then. */
  filtered: boolean
  formulas: Formulas
  notice?: string
}

export function buildWindow(a: WindowInput): SheetWindow {
  const { layout: l, views } = a
  const lastCol = a.col + a.cols - 1
  const win: SheetWindow = {
    row: a.row,
    col: a.col,
    text: [],
    values: [],
    formulas: {},
    colors: {},
    styleIds: {},
    styles: a.styles,
    widths: [],
    rowHeights: {},
    merges: [],
    tables: l.tables.map((t) => ({ name: t.name, range: t.range, headerRow: t.headerRow })),
    links: {},
    notes: {},
    ...(a.notice ? { notice: a.notice } : {}),
  }
  for (let c = a.col; c <= lastCol; c++) win.widths.push(widthOf(l, c))
  for (const r of a.rows) {
    const h = heightOf(l, r)
    if (h !== undefined) win.rowHeights[r] = h
    const text: string[] = []
    const values: SheetWindow['values'][number] = []
    for (let c = a.col; c <= lastCol; c++) {
      const addr = cellAddress(r, c)
      const v = views[addr]
      text.push(v?.text ?? '')
      values.push(v ? v.value : null)
      const style = v?.style ?? lineLook(l, r, c).style
      if (style !== undefined) win.styleIds[addr] = style
      if (!v) continue
      const formula = a.formulas.of(l, r, c, v)
      if (formula !== null) win.formulas[addr] = formula
      if (v.color) win.colors[addr] = v.color
      if (v.note) win.notes![addr] = v.note
      if (v.link) {
        const { url, location, tooltip } = v.link
        win.links![addr] = { ...(url ? { href: url } : {}), ...(location ? { location } : {}), ...(tooltip ? { tooltip } : {}) }
      }
    }
    win.text.push(text)
    win.values.push(values)
  }
  if (a.rows.length && a.cols > 0) {
    if (a.filtered) {
      const shown = new Set(a.rows)
      win.merges = l.merges.filter((m) => m.rect.r1 === m.rect.r2 && shown.has(m.rect.r1) && m.rect.c1 <= lastCol && m.rect.c2 >= a.col).map((m) => m.text)
    } else {
      const area = { r1: a.rows[0]!, c1: a.col, r2: a.rows[a.rows.length - 1]!, c2: lastCol }
      win.merges = l.merges.filter((m) => rectsMeet(m.rect, area)).map((m) => m.text)
    }
  }
  return win
}
