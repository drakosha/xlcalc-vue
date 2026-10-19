// A sheet window from xlcalc's viewport: displayed text, values, formulas in the workbook's locale, the text each cell
// opens with, styles and conditional formatting, sizes, merged ranges, tables, notes and hyperlinks. Under a display
// filter only merged ranges on a single visible row stay, as Excel displays them; that rule is the only part of the window
// xlcalc does not provide.
import type { CellStyle, Viewport } from 'xlcalc'
import { cellAddress, parseRange } from 'xlcalc/a1'
import { DEFAULT_ROW_HEIGHT } from 'xlcalc/text'
import type { SheetWindow } from '../contract/source'

interface WindowShape {
  /** The window's first row as requested: a sheet row, or a display index under a filter. */
  row: number
  /** The window's sheet rows in display order. */
  rows: number[]
  col: number
  cols: number
  /** The rows come from a display filter: merged ranges over several rows are omitted, as Excel displays them then. */
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
    numberFormats: {},
    widths: v.columns.slice(0, a.cols).map((c) => (c.hidden ? 0 : c.width)),
    rowHeights: {},
    merges: [],
    tables: v.tables.map((t) => ({ name: t.name, range: t.range, headerRow: t.headers ? (parseRange(t.range)?.top ?? 0) : 0 })),
    appearance: {},
    links: {},
    notes: {},
    ...(a.notice ? { notice: a.notice } : {}),
  }
  // A row at the default height (xlcalc's for a row without one) is omitted from `rowHeights`.
  for (const r of v.rows) if (r.hidden || r.height !== DEFAULT_ROW_HEIGHT) win.rowHeights[r.row] = r.hidden ? 0 : r.height
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
      if (cell.numberFormat !== undefined) win.numberFormats![addr] = cell.numberFormat
      if (cell.formulaLocal !== undefined) win.formulas[addr] = cell.formulaLocal
      if (cell.editText) win.editText![addr] = cell.editText
      if (cell.color) win.colors[addr] = cell.color
      if (cell.note) win.notes![addr] = cell.note
      if (cell.appearance) win.appearance![addr] = cell.appearance
      if (cell.link) {
        // with an address the location belongs to what it points to (Excel's SubAddress), not to this workbook
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
