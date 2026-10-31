// What a PivotTable's cells show, from the layout xlcalc saved with it (`PivotLocation`, `rowLines`, `columnLines`): the
// labels of its fields' items, the cells with Excel's filter buttons, and what a cell is (an item, a total, a value, a
// header), for the sheet's cell menu, filter buttons and dialogs. TODO(xlcalc): xlcalc 0.72's pivotCellAt and pivotCells in
// xlcalc/text are to take the place of `cellOf` and `buttonsOf`, the one place this is worked out here.
import { builtinFormatCode, dateToSerial, displayValue, toLocal, type BookLayout, type CellValue, type PivotAxisField, type PivotCacheField, type PivotLine, type PivotTable, type PivotValue } from 'xlcalc'
import { cellAddress, parseCell, sameName } from 'xlcalc/a1'
import type { PivotAxis, SheetPivotButton, SheetPivotCalculated, SheetPivotCell, SheetPivotCells } from '../contract/source'

/** Excel's short date (built-in format 14), in which a PivotTable shows a date without a format of its own. */
const SHORT_DATE = 14
/** The label of the blank item, as xlcalc lays it out in the workbook's language. */
const BLANK = { en: '(blank)', ru: '(пусто)' }

/** A value of the cache as a cell holds it. */
function cellValueOf(v: PivotValue): CellValue {
  if (v === null) return null
  if (typeof v === 'object' && 'date' in v) return dateToSerial(v.date)
  return v as CellValue
}

/** A field's items' labels in the PivotTable's order: the caption, or the value in the field's number format. */
function labelsOf(t: PivotTable, f: PivotCacheField, index: number, locale: string): string[] {
  const own = t.fields[index]
  const values = f.group?.items ?? f.items ?? []
  const format = own?.numberFormat ?? f.numberFormat
  const blank = locale.toLowerCase().startsWith('ru') ? BLANK.ru : BLANK.en
  return (own?.items ?? []).map((x) => {
    if (x.caption !== undefined) return x.caption
    const v = x.item === undefined ? undefined : values[x.item]
    if (v === undefined) return ''
    if (v === null) return blank
    const date = typeof v === 'object' && 'date' in v
    return displayValue(cellValueOf(v), { locale, numberFormat: date ? (format ?? SHORT_DATE) : format }).text
  })
}

/** The field of an axis at a position (`values` for the value fields' place), with the form it is shown in. */
const layoutOf = (t: PivotTable, f: PivotAxisField) => (f === 'values' ? undefined : t.fields[f]?.layout) ?? t.layout ?? 'compact'

/** The column of each row field among the row labels: a compact field shares the column of a compact field before it. */
function rowColumns(t: PivotTable): number[] {
  const rows = t.rows ?? []
  const out: number[] = []
  rows.forEach((f, p) => out.push(p === 0 ? 0 : out[p - 1]! + (layoutOf(t, rows[p - 1]!) === 'compact' && layoutOf(t, f) === 'compact' ? 0 : 1)))
  return out
}

/** The items each line shows, the fields before `repeat` taken from the line before. */
function lineItems(lines: PivotLine[]): number[][] {
  let last: number[] = []
  return lines.map((l) => (last = [...last.slice(0, l.repeat ?? 0), ...l.items]))
}

/** A label of an axis line at the field position `p`: the item, or the line's total. */
function labelOf(fields: PivotAxisField[], line: PivotLine, items: number[], p: number, axis: PivotAxis): SheetPivotCell {
  if (line.type === 'grand') return { kind: 'grand', axis, ...(line.value !== undefined ? { value: line.value } : {}) }
  if (line.type === 'blank') return { kind: 'blank' }
  const f = fields[p]
  if (f === 'values') return { kind: 'valueHeader', value: items[p] ?? line.value ?? 0 }
  if (f === undefined) return { kind: 'blank' }
  return { kind: line.type ? 'subtotal' : 'item', axis, field: f, ...(items[p] !== undefined ? { item: items[p] } : {}) }
}

/** The Filters area's field at a cell: fields go down a column of `pageRows` then over (or over then down), three columns each. */
function pageAt(t: PivotTable, row: number, col: number, top: number, left: number): SheetPivotCell | null {
  const { pageRows = 0, pageColumns = 1 } = t.location
  const pages = t.pages ?? []
  const r = row - (top - pageRows - 1)
  const block = Math.floor((col - left) / 3)
  const within = (col - left) % 3
  if (!pages.length || r < 0 || r >= pageRows || block < 0 || within > 1) return null
  const k = t.pagesAcrossFirst ? r * pageColumns + block : block * pageRows + r
  const page = pages[k]
  return page ? { kind: within === 0 ? 'pageField' : 'pageItem', axis: 'pages', field: page.field } : null
}

/** What the cell at `row` and `col` of the sheet is in the PivotTable `t`; null outside it. */
function cellOf(t: PivotTable, row: number, col: number): SheetPivotCell | null {
  const x = rectOf(t)
  if (!x) return null
  if (row < x.top) return pageAt(t, row, col, x.top, x.left)
  if (row > x.bottom || col < x.left || col > x.right) return null
  const { firstHeaderRow: headRow, firstDataRow: dataRow, firstDataColumn: dataCol } = t.location
  const r = row - x.top
  const c = col - x.left
  const rows = t.rows ?? []
  const columns = t.columns ?? []
  const rowLines = t.rowLines ?? []
  const columnLines = t.columnLines ?? []
  const cols = rowColumns(t)
  if (r >= dataRow && c < dataCol) {
    const line = rowLines[r - dataRow]
    if (!line) return null
    // no row field: the value fields' names left of their values
    if (!rows.length) return t.values?.length ? { kind: 'valueHeader', value: line.value ?? 0 } : { kind: 'blank' }
    const items = lineItems(rowLines)[r - dataRow]!
    const start = line.repeat ?? 0
    const last = start + line.items.length - 1
    // a total's label stands in its field's column (the grand total's in the first), the cells beside it blank
    if (line.type === 'grand') return c === 0 ? labelOf(rows, line, items, 0, 'rows') : { kind: 'blank' }
    if (line.type) return cols[last] === c ? labelOf(rows, line, items, last, 'rows') : { kind: 'blank' }
    // the deepest field of the line in this column; before the line's own fields, a label repeated (Repeat Item Labels)
    const here = cols.flatMap((cc, p) => (cc === c && p >= start && p < items.length ? [p] : []))
    if (here.length) return labelOf(rows, line, items, here.at(-1)!, 'rows')
    const repeated = cols.flatMap((cc, p) => {
      const f = rows[p]
      const repeats = f !== 'values' && f !== undefined && (t.fields[f]?.repeatLabels ?? t.repeatLabels) && layoutOf(t, f) !== 'compact'
      return cc === c && p < start && repeats ? [p] : []
    })
    return repeated.length ? labelOf(rows, line, items, repeated.at(-1)!, 'rows') : { kind: 'blank' }
  }
  if (r >= dataRow) {
    // a value: its value field from the line of the axis the value fields lie across
    const rl = rowLines[r - dataRow]
    const cl = columnLines[c - dataCol]
    const value = rl?.value ?? cl?.value ?? 0
    return { kind: 'value', ...(t.values?.length ? { value } : {}) }
  }
  if (c >= dataCol && columns.length && r >= headRow) {
    // the column labels: a row per column field
    const lines = lineItems(columnLines)
    const line = columnLines[c - dataCol]
    if (!line) return null
    if (line.type === 'grand') return labelOf(columns, line, lines[c - dataCol]!, 0, 'columns')
    return labelOf(columns, line, lines[c - dataCol]!, r - headRow, 'columns')
  }
  if (c >= dataCol && r === dataRow - 1) {
    // no column field: the value fields' names over their values
    return t.values?.length ? { kind: 'valueHeader', value: columnLines[c - dataCol]?.value ?? 0 } : { kind: 'blank' }
  }
  if (c >= dataCol && r === headRow - 1) {
    // Column Labels in compact form, the column fields' names in the others
    const compact = (t.layout ?? 'compact') === 'compact'
    const f = compact ? (c === dataCol ? columns.find((x) => x !== 'values') : undefined) : columns[c - dataCol]
    return f === undefined || f === 'values' ? { kind: 'blank' } : { kind: 'header', axis: 'columns', field: f }
  }
  if (r === dataRow - 1 && c < dataCol) {
    // Row Labels in compact form, the row fields' names in the others
    const f = rows.find((x, p) => cols[p] === c && x !== 'values')
    return f === undefined || f === 'values' ? { kind: 'blank' } : { kind: 'header', axis: 'rows', field: f }
  }
  // the top left: one value field's name
  return r === 0 && c === 0 && t.values?.length === 1 && columns.length && rows.length ? { kind: 'valueHeader', value: 0 } : { kind: 'blank' }
}

function rectOf(t: PivotTable) {
  const [a, b] = t.location.range.split(':')
  const tl = parseCell(a!)
  const br = parseCell(b ?? a!)
  return tl && br ? { top: tl.row, left: tl.col, bottom: br.row, right: br.col } : null
}

/** The filter buttons: the row fields' headers, the column fields' (Column Labels alone in compact form), the Filters area's items. */
function buttonsOf(t: PivotTable): SheetPivotButton[] {
  const x = rectOf(t)
  if (!x || t.fieldHeaders === false || !t.rowLines) return []
  const out: SheetPivotButton[] = []
  const { firstHeaderRow: headRow, firstDataRow: dataRow, firstDataColumn: dataCol } = t.location
  const rows = t.rows ?? []
  const cols = rowColumns(t)
  const seen = new Set<number>()
  rows.forEach((f, p) => {
    if (f === 'values' || seen.has(cols[p]!) || dataRow < 1) return
    seen.add(cols[p]!)
    out.push({ cell: cellAddress(x.top + dataRow - 1, x.left + cols[p]!), axis: 'rows', field: f })
  })
  const columns = t.columns ?? []
  if (headRow >= 1) {
    const compact = (t.layout ?? 'compact') === 'compact'
    columns.forEach((f, p) => {
      if (f === 'values' || (compact && out.some((b) => b.axis === 'columns'))) return
      out.push({ cell: cellAddress(x.top + headRow - 1, x.left + dataCol + (compact ? 0 : p)), axis: 'columns', field: f })
    })
  }
  const { pageRows = 0, pageColumns = 1 } = t.location
  ;(t.pages ?? []).forEach((page, k) => {
    const r = t.pagesAcrossFirst ? Math.floor(k / pageColumns) : k % Math.max(1, pageRows)
    const block = t.pagesAcrossFirst ? k % pageColumns : Math.floor(k / Math.max(1, pageRows))
    out.push({ cell: cellAddress(x.top - pageRows - 1 + r, x.left + 3 * block + 1), axis: 'pages', field: page.field })
  })
  return out
}

/** A formula of the cache as Excel's dialog shows it: in the workbook's locale, with `=`; as stored where it does not translate. */
function shown(formula: string, locale: string): string {
  const got = toLocal(formula, locale)
  return `=${got.ok ? got.formula : formula}`
}
/** The cache's calculated fields and items; an item's formula without its field's name before each item (`Region[East]` as `East`). */
function calculatedOf(cache: { fields: PivotCacheField[]; calculatedItems?: { field: number; item: number; formula: string }[] }, locale: string): SheetPivotCalculated[] {
  const out: SheetPivotCalculated[] = []
  cache.fields.forEach((f, field) => f.formula !== undefined && out.push({ field, name: f.name, formula: shown(f.formula, locale) }))
  for (const c of cache.calculatedItems ?? []) {
    const name = cache.fields[c.field]?.name ?? ''
    const plain = c.formula.split(`${name}[`).join('').replace(/\]/g, '')
    out.push({ field: c.field, item: c.item, name: String(cache.fields[c.field]?.items?.[c.item] ?? ''), formula: shown(plain, locale) })
  }
  return out
}

/** `SheetSource.pivotCells` over the session's layout; null when the sheet has no PivotTable `name`. */
export function pivotCellsOf(layout: BookLayout, sheet: string, name: string, cells: string[] | undefined, locale: string): SheetPivotCells | null {
  const t = layout.sheets.find((x) => sameName(x.name, sheet))?.pivotTables?.find((p) => sameName(p.name, name))
  const cache = t && layout.pivotCaches?.find((c) => c.id === t.cache)
  if (!t || !cache) return null
  const onAxis = new Set<number>([...(t.rows ?? []), ...(t.columns ?? []), ...(t.pages ?? []).map((p) => p.field)].filter((f): f is number => f !== 'values'))
  const labels = cache.fields.map((f, i) => (onAxis.has(i) || t.fields[i]?.items ? labelsOf(t, f, i, locale) : null))
  const at = (cell: string) => {
    const a = parseCell(cell)
    return a ? cellOf(t, a.row, a.col) : null
  }
  const formats = (t.values ?? []).map((v) => (v.numberFormat === undefined ? null : typeof v.numberFormat === 'string' ? v.numberFormat : (builtinFormatCode(v.numberFormat, locale) ?? null)))
  return { labels, formats, calculated: calculatedOf(cache, locale), buttons: buttonsOf(t), ...(cells ? { cells: cells.map(at) } : {}) }
}
