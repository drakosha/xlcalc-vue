// Cell inputs for the session: what a person typed read by `parseEntry`, and a cell's input rebuilt from its view.
import { parseEntry, type CellInput, type CellView, type NumberFormat } from 'xlcalc'
import { colLabel, colNumber, colsAddress, parseRect, type Rect } from '../core/a1'

/**
 * What a range of the sheet source names: a rectangle, whole columns (`B:D`) or whole rows (`2:9`). Lines carry their
 * text twice: `range` as `cells` takes it (`B:B`), `span` as the layout takes it (`B`, `B:D`).
 */
export type Target = { kind: 'cells'; rect: Rect } | { kind: 'columns' | 'rows'; lo: number; hi: number; range: string; span: string }

export function parseTarget(range: string): Target | null {
  const text = range.trim().toUpperCase()
  const cols = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(text)
  if (cols) {
    const [lo, hi] = [colNumber(cols[1]!), colNumber(cols[2]!)].sort((x, y) => x - y) as [number, number]
    return { kind: 'columns', lo, hi, range: `${colLabel(lo)}:${colLabel(hi)}`, span: colsAddress(lo, hi) }
  }
  const rows = /^(\d{1,7}):(\d{1,7})$/.exec(text)
  if (rows) {
    const [lo, hi] = [Number(rows[1]), Number(rows[2])].sort((x, y) => x - y) as [number, number]
    return lo > 0 ? { kind: 'rows', lo, hi, range: `${lo}:${hi}`, span: lo === hi ? String(lo) : `${lo}:${hi}` } : null
  }
  const rect = parseRect(text)
  return rect ? { kind: 'cells', rect } : null
}

/**
 * A cell's input as the session holds it, from its view: the formula or the value (a cell an array spilled into holds
 * none), the format, style, meta, note and link. `exact`, the book's own input, takes its place where a view cannot
 * tell it (an old-kind array formula's area).
 */
export function inputOf(view: CellView | undefined, exact?: CellInput): CellInput {
  if (exact) return { ...exact }
  if (!view) return {}
  const out: CellInput = {}
  if (view.formula !== undefined) out.formula = view.formula
  else if (view.spillOf === undefined && view.value !== null && typeof view.value !== 'object') out.value = view.value
  if (view.format !== undefined) out.format = view.format
  if (view.style !== undefined) out.style = view.style
  if (view.meta !== undefined) out.meta = view.meta
  if (view.note !== undefined) out.note = view.note
  if (view.link !== undefined) out.link = view.link
  return out
}

export interface EntryContext {
  locale: string
  dateSystem?: 1900 | 1904
  /** The book's clock: the year of a date typed without one. */
  now: string
  /** The format the cell shows without one of its own (its row's or column's). */
  inherited?: NumberFormat
}

/** Typed text as the cell's new input, keeping its style, meta, note and link; a formula that does not parse gives its position. */
export function typedInput(text: string | null, view: CellView | undefined, ctx: EntryContext): { ok: true; input: CellInput | null } | { ok: false; message: string; at: number } {
  const own = view?.format
  const got = parseEntry(text ?? '', {
    locale: ctx.locale,
    now: ctx.now,
    ...(ctx.dateSystem ? { dateSystem: ctx.dateSystem } : {}),
    cell: { format: own ?? ctx.inherited, style: view?.style, meta: view?.meta },
  })
  if (!got.ok) return { ok: false, message: got.message, at: got.at }
  const input: CellInput = { ...got.input }
  // A format the cell only shows from its row or column stays theirs.
  if (own === undefined && input.format !== undefined && input.format === ctx.inherited) delete input.format
  // A note and a link go with the cell, as its style does.
  if (view?.note) input.note = view.note
  if (view?.link) input.link = view.link
  return { ok: true, input: Object.keys(input).length ? input : null }
}

/** A cell's input with a style and a number format set (a number), removed (null) or kept (undefined). */
export function styledInput(base: CellInput, style: number | null | undefined, format: NumberFormat | null | undefined): CellInput {
  const out = { ...base }
  if (style === null) delete out.style
  else if (style !== undefined) out.style = style
  if (format === null) delete out.format
  else if (format !== undefined) out.format = format
  return out
}
