// Cell inputs for the session: what a person typed read by `parseEntry`, and the ranges the source's methods name.
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
  const own = view?.numberFormat
  const got = parseEntry(text ?? '', {
    locale: ctx.locale,
    now: ctx.now,
    ...(ctx.dateSystem ? { dateSystem: ctx.dateSystem } : {}),
    cell: { numberFormat: own ?? ctx.inherited, style: view?.style, meta: view?.meta },
  })
  if (!got.ok) return { ok: false, message: got.message, at: got.at }
  const input: CellInput = { ...got.input }
  // A format the cell only shows from its row or column stays theirs.
  if (own === undefined && input.numberFormat !== undefined && input.numberFormat === ctx.inherited) delete input.numberFormat
  // A note and a link go with the cell, as its style does.
  if (view?.note) input.note = view.note
  if (view?.link) input.link = view.link
  return { ok: true, input: Object.keys(input).length ? input : null }
}
