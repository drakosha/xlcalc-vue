// What the session answers, as the sheet source answers it: refusals as outcomes, cells as updates, whether the
// component must read the sheets and windows again, and the rectangles that hold edited cells.
import type { CellRect, CellRef, EditResult, ViewportCell } from 'xlcalc'
import { cellAddress, nameKey, parseCell } from 'xlcalc/a1'
import type { CellUpdate, EditOutcome } from '../source'

export const failed = (message: string): EditOutcome => ({ ok: false, reason: 'failed', message })

/** An exception of the session as an outcome: input the book refuses, a busy calculator, anything else failed. */
export function outcomeOf(e: unknown, input: 'refused' | 'failed'): EditOutcome {
  const err = e as { name?: string; code?: string; message?: string }
  const message = err?.message ?? String(e)
  if (err?.name === 'InputError') return { ok: false, reason: input, message }
  if (err?.name === 'CalcError' && (err.code === 'soft_timeout' || err.code === 'queue_full' || err.code === 'aborted')) return { ok: false, reason: 'busy', message }
  return failed(message)
}

/** Codes of a formula Excel does not take as typed: its window says where (`invalid` with `at`). */
const INVALID = new Set(['syntax', 'empty', 'too_long', 'too_deep'])

/** A text the session did not take, with its cell: a formula that does not read is invalid at its place, the rest refused. */
export function refusalOf(r: { code: string; message: string; at?: number; sheet: string; cell: string }): EditOutcome {
  const cells = [{ sheet: r.sheet, cell: r.cell }]
  if (INVALID.has(r.code)) return { ok: false, reason: 'invalid', message: r.message, ...(r.at !== undefined ? { at: r.at } : {}), cells }
  return { ok: false, reason: 'refused', message: r.message, cells }
}

/** A cell of a window as the update of an edited cell: its formula as shown and the text it opens with. */
export function updateOf(sheet: string, addr: string, cell: ViewportCell | undefined): CellUpdate {
  return {
    sheet,
    cell: addr,
    text: cell?.text ?? '',
    value: cell ? cell.value : null,
    ...(cell?.color ? { color: cell.color } : {}),
    formula: cell?.formulaLocal ?? null,
    editText: cell?.editText ?? '',
    numberFormat: cell?.numberFormat ?? null,
  }
}

/** The change moved, laid out or restyled more than cell updates carry: the sheets and windows are read again. */
export const needsReload = (got: EditResult): boolean => !!(got.layout?.length || got.styles || got.moved?.length || got.changes.some((c) => c.style !== undefined))

/** A key of a cell that matches sheets as the library does, in any case. */
export const cellKey = (sheet: string, cell: string): string => `${nameKey(sheet)}!${cell}`

/** Cells on a sheet: the rectangle that holds them and their addresses. */
export interface CellsBox {
  sheet: string
  rect: CellRect
  only?: Set<string>
}

/** Cells by sheet, each sheet's in the rectangle that holds them; an address that is no cell is left out. */
export function boxesOf(refs: CellRef[]): CellsBox[] {
  const boxes = new Map<string, Required<CellsBox>>()
  for (const ref of refs) {
    const at = parseCell(ref.cell)
    if (!at) continue
    const key = nameKey(ref.sheet)
    const b = boxes.get(key)
    if (!b) boxes.set(key, { sheet: ref.sheet, rect: { top: at.row, left: at.col, bottom: at.row, right: at.col }, only: new Set([cellAddress(at.row, at.col)]) })
    else {
      b.rect = { top: Math.min(b.rect.top, at.row), left: Math.min(b.rect.left, at.col), bottom: Math.max(b.rect.bottom, at.row), right: Math.max(b.rect.right, at.col) }
      b.only.add(cellAddress(at.row, at.col))
    }
  }
  return [...boxes.values()]
}
