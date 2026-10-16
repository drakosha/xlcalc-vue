// What the session answers, as the sheet source answers it: refusals as outcomes, cells as updates, whether the
// component must read the sheets and windows again, and the rectangles that hold edited cells.
import type { CellRect, CellRef, CellTextsRefusal, EditResult, PasteTextRefusal, ViewportCell } from 'xlcalc'
import { cellAddress, nameKey, parseCell } from 'xlcalc/a1'
import { cellRect, unionRect } from '../rect'
import type { CellUpdate, EditOutcome } from '../source'
import { alertRefusalOf } from './validation'

export const failed = (message: string): EditOutcome => ({ ok: false, reason: 'failed', message })

/** An exception of the session as an outcome: input the book refuses as `input`, anything else (a busy calculator too) failed. */
export function outcomeOf(e: unknown, input: 'refused' | 'failed'): EditOutcome {
  const err = e as { name?: string; message?: string }
  const message = err?.message ?? String(e)
  return err?.name === 'InputError' ? { ok: false, reason: input, message } : failed(message)
}

/** Codes of a formula Excel does not take as typed: its window says where (`invalid` with `at`). */
const INVALID = new Set(['syntax', 'empty', 'too_long', 'too_deep'])

/** A text the session did not take: a value a rule refuses with the rule's alert, a formula that does not read invalid at its place with the correction Excel offers, the rest refused. */
export function refusalOf(r: CellTextsRefusal | PasteTextRefusal): EditOutcome {
  if (r.code === 'validation') return alertRefusalOf(r)
  if (!INVALID.has(r.code)) return { ok: false, reason: 'refused', message: r.message }
  // a paste's refusals are none of these codes
  const { at, correction } = r as { at?: number; correction?: string }
  return { ok: false, reason: 'invalid', message: r.message, ...(at !== undefined ? { at } : {}), ...(correction !== undefined ? { correction } : {}) }
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

/** The change moved, laid out, restyled or changed a note or a hyperlink, which cell updates do not carry: the sheets and windows are read again. */
export const needsReload = (got: EditResult): boolean =>
  !!(got.layout?.length || got.styles || got.moved?.length || got.changes.some((c) => c.style !== undefined || c.note !== undefined || c.link !== undefined))

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
    if (!b) boxes.set(key, { sheet: ref.sheet, rect: cellRect(at.row, at.col), only: new Set([cellAddress(at.row, at.col)]) })
    else {
      b.rect = unionRect(b.rect, cellRect(at.row, at.col))
      b.only.add(cellAddress(at.row, at.col))
    }
  }
  return [...boxes.values()]
}
