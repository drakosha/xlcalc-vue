// The ranges the source's methods name: rectangles, whole columns and whole rows.
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
